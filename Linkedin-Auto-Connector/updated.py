from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser
from selenium.webdriver.common.action_chains import ActionChains
from colorama import Fore, Style, init
import time, requests
from urllib.parse import quote

# Google Sheets imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# Initialize colorama
init(autoreset=True)

# Initialize config parser
config = ConfigParser()
config_file = 'setup.ini'
config.read(config_file)

remote_webdriver_url = "http://localhost:52613"

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--log-level=2")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    driver = webdriver.Remote(options=chrome_options, command_executor=remote_webdriver_url)
    return driver

def save_cookie(driver:webdriver.Chrome):
    li_at_cookie = driver.get_cookie('li_at')['value']
    config.set('LinkedIn', 'li_at', li_at_cookie)
    with open(config_file, 'w') as f:
        config.write(f)

def login_with_cookie(driver:webdriver.Chrome, li_at):
    print(Fore.YELLOW + "Attempting to log in with cookie...")
    driver.get("https://www.linkedin.com")
    driver.add_cookie({
        "name": "li_at",
        "value": f"{li_at}",
        "path": "/",
        "secure": True,
    })
    driver.refresh()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    print(Fore.GREEN + "[INFO] Logged in with cookie successfully.")

def select_location(driver:webdriver.Chrome, location:str):
    try:
        print("Selecting location")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchFilter_geoUrn"))).click()
        time.sleep(1)
        location_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Add a location']")))
        location_input.send_keys(location)
        time.sleep(2)
        driver.find_element(By.XPATH,f"//*[text()='{location.title()}']").click()
        time.sleep(1)
        driver.find_element(By.XPATH,"//button[@aria-label='Apply current filter to show results']").click()
        time.sleep(3)
    except Exception as e:
        print(Fore.RED + f"[INFO] Error selecting location: {e}")

def login_with_credentials(driver:webdriver.Chrome, email:str, password:str):
    print(Fore.YELLOW + "Logging in with credentials...")
    driver.get("https://www.linkedin.com/login")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(email)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(By.ID, "global-nav-typeahead") or 
        "Enter the code" in d.page_source
    )
    if "Enter the code" in driver.page_source:
        verification_code = input("[+] Enter the verification code sent to your email: ")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "input__email_verification_pin")))
        driver.find_element(By.ID, "input__email_verification_pin").send_keys(verification_code)
        driver.find_element(By.ID, "email-pin-submit-button").click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    print(Fore.GREEN + "[INFO] Logged in with credentials successfully.")
    save_cookie(driver)

# --- Google Sheets Integration ---
def get_profile_urls_from_sheet(sheet_url, column_name, creds_path):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url).sheet1
    records = sheet.get_all_records()
    urls = [row[column_name] for row in records if row.get(column_name)]
    return urls

# --- Robust click helper ---
def robust_click(driver, element):
    try:
        element.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", element)
        except Exception:
            ActionChains(driver).move_to_element(element).pause(0.5).click().perform()

# --- UPDATED: Robust Connect Button Handling & Confirmation ---
def send_connection_request_to_urls(driver, urls, letter, include_notes):
    successful = 0
    for url in urls:
        try:
            driver.get(url)
            time.sleep(3)
            wait = WebDriverWait(driver, 15)
            connect_button = None

            # Try direct Connect button
            try:
                connect_button = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='Connect'] and not(@disabled)]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", connect_button)
                driver.execute_script("window.scrollBy(0, -150);")  # Offset for sticky header
                time.sleep(0.5)
                robust_click(driver, connect_button)
            except TimeoutException:
                # Try "More" menu
                try:
                    more_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[normalize-space()='More']]"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView(true);", more_button)
                    more_button.click()
                    time.sleep(1)
                    connect_menuitem = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@role='menu']//span[text()='Connect']/ancestor::*[@role='menuitem' or @role='button'][1]"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView(true);", connect_menuitem)
                    driver.execute_script("window.scrollBy(0, -150);")
                    time.sleep(0.5)
                    robust_click(driver, connect_menuitem)
                except TimeoutException:
                    print(Fore.RED + f"[INFO] Connect button not found, skipping profile: {url}")
                    continue

            time.sleep(2)
            # --- Add note if needed ---
            if include_notes:
                try:
                    add_note_button = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//button[@aria-label="Add a note"]')))
                    add_note_button.click()
                    message_box = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//textarea[@name="message"]')))
                    try:
                        name_elem = driver.find_element(By.XPATH, "//h1[contains(@class,'text-heading-xlarge')]")
                        name = name_elem.text.split(' ')[0]
                    except Exception:
                        name = "there"
                    message_box.send_keys(letter.replace("{name}", name).replace("{fullName}", name))
                    time.sleep(1)
                    send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send invitation"]')
                    driver.execute_script("arguments[0].click();", send_button)
                    time.sleep(2)
                except Exception as e:
                    print(Fore.YELLOW + f"Could not add a note: {e}")
                    # Try sending without a note
                    try:
                        send_button = wait.until(EC.element_to_be_clickable(
                            (By.XPATH, '//button[@aria-label="Send now" or @aria-label="Send without a note"]')))
                        send_button.click()
                        time.sleep(2)
                    except Exception as e2:
                        print(Fore.RED + f"Could not send invitation: {e2}")
                        continue
            else:
                try:
                    send_button = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, '//button[@aria-label="Send now" or @aria-label="Send without a note"]')))
                    send_button.click()
                    time.sleep(2)
                except Exception as e:
                    print(Fore.RED + f"Could not send invitation: {e}")
                    continue

            # --- Confirmation check: Only count as successful if Connect button is gone ---
            try:
                driver.find_element(By.XPATH, "//button[.//span[text()='Connect']]")
                print(Fore.RED + f"[INFO] Connection request may NOT have been sent for {url} (Connect button still present).")
                continue
            except:
                print(Fore.GREEN + f"[INFO] Connection request sent successfully to {url}")
                successful += 1
                print("---------------------------------------------------------------------------------------------------------------")
            time.sleep(2)
        except StaleElementReferenceException:
            print(Fore.YELLOW + f"[INFO] Stale element, retrying profile: {url}")
            continue
        except Exception as e:
            print(Fore.RED + f"[INFO] Failed to send request to {url}: {e}")
    print(Fore.YELLOW + f"Total successful connections: {successful}")

# --- Existing robust search-based connection logic (unchanged) ---
def send_connection_request(driver: webdriver.Chrome, limit: int, letter: str, include_notes: bool, message_letter: str):
    successful_connections = 0
    while successful_connections < limit:
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            try:
                driver.find_element(By.XPATH, "//h2[text()='No free personalized invitations left']")
                print(Fore.RED + "[ERROR] No free personalized invitations left.")
                return
            except:
                pass
            if message_letter == "":
                connect_buttons = driver.find_elements(By.XPATH, "//*[text()='Connect']/..")
            else:
                connect_buttons = driver.find_elements(By.XPATH, "//*[text()='Message']/..")
            print(f"Number of connect buttons found: {len(connect_buttons)}")
            if not connect_buttons:
                print("No connect buttons found, moving to next page...")
                try:
                    driver.find_element(By.XPATH, "//button[@aria-label='Next']").click()
                    time.sleep(2)
                    continue
                except:
                    print("No next page available")
                    break
            for i, connect_button in enumerate(connect_buttons):
                if successful_connections >= limit:
                    break
                try:
                    actions = ActionChains(driver)
                    if message_letter == "":
                        try:
                            connect_container = connect_button.find_element(By.XPATH, "./ancestor::div[contains(@class, 'entity-result')]")
                            linkedin_container = connect_container.find_element(By.XPATH, ".//a[contains(@href, '/in/')]")
                            linkedin_url = linkedin_container.get_attribute('href')
                            name = linkedin_container.text.split(' ')[0].title()
                        except Exception as e:
                            print(f"Method 1 failed, trying Method 2: {e}")
                            try:
                                connect_container = connect_button.find_element(By.XPATH, "./ancestor::li")
                                linkedin_container = connect_container.find_element(By.XPATH, ".//span[contains(@class, 'entity-result__title')]//a")
                                linkedin_url = linkedin_container.get_attribute('href')
                                name = linkedin_container.text.split(' ')[0].title()
                            except Exception as e2:
                                print(f"Method 2 also failed, trying Method 3: {e2}")
                                try:
                                    connect_container = connect_button.find_element(By.XPATH, "./ancestor::*[contains(@class, 'result')]")
                                    linkedin_container = connect_container.find_element(By.XPATH, ".//a[contains(@href, 'linkedin.com/in/')]")
                                    linkedin_url = linkedin_container.get_attribute('href')
                                    name = linkedin_container.get_attribute('aria-label') or linkedin_container.text
                                    if name:
                                        name = name.split(' ')[0].title()
                                    else:
                                        name = "Unknown"
                                except Exception as e3:
                                    print(f"All methods failed for button {i}: {e3}")
                                    continue
                        actions.move_to_element(connect_button).perform()
                        time.sleep(1)
                        connect_button.click()
                        time.sleep(1)
                        if not include_notes:
                            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Send without a note"]'))).click()
                        else:
                            add_note_button = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="Add a note"]')))
                            add_note_button.click()
                            message_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//textarea[@name="message"]')))
                            message_box.send_keys(letter.replace("{name}", name).replace("{fullName}", name))
                            time.sleep(1)
                            send_button = driver.find_element(By.XPATH, '//button[@aria-label="Send invitation"]')
                            driver.execute_script("arguments[0].click();", send_button)
                    successful_connections += 1
                    print(Fore.GREEN + f"[INFO] Connection request sent successfully to {linkedin_url}")
                    print("---------------------------------------------------------------------------------------------------------------")
                    time.sleep(3)
                except Exception as e:
                    print(f"Error with button {i}: {e}")
                    continue
            if successful_connections < limit:
                try:
                    driver.find_element(By.XPATH, "//button[@aria-label='Next']").click()
                    time.sleep(1)
                except:
                    print("No more pages available")
                    break
        except Exception as e:
            print(Fore.RED + f"[INFO] Error occurred: {e}")
            break

def main():
    print(Fore.CYAN + "[-] LinkedIn Auto Connector - Enhanced with Google Sheets Option")
    use_sheet = input(Fore.MAGENTA + "[+] Do you want to import LinkedIn profile URLs from a Google Sheet? (y/n): " + Fore.RESET).strip().lower()
    message = ''
    message_letter = ''
    include_note = False
    if use_sheet == 'y':
        sheet_url = input(Fore.MAGENTA + "[+] Enter your Google Sheet URL: " + Fore.RESET).strip()
        column_name = input(Fore.MAGENTA + "[+] Enter the column name containing LinkedIn profile URLs: " + Fore.RESET).strip()
        creds_path = input(Fore.MAGENTA + "[+] Enter the path to your Google service account credentials JSON file: " + Fore.RESET).strip()
        letter = input(Fore.MAGENTA + "[+] Enter the message letter for the connection request (use {name} for personalization): " + Fore.RESET)
        include_notes = input(Fore.MAGENTA + "[+] Do you want to include a note in the connection request? (y/n): " + Fore.RESET).strip().lower() == 'y'
        limit = int(input(Fore.MAGENTA + "[+] Enter the maximum number of connection requests to send: " + Fore.RESET))
        li_at = input(Fore.MAGENTA + "[+] Enter the li_at of Linkedin: " + Fore.RESET)
        print("----------------------------------------------------------------")
        driver = setup_driver()
        try:
            login_with_cookie(driver, li_at)
        except Exception as e:
            print(Fore.RED + f"[INFO] Cookie login failed: {e}\n" + Fore.YELLOW + "Attempting login with credentials.")
            email = config.get('LinkedIn', 'email')
            password = config.get('LinkedIn', 'password')
            login_with_credentials(driver, email, password)
        profile_urls = get_profile_urls_from_sheet(sheet_url, column_name, creds_path)
        print(Fore.YELLOW + f"[INFO] Found {len(profile_urls)} profile URLs in the sheet. Sending requests (up to your limit)...")
        send_connection_request_to_urls(driver, profile_urls[:limit], letter, include_notes)
        driver.quit()
        return
    # --- Existing flow ---
    connection_degree = input(Fore.MAGENTA + "[+] Enter the connection degree (1st, 2nd, 3rd): " + Fore.RESET)
    if connection_degree.lower() not in ['1st', '2nd', '3rd']:
        print(Fore.RED + "[ERROR] Invalid connection degree. Please enter 1st, 2nd, or 3rd.")
        connection_degree = input(Fore.MAGENTA + "[+] Enter the connection degree (1st, 2nd, 3rd): " + Fore.RESET)
    keyword = input(Fore.MAGENTA + "[+] Enter the keyword for the search: " + Fore.RESET)
    location = input(Fore.MAGENTA + "[+] Enter the location: " + Fore.RESET)
    if connection_degree.lower() == '1st':
        message_letter = input(Fore.MAGENTA + "[+] Enter the message letter for the connection request: " + Fore.RESET)
    if message_letter == "":
        include_note = input(Fore.MAGENTA + "[+] Do you want to include a note in the connection request? (y/n): " + Fore.RESET)
        if include_note.lower() == 'y':
            include_note = True
            message = input(Fore.MAGENTA + "[+] Enter the personalized message to send with connection requests: " + Fore.RESET)
        else:
            include_note = False
    limit = int(input(Fore.MAGENTA + "[+] Enter the maximum number of connection requests to send: " + Fore.RESET))
    li_at = input(Fore.MAGENTA + "[+] Enter the li_at of Linkedin: " + Fore.RESET)
    print("----------------------------------------------------------------")
    driver = setup_driver()
    try:
        login_with_cookie(driver, li_at)
    except Exception as e:
        print(Fore.RED + f"[INFO] Cookie login failed: {e}\n" + Fore.YELLOW + "Attempting login with credentials.")
        email = config.get('LinkedIn', 'email')
        password = config.get('LinkedIn', 'password')
        login_with_credentials(driver, email, password)
    network_mapping = {
        "1st": "%5B%22F%22%5D",  
        "2nd": "%5B%22S%22%5D",  
        "3rd": "%5B%22O%22%5D"   
    }
    network_code = network_mapping.get(connection_degree, "")
    search_url = f"https://www.linkedin.com/search/results/people/?keywords={keyword.replace(' ','%20').lower()}&locations={location.replace(' ','%20')}&network={network_code}&origin=FACETED_SEARCH"
    print(Fore.YELLOW + f"[INFO] Navigating to search URL: {search_url}")
    driver.get(search_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "global-nav-typeahead")))
    if location != "":
        select_location(driver, location)
    send_connection_request(driver=driver, limit=limit, letter=message, include_notes=include_note, message_letter=message_letter)
    driver.quit()

if __name__ == "__main__":
    main()
