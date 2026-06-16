import os
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Muat variabel dari file .env di root project
# Buat file .env dari .env.example lalu isi dengan kredensial kamu.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(_ROOT, '.env'))

LOGIN_URL     = "https://simkeswa.kemkes.go.id/login"
DATA_PAGE_URL = "https://simkeswa.kemkes.go.id/form_skrining/list_data?per_page=1350"

USERNAME      = os.getenv("SIMKESWA_USERNAME")
PASSWORD      = os.getenv("SIMKESWA_PASSWORD")
MAX_DELETIONS = int(os.getenv("MAX_DELETIONS", "100"))

if not USERNAME or not PASSWORD:
    raise EnvironmentError(
        "Kredensial belum diset. Buat file .env dari .env.example "
        "dan isi SIMKESWA_USERNAME serta SIMKESWA_PASSWORD."
    )

def setup_driver():
    """Initializes and returns a Selenium WebDriver instance."""
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)
    return driver

def login(driver, url, username, password):
    """
    Handles the login process for a given website.
    
    Args:
        driver: The Selenium WebDriver instance.
        url: The login page URL.
        username: The username credential.
        password: The password credential.
        
    Returns:
        True if login appears successful, False otherwise.
    """
    print(f"--- Attempting Login ---")
    try:
        driver.get(url)
        print("Login page opened.")
        
        # --- YOU MUST UPDATE THESE SELECTORS ---
        # Find the correct selectors for the username, password, and login button.
        # Common selectors are 'name', 'id', or XPath.
        # The 'name' attributes from your example ('token' for username, 'password' for password) are used here.
        username_field_name = "token"
        password_field_name = "password"
        # The login button selector needs to be found by inspecting the element.
        # It could be a button with text 'Login', 'Sign In', etc.
        login_button_xpath = "//button[@type='submit']" 
        
        print(f"Locating username field with name='{username_field_name}'...")
        user_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, username_field_name))
        )
        
        print(f"Locating password field with name='{password_field_name}'...")
        pass_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, password_field_name))
        )
        
        print("Entering credentials...")
        user_field.send_keys(username)
        pass_field.send_keys(password)
        
        print("Locating and clicking login button...")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, login_button_xpath))
        )
        login_button.click()
        
        # --- VERIFY SUCCESSFUL LOGIN ---
        # A robust way to check for success is to wait for an element that ONLY
        # appears after logging in, like a "Logout" button or a user dashboard element.
        # Or, you can check that the URL has changed from the login page.
        print("Waiting for post-login page to load...")
        WebDriverWait(driver, 15).until(
            EC.url_changes(url)
        )
        
        if driver.current_url != url:
            print("Login successful. Redirected to a new page.")
            return True
        else:
            print("Login may have failed. URL did not change.")
            return False

    except TimeoutException:
        print("Login failed. A required element was not found on the page in time.")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during login: {e}")
        return False

def auto_delete(driver, url, max_deletions):
    """
    Navigates to the data page and repeatedly deletes items.
    (This function remains the same as before)
    """
    print(f"\n--- Starting Deletion Process ---")
    driver.get(url)
    print(f"Navigated to {url}")
    time.sleep(3) # Wait for page to settle

    deleted_count = 0
    for i in range(max_deletions):
        print(f"\n--- Starting Deletion Cycle #{i + 1} ---")
        try:
            page_body = driver.find_element(By.TAG_NAME, 'html')
            
            print("Scrolling to the bottom of the page...")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

            delete_button_xpath = "(//a[contains(@href, 'delete?id=')])[last()]"
            print("Locating the delete button...")
            delete_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, delete_button_xpath))
            )
            
            print("Clicking delete button...")
            driver.execute_script("arguments[0].click();", delete_button)

            print("Waiting for confirmation pop-up...")
            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                print(f"Alert found. Accepting...")
                alert.accept()
            except TimeoutException:
                confirm_button_xpath = "//button[contains(text(), 'Confirm') or contains(text(), 'OK')]"
                confirm_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, confirm_button_xpath))
                )
                print("HTML modal confirmation button found. Clicking it...")
                confirm_button.click()

            print("Confirmation sent. Waiting for page to reload...")
            WebDriverWait(driver, 15).until(EC.staleness_of(page_body))
            
            deleted_count += 1
            print(f"Page reloaded. Item #{deleted_count} deleted successfully.")
            time.sleep(1)

        except (NoSuchElementException, TimeoutException):
            print("\nCould not find any more delete buttons. The process is likely complete.")
            break
        except Exception as e:
            print(f"\nAn unexpected error occurred: {e}")
            break

    print(f"\n--- Automation Finished ---")
    print(f"Total items deleted: {deleted_count}")


if __name__ == "__main__":
    my_driver = None
    try:
        my_driver = setup_driver()
        # First, perform the login
        if login(my_driver, LOGIN_URL, USERNAME, PASSWORD):
            # If login is successful, proceed to the auto-delete function
            auto_delete(my_driver, DATA_PAGE_URL, MAX_DELETIONS)
            print("SUCCES!")
        else:
            print("\nCould not proceed to deletion due to login failure.")
            
    finally:
        if my_driver:
            print("Closing browser in 10 seconds...")
            time.sleep(10)
            my_driver.quit()
            print("Browser closed.")