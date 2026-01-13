import os
import time
import getpass
import keyring  # For secure password storage
import sys # For exiting cleanly if needed
import shutil # For moving files
import logging # <<< Add this import
import tempfile # <<< Add tempfile import
import traceback 
import atexit 

# --- GUI Imports ---
import tkinter as tk
from tkinter import simpledialog, messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.microsoft import EdgeChromiumDriverManager # Handles driver download/management

# --- Configuration ---
# !!! IMPORTANT: Replace with your actual email and the correct starting URL !!!
USER_EMAIL = ""   
INITIAL_LOGIN_URL = "" # Replace with the actual initial login page URL

# Use a unique service name for storing the password in the OS keychain
KEYRING_SERVICE_NAME = "smartdb_sso_automation" 

# --- Basic Logging Setup (can be overridden by TaskManager's setup) ---
# This provides a fallback if the script is run standalone
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# ---------------------------------------------------------------------

# --- Field Definitions ---
# Map user-friendly names to the itemkey attributes found in the HTML
FIELDS_TO_EXTRACT = {
    "Request No": "Application_number",
    "Requested Date": "id_shinseibi",
    "Deadline": "id_yoteibi",
    "Revision No": "REV_No",
    "Applied Vessel": "Ship_No",
    "Drawing No": "DWG_No",
    "Description": "Drawing_Name_Common_1",
    "Note": "id_bikouran",
    "Modeling Staff": "acModeler",
    # Add itemkeys for any potential file fields if known
    "PDF File": "fil_PDF_File", 
    # Add more like: "Other File": "fil_Other_Attachment" 
}

# Define which itemkeys represent file attachments
FILE_ITEMKEYS = ["fil_PDF_File"] # Add other known file itemkeys here

# --- Default Download Folder (Heuristic) ---
# This is used to find the file *after* Selenium downloads it. Adjust if needed.
DEFAULT_DOWNLOAD_FOLDER = os.path.join(os.path.expanduser('~'), 'Downloads')
# Use logging instead of print
logging.info(f"Assuming default system download folder is: {DEFAULT_DOWNLOAD_FOLDER}")
logging.info("  (Will look for downloaded files here before moving)")

# --- Secure Password Handling ---

# Updated function using tkinter dialogs
def store_or_get_password(service_name: str, username: str) -> str | None:
    """
    Retrieves the password from the OS credential manager (keyring).
    Allows the user to confirm usage or enter a new password via GUI dialogs.
    """
    # Setup tkinter root window and hide it
    # Ensure this runs even if tkinter isn't the main app loop
    try:
        root = tk.Tk()
        root.withdraw()
    except tk.TclError as e:
        logging.warning(f"Could not initialize Tkinter GUI environment for password: {e}. Falling back to console.")
        # Fallback to original getpass logic if GUI fails
        password = keyring.get_password(service_name, username)
        if not password:
            # Cannot use getpass reliably in a GUI/packaged app context usually
            logging.error("GUI fallback to console failed, and no stored password. Cannot get password via console.")
            # Maybe prompt differently or indicate failure clearly
            # For now, return None as we can't reliably get console input here
            return None
            # password = getpass.getpass(f"Enter COMPANY SSO Password for {username} (console): ") # This line is problematic
            # if password:
            #     try:
            #          # Ask to save via console as well - Also problematic
            #          save_choice = input("Save this password to keyring? (y/n): ").lower()
            #          if save_choice == 'y':
            #               keyring.set_password(service_name, username, password)
            #               logging.info("Password saved to keyring (console fallback).")
            #     except Exception as ke:
            #          logging.error(f"Error saving password to keyring (console fallback): {ke}")
        return password # Return stored password if found, otherwise None

    password_to_use: str | None = None
    stored_password = None
    try:
        stored_password = keyring.get_password(service_name, username)
    except Exception as e:
        logging.warning(f"Could not read password from keyring: {e}")
        messagebox.showwarning("Keyring Read Error", f"Could not read password from keyring:\n{e}\nPlease enter it manually.")

    if stored_password:
        logging.info(f"Found a stored password for {username} in system keyring.")
        use_stored = messagebox.askyesno(
            title="Password Found",
            message=f"Use stored password for {username}?\n\n(Click 'No' to enter a new one)",
            parent=root # Associate with hidden root
        )
        if use_stored:
            logging.info("User opted to use stored password.")
            password_to_use = stored_password
        else:
            logging.info("User opted to re-enter password.")
            # Fall through to prompt for new password
            stored_password = None # Ensure we enter the next block correctly

    if not stored_password: # Either not found initially, or user selected 'No'
        logging.info("Prompting for password via GUI dialog...")
        new_password = simpledialog.askstring(
            "Password Required",
            f"Enter COMPANY SSO Password for {username}:",
            show='*',
            parent=root # Associate dialog with hidden root
        )

        if new_password: # If user entered something and didn't cancel
            save_new = messagebox.askyesno(
                 title="Save Password?",
                 message="Save this password to the system keyring for future use?",
                 parent=root
            )
            if save_new:
                try:
                    keyring.set_password(service_name, username, new_password)
                    logging.info("Password saved securely to system keyring.")
                except Exception as e:
                    logging.error(f"Error saving password to keyring: {e}")
                    messagebox.showerror("Keyring Save Error", f"Could not save password to keyring:\n{e}")
            else:
                 logging.info("Password NOT saved to keyring (will use for this session only).")
            password_to_use = new_password
        else:
            # User likely cancelled the dialog
            messagebox.showerror("Password Required", "Password entry cancelled or empty. Cannot proceed.", parent=root)
            logging.warning("Password entry cancelled or empty.")
            password_to_use = None

    # Clean up the tkinter window
    root.destroy()
    return password_to_use

# --- Email Handling ---
def get_email_address(default_email: str | None = None) -> str | None:
    """
    Retrieves the email address to use.
    Confirms a default email or prompts the user via GUI dialogs.
    """
    try:
        root = tk.Tk()
        root.withdraw()
    except tk.TclError as e:
        logging.warning(f"Could not initialize Tkinter GUI environment for email: {e}. Falling back to console/default.")
        # Fallback to console logic - might not work well in exe
        if default_email:
            logging.info(f"Using default email due to GUI fallback: {default_email}")
            # Cannot reliably ask via input() here
            return default_email
        else:
             logging.error("GUI fallback for email failed, and no default email provided.")
             return None # Indicate failure

    email_to_use: str | None = None

    if default_email:
        logging.info(f"Found default email: {default_email}")
        use_default = messagebox.askyesno(
            title="Email Found",
            message=f"Use default email: {default_email}?\n\n(Click 'No' to enter a different one)",
            parent=root
        )
        if use_default:
            logging.info("User opted to use default email.")
            email_to_use = default_email
        else:
            logging.info("User opted to enter email.")
            # Fall through to prompt
            pass # Ensures we enter the next block correctly

    if not email_to_use: # Either no default, or user selected 'No'
        logging.info("Prompting for email via GUI dialog...")
        entered_email = simpledialog.askstring(
            "Email Required",
            "Enter Email Address:",
            parent=root
        )
        if entered_email:
            email_to_use = entered_email.strip()
        else:
            messagebox.showerror("Email Required", "Email entry cancelled or empty. Cannot proceed.", parent=root)
            logging.warning("Email entry cancelled or empty.")
            email_to_use = None

    root.destroy()
    return email_to_use

# --- Helper Function for Data Extraction (within iframe) ---
def extract_item_value(wait: WebDriverWait, itemkey: str) -> str:
    """
    Attempts to extract the text value associated with an itemkey/name
    using common patterns found in the SmartDB HTML structure.
    Prioritizes known specific locators before falling back to generic ones.
    Assumes the driver context is already switched INTO the correct iframe.

    Args:
        wait: The WebDriverWait instance (used for driver access and its default timeout).
        itemkey: The itemkey or name attribute value to look for.

    Returns:
        The extracted text value, or 'Not Found' if unsuccessful.
    """
    value = "Not Found" # Default value
    driver = wait._driver # Get the underlying driver instance

    # <<< Map of specific locators for known problematic/unique fields >>>
    specific_locators = {
        "Modeling Staff": (By.CSS_SELECTOR, "div[itemkey='acModeler'] div.view-mode a"),
        # Add other known specific locators here if identified
        # e.g., "SomeOtherField": (By.ID, "specificIdForField"),
    }

    # <<< Try specific locator first if available >>>
    if itemkey in specific_locators:
        locator_type, locator_value = specific_locators[itemkey]
        specific_timeout = 10 # Allow a decent time for the specific element
        logging.debug(f"    Trying specific locator first for '{itemkey}': {locator_type}='{locator_value}' (Timeout: {specific_timeout}s)")
        try:
            element = WebDriverWait(driver, specific_timeout).until(
                 EC.presence_of_element_located((locator_type, locator_value))
            )
            # Extract text based on tag (same logic as below)
            if element.tag_name == 'textarea': value = element.text
            elif element.tag_name == 'input': value = element.get_attribute('value')
            else: value = element.text
            value = value.strip() if value else ""
            logging.debug(f"    Found value using specific locator: '{value}'")
            return value if value else "" # Return immediately if found
        except TimeoutException:
            logging.debug(f"    Specific locator for '{itemkey}' timed out. Falling back to generic locators.")
        except Exception as e:
            logging.error(f"    Error with specific locator for {itemkey}: {e}. Falling back.")

    # <<< Fallback to generic locators >>>
    logging.debug(f"    Trying generic locators for '{itemkey}'...")
    base_locators = [
        (By.CSS_SELECTOR, f"td[name='{itemkey}'] div.VCenter"),
        (By.CSS_SELECTOR, f"div[itemkey='{itemkey}'] div.VCenter"),
        # Specific pattern for AccountList already handled above, keep here as ultimate fallback maybe? Or remove? Let's keep for now.
        (By.CSS_SELECTOR, f"div[itemkey='{itemkey}'] div.view-mode a"),
        (By.CSS_SELECTOR, f"textarea[name='{itemkey}']"),
        (By.CSS_SELECTOR, f"input[name='{itemkey}']"),
    ]

    # <<< Reduced timeout for generic locator attempts >>>
    generic_timeout_seconds = 3 # Reduced from 5s

    for locator_type, locator_value in base_locators:
        # Skip the specific locator if it was already tried and failed
        if itemkey in specific_locators and (locator_type, locator_value) == specific_locators[itemkey]:
             continue

        try:
            # Use the reduced generic timeout
            logging.debug(f"    Trying generic locator: {locator_type} = {locator_value} with timeout {generic_timeout_seconds}s")
            element = WebDriverWait(driver, generic_timeout_seconds).until(
                 EC.presence_of_element_located((locator_type, locator_value))
            )

            # Extract text differently based on tag type
            if element.tag_name == 'textarea': value = element.text
            elif element.tag_name == 'input': value = element.get_attribute('value')
            else: value = element.text
            value = value.strip() if value else ""

            logging.debug(f"    Found value with generic locator: '{value}'")
            if value or value == "": # If found (even if empty string), break the loop
                return value if value else ""
        except TimeoutException:
            logging.debug(f"    Generic locator timed out.")
            continue # Try the next locator
        except Exception as e:
            logging.error(f"    Error extracting value for {itemkey} with generic locator {locator_value}: {e}")
            return f"Error extracting: {e}"

    logging.debug(f"  -> Final value for {itemkey} after all locators: '{value}'") # Debugging log
    return value # Return 'Not Found' if all locators failed

# --- Main Login Logic ---

def login_to_smartdb(start_url: str, email: str, sso_password: str, target_download_folder: str | None = None) -> tuple[bool, dict, webdriver.Edge | None, str | None]:
    """
    Automates login, extracts text data, and optionally downloads attachments
    directly to a specified folder.

    Args:
        start_url: The initial login page URL.
        email: The user's email address.
        sso_password: The password for the company SSO page.
        target_download_folder: Optional. Absolute path to download files to.
                                If None, download step is skipped.

    Returns:
        A tuple: (bool indicating login/extraction success,
                  dict containing extracted text data,
                  WebDriver instance (if successful, for closing) or None if failed,
                  str path to the temporary user data directory used, or None if failed)
    """
    driver = None # Initialize driver to None for finally block
    login_successful = False
    extracted_data = {} # Dictionary to store results
    temp_user_data_dir = None # <<< Initialize user data dir variable
    service = None # Initialize service to None

    # <<< Enhance Logging: Log inputs >>>
    logging.info("-" * 30)
    logging.info("Attempting SmartDB Login Sequence...")
    logging.info(f"  Start URL: {start_url}")
    logging.info(f"  Email: {email}")
    logging.info(f"  Target Download Folder: {target_download_folder if target_download_folder else 'Not specified'}")
    logging.info("-" * 30)

    if not start_url:
        logging.error("Initial Login URL (start_url) is empty. Cannot proceed.")
        return False, {}, None, None

    try:
        # --- Create a unique temporary user data directory ---
        try:
            temp_user_data_dir = tempfile.mkdtemp(prefix="edge_user_data_")
            logging.info(f"Created temporary user data directory: {temp_user_data_dir}")
            # <<< Add check if directory exists after creation >>>
            if not os.path.isdir(temp_user_data_dir):
                 logging.error("Temporary directory was reported as created, but does not exist. Check permissions/disk space.")
                 # Fallback? For now, return failure.
                 return False, {}, None, temp_user_data_dir # Return path for potential cleanup
        except Exception as temp_err:
            logging.error(f"Failed to create temporary user data directory: {temp_err}", exc_info=True)
            # Propagate failure: return None for driver and the temp dir path
            return False, {}, None, None # <<< Update return signature

        # --- WebDriver Setup ---
        # <<< Enhance Logging: Driver Manager >>>
        driver_path = None
        try:
            logging.info("Initializing WebDriver Manager (EdgeChromiumDriverManager)...")
            # <<< Log driver path resolution >>>
            driver_path = EdgeChromiumDriverManager().install()
            logging.info(f"WebDriver Manager located/installed Edge driver at: {driver_path}")
            if not os.path.exists(driver_path) or not os.access(driver_path, os.X_OK):
                 logging.error(f"Driver file at '{driver_path}' is missing or not executable after installation. Check permissions or antivirus.")
                 # Attempt cleanup of temp dir before exiting
                 if temp_user_data_dir and os.path.exists(temp_user_data_dir):
                      shutil.rmtree(temp_user_data_dir, ignore_errors=True)
                 return False, {}, None, None
        except Exception as driver_manager_err:
            logging.error(f"Failed during WebDriver Manager driver installation/location: {driver_manager_err}", exc_info=True)
            # Attempt cleanup of temp dir before exiting
            if temp_user_data_dir and os.path.exists(temp_user_data_dir):
                 shutil.rmtree(temp_user_data_dir, ignore_errors=True)
            return False, {}, None, None


        # <<< Enhance Logging: Service Setup >>>
        try:
            logging.info(f"Setting up Edge driver service using driver path: {driver_path}...")
            # <<< Pass the explicit path to the service >>>
            service = EdgeService(executable_path=driver_path)
            logging.info("Edge driver service initialized.")
        except Exception as service_err:
            logging.error(f"Failed to initialize Edge driver service: {service_err}", exc_info=True)
            # Attempt cleanup of temp dir before exiting
            if temp_user_data_dir and os.path.exists(temp_user_data_dir):
                 shutil.rmtree(temp_user_data_dir, ignore_errors=True)
            return False, {}, None, None

        # <<< Enhance Logging: Options Setup >>>
        logging.info("Setting up Edge options...")
        options = EdgeOptions()
        options.add_argument("--disable-extensions")
        # options.add_argument("--inprivate") # --user-data-dir replaces the need for --inprivate
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--headless") # Make the browser run headlessly <<< COMMENTED OUT
        options.add_argument("--log-level=0") # Reduce default browser logging clutter

        # --- Add the unique user data directory argument ---
        logging.info(f"Adding user data directory argument: --user-data-dir={temp_user_data_dir}")
        options.add_argument(f"--user-data-dir={temp_user_data_dir}")

        # --- ADD OPTIONS TO IGNORE SSL ERRORS ---
        # Consider uncommenting these if SSL errors are suspected on user machines,
        # but be aware of the security implications.
        # options.add_argument('--ignore-certificate-errors')
        # options.add_argument('--allow-running-insecure-content')
        # options.add_argument('--disable-web-security') # Use with extreme caution
        # -------------------------------------------

        # --- Configure Download Preferences IF target folder provided ---
        prefs = {
            "download.prompt_for_download": False, # Never ask
            "download.directory_upgrade": True, # Allow download to specified dir
            "plugins.always_open_pdf_externally": True # Try to force download over preview
        }
        if target_download_folder:
            logging.info(f"Configuring download directory preference: {target_download_folder}")
            prefs["download.default_directory"] = target_download_folder
        else:
            logging.info("No target download folder provided. Download preferences set for no prompts.")

        options.add_experimental_option("prefs", prefs)
        # -------------------------------------------------------------

        # <<< Enhance Logging: Driver Instantiation >>>
        logging.info("Instantiating webdriver.Edge with configured service and options...")
        logging.debug(f"  Options arguments: {options.arguments}")
        logging.debug(f"  Options experimental options: {options.experimental_options}")
        try:
            driver = webdriver.Edge(service=service, options=options) # Pass updated options
            logging.info("webdriver.Edge instance created successfully.")
            # <<< Add immediate check after instantiation >>>
            if not driver:
                # This case is unlikely as webdriver.Edge usually raises an exception on failure, but check anyway.
                logging.error("webdriver.Edge call completed but returned a None driver object.")
                raise RuntimeError("WebDriver instantiation unexpectedly returned None.")

            # <<< Check initial browser state (optional, might fail if browser closed quickly) >>>
            try:
                initial_handles = driver.window_handles
                logging.info(f"Initial window handles: {initial_handles} (Count: {len(initial_handles)})")
                initial_url = driver.current_url
                logging.info(f"Initial browser URL after launch: '{initial_url}'") # Expect 'data:,' or similar
            except Exception as post_launch_check_err:
                logging.warning(f"Could not perform post-launch checks on browser state: {post_launch_check_err}")

        except Exception as driver_init_err:
            logging.error(f"Failed to instantiate webdriver.Edge: {driver_init_err}", exc_info=True)
            # Attempt cleanup of temp dir before exiting
            if temp_user_data_dir and os.path.exists(temp_user_data_dir):
                 shutil.rmtree(temp_user_data_dir, ignore_errors=True)
            # Return None driver explicitly
            return False, {}, None, temp_user_data_dir

        # --- Setup Waits ---
        driver.implicitly_wait(0) # <<< SET IMPLICIT WAIT TO ZERO
        wait = WebDriverWait(driver, 15) # <<< Reduced standard wait to 15s
        short_wait = WebDriverWait(driver, 7)  # <<< Added shorter wait for optional elements (7s)

        # --- Initial Navigation ---
        logging.info(f"Navigating to initial login page: {start_url}")
        try:
            driver.get(start_url)
            logging.info("driver.get() command executed.")
            # <<< Add check immediately after get() >>>
            current_url_after_get = driver.current_url
            logging.info(f"Current URL immediately after get('{start_url}'): '{current_url_after_get}'")
            # If the URL is still blank ('', 'data:,', etc.) it might indicate an immediate navigation error
            # or a very fast redirect to an error page that wasn't caught.
            if not current_url_after_get or current_url_after_get.startswith('data:'):
                 logging.warning(f"URL after driver.get() is '{current_url_after_get}'. Navigation might have failed silently or blocked. Proceeding with checks...")
                 # Consider taking a screenshot here if it fails later
                 # driver.save_screenshot("post_initial_get_blank.png")

        except Exception as get_err:
            logging.error(f"Error during driver.get('{start_url}'): {get_err}", exc_info=True)
            if driver:
                 try: driver.save_screenshot("driver_get_error.png"); logging.info("Saved screenshot: driver_get_error.png")
                 except Exception as ss_err: logging.error(f"Failed to save screenshot on driver.get() error: {ss_err}")
            raise # Re-raise the exception to be caught by the outer block

        # --- Step 1: Initial Email Entry ---
        logging.info("Waiting for initial email input...")
        email_input_locator = (By.CSS_SELECTOR, '[data-testid="username-input"]')
        email_input = wait.until(EC.visibility_of_element_located(email_input_locator), "Could not find initial email input")
        logging.info("Entering email on initial page...")
        email_input.send_keys(email)

        # Wait for the 'Next' button to become enabled/clickable
        next_button_locator = (By.CSS_SELECTOR, '[data-testid="next-button"]')
        next_button = wait.until(EC.element_to_be_clickable(next_button_locator), "Could not find or click initial 'Next' button")
        logging.info("Clicking 'Next'...")
        # Sometimes a small delay helps before clicking after enabling
        time.sleep(0.5)
        next_button.click()

        # --- Step 2: Click 'Sign in with Microsoft' ---
        logging.info("Looking for 'Sign in with Microsoft' button...")
        ms_signin_button_locator = (By.CSS_SELECTOR, '[data-testid="azureAuthButton"]')
        ms_signin_button = wait.until(EC.element_to_be_clickable(ms_signin_button_locator), "Could not find or click 'Sign in with Microsoft' button")
        logging.info("Clicking 'Sign in with Microsoft'...")
        ms_signin_button.click()

        # --- Step 3: Microsoft Login Page ---
        logging.info("Waiting for Microsoft login page (login.microsoftonline.com)...")
        ms_email_input_locator = (By.CSS_SELECTOR, 'input[type="email"]')
        try:
            logging.info("Attempting to find email input on Microsoft page...")
            # <<< Use short_wait for finding the optional MS email input >>>
            ms_email_input = short_wait.until(EC.presence_of_element_located(ms_email_input_locator), "Could not find email input presence on Microsoft page within 7s")
            ms_email_input = short_wait.until(EC.visibility_of(ms_email_input), "Email input on Microsoft page not visible within 7s")
            logging.info("Entering email on Microsoft page...")
            ms_email_input.send_keys(email)

            logging.info("Finding Microsoft 'Next' button...")
            ms_next_button_locator = (By.CSS_SELECTOR, 'input[type="submit"][value="Next"]')
            # <<< Use short_wait for the button after input is found >>>
            ms_next_button = short_wait.until(EC.element_to_be_clickable(ms_next_button_locator), "Could not find 'Next' button on Microsoft page within 7s")
            logging.info("Clicking 'Next' on Microsoft page...")
            ms_next_button.click()
        except TimeoutException:
            # <<< Update log message to reflect shorter wait >>>
            logging.warning("Did not find the standard Microsoft email input page/flow within 7s.")
            logging.warning("Maybe auto-login or different flow occurred. Proceeding...")
            if driver:
                try:
                     driver.save_screenshot("ms_login_page_skipped_or_changed.png")
                     logging.info("Saved screenshot: ms_login_page_skipped_or_changed.png")
                except Exception as ss_err:
                     logging.error(f"Failed to save screenshot during MS login skip: {ss_err}")

        # --- ADD CHECK: Are we already on the target site? ---
        already_on_smartdb = False
        try:
            # <<< Use a minimal delay >>>
            time.sleep(0.5)
            current_url_after_ms = driver.current_url
            if ("smartdb.jp" in current_url_after_ms and
                "login" not in current_url_after_ms.lower() and
                "microsoft" not in current_url_after_ms.lower() and
                "sso.osgapp" not in current_url_after_ms.lower()):
                 logging.info(f"Detected already on SmartDB page after Microsoft step: {current_url_after_ms}. Skipping SSO step.")
                 already_on_smartdb = True
            # else: # No explicit log needed if proceeding, Step 4 log covers it
            #     logging.info(f"Not yet on SmartDB page after Microsoft step. Current URL: {current_url_after_ms}. Proceeding to SSO step.")
        except Exception as url_check_err:
             logging.warning(f"Could not reliably check URL after Microsoft step: {url_check_err}. Assuming SSO step is needed.")
        # --- END CHECK ---

        # --- Step 4: Company SSO Page (sso.osgapp.net) ---
        if not already_on_smartdb:
            logging.info("Waiting for Company SSO page (sso.osgapp.net)...")
            sso_email_locator = (By.ID, 'Upn')
            sso_password_locator = (By.ID, 'UpnPassword')
            sso_login_button_locator = (By.CSS_SELECTOR, 'input[type="submit"][name="mailSignIn"]')

            try:
                # <<< Use standard wait (15s) here, as SSO page is necessary if reached >>>
                logging.info("Waiting for SSO email input (should be pre-filled)...")
                sso_email_input = wait.until(EC.visibility_of_element_located(sso_email_locator), "Could not find email input on SSO page")
                logging.info(f"Found email field on SSO page (ID: {sso_email_locator[1]}). Value: '{sso_email_input.get_attribute('value')}'")

                logging.info("Waiting for SSO password input...")
                sso_password_input = wait.until(EC.visibility_of_element_located(sso_password_locator), "Could not find password input on SSO page")
                logging.info("Entering password on SSO page...")
                sso_password_input.send_keys(sso_password)

                logging.info("Waiting for SSO login button...")
                sso_login_button = wait.until(EC.element_to_be_clickable(sso_login_button_locator), "Could not find or click Login button on SSO page")
                logging.info("Clicking Login on SSO page...")
                sso_login_button.click()

                # --- Step 4.5: Handle 'Stay signed in?' Prompt ---
                logging.info("Checking for 'Stay signed in?' prompt (after SSO)...")
                try:
                    stay_signed_in_no_button_locator = (By.ID, 'idBtn_Back')
                    # <<< Use short_wait for this optional prompt >>>
                    no_button = short_wait.until(
                        EC.element_to_be_clickable(stay_signed_in_no_button_locator),
                        "'Stay signed in?' prompt 'No' button not found or clickable within 7s." # Update message
                    )
                    logging.info("'Stay signed in?' prompt found. Clicking 'No'...")
                    no_button.click()
                except TimeoutException:
                    # <<< Update log message >>>
                    logging.info("'Stay signed in?' prompt not detected or timed out within 7s after SSO. Continuing...")
                except Exception as e:
                    logging.warning(f"An unexpected error occurred while handling 'Stay signed in?' prompt after SSO: {e}")
                # --- End of Step 4.5 ---

            except TimeoutException as e:
                logging.error(f"Timeout interacting with expected SSO page elements: {e}")
                logging.error("This happened even after checking if we were already on SmartDB. The page might be unexpected.")
                if driver:
                     try:
                          driver.save_screenshot("sso_page_unexpected_error.png")
                          logging.info("Saved screenshot: sso_page_unexpected_error.png")
                     except Exception as ss_err:
                          logging.error(f"Failed to save screenshot during SSO unexpected error: {ss_err}")
                raise
        # <<< End of "if not already_on_smartdb" block >>>

        # --- Step 5: Wait for Final Dashboard ---
        logging.info("Waiting for final application page confirmation...")
        try:
            # <<< Use standard wait (15s) for the crucial final URL confirmation >>>
            final_url_loaded = wait.until(
                lambda drv: (
                    "smartdb.jp" in drv.current_url and
                    "login" not in drv.current_url.lower() and
                    "microsoft" not in drv.current_url.lower() and
                    "sso.osgapp" not in drv.current_url.lower()
                ),
                message="URL did not change to the expected application domain after login redirects within 15s." # Update message
            )

            if final_url_loaded:
                final_url = driver.current_url
                logging.info(f"Login successful. Final URL is on smartdb.jp: {final_url}")
                login_successful = True

        except TimeoutException as e:
             final_url = driver.current_url if driver else "N/A"
             logging.error(f"Login likely failed. Timed out waiting for URL to change. Current URL: {final_url}")
             if driver:
                  try:
                      driver.save_screenshot("final_redirect_error.png")
                      logging.info("Saved screenshot: final_redirect_error.png")
                  except Exception as ss_err:
                       logging.error(f"Failed to save screenshot during final redirect error: {ss_err}")
             login_successful = False
             # No need to quit driver here, let the outer block handle it
             raise # Re-raise TimeoutException
        except Exception as e:
             logging.error(f"An unexpected error occurred during final URL check: {e}", exc_info=True)
             if driver:
                  try:
                       driver.save_screenshot("final_redirect_unexpected_error.png")
                       logging.info("Saved screenshot: final_redirect_unexpected_error.png")
                  except Exception as ss_err:
                       logging.error(f"Failed to save screenshot during final redirect unexpected error: {ss_err}")
             login_successful = False
             raise # Re-raise the exception


        # --- Step 6: Extract Text Data & Download Files (if Login Successful) ---
        if login_successful:
            logging.info("Login successful. Attempting data extraction and downloads...")
            try:
                # --- Switch to the IFrame ---
                iframe_locator = (By.CSS_SELECTOR, 'iframe[title="binderTopPage"]')
                logging.info(f"Waiting for and switching to iframe: {iframe_locator}")
                wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
                logging.info("Successfully switched to iframe.")

                # --- Extract TEXT fields ---
                for field_name, itemkey in FIELDS_TO_EXTRACT.items():
                    if itemkey not in FILE_ITEMKEYS:
                        logging.info(f"  Extracting '{field_name}' (itemkey={itemkey})...")
                        value = extract_item_value(wait, itemkey)
                        extracted_data[field_name] = value
                        logging.info(f"    -> Value: '{value}'")

                # --- Download Files (if target folder is set) ---
                downloaded_files_info = {}
                if target_download_folder:
                    logging.info(f"\nAttempting file downloads to temp folder: {target_download_folder}")
                    # Ensure the target folder exists, create if not
                    try:
                        os.makedirs(target_download_folder, exist_ok=True)
                    except OSError as e:
                         logging.error(f"  ERROR: Could not create target download directory '{target_download_folder}': {e}")
                         # Store error for all potential files if directory fails
                         for itemkey in FILE_ITEMKEYS:
                             downloaded_files_info[itemkey] = {"error": f"Cannot create directory: {e}", "status": "Failed"}
                         target_download_folder = None # Prevent further download attempts

                    if target_download_folder:
                        for itemkey in FILE_ITEMKEYS:
                            file_downloaded_path = None
                            file_download_error = None
                            expected_filename_from_link = "Unknown Filename"

                            try:
                                logging.info(f"  Looking for download links for itemkey '{itemkey}'...")
                                link_elements = driver.find_elements(By.CSS_SELECTOR, f"div[itemkey='{itemkey}'] div.view-mode a.TextLink")

                                if not link_elements:
                                    logging.info(f"  No file links found for itemkey '{itemkey}'.")
                                    continue

                                if len(link_elements) > 1:
                                    logging.warning(f"  Found {len(link_elements)} links for itemkey '{itemkey}'. Downloading only the first.")
                                link_element = link_elements[0]

                                expected_filename_from_link = link_element.text.strip()
                                if not expected_filename_from_link:
                                    logging.warning(f"  Link for itemkey '{itemkey}' has no text. Using fallback filename.")
                                    href = link_element.get_attribute('href')
                                    if href:
                                        try:
                                            expected_filename_from_link = os.path.basename(href.split('?')[0])
                                            logging.info(f"  Using filename from href: {expected_filename_from_link}")
                                        except Exception:
                                             expected_filename_from_link = f"downloaded_file_{itemkey}"
                                    else:
                                         expected_filename_from_link = f"downloaded_file_{itemkey}"

                                logging.info(f"  Preparing download for: '{expected_filename_from_link}' (from itemkey '{itemkey}')")

                                files_before = set(os.listdir(target_download_folder))
                                try:
                                    logging.info("    Clicking download link...")
                                    driver.execute_script("arguments[0].click();", link_element)
                                    logging.info("    Link click executed.")
                                except Exception as click_err:
                                    logging.error(f"    ERROR clicking link for {expected_filename_from_link}: {click_err}", exc_info=True)
                                    file_download_error = f"Click Failed: {click_err}"

                                if not file_download_error:
                                    logging.info(f"    Waiting for download to complete in '{target_download_folder}' (expected related to: '{expected_filename_from_link}')...")
                                    download_wait_timeout = 120
                                    poll_interval = 1
                                    stability_delay = 2
                                    start_wait_time = time.monotonic()
                                    new_file_path = None
                                    last_size = -1

                                    while time.monotonic() - start_wait_time < download_wait_timeout:
                                        try:
                                            current_files = set(os.listdir(target_download_folder))
                                        except FileNotFoundError:
                                            logging.error(f"    ERROR: Target download folder '{target_download_folder}' seems to have disappeared.")
                                            file_download_error = "Temporary download folder vanished"
                                            break
                                        except Exception as list_err:
                                            logging.warning(f"    WARNING: Error listing files in temporary folder: {list_err}. Retrying...")
                                            time.sleep(poll_interval)
                                            continue

                                        new_files = current_files - files_before
                                        potential_final_file_path = None
                                        found_temp = False

                                        # logging.debug(f"    DEBUG: Time {int(time.monotonic() - start_wait_time)}s. New files found: {new_files}") # Verbose Debug

                                        if new_files:
                                            non_temp_files = [os.path.join(target_download_folder, fname) for fname in new_files if not fname.lower().endswith((".crdownload", ".tmp", ".part", ".download"))]

                                            if len(non_temp_files) == 1:
                                                potential_final_file_path = non_temp_files[0]
                                                # logging.debug(f"    DEBUG: Found potential final file: {os.path.basename(potential_final_file_path)}")
                                            elif len(non_temp_files) > 1:
                                                logging.warning(f"    WARNING: Multiple non-temporary files appeared: { [os.path.basename(f) for f in non_temp_files] }. Attempting fallback match...")
                                                for fpath in non_temp_files:
                                                    if expected_filename_from_link.lower() in os.path.basename(fpath).lower():
                                                         potential_final_file_path = fpath
                                                         # logging.debug(f"    DEBUG: Using fallback match: {os.path.basename(potential_final_file_path)}")
                                                         break
                                                if not potential_final_file_path:
                                                     logging.warning(f"    WARNING: Fallback match failed. Cannot determine correct downloaded file.")

                                            if not potential_final_file_path:
                                                for fname in new_files:
                                                     if fname.lower().endswith((".crdownload", ".tmp", ".part", ".download")):
                                                          found_temp = True
                                                          break

                                            if potential_final_file_path:
                                                try:
                                                    current_size = os.path.getsize(potential_final_file_path)
                                                    # logging.debug(f"    DEBUG: Checking stability for {os.path.basename(potential_final_file_path)}. Size: {current_size}")

                                                    if current_size == last_size and current_size > 0:
                                                        logging.info(f"    ...size stable ({current_size} bytes). Waiting {stability_delay}s for final check...")
                                                        time.sleep(stability_delay)
                                                        final_size = os.path.getsize(potential_final_file_path)
                                                        if final_size == current_size:
                                                             new_file_path = potential_final_file_path
                                                             logging.info(f"    DEBUG: File confirmed stable: {os.path.basename(new_file_path)}")
                                                             break # Exit while loop
                                                        else:
                                                             logging.info(f"    ...size changed after stability delay ({current_size} -> {final_size}). Resetting check.")
                                                             last_size = final_size
                                                    elif current_size > 0:
                                                         last_size = current_size
                                                    else:
                                                         last_size = 0

                                                except FileNotFoundError:
                                                    logging.warning(f"    WARNING: Potential file '{os.path.basename(potential_final_file_path)}' disappeared during size check. Resetting.")
                                                    potential_final_file_path = None
                                                    last_size = -1
                                                except OSError as stat_err:
                                                    logging.warning(f"    WARNING: Error getting size/status for '{os.path.basename(potential_final_file_path)}': {stat_err}. Might be locked, retrying...")
                                                    last_size = -1
                                                except Exception as stat_err:
                                                     logging.warning(f"    WARNING: Unexpected error getting size/status for '{os.path.basename(potential_final_file_path)}': {stat_err}. Retrying...")
                                                     last_size = -1

                                            elif found_temp:
                                                logging.info(f"    ...download in progress (found temp file)...")
                                                last_size = -1

                                        time.sleep(poll_interval)
                                    # --- End of while loop ---

                                    if new_file_path:
                                        file_downloaded_path = new_file_path
                                        expected_filename_from_link = os.path.basename(file_downloaded_path) # Use actual name
                                        logging.info(f"    SUCCESS: Download completed and file stable. Detected file: {expected_filename_from_link}")
                                    elif file_download_error: # Folder disappeared error
                                        logging.error(f"    ERROR: {file_download_error}")
                                    else: # Timeout
                                        elapsed_wait = time.monotonic() - start_wait_time
                                        file_download_error = f"Download Timeout/Error: File download did not complete or stabilize in '{target_download_folder}' within {download_wait_timeout}s."
                                        logging.error(f"    ERROR: {file_download_error}")
                                        try:
                                             logging.error(f"    Files present in '{target_download_folder}' on timeout: {os.listdir(target_download_folder)}")
                                        except Exception as list_err:
                                             logging.error(f"    Could not list files in target folder for debugging on timeout: {list_err}")

                            except Exception as find_err:
                                logging.error(f"    ERROR finding or processing link for itemkey {itemkey}: {find_err}", exc_info=True)
                                file_download_error = f"Error finding/processing link: {find_err}"

                            # --- Store result for this file itemkey ---
                            if file_downloaded_path:
                                actual_filename = os.path.basename(file_downloaded_path)
                                downloaded_files_info[itemkey] = {"path": file_downloaded_path, "status": "Success", "filename_actual": actual_filename, "filename_expected": expected_filename_from_link}
                            else:
                                error_msg = file_download_error or "Unknown download error"
                                downloaded_files_info[itemkey] = {"error": error_msg, "status": "Failed", "filename_expected": expected_filename_from_link}

                    # --- Consolidate download results into extracted_data ---
                    if downloaded_files_info:
                         extracted_data["Downloaded Files"] = downloaded_files_info
                         # Set Primary Path ONLY IF the file was moved successfully (done in scrape_data now)
                         # if any(info.get("status") == "Success (Moved)" for info in downloaded_files_info.values()):
                         #     extracted_data["Primary Downloaded File Path"] = target_download_folder # Use the *final* target folder
                         #     logging.info(f"  Set 'Primary Downloaded File Path' to final target folder: {target_download_folder}")


                    # --- Report overall download status ---
                    if downloaded_files_info:
                        success_count = sum(1 for info in downloaded_files_info.values() if info.get("status") == "Success") # Before move attempt
                        fail_count = len(downloaded_files_info) - success_count
                        if fail_count > 0:
                            logging.warning(f"\nInitial file download process completed with {success_count} success(es) and {fail_count} failure(s).")
                            for key, info in downloaded_files_info.items():
                                 if info.get("status") == "Failed":
                                      logging.warning(f"  - Failed initial download itemkey '{key}' (Expected: {info.get('filename_expected', 'N/A')}): {info.get('error', 'Unknown reason')}")
                        elif success_count > 0:
                            logging.info("\nInitial file download process completed successfully (before move).")
                        else:
                            logging.info("\nInitial file download process finished, but no files were successfully downloaded (or none were found).")
                else: # if not target_download_folder
                    logging.info("\nSkipping file download step (no valid target folder specified).")

            except TimeoutException as frame_e:
                logging.error(f"Could not find or switch to the iframe (title='binderTopPage'): {frame_e}", exc_info=True) # Log traceback
                if driver:
                     try:
                          driver.save_screenshot("iframe_switch_error.png")
                          logging.info("Saved screenshot: iframe_switch_error.png")
                     except Exception as ss_err:
                          logging.error(f"Failed to save screenshot during iframe error: {ss_err}")
                login_successful = False
                raise # Re-raise exception
            except Exception as e:
                logging.error(f"An error occurred during data extraction or download: {e}", exc_info=True)
                if driver:
                    try:
                         driver.save_screenshot("data_extraction_error.png")
                         logging.info("Saved screenshot: data_extraction_error.png")
                    except Exception as ss_err:
                         logging.error(f"Failed to save screenshot during data extraction error: {ss_err}")
                login_successful = False
                raise # Re-raise exception
            finally:
                # --- Switch back BUT keep driver open (unless error occurred) ---
                try:
                    # Only switch back if driver is still valid
                    if driver and login_successful:
                        logging.info("Switching back to default content from iframe...")
                        driver.switch_to.default_content()
                        logging.info("Switched back successfully.")
                except Exception as switch_back_err:
                    logging.error(f"Error switching back from iframe: {switch_back_err}", exc_info=True) # Log traceback
                    # Consider this a failure if switch back doesn't work
                    login_successful = False
                    # Let the main exception handler decide whether to quit driver


    except TimeoutException as e:
        logging.error(f"A timeout occurred during the main login flow.")
        logging.error(f"Timeout details: {e}")
        logging.error(traceback.format_exc()) # <<< Log full traceback
        if driver:
             current_url = "N/A"
             try: current_url = driver.current_url
             except: pass
             logging.error(f"Current URL at timeout: {current_url}")
             try: driver.save_screenshot("timeout_error.png"); logging.info("Saved screenshot: timeout_error.png")
             except Exception as ss_err: logging.error(f"Failed to save screenshot on TimeoutException: {ss_err}")
        # Return None for driver and the temp_user_data_dir path
        return False, extracted_data, None, temp_user_data_dir # <<< Update return
    except NoSuchElementException as e:
        logging.error(f"A required element was not found during login.")
        logging.error(f"Element details: {e}")
        logging.error(traceback.format_exc()) # <<< Log full traceback
        if driver:
             current_url = "N/A"
             try: current_url = driver.current_url
             except: pass
             logging.error(f"Current URL at NoSuchElementException: {current_url}")
             try: driver.save_screenshot("element_not_found_error.png"); logging.info("Saved screenshot: element_not_found_error.png")
             except Exception as ss_err: logging.error(f"Failed to save screenshot on NoSuchElementException: {ss_err}")
        return False, extracted_data, None, temp_user_data_dir # <<< Update return
    except Exception as e:
        # Catch any other exception during the process, including WebDriver instantiation errors if not caught earlier
        logging.error(f"An unexpected error occurred during login/scraping.")
        # <<< Log the specific exception type and message >>>
        logging.error(f"Error Type: {type(e).__name__}")
        logging.error(f"Error Message: {e}")
        logging.error(traceback.format_exc()) # <<< Log full traceback
        if driver:
            current_url = "N/A"
            try: current_url = driver.current_url
            except: pass
            logging.error(f"Current URL at unexpected error: {current_url}")
            try:
                driver.save_screenshot("unexpected_error.png"); logging.info("Saved screenshot: unexpected_error.png")
            except Exception as ss_err: logging.error(f"Failed to save screenshot on general Exception: {ss_err}")
        # Return potentially partial data, None for driver, but include the temp_user_data_dir path for cleanup
        return False, extracted_data, None, temp_user_data_dir # <<< Update return
    finally:
        # <<< Add logging to the finally block >>>
        logging.info("Entering login_to_smartdb finally block...")
        # The existing logic to switch back from iframe is fine here.
        # The driver is NO LONGER quit in this function's finally block.
        # It's returned to the caller for cleanup.
        logging.info("Exiting login_to_smartdb finally block.")
        # --- Switch back BUT keep driver open (unless error occurred) ---
        try:
            # Only switch back if driver is still valid
            if driver and login_successful:
                logging.info("Switching back to default content from iframe...")
                driver.switch_to.default_content()
                logging.info("Switched back successfully.")
        except Exception as switch_back_err:
            logging.error(f"Error switching back from iframe: {switch_back_err}", exc_info=True) # Log traceback
            # Consider this a failure if switch back doesn't work
            login_successful = False
            # Let the main exception handler decide whether to quit driver


    # --- Return Results ---
    if login_successful:
        logging.info("Login and text extraction successful. Returning driver instance.")
        return True, extracted_data, driver, temp_user_data_dir # <<< Update return
    else:
        # This part should theoretically not be reached if exceptions are raised correctly above
        logging.error("Login or text extraction failed (unexpected non-exception failure).")
        if driver:
             logging.warning("Login failed but driver instance exists - attempting quit.")
             try: driver.quit()
             except Exception as q_err: logging.error(f"Error quitting driver on non-exception failure: {q_err}")
        # Return failure, potentially partial data, None driver, but include temp_user_data_dir
        return False, extracted_data, None, temp_user_data_dir # <<< Update return


# --- Removed Download and Move Function ---
# ... (function completely removed) ...


# --- Script Execution (if run directly) ---
def main():
    # Use logging instead of print
    log_filename = "smartdb_automation.log"
    file_handler = None # Initialize file_handler variable
    try:
        if getattr(sys, 'frozen', False):
            # Running as bundled executable (e.g., PyInstaller)
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as a script
            base_path = os.path.dirname(os.path.abspath(__file__))
        log_file_path = os.path.join(base_path, log_filename)
        
        with open(log_file_path, 'a') as f: # Use 'a' append mode
             f.write("") # Test write permission
             
    except Exception as log_path_err:
         print(f"[ERROR] Could not determine or write to script/exe directory log path: {log_path_err}")
         log_file_path = log_filename 
         print(f"[CRITICAL] Falling back to logging in current working directory: {os.path.abspath(log_file_path)}")

    # <<< Explicit handler setup >>>
    root_logger = logging.getLogger()
    # Remove existing handlers (important if this script is run multiple times in same process)
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
        handler.close()

    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s')

    # File Handler
    try:
        file_handler = logging.FileHandler(log_file_path, mode='w') # Overwrite
        file_handler.setFormatter(log_formatter)
        root_logger.addHandler(file_handler)
    except Exception as e:
        print(f"[ERROR] Failed to create file handler for '{log_file_path}': {e}")
        file_handler = None

    # Stream Handler
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(log_formatter)
    root_logger.addHandler(stream_handler)

    # Set level
    root_logger.setLevel(logging.INFO)

    # <<< Register cleanup function >>>
    def cleanup_logging():
        logging.info("Flushing log handlers during exit (atexit)...")
        for handler in logging.getLogger().handlers:
            try:
                handler.flush()
                handler.close()
            except: # Ignore errors during cleanup
                pass
    atexit.register(cleanup_logging)
    # -----------------------------------

    logging.info(f"--- Log Start (SmartDB Standalone) ---")
    logging.info(f"Logging to: {log_file_path}") 
    # <<< Add immediate test message and flush >>>
    logging.info("--- Initial logging setup complete (standalone). This message should appear. ---")
    if file_handler:
        file_handler.flush()
    # ------------------------------------------------------------------------------------

    # --- Get Email ---
    email_to_login = get_email_address(USER_EMAIL) # Use constant as default
    if not email_to_login:
        logging.error("No email address provided. Exiting.")
        sys.exit(1)
    logging.info(f"Using email: {email_to_login}")

    # --- Get Password (using the obtained email) ---
    sso_password = store_or_get_password(KEYRING_SERVICE_NAME, email_to_login)
    if not sso_password:
        logging.error("Could not get SSO password (user cancelled or entry failed). Exiting.")
        sys.exit(1) # Exit if password fails

    # --- Define a target directory for standalone testing ---
    test_download_target = os.path.join(DEFAULT_DOWNLOAD_FOLDER, "smartdb_standalone_test_downloads")
    try:
         os.makedirs(test_download_target, exist_ok=True)
         logging.info(f"Using standalone download target: {test_download_target}")
    except Exception as e:
         logging.error(f"Could not create standalone download target folder: {e}. Downloads may fail.")
         test_download_target = None
    # ---------------------------------------------------------

    # Attempt the login and get data
    success = False
    data = {}
    driver = None
    temp_user_dir_path_standalone = None # For cleanup
    try:
        success, data, driver, temp_user_dir_path_standalone = login_to_smartdb(
            INITIAL_LOGIN_URL, email_to_login, sso_password, target_download_folder=test_download_target
        )

        if success:
            logging.info("\nStandalone Login and data extraction process completed successfully.")
            logging.info(f"Extracted Data: {data}")
            # <<< Add a small delay before closing in standalone for visibility >>>
            logging.info("Pausing for 5 seconds before closing browser...")
            time.sleep(5)
        else:
            # <<< Improve message when success is False >>>
            logging.error("\nStandalone Login process failed or did not complete successfully.")
            logging.error("Please check the log file ('smartdb_automation.log') for detailed error messages.")
            # <<< Add pause on failure too, so user can see the empty browser/last state >>>
            logging.error("Pausing for 10 seconds before attempting cleanup...")
            time.sleep(10)

    except Exception as standalone_e:
         logging.error(f"A critical error occurred in the main execution block: {standalone_e}", exc_info=True)
         driver = None 
         logging.error("Pausing for 10 seconds after critical error...")
         if file_handler: file_handler.flush() # Flush logs before sleep
         time.sleep(10)
    finally:
        logging.info("Entering main finally block for cleanup (standalone)...")
        # --- Close driver if open ---
        if driver:
            logging.info("Closing WebDriver (standalone)...")
            try:
                driver.quit()
                logging.info("WebDriver quit successfully.")
            except Exception as q_err:
                 logging.error(f"Error quitting WebDriver in standalone finally: {q_err}")

        # --- Clean up temporary user data directory (standalone) ---
        if temp_user_dir_path_standalone and os.path.exists(temp_user_dir_path_standalone):
            logging.info(f"Cleaning up temporary user data directory (standalone): {temp_user_dir_path_standalone}")
            try:
                # <<< Add retry logic for shutil.rmtree, sometimes needed on Windows >>>
                retries = 3
                delay = 1
                for i in range(retries):
                    try:
                        shutil.rmtree(temp_user_dir_path_standalone)
                        logging.info(f"Successfully removed temporary directory.")
                        break # Success
                    except OSError as clean_err:
                         logging.warning(f"Attempt {i+1}/{retries}: Could not remove temporary user data directory '{temp_user_dir_path_standalone}': {clean_err}. Retrying in {delay}s...")
                         time.sleep(delay)
                else: # If loop finished without break
                     logging.error(f"Failed to remove temporary directory after {retries} attempts: {temp_user_dir_path_standalone}")

            except Exception as clean_err: # Catch other potential errors
                logging.error(f"An unexpected error occurred during temporary directory cleanup: {clean_err}")
        elif temp_user_dir_path_standalone:
             logging.warning(f"Temporary user data directory was specified ({temp_user_dir_path_standalone}) but not found for cleanup.")
        else:
             logging.info("No temporary user data directory path recorded for cleanup.")

        logging.info("Main finally block cleanup finished (standalone).")
        # <<< Flush again here just before exiting the main function's finally block >>>
        # Although atexit should handle it, adding it here provides extra safety
        if file_handler:
             try:
                 file_handler.flush()
             except: pass


if __name__ == "__main__":
    # --- IMPORTANT PRE-REQUISITES ---
    # 1. Ensure Microsoft Edge browser is installed.
    # 2. You WILL need to inspect the Company SSO page (sso.osgapp.net)
    #    and potentially the Microsoft page (login.microsoftonline.com)
    #    using your browser's Developer Tools (F12 -> Elements tab) to find the
    #    correct CSS SELECTORS or IDs for the input fields and buttons.
    #    Update the placeholder locators in the `login_to_smartdb` function.
    #    Check if the SSO form is inside an <IFRAME>. If so, uncomment and adjust
    #    the iframe handling code in Step 4.
    # 3. You also need a reliable selector for the final dashboard page
    #    (replace 'dashboard-element-id').
    # 4. Run `pip install selenium webdriver-manager keyring` if you haven't already.
    # <<< Add instruction about checking the log file >>>
    # 5. If the script fails, check the generated 'smartdb_automation.log' file in the same directory
    #    as the script/executable for detailed error information.
    # ---------------------------------
    main()