import os
import time
import shutil
from tkinter import Tk, filedialog
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager

def select_json_files():
    """Prompt the user to select JSON files until they cancel."""
    Tk().withdraw()  # Hide the root Tkinter window
    all_selected_files = []  # List to store all selected files

    while True:
        # Prompt the user to select JSON files
        input_json_files = filedialog.askopenfilenames(
            title="Select JSON Files",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not input_json_files:
            # If the user cancels, break the loop
            print("File selection canceled.")
            break
        else:
            print(f"Selected files: {input_json_files}")
            all_selected_files.extend(input_json_files)

    if not all_selected_files:
        print("No JSON files selected. Exiting.")
        return None

    print(f"All selected JSON files: {all_selected_files}")
    return all_selected_files

def open_website_and_interact(input_json_files):
    """Interact with the website to upload the JSON files and download the JSON output files."""
    # Set a hardcoded download directory
    download_folder = os.path.join(os.getcwd(), "downloads")
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)  # Create the folder if it doesn't exist

    print(f"Download directory set to: {download_folder}")

    # Set up Edge options for dynamic download directory
    edge_options = Options()
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--headless")  # Run in headless mode (hide the browser)

    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)

    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        driver.get("https://tools.finexo.in/gstr-1")
        print(f"Website Title: {driver.title}")

        # Wait for the file input element to appear
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        print(f"Uploading JSON files: {input_json_files}")
        combined_files = "\n".join(input_json_files)
        file_input.send_keys(combined_files)
        print("JSON files uploaded successfully.")

        # Wait for the "Download JSON" button to appear
        download_json_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Download JSON')]"))
        )
        print("Download JSON button is ready. Clicking it...")
        download_json_button.click()

        # Wait for the "Download GSTR-1 JSON" button to appear
        download_gstr1_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Download GSTR-1 JSON')]"))
        )
        print("Download GSTR-1 JSON button is ready. Clicking it...")
        download_gstr1_button.click()

        # Wait for the files to download
        print("Waiting for the files to download...")
        time.sleep(5)

        # Verify if the files exist in the hardcoded download folder
        downloaded_files = [f for f in os.listdir(download_folder) if f.endswith(".json")]
        if downloaded_files:
            print(f"JSON files downloaded successfully: {downloaded_files}")
        else:
            print("No JSON files found in the download folder.")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        print("Clearing browser cache and cookies...")
        driver.delete_all_cookies()

        # Close the browser
        driver.quit()

def upload_json_and_download_excel(json_file):
    """Upload the moved JSON file and download the output Excel file."""
    # Set a hardcoded download directory for Excel files
    excel_download_folder = os.path.join(os.getcwd(), "excel_downloads")
    if not os.path.exists(excel_download_folder):
        os.makedirs(excel_download_folder)  # Create the folder if it doesn't exist

    print(f"Excel download directory set to: {excel_download_folder}")

    # Set up Edge options for dynamic download directory
    edge_options = Options()
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--headless")  # Run in headless mode (hide the browser)

    prefs = {
        "download.default_directory": excel_download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)

    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        driver.get("https://tools.finexo.in/gstr-1")
        print(f"Website Title: {driver.title}")

        # Wait for the file input element to appear
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        print(f"Uploading JSON file: {json_file}")
        file_input.send_keys(json_file)
        print("JSON file uploaded successfully.")

        # Wait for the "Download Excel" button to appear
        download_excel_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Download Excel')]"))
        )
        print("Download Excel button is ready. Clicking it...")
        download_excel_button.click()

        # Wait for the file to download
        print("Waiting for the Excel file to download...")
        time.sleep(5)

        # Verify if the Excel file exists in the hardcoded download folder
        downloaded_files = [f for f in os.listdir(excel_download_folder) if f.endswith(".xlsx")]
        if downloaded_files:
            print(f"Excel file downloaded successfully: {downloaded_files}")
        else:
            print("No Excel files found in the download folder.")

    except Exception as e:
        print(f"An error occurred during the second phase: {e}")

    finally:
        print("Clearing browser cache and cookies...")
        driver.delete_all_cookies()

        # Close the browser
        driver.quit()

def move_files_to_common_folder(json_file, excel_file):
    """Move the last downloaded JSON and Excel files to a common folder."""
    # Prompt the user for a destination folder
    Tk().withdraw()  # Hide the root Tkinter window
    destination_folder = filedialog.askdirectory(title="Select Destination Folder")
    if not destination_folder:
        print("No destination folder selected. Files not moved.")
        return

    # Move the JSON file
    if os.path.exists(json_file):
        json_destination = os.path.join(destination_folder, os.path.basename(json_file))
        shutil.move(json_file, json_destination)
        print(f"Moved JSON file to: {json_destination}")
    else:
        print(f"JSON file not found: {json_file}")

    # Move the Excel file
    if os.path.exists(excel_file):
        excel_destination = os.path.join(destination_folder, os.path.basename(excel_file))
        shutil.move(excel_file, excel_destination)
        print(f"Moved Excel file to: {excel_destination}")
    else:
        print(f"Excel file not found: {excel_file}")

if __name__ == "__main__":
    # Step 1: Select the JSON files
    selected_json_files = select_json_files()
    if selected_json_files:
        # Step 2: Proceed with site operations to download JSON files
        open_website_and_interact(selected_json_files)

        # Step 3: Get the last downloaded JSON file
        download_folder = os.path.join(os.getcwd(), "downloads")
        json_files = [os.path.join(download_folder, f) for f in os.listdir(download_folder) if f.endswith(".json")]
        if json_files:
            # Sort files by modification time (most recent first)
            json_files.sort(key=os.path.getmtime, reverse=True)
            last_downloaded_json = json_files[0]

            # Step 4: Upload the moved JSON file and download the Excel file
            upload_json_and_download_excel(last_downloaded_json)

            # Step 5: Get the last downloaded Excel file
            excel_download_folder = os.path.join(os.getcwd(), "excel_downloads")
            excel_files = [os.path.join(excel_download_folder, f) for f in os.listdir(excel_download_folder) if f.endswith(".xlsx")]
            if excel_files:
                # Sort files by modification time (most recent first)
                excel_files.sort(key=os.path.getmtime, reverse=True)
                last_downloaded_excel = excel_files[0]

                # Step 6: Move both files to a common folder
                move_files_to_common_folder(last_downloaded_json, last_downloaded_excel)
            else:
                print("No Excel files found to move.")
        else:
            print("No JSON files found to move.")