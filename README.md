Kaycha COA Automation Scripts

This repository contains a set of Python scripts designed to automate the process of retrieving, processing, and storing Certificate of Analysis (COA) data from the Kaycha website. The scripts use Selenium for web automation and OpenPyXL for Excel file manipulation.
Scripts Overview. This is my first go at this so be patient with me
... please.

    main.py: Automates the login process on the Kaycha website, selects a specific date range, and exports COA data as a CSV file.
    work_with_export.py: Processes the exported CSV file by replacing specific text values with zeros and converts it into an XLSX file.
    add_to_sp.py: Originally intended for adding data to an online SharePoint file, this script now appends the processed data to a local Excel file.

Installation

    Clone the repository to your local machine.
    Ensure Python is installed on your system.
    Install required Python packages using the requirements.txt file:

    bash

    pip install -r requirements.txt

    You'll need a WebDriver compatible with the version of your browser for Selenium to work (e.g., geckodriver for Firefox).

Configuration

    Create a confidential_info.py file in the config directory with your credentials, file paths, script paths, and website URLs. This file should not be uploaded to version control. Example structure:

    python

    class Passwords:
        def __init__(self):
            self.password = "your_password"

    class Usernames:
        def __init__(self):
            self.username = "your_username"

    class FilePaths:
        def __init__(self):
            self.operating_file_path = "/path/to/operating/files"
            self.excel_filename = "your_excel_file.xlsm"
            self.add_to_local = "python3 /path/to/add_to_local.py"
            self.work_with_export = "python3 /path/to/work_with_export.py"

    class WebsiteURLs:
        def __init__(self):
            self.kaycha_website = "your_state_login_page"

    # Remember to replace the paths, credentials, and URLs with your own

    Modify the file paths, credentials, and URLs in confidential_info.py according to your requirements.

Usage

    Running main.py:
        Initiates the web automation process.
        Ensure the correct URL and element IDs are set for your use case.
    Running work_with_export.py:
        Automatically finds and processes the downloaded COA data file.
    Running add_to_sp.py:
        Appends the processed data to a specified Excel file on your local machine.

Contributing

Feel free to fork this repository and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.
License

This project is licensed under the MIT License - see the LICENSE file for details.