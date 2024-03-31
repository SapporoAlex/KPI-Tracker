# KPI Automation System

<img src="https://github.com/SapporoAlex/KPI-Tracker/blob/main/preview.jpg">

## Description
This system was designed for Andrew Statter, CEO of Titan Greentech, as a visualisation tool to view the KPIs of his staff. This project is an automated system designed for extracting, processing, and visualizing Key Performance Indicator (KPI) data for employees of GreenTech, utilizing the Tamago platform. The script navigates through web interfaces, scrapes relevant data, processes it, updates and creates visualizations in Excel spreadsheets, and finally sends out an email with the updated spreadsheet attached.

## Features
- **Data Scraping**: Utilizes Playwright to automate web navigation and scrape KPI data from the Tamago platform.
- **Data Processing**: Processes and organizes scraped data in Python.
- **Excel Automation**: Uses Openpyxl to update and create visual charts in Excel workbooks for individual and team KPI tracking.
- **Email Automation**: Automatically sends updated KPI data to specified recipients via email with the Excel file attached.

## Requirements
- Python 3.x
- Playwright
- BeautifulSoup4
- Openpyxl
- smtplib and email for email functionality

Ensure you have the required Python version and all dependencies installed. Dependencies can be installed using pip:
```bash
pip install playwright beautifulsoup4 openpyxl
```
For email functionality, ensure your email server details are correctly configured within the script.

## Usage
Clone the repository to your local machine.
Navigate to the script directory.
Run the script using Python:
```bash
python KPI Tracker.py
```
Ensure you have the necessary login credentials and URLs configured within the script for accessing the Tamago platform.

## Configuration
Before running the script, make sure to fill in the required fields marked with **********, such as Tamago platform URLs, login credentials, and email details.

- Note
This script is designed for GreenTech's specific use case and might need adjustments for other environments or platforms.

## License
MIT License

## Contributing
Contributions to the KPI Automation System are welcome. Please feel free to fork the repository, make your changes, and submit a pull request.

## Disclaimer
This project is not affiliated with Tamago. Use at your own risk and ensure compliance with Tamago's terms of service.

## Special Thanks
I would sincerely like to thank Andrew Statter, CEO of Titan Greentech, for reaching out to me on this project and collaborating with me during the production process.
