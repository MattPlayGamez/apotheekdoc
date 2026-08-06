# ApotheekDoc Auto-Generator

A Python automation script that fetches the current on-duty pharmacy (apotheek van wacht) in Oostende, Belgium, and automatically generates a formatted Word document (`.docx`) with the pharmacy's details. 

## Features

* **Automated Web Scraping:** Uses `cloakbrowser` to silently scrape `apotheek.be` for the on-duty pharmacy in Oostende.
* **Smart Template Management:** Automatically downloads the required Word document template (`sjabloon.docx`) from GitHub if it's not already on your machine.
* **Document Generation:** Uses `python-docx` and `docxedit` to find and replace specific placeholders with real-time scraped data (name, address, phone number).
* **Localized Dates:** Automatically calculates and inserts today's date and weekday in Dutch.

## Prerequisites

Before running the script, ensure you have Python 3 installed on your system along with the required dependencies.

Install the required Python packages:

```bash
pip install cloakbrowser python-docx docxedit