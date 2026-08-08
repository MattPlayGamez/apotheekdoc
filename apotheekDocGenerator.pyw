import os
import requests
from bs4 import BeautifulSoup
from docx import Document
import docxedit
from urllib.request import urlretrieve
from datetime import date

# 1. Date setup
today = date.today()
days = ["Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag", "Zondag"]
print(today.strftime("%d/%m/%Y"))
todayDay = days[today.weekday()]
todayDate = today.strftime("%d/%m/%Y")

# 2. File and URL setup
APOTHEEK_URL = "https://www.apotheek.be/PharmacySearch?QueryDesktop=Oostende&QueryMobile=&OnDuty=true"
SJABLOON_PATH = os.path.expandvars("%USERPROFILE%\\Documents\\sjabloon.docx")
SJABLOON_URL = "https://github.com/MattPlayGamez/apotheekdoc/raw/refs/heads/main/sjabloon.docx"
NEW_FILE_PATH = os.path.expandvars("%USERPROFILE%\\Desktop\\apotheekblad.docx")

PLACEHOLDERS = ["[WEEKDAG]", "[DATUM]", "[NAAM]", "[STRAAT]", "[STAD]", "[TELEFOON]"]

# 3. Download template if it doesn't exist
if not os.path.isfile(SJABLOON_PATH):
    urlretrieve(SJABLOON_URL, SJABLOON_PATH)

# 4. Fetch the page with requests
print("Fetching page...")
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}
response = requests.get(APOTHEEK_URL, headers=headers)
response.raise_for_status()

# 5. Parse the HTML with BeautifulSoup
soup = BeautifulSoup(response.text, 'html.parser')
print("Page loaded")

# Extract pharmacy info
apo_div = soup.select_one("div.pharmacy-search-result-accordion-button")

if not apo_div:
    print("Error: Could not find pharmacy data.")
    print("Note: If apotheek.be uses JavaScript to load these results dynamically, requests/bs4 won't see them. You might need to query their hidden API directly.")
    exit(1)

# Get text, separate by newlines, and filter out empty strings to mimic inner_text().split('\n')
apo_text_lines = [line.strip() for line in apo_div.get_text(separator='\n').split('\n') if line.strip()]
print(apo_text_lines)

name = apo_text_lines[0]
street = apo_text_lines[1].split(',')[0].strip()
city = apo_text_lines[1].split(',')[1].strip()

# Extract phone number
phone_node = soup.select_one("a.btn.quaternary-button.text-nowrap.tel")
if not phone_node:
    print("Error: Could not find the phone number.")
    exit(1)

phone_raw = phone_node.get_text(strip=True).replace(' ', '')
phone = f"{phone_raw[:3]}/{phone_raw[3:9]}"

NEW_VALUES = [todayDay, todayDate, name, street, city, phone]

# 6. Edit Document
document = Document(SJABLOON_PATH)

for i, placeholder in enumerate(PLACEHOLDERS):
    docxedit.replace_string(document, placeholder, NEW_VALUES[i])

# 7. Save Document
if os.path.exists(NEW_FILE_PATH):
    os.remove(NEW_FILE_PATH)
    
document.save(NEW_FILE_PATH)
print(f"Success! Document saved to {NEW_FILE_PATH}")