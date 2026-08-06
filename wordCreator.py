from cloakbrowser import launch
from docx import Document
import docxedit
from time import sleep
from urllib.request import urlretrieve
from datetime import date
import os

today = date.today()
days = ["Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag", "Zondag"]
print(today.strftime("%d/%m/%Y"))
todayDay = days[today.weekday()]
todayDate = today.strftime("%d/%m/%Y")

APOTHEEK_URL="https://www.apotheek.be/PharmacySearch?QueryDesktop=Oostende&QueryMobile=&OnDuty=true"
SJABLOON_PATH = os.path.expandvars("%USERPROFILE%\\Documents\\sjabloon.docx")
SJABLOON_URL = "https://github.com/MattPlayGamez/apotheekdoc/raw/refs/heads/main/sjabloon.docx"
NEW_FILE_PATH = os.path.expandvars("%USERPROFILE%\\Desktop\\apotheekblad.docx")

PLACEHOLDERS = ["[WEEKDAG]", "[DATUM]", "[NAAM]", "[STRAAT]", "[STAD]", "[TELEFOON]"]

browser = launch(headless=True, humanize=True, slow_mo=50)
page = browser.new_page()
sleep(1)

urlretrieve(SJABLOON_URL, SJABLOON_PATH) if not os.path.isfile(SJABLOON_PATH) else None

page.goto(APOTHEEK_URL)
print("Page loaded")
    

apo = page.query_selector_all("div.pharmacy-search-result-accordion-button")[0].inner_text().split('\n')
print(apo)
name = apo[0].strip()
street = apo[1].split(',')[0].strip()
city = apo[1].split(',')[1].strip()
phone_raw = page.query_selector_all("a.btn.quaternary-button.text-nowrap.tel")[0].inner_text().strip().replace(' ', '')
phone = f"{phone_raw[:3]}/{phone_raw[3:9]}"
NEW_VALUES = [todayDay, todayDate, name, street, city, phone]


document = Document(SJABLOON_PATH)

for i, placeholder in enumerate(PLACEHOLDERS):
    docxedit.replace_string(document, placeholder, NEW_VALUES[i])


os.remove(NEW_FILE_PATH) if os.path.exists(NEW_FILE_PATH) else None
document.save(NEW_FILE_PATH)

browser.close()