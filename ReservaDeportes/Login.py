from requests import Session
from bs4 import BeautifulSoup as bs
with Session() as s:
    site = s.get("https://reservadeportes.com/ADI.html")
    print(site.content)