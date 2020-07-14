# urllib.request for opening and reading URLs
import urllib.request
from bs4 import BeautifulSoup
url = urllib.request.urlopen('https://www.alkosto.com/tv/televisores/ver/samsung/?mode=grid').read()
soup = BeautifulSoup(url)
tags = soup('a')
for tag in tags:
    print(tag.get('href'))