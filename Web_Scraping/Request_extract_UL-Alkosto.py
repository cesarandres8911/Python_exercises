import requests
from bs4 import BeautifulSoup
req = requests.get('https://www.alkosto.com/tv/televisores/ver/samsung/?mode=grid')
soup = BeautifulSoup(req.text, 'lxml')

tag = soup.findAll('h2', {'class': 'product-name'})[0:5]
for tags in tag:
    print (tags)