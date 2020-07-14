# BeatifulSoup extract the desired data when provided with an HTML or XML document
# urllib.request for opening and reading URLs
from bs4 import BeautifulSoup
import urllib.request

# Get and assign all contents from web page to a variable.
url = urllib.request.urlopen('https://www.alkosto.com/tv/televisores/ver/samsung/?mode=grid').read()

# Use BeatifulSoup library for to manipulate the data on the website.
soup = BeautifulSoup(url)
tags = soup.find_all("ul")
for tag in tags:
    print("{0}: {1}".format(tag.name, tag.text))