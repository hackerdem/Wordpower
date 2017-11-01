#retrieve example sentences from collins online dictionary

from bs4 import BeautifulSoup
import socks
import requests
url='https://www.collinsdictionary.com/search/?dictCode=french-english&q=pendant'
req=requests.get(url)
data=req.text
soup=BeautifulSoup(data)
soup.findAll('blockquote')

