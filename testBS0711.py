from bs4 import BeautifulSoup
import requests

url = "https://www.example.com"
res = requests.get(url)
soup = BeautifulSoup(res.text, "html.parser")
print("ページタイトル：", soup.title.string)
