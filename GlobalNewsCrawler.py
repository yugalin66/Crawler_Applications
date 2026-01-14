# 抓取全球新聞的網頁原始碼(HTML)
import bs4
import urllib.request as req


def getData(url):
    # 建立一個Reqest物件，附加Request Headers的資訊
    # 按F12 -> Network，重新整理網頁 -> index.html -> Headers -> Request Headers ->複製需要的資訊
    request = req.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    })
    with req.urlopen(request) as response:
        data = response.read().decode("utf-8")

    # 解析原始碼，取得每篇文章的標題
    root = bs4.BeautifulSoup(data, "html.parser")
    # 尋找所有 class="title" 的div標籤

    titles = root.find_all("div", class_="subcate-list__link__text")
    for title in titles:
        print(title.a["title"])


# main : 抓取多個網頁的標題
pageURL = "https://www.worldjournal.com/wj/cate/business/121209"  # 複製要抓取資料網頁的網址
getData(pageURL)
