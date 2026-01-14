# 抓取台灣財經新聞的網頁原始碼(HTML)
import bs4
import urllib.request as req


def getTitle(titleURL):
    request = req.Request(titleURL, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    })
    with req.urlopen(request) as response:
        data = response.read().decode("utf-8")
    root = bs4.BeautifulSoup(data, "html.parser")
    titles = root.find_all("span", class_="post-title")
    for title in titles:
        print(title.string)


def getData(url):
    request = req.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    })
    with req.urlopen(request) as response:
        data = response.read().decode("utf-8")
    root = bs4.BeautifulSoup(data, "html.parser")

    ###### 修改這部分(想抓的資料) #########
    # 按右鍵進入檢查畫面，看HTML碼，找到要抓的部分

    titles = root.find_all("a", style="color:#777777;text-decoration: none;")
    for title in titles:
        link = title["href"]
        getTitle(link)
    NextPage = root.find("a", string="下一頁 »")
    NextLink = NextPage["href"]
    return NextLink
    ###### 修改這部分 #########


# main : 抓取多個網頁的標題
pageURL = "https://ctee.com.tw/livenews/aj"  # 複製要抓取資料網頁的網址
counts = 3  # 總共要抓幾頁的標題
for count in range(counts):
    pageURL = getData(pageURL)
