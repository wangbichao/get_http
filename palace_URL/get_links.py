import urllib.request
import re


url1 = 'http://hdzx.xicp.cn:56751/sjbmxt/wap.asp'
url2 = 'http://hdzx.xicp.cn:56751/sjbmxt/banjilist.asp'

# httpre = re.compile(r"(http://\S*?)[\"|>|)]", re.IGNORECASE)


def get_links(url):
    # connect to a URL
    website = urllib.request.urlopen(url)
    # read html code    
    html = website.read()
    print(html)
    # use re.findall to get all the links
    links = re.findall('"((location|http|ftp)s?://.*?)"', html)
    print(links)


if __name__ == "__main__":
    get_links(url1)     # call get_links() method
