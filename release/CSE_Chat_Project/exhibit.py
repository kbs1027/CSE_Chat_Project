import urllib.request
import urllib.parse
from bs4 import BeautifulSoup

web_url = "https://www.contestkorea.com/sub/list.php?int_gbn=1&Txt_bcode=030510001"
row_url = "https://www.contestkorea.com/sub/"
before = ""
def create_before(i):
    global before
    with urllib.request.urlopen(web_url) as response:
        html = response.read() 
        soup = BeautifulSoup(html, 'html.parser')
        before = soup.find_all("div", attrs={'class': 'title'})[i]

def create_feed():
    global before
    with urllib.request.urlopen(web_url) as response:
        html = response.read()
        soup = BeautifulSoup(html, 'html.parser')
        return soup

def crawl_exhibit(soup, num):
    all_divs = soup.find_all("div", attrs={'class': 'title'})
    all_exhibit_names = soup.find_all("span", attrs={'class': 'txt'})
    all_exhibit_host = soup.find_all('li', attrs={'class': 'icon_1'})
    all_exhibit_target = soup.find_all('li', attrs={'class': 'icon_2'})
    all_exhibit_date = soup.find_all('span', attrs={'class': 'step-1'})
    all_exhibit_date2 = soup.find_all('span', attrs={'class': 'step-2'})
    all_exhibit_date3 = soup.find_all('span', attrs={'class': 'step-3'})
    new_url = row_url + all_divs[num].select_one('a')['href']
    host = "".join(all_exhibit_host[num].text.split(".")[0]) +": "+ "".join(all_exhibit_host[num].text.split(".")[1])
    target = "".join(all_exhibit_target[num].text.split(".")[0]) +": "+ "".join(all_exhibit_target[num].text.split(".")[1].split())
    date = all_exhibit_date[num].text.split()
    date2 = all_exhibit_date2[num].text.split()
    date3 = all_exhibit_date3[num].text.split()
    

    return (f"공모전 이름 : {all_exhibit_names[num].text}\nurl : {new_url}\n{date[0]} : {date[1]}\n{date2[0]} : {date2[1]}\n{date3[0]} : {date3[1]}\n{host}\n{target}")

def exhibit_compare(soup):
    global before
    all_divs = soup.find_all("div", attrs={'class': 'title'})
    for i in range(0, len(all_divs)):
        if before == all_divs[i]:
            before = all_divs[0]
            return i
    before = all_divs[0]
