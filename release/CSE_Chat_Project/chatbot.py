import feedparser
import schedule
import json
import selenium
import pyautogui
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from io import BytesIO
import win32clipboard
from PIL import Image
import logging
import logging.handlers
import os
import chromedriver_autoinstaller as AutoChrome
import shutil

LOG_MAX_SIZE = 1024 *1024 * 10

LOG_FILE_CNT = 10


logger = logging.getLogger(__name__)

logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('[%(asctime)s - %(name)s - %(levelname)s] :line %(lineno)d - %(message)s')
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)
file_handler = logging.handlers.RotatingFileHandler('log/CSE_Chatbot.log',maxBytes=LOG_MAX_SIZE, backupCount=LOG_FILE_CNT,encoding='utf-8')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

#chatnames = ['src/PNU_ChatBot.PNG']
chatnames = ['src/test.PNG']


def driverUpdate():
    try:
        chrome_list = []
        chrome_ver = AutoChrome.get_chrome_version().split('.')[0]
        current_list = os.listdir(os.getcwd())
        for i in current_list:
            path = os.path.join(os.getcwd(), i)
            if os.path.isdir(path):
                if 'chromedriver.exe' in os.listdir(path):
                    chrome_list.append(i)
        old_version = list(set(chrome_list)-set([chrome_ver]))
        for i in old_version:
            path = os.path.join(os.getcwd(),i)
            shutil.rmtree(path)
        if not chrome_ver in current_list:
            logger.info(f'DriverUpdate-Before{old_version} => Latest[{chrome_ver}]')
            AutoChrome.install(True)
        else :
            logger.info(f'DriverUpdate-Release[{chrome_ver}]')
            pass
    except Exception as e:
        logger.error("UpdateError")
        logger.error(e)
        pass

def loadDriver():
    global driver
    chrome_ver = AutoChrome.get_chrome_version().split('.')[0]
    path = os.path.join(os.getcwd(),chrome_ver)
    path = os.path.join(path,'chromedriver.exe')

    options = Options()
    options.add_argument("lang=ko_KR")
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument('--window-size=1280,720')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36')
    driver = webdriver.Chrome(str(path), options=options)
    driver.set_page_load_timeout(15) # timeout 10 sec
    logger.info(f'LoadDriver-Release[{chrome_ver}]')
    return driver
    
def sendContent(path):
    image = Image.open(path)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)
    chat_btn = pyautogui.locateCenterOnScreen('src/chatblock.PNG', confidence=0.8)
    pyautogui.moveTo(chat_btn.x, chat_btn.y)
    pyautogui.doubleClick()
    pyautogui.hotkey("ctrl","v")
    time.sleep(1)
    pyautogui.hotkey("enter")
    logger.info("sendContent")
    

def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def macroSend(chatname_path, contents):
    SendCheck = False
    for i in range(3):
        logger.info(f'==try count:[{i+1}]')
        try:
            room_btn = pyautogui.locateCenterOnScreen(chatname_path, confidence=0.9)
            pyautogui.moveTo(room_btn.x, room_btn.y)
            pyautogui.rightClick()
            time.sleep(1)
            roomopen_btn = pyautogui.locateCenterOnScreen('src/chatopen.PNG', confidence=0.9)
            pyautogui.moveTo(roomopen_btn.x, roomopen_btn.y)
            pyautogui.doubleClick()
            time.sleep(1)
            chat_btn = pyautogui.locateCenterOnScreen('src/chatblock.PNG', confidence=0.9)
            pyautogui.moveTo(chat_btn.x, chat_btn.y)
            pyautogui.doubleClick()
            pyperclip.copy(contents)
            pyautogui.hotkey("ctrl","v")
            time.sleep(1)
            send_btn = pyautogui.locateCenterOnScreen('src/send.PNG', confidence=0.9)
            pyautogui.moveTo(send_btn.x, send_btn.y)
            pyautogui.doubleClick()
            logger.info("macroSendMessage")
            logger.debug(contents)
            SendCheck = True
        except Exception as ex:
            logger.info(f'Message==try count:[{i+1}] => exception:\n{ex}')
        if SendCheck == True:
            break
        time.sleep(5)
    if SendCheck == False:
        logger.error("failSendMessage")
        raise AirflowException('retry gogo!')

def closeChat():
    close_btn = pyautogui.locateCenterOnScreen('src/close.PNG')
    pyautogui.moveTo(close_btn.x, close_btn.y)
    pyautogui.doubleClick()
    logger.info("closeChat")
    
def screenshot(URL,Type):
    global driver
    dirver = loadDriver()
    crawling_success = False
    for i in range(5):
        logger.info(f'ScreenShot==try count:[{i+1}]')
        try:
            driver.get(url=URL)
            crawling_success = True
        except Exception as ex:
            logger.info(f'==try count:[{i+1}] => exception:\n{ex}')
        if crawling_success == True:
            break
        time.sleep(5)
    if crawling_success == False:
        logger.error("failSendScreenshot")
        raise AirflowException('retry gogo!') # Airflow Exception
    driver.save_screenshot(f'{Type}.png')
    del driver
    
def Load(Case):
    if Case == "Notice" :
        with open("data/Notice_recent.json", 'r') as f:
            rss = json.load(f)
            
    elif Case == "Free" :
        with open("data/Free_recent.json", 'r') as f:
            rss = json.load(f)
            
    elif Case == "Employment" :
        with open("data/Employment_recent.json", 'r') as f:
            rss = json.load(f)
            
    elif Case == "Contest":
        with open("data/Contest_recent.json", 'r') as f:
            rss = json.load(f)
            
    else:
        with open("data/Others_recent.json", 'r') as f:
            rss = json.load(f)
    
    f.close()
    logger.info(f"Load{Case}")
    return rss

def Save(Case):
    if Case == "Notice" :
        Notice_rss = feedparser.parse("https://cse.pusan.ac.kr/bbs/cse/2605/rssList.do?row=3")
        if Notice_rss.bozo:
            raise Exception(f"{Case} parse error") 
        else:
            with open("data/Notice_recent.json", 'w') as f:
                f.write(json.dumps(Notice_rss, ensure_ascii=False))
    
    elif Case == "Free" :
        Free_rss = feedparser.parse("https://cse.pusan.ac.kr/bbs/cse/2618/rssList.do?row=3")
        if Free_rss.bozo:
            raise Exception(f"{Case} parse error") 
        else:
            with open("data/Free_recent.json", 'w') as f:
                f.write(json.dumps(Free_rss, ensure_ascii=False))
    
    elif Case == "Employment" :
        Employment_rss = feedparser.parse("https://cse.pusan.ac.kr/bbs/cse/2616/rssList.do?row=3")
        if Employment_rss.bozo:
            raise Exception(f"{Case} parse error") 
        else:
            with open("data/Employment_recent.json", 'w') as f:
                f.write(json.dumps(Employment_rss, ensure_ascii=False))
            
    
    elif Case == "Contest":
        Contest_rss = feedparser.parse("https://cse.pusan.ac.kr/bbs/cse/12278/rssList.do?row=3")
        if Contest_rss.bozo:
            raise Exception(f"{Case} parse error") 
        else:
            with open("data/Contest_recent.json", 'w') as f:
                f.write(json.dumps(Contest_rss, ensure_ascii=False))
    
    else:
        Others_rss = feedparser.parse("https://cse.pusan.ac.kr/bbs/cse/2617/rssList.do?row=3")
        if Others_rss.bozo:
            raise Exception(f"{Case} parse error") 
        else:
            with open("data/Others_recent.json", 'w') as f:
                f.write(json.dumps(Others_rss, ensure_ascii=False))
    
    f.close()
    logger.info(f"Save{Case}")

try:
    Save("Notice")
    Save("Free")
    Save("Employment")
    Save("Contest")
    Save("Others")
    
    Notice_rss = Load("Notice")
    Free_rss = Load("Free")
    Employment_rss = Load("Employment")
    Contest_rss = Load("Contest")
    Others_rss = Load("Others")
    logger.info("chatbot initialized")
except:
    Notice_rss = Load("Notice")
    Free_rss = Load("Free")
    Employment_rss = Load("Employment")
    Contest_rss = Load("Contest")
    Others_rss = Load("Others")
    logger.info("chatbot initialized")
    logger.error("internetConnectError")
    
Notice_present_state = Notice_rss["entries"][0]
Free_present_state = Free_rss["entries"][0]
Employment_present_state = Employment_rss["entries"][0]
Contest_present_state = Contest_rss["entries"][0]
Others_present_state = Others_rss["entries"][0]

def Notice():
    logger.info("searching 공지사항")
    global Notice_present_state
    try:
        Save("Notice")
    except:
        logger.error("SaveError 공지사항")
        pass
    try:
        Notice_new_rss = Load("Notice")
        for i in range(len(Notice_new_rss["entries"])):
            Notice_new_state = Notice_new_rss["entries"][i]
            if Notice_new_state == Notice_present_state:
                break
            else:
                if Notice_present_state != Notice_new_state:
                    logger.info("search new 공지사항")
                    date = Notice_new_state['published']
                    category = '[공지사항]{}\n'.format(date.split(" ")[0])
                    content = "{}\n{}\n".format(Notice_new_state['title'],Notice_new_state['link'])
                    Allcontent = category + content
                    for chatname in chatnames:
                        logger.info(chatname)
                        try:
                            macroSend(chatname, Allcontent)
                        except Exception as e:
                            logger.error("SendError")
                            logger.error(e)
                            pass
                        try:
                            screenshot(Notice_new_state["link"],'Notice')
                            filepath = 'Notice.png'
                            sendContent(filepath)
                        except Exception as e:
                            logger.error("ScreenshotError")
                            logger.error(e)
                            pass
                        try:
                            closeChat()
                        except Exception as e:
                            logger.error("ClosechatError")
                            logger.error(e)
                            pass 

        Notice_present_state = Notice_new_rss["entries"][0]
    except Exception as e:
        logger.error("LoadError 공지사항")
        logger.error(e)
        pass
    
def Free():
    logger.info("searching 자유게시판")
    global Free_present_state
    try:
        Save("Free")
    except:
        logger.error("SaveError 자유게시판")
        pass
    try:
        Free_new_rss = Load("Free")
        for i in range(len(Free_new_rss["entries"])):
            Free_new_state = Free_new_rss["entries"][i]
            if Free_new_state == Free_present_state:
                break
            else:
                if Free_present_state != Free_new_state:
                    logger.info("search new 자유게시판")
                    date = Free_new_state['published']
                    category = '[자유게시판]{}\n'.format(date.split(" ")[0])
                    content = "{}\n{}\n".format(Free_new_state['title'],Free_new_state['link'])
                    Allcontent = category + content
                    for chatname in chatnames:
                        logger.info(chatname)
                        try:
                            macroSend(chatname, Allcontent)
                        except Exception as e:
                            logger.error("SendError")
                            logger.error(e)
                            pass
                        try:
                            screenshot(Free_new_state["link"],'Free')
                            filepath = 'Free.png'
                            sendContent(filepath)
                        except Exception as e:
                            logger.error("ScreenshotError")
                            logger.error(e)
                            pass
                        try:
                            closeChat()
                        except Exception as e:
                            logger.error("ClosechatError")
                            logger.error(e)
                            pass 
        Free_present_state = Free_new_rss["entries"][0]
    except Exception as e:
        logger.error("LoadError 자유게시판")
        logger.error(e)
        pass
        
def Employoment():
    logger.info("searching 채용 게시판")
    global Employment_present_state
    try:
        Save("Employment")
    except:
        logger.error("SaveError 채용게시판")
        pass
    try:
        Employment_new_rss = Load("Employment")
        for i in range(len(Employment_new_rss["entries"])):
            Employment_new_state = Employment_new_rss["entries"][i]
            if Employment_new_state == Employment_present_state:
                break
            else:
                if Employment_present_state != Employment_new_state:
                    logger.info("search new 채용게시판")
                    date = Employment_new_state['published']
                    category = '[채용게시판]{}\n'.format(date.split(" ")[0])
                    content = "{}\n{}\n".format(Employment_new_state['title'], Employment_new_state['link'])
                    Allcontent = category + content
                    for chatname in chatnames:
                        logger.info(chatname)
                        try:
                            macroSend(chatname, Allcontent)
                        except Exception as e:
                            logger.error("SendError")
                            logger.error(e)
                            pass
                        try:
                            screenshot(Employment_new_state["link"],'Employment')
                            filepath = 'Employment.png'
                            sendContent(filepath)
                        except Exception as e:
                            logger.error("ScreenshotError")
                            logger.error(e)
                            pass
                        try:
                            closeChat()
                        except Exception as e:
                            logger.error("ClosechatError")
                            logger.error(e)
                            pass
        Employment_present_state = Employment_new_rss["entries"][0]
    except Exception as e:
        logger.error("LoadError 채용게시판")
        logger.error(e)
        pass
        
def Contest():
    logger.info("searching 경진대회게시판")
    global Contest_present_state
    try:
        Save("Contest")
    except:
        logger.error("SaveError 경진대회게시판")
        pass
    try:
        Contest_new_rss = Load("Contest")
        for i in range(len(Contest_new_rss["entries"])):
            Contest_new_state = Contest_new_rss["entries"][i]
            if Contest_new_state == Contest_present_state:
                break
            else:
                Contest_new_state = Contest_new_rss["entries"][0]
                if Contest_present_state != Contest_new_state:
                    logger.info("search new 경진대회게시판")
                    date = Contest_new_state['published']
                    category = '[경진대회게시판]{}\n'.format(date.split(" ")[0])
                    content = "{}\n{}\n".format(Contest_new_state['title'],Contest_new_state['link'])
                    Allcontent = category + content
                    for chatname in chatnames:
                        logger.info(chatname)
                        try:
                            macroSend(chatname, Allcontent)
                        except Exception as e:
                            logger.error("SendError")
                            logger.error(e)
                            pass
                        try:
                            screenshot(Contest_new_state["link"],'Contest')
                            filepath = 'Contest.png'
                            sendContent(filepath)
                        except Exception as e:
                            logger.error("ScreenshotError")
                            logger.error(e)
                            pass
                        try:
                            closeChat()
                        except Exception as e:
                            logger.error("ClosechatError")
                            logger.error(e)
                            pass
        Contest_present_state = Contest_new_rss["entries"][0]
    except Exception as e:
        logger.error("LoadError 경진대회게시판")
        logger.error(e)
        pass

def Others():
    logger.info("searching 기타게시판")
    global Others_present_state
    try:
        Save("Others")
    except:
        logger.error("SaveError 기타게시판")
        pass
    try:
        Others_new_rss = Load("Others")
        for i in range(len(Others_new_rss["entries"])):
            Others_new_state = Others_new_rss["entries"][i]
            if Others_new_state == Others_present_state:
                break
            else:
                Others_new_state = Others_new_rss["entries"][0]
                if Others_present_state != Others_new_state:
                    logger.info("search new 기타게시판")
                    date = Others_new_state['published']
                    category = '[기타게시판]{}\n'.format(date.split(" ")[0])
                    content = "{}\n{}\n".format(Others_new_state['title'],Others_new_state['link'])
                    Allcontent = category + content
                    for chatname in chatnames:
                        logger.info(chatname)
                        try:
                            macroSend(chatname, Allcontent)
                        except Exception as e:
                            logger.error("SendError")
                            logger.error(e)
                            pass
                        try:
                            screenshot(Others_new_state["link"],'Others')
                            filepath = 'Others.png'
                            sendContent(filepath)
                        except Exception as e:
                            logger.error("ScreenshotError")
                            logger.error(e)
                            pass
                        try:
                            closeChat()
                        except Exception as e:
                            logger.error("ClosechatError")
                            logger.error(e)
                            pass
        Others_present_state = Others_new_rss["entries"][0]
    except Exception as e:
        logger.error("LoadError 기타게시판")
        logger.error(e)
        pass
        
        
        
schedule.every().day.at("00:00").do(driverUpdate)
schedule.every(300).seconds.do(Contest)
schedule.every(300).seconds.do(Employoment)
schedule.every(300).seconds.do(Free)
schedule.every(300).seconds.do(Notice)
schedule.every(300).seconds.do(Others)

if __name__ == "__main__":
    while True:
        schedule.run_pending()
