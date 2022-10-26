import datetime
import getopt
import math
import os
import re
import signal
import sys
import time
from time import sleep
import pywhatkit
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

EXCEL_SOURCE_FOLDER = os.getcwd()
LESSON_END_MAPPINGS = {1: "8:35", 2: "9:20", 3: "10:25", 4: "11:15", 5: "12:20", 6: "13:10", 7: "14:25", 8: "15:10", 9: "16:00", 10: "16:45"}
LESSON_COUNT = 10
SLEEP_TIME = 60 * 5  # seconds
EXTRA_OVERDUE_TIME_AFTER_LESSON = 20  # minutes

REDIRECT_ALL_MESSAGES = False
HEADLESS = True
KILLER = False
SAVE_IMAGES = False
TIMEOUT = 20
DRIVER_PATH = "chromedriver.exe"
VOID_MESSAGES = False

WHATSAPP_URL = r"https://web.whatsapp.com/send?phone={number}&text={text}&app_absent=0"
COUNTRY_CODE = "+972"
numbers = {"daniel": "", "dvora": "", "multimedia": "", "olga": ""}

current_overdue_lesson = 0
current_day_num = 0
messaged = []
cwd = os.getcwd()


def get_stealthy_driver_options(headless=True, spoof_user_agent=True, stealth=True, reload_data=True):
    # based on https://stackoverflow.com/questions/53039551/selenium-webdriver-modifying-navigator-webdriver-flag-to-prevent-selenium-detec
    # and https://stackoverflow.com/a/70201673
    options = Options()
    if reload_data:
        options.add_argument(rf'user-data-dir={cwd}/python/wa_data')
    options.headless = headless
    options.add_argument("window-size=1920,1080")
    # options.add_argument("start-maximized")
    # options.add_argument("nogpu")
    # options.add_argument("disable-gpu")
    # options.add_argument("no-sandbox")
    if stealth:
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('disable-blink-features=AutomationControlled')
    if spoof_user_agent:
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36')
    return options


def send_message_to_user(number: str, message: str):
    driver.get(WHATSAPP_URL.format(**{"number": COUNTRY_CODE + number, "text": message}))
    wait_until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, ".tvf2evcx.oq44ahr5.lb5m6g5c.svlsagor.p2rjqpw5.epia9gcq"))).click()
    print(f"sent to {number}")
    sleep(2)
    # wait_until(EC.presence_of_element_located((By.CSS_SELECTOR, ".tvf2evcx.oq44ahr5.lb5m6g5c.svlsagor.p2rjqpw5.epia9gcq"))).click()
    # wait_until(EC.presence_of_element_located((By.XPATH, ".reply-button"))).click()
    # wait_until(EC.presence_of_element_located((By.CLASS_NAME, "_3xTHG"))).click()


def check_time():
    # check if current day is the same as current_day_num, if not update it, and clean the "messaged" list.
    global current_day_num, messaged, current_overdue_lesson
    c_time = datetime.datetime.now()
    if datetime.datetime.now().day != current_day_num:
        current_day_num = c_time.day
        current_overdue_lesson = 0
        messaged = []
    # also update CURRENT_LESSON, accounting for the extra 20m headroom.
    for i in range(1, 11):
        lesson_time = datetime.datetime.strptime(LESSON_END_MAPPINGS[i], "%H:%M") + datetime.timedelta(minutes=EXTRA_OVERDUE_TIME_AFTER_LESSON)
        if lesson_time.time() < c_time.time():
            current_overdue_lesson = i


def lesson_parser(data: str):
    data = str(data)
    data = data.lower()
    if 'x' in data:
        return False
    split = re.match(r"(\d*)(?:\D+(\d*))*", data)
    print(split.groups())
    if split[2] is not None:
        # split = [num for num in map(int, split)]
        # print(split[2])
        return int(split[2]) <= current_overdue_lesson
    else:
        num = int(split[1])
        while num > LESSON_COUNT:
            num = num % (10 ** int(math.log10(num)))  # remove leading number
        try:
            # print(num)
            return num <= current_overdue_lesson
        except Exception:
            pass
    # print(f"failed, returning false.")
    return False


def get_return_message(row):
    name = str(row.NameSurn).lower()
    he = {chr(x): 0 for x in range(1488, 1515)}
    en = {chr(x): 0 for x in range(97, 97 + 26)}
    for char in name:
        if char in he:
            he[char] += 1
        elif char in en:
            en[char] += 1

    if sum(x for x in he.values()) < sum(x for x in en.values()):  # if there are more english chars
        return get_return_message_eng(row)
    # else get hebrew message
    return get_return_message_heb(row)


def get_return_message_heb(row):
    str_builder = ["ברכות" + f" {row.NameSurn},",
                   "הרינו לעדכנך שלא החזרת את הפריטים הבאים במסגרת הזמן המבוקש: "]
    if pd.notna(row.KB):
        str_builder.append("מקלדת מספר" + f" {row.KB} " + "(אל תשכח את השבב!)")
    if pd.notna(row.Power):
        str_builder.append(f"{int(row.Power)} " + "כבל חשמל")
    if pd.notna(row.HeadSET):
        str_builder.append("אוזניות" + f" {int(row.HeadSET)}")
    if pd.notna(row.USB):
        str_builder.append(f"{int(row.USB)} " + "כבל" + " USB")
    if pd.notna(row.HDMI):
        str_builder.append(f"{int(row.HDMI)} " + "כבל" + " HDMI")
    if pd.notna(row.Projector):
        str_builder.append("מקרן" + f" {int(row.Projector)}")
    if any((pd.notna(row.LTZoom), pd.notna(row.LTEclipse), pd.notna(row.LTZoom))):
        laptop_builder = ["מחשב/ים נייד/ים" + ": "]
        if pd.notna(row.LTZoom):
            laptop_builder.append(f"{row.LTZoom}")
        if pd.notna(row.LTEclipse):
            laptop_builder.append(f"{row.LTEclipse}")
        if pd.notna(row.LPITest):
            laptop_builder.append(f"{row.LPITest}")
        final_laptop_str = laptop_builder[0] + ", ".join(laptop_builder[1:])
        str_builder.append(final_laptop_str)
    if pd.notna(row.OTHER):
        str_builder.append("וגם את הפריטים הבאים" + f": {row.OTHER}")
    str_builder.append("")
    str_builder.append("נא להחזיר אותם בהקדם, ")
    str_builder.append('בברכה, מחלקת מולטימדיה של תיכון "ליד"ה".')
    list_1 = [*str_builder[:2], ", %0a".join(str_builder[2:-3]), "%0a".join(str_builder[-3:])]
    return "%0a".join(list_1)


def get_return_message_eng(row):
    str_builder = [f"Greetings {row.NameSurn}, ",
                   "We couldn't help but notice you have failed to return the following items within your requested timeframe: "]
    if pd.notna(row.KB):
        str_builder.append(f"Keyboard number {row.KB} (Don't forget the chip!)")
    if pd.notna(row.Power):
        str_builder.append(f"{int(row.Power)} power cable/s")
    if pd.notna(row.HeadSET):
        str_builder.append(f"{int(row.HeadSET)} headset/s")
    if pd.notna(row.USB):
        str_builder.append(f"{int(row.USB)} USB cable/s")
    if pd.notna(row.HDMI):
        str_builder.append(f"{int(row.HDMI)} HDMI cable/s")
    if pd.notna(row.Projector):
        str_builder.append(f"{int(row.Projector)} Projector/s")
    if any((pd.notna(row.LTZoom), pd.notna(row.LTEclipse), pd.notna(row.LTZoom))):
        laptop_builder = ["Laptop/s: "]
        if pd.notna(row.LTZoom):
            laptop_builder.append(f"{row.LTZoom}")
        if pd.notna(row.LTEclipse):
            laptop_builder.append(f"{row.LTEclipse}")
        if pd.notna(row.LPITest):
            laptop_builder.append(f"{row.LPITest}")
        final_laptop_str = laptop_builder[0] + ", ".join(laptop_builder[1:])
        str_builder.append(final_laptop_str)
    if pd.notna(row.OTHER):
        str_builder.append(f"As well as the following various items: {row.OTHER}")
    str_builder.append("")
    str_builder.append("Please do return them posthaste, ")
    str_builder.append("Regards, the Multimedia department of the Leyada High School.")
    list_1 = [*str_builder[:2], ", %0a".join(str_builder[2:-3]), "%0a".join(str_builder[-3:])]
    return "%0a".join(list_1)
    # return "Really cool stuff, but now with a newline for realsies: %0a" + str(row)


def get_this_day_from_excel():
    excel_sheet = None
    for file in os.listdir(EXCEL_SOURCE_FOLDER):
        file = EXCEL_SOURCE_FOLDER + os.sep + file
        if os.path.isfile(file) and os.path.splitext(file)[1] == '.xlsm':
            if excel_sheet is None or os.path.getmtime(excel_sheet) < os.path.getmtime(file):
                excel_sheet = file
    return pd.read_excel(excel_sheet, "ThisDay")


def get_missed_returns():
    check_time()
    messages = []
    while True:
        try:
            df = get_this_day_from_excel()  # reading from file can throw errno13
            break
        except Exception as ex:
            print(f"Exception occurred: {ex}")
            sleep(1)
            continue
    late = df[df["Lessons"].apply(lesson_parser)].reset_index(drop=True)
    print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}    we have these late returns:")
    for row in late.itertuples():
        print(row)
        if row.TelNumber not in messaged:
            if VOID_MESSAGES:
                print(f"theoretically messaged {row.TelNumber}")
                messaged.append(row.TelNumber)
                continue
            print(f"messaged {row.TelNumber}")
            if REDIRECT_ALL_MESSAGES:
                messages.append((numbers["daniel"], get_return_message(row)))
            else:
                messages.append((row.TelNumber, get_return_message(row)))
            messaged.append(row.TelNumber)
    print(f"Note that as of now, {messaged} have been messaged.\n")
    return messages


def main():
    # print(driver.execute_script("return navigator.userAgent;"))
    # nuke the webdriver property for stealth
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    while True:
        returns = get_missed_returns()
        for tel, message in returns:
            send_message_to_user(str(tel), message)
            if SAVE_IMAGES:
                driver.save_screenshot(f'python/current_page_state_{tel}.png')
        print("", end="", flush=True)
        sleep(SLEEP_TIME)


def catch_sigint(param1, param2):
    # disgusting
    print("caught!")
    driver.quit()
    if KILLER:
        os.system("taskkill /f /im chrome.exe")
    sys.exit(0)


def parse_args():
    global REDIRECT_ALL_MESSAGES, EXCEL_SOURCE_FOLDER, HEADLESS, KILLER, DRIVER_PATH, SAVE_IMAGES, SLEEP_TIME, EXTRA_OVERDUE_TIME_AFTER_LESSON, TIMEOUT, VOID_MESSAGES
    opts, left_over = getopt.getopt(sys.argv[1:], "p:c:hkis:e:t:v", ["bother_daniel", "help", "driver="])
    for i, arg in enumerate(opts):
        if arg[0] == "--bother_daniel":
            REDIRECT_ALL_MESSAGES = True
        elif arg[0] == "--driver":
            DRIVER_PATH = arg[1]
        elif arg[0] == "-p":
            if os.path.exists(arg[1]) and not os.path.isfile(arg[1]):
                EXCEL_SOURCE_FOLDER = arg[1]
            else:
                exit("Path error")
        elif arg[0] == "-h":
            HEADLESS = False
        elif arg[0] == "-i":
            SAVE_IMAGES = True
        elif arg[0] == "-s":
            SLEEP_TIME = int(arg[1])
        elif arg[0] == "-t":
            TIMEOUT = int(arg[1])
        elif arg[0] == "-e":
            EXTRA_OVERDUE_TIME_AFTER_LESSON = int(arg[1])
        elif arg[0] == "-k":
            KILLER = True
        elif arg[0] == "-v":
            VOID_MESSAGES = True
        elif arg[0] == "--help":
            print("Welcome to 'The Thing'\n"
                  "     -p [path] defines the path to the folder containing the excel spreadsheets, defaults to the cwd\n"
                  "     -h disables headless mode\n"
                  "     -k enables the killing of all chrome instances once it terminates\n"
                  "     -i enables the saving of debug screenshots\n"
                  "     -s [seconds] sleep time between attempts, defaults to 300 (5m) \n"
                  "     -e [minutes] the time to wait after a lesson before messaging someone about their return, defaults to 20 \n"
                  "     -t [time] timeout delay while waiting for pages to load"
                  "\n"
                  "     --bother_daniel enables message redirection to a single phone number\n"
                  "     --driver sets the chromedriver path")
            sys.exit(0)


if __name__ == '__main__':  # compile w/: pyinstaller --onefile .\main.py -n whatsapp_messager
    print("Starting script.")
    # driver.get("https://recaptcha-demo.appspot.com/recaptcha-v3-request-scores.php") # recaptcha test page
    parse_args()
    signal.signal(signal.SIGINT, catch_sigint)

    opts = get_stealthy_driver_options(headless=HEADLESS)
    driver = webdriver.Chrome(executable_path=DRIVER_PATH, options=opts)
    wait_until = WebDriverWait(driver, TIMEOUT).until

    try:
        main()
    except Exception as e:
        print(e)
        driver.save_screenshot("python/error_report.png")
        driver.quit()
        if KILLER:
            os.system("taskkill /f /im chrome.exe")

# TODO add automatic detection of login page.
