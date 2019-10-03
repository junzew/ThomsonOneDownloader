from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pywinauto.application import Application

import time
import datetime
import pandas as pd
import copy
import pyautogui as gui

# import pickle
# https://stackoverflow.com/questions/48880646/python-selenium-use-a-browser-that-is-already-open-and-logged-in-with-login-cre
# https://stackoverflow.com/questions/24925095/selenium-python-internet-explorer
# https://github.com/SeleniumHQ/selenium/wiki/InternetExplorerDriver#required-configuration
# Enable 64-bit process for Enhanced Protected Mode
# https://stackoverflow.com/questions/40626810/how-to-fix-the-slow-sendkeys-on-ie-11-with-selenium-webdriver-3-0-0

folder = "C:\\Users\\admin\\Documents\\ThomsonOne\\" # UPDATE THIS
excel_file_name = folder + "symbols.xlsx"
symbols_file_name = folder + "symbols.txt"
driver_location = folder + "IEDriverServer.exe" # needs to be placed in PATH

login_url = "https://login.ezproxy.library.ubc.ca/login?qurl=http%3a%2f%2fezproxy.library.ubc.ca%2flogin%2fthomsonone"
url = "https://www.thomsonone.com/Workspace/Main.aspx?View=Action%3dOpen&BrandName=www.thomsonone.com&IsSsoLogin=True"
driver = webdriver.Ie(driver_location)
timeout = 5 #seconds
driver.set_page_load_timeout(timeout)
success = []

contributors = [
"BARCLAYS",
"BMO CAPITAL MARKET",
"CREDIT SUISSE",
"DEUTSCHE BANK",
"EVERCORE ISI",
"HSBC GLOBAL RESEARCH",
"JEFFERIES",
"JPMORGAN",
"MORGAN STANLEY",
"RBC CAPITAL MARKETS (CANADA)",
"UBS RESEARCH",
"WELLS FARGO SECURITIES, LLC",
]
# UPDATE IF NECESSARY
# position of the 'inverse triangle' button in the download pop up
save_as_x = 1354
save_as_y = 164

def load_symbols():
    """
    Load company names from saved symbols.txt
    If no symbols.txt found, read from symbols.xlsx
    returns a list of string symbols (company names)
    """
    try: # to load saved symbols
        with open(symbols_file_name) as f:
            symbols = list(map(str.strip, f.readlines()))
        print("{} symbols read from saved file".format(len(symbols)))
    except: # that if symbols are not saved, then read from Excel
        # Parse the Excel file
        xl_file = pd.ExcelFile(excel_file_name)
        dfs = {sheet_name: xl_file.parse(sheet_name)
            for sheet_name in xl_file.sheet_names}
        sheet1 = dfs["Sheet1"]
        # set of symbols: total count:2926
        symbols = set(list(map(str.strip, sheet1["symbol"])))
        symbols = sorted(list(symbols))
        with open(symbols_file_name, "w+") as f:
            for s in symbols:
                f.write(s + '\n')
        print("{} symbols read from Excel".format(len(symbols)))
    return symbols

def close_unwanted_windows(driver, keep):
    """Close all windows except the window 'keep'
    """
    handles = driver.window_handles
    print("windows:", handles)
    for h in handles:
        if h != keep:
            driver.switch_to.window(h)
            print("closing", h)
            driver.close()
            print("closed", h)

def execute_app():
    start_time = time.time()
    companies = load_symbols()
    # Start interacting with the browser
    driver.maximize_window()
    driver.get(login_url)
    print("Please log in")
    print("Go to Company View->Research")
    print("Do a sample search")
    print("Then press Enter")
    foo = input()
    # Manually login with CWL
    print("Logged in")
    assert "Thomson One" in driver.title # sanity check
    print("Ready to start jobs")

    companies = ['AAPL-US', 'ADBE-US'] # DEBUG
    total_count = len(companies)
    for i, company in enumerate(companies):
        print("{} Processing {}".format(datetime.datetime.now(), company))
        print("Overal Elapsed Time: {:.1f} seconds".format(time.time() - start_time))
        print("Overal Progress: {}/{}={:.2f}%".format(i, total_count, i/total_count * 100))
        # for f in ["frameWinF7", "PCPWinF6", "PCPWinF7", "frameWinF6"]:
        try:
            # print("try F7")
            driver.switch_to.default_content()
            driver.switch_to.frame("frameWinF7")
            driver.switch_to.frame("PCPWinF7")
            # print("F7!")
        except:
            try:
                # print("try F6")
                driver.switch_to.default_content()
                driver.switch_to.frame("frameWinF6")
                driver.switch_to.frame("PCPWinF6")
                # print("F6!")
            except:
                print("Bad, can't switch into any frame")
                exit()
        try:
            print("trying to find trackInput")
            trackinput = driver.find_element_by_id("trackInput")
            trackinput.clear()
            trackinput.send_keys(company)
            print("Entered company name {}".format(company))
            # time.sleep(1)

            go = driver.find_element_by_id("go")
            go.send_keys(u'\ue007')
            print("Go!")
            time.sleep(5) #TODO

            t = time.time()
            n = len(contributors)
            for j, contributor in enumerate(contributors):
                try:
                    print("Processing {}_{}".format(company, contributor))
                    print(" Elapsed Time: {:.1f} seconds".format(time.time() - t))
                    print(" Progress    : {}/{}={:.1f}%".format(j, n, j/n * 100))
                    print(driver.window_handles)

                    driver.switch_to.default_content()
                    driver.switch_to.frame("frameWinC1-4054-8135")
                    # print("switched to frameWinC1-4054-8135")
                    driver.switch_to.frame("PCPWinC1-4054-8135") #id =3234
                    # print("switched to frame PCPWinC1-4054-8135")

                    if j == 0: #only need to set dates once for each company
                        # range = driver.find_element_by_id("ctl00__criteria_dateRange")
                        driver.execute_script('document.getElementById("ctl00__criteria_dateRange").selectedIndex = 7') # Custom
                        # Enter dates
                        driver.execute_script('document.getElementById("ctl00__criteria__fromDate").value = "01/01/15"')
                        print("Entered from date")
                        driver.execute_script('document.getElementById("ctl00__criteria__toDate").value = "01/01/18"')
                        print("Entererd to date")

                    ctb_input = driver.find_element_by_id("contributors.searchText")
                    ctb_input.clear() # delete
                    for k in range(4):
                        ctb_input.send_keys('\b\b\b\b\b\b\b')
                    ctb_input.send_keys(contributor)
                    print("Entered {}".format(contributor))
                    # wait for pop up to show and choose first by default
                    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, "dijitMenuItem")))
                    option = driver.find_element_by_class_name("dijitMenuItem")
                    # Strange offset problem with Dojo combobox
                    h = ActionChains(driver).move_to_element(option).move_by_offset(0,-60)
                    h.click().perform()
                    print("Added contibutor {}".format(contributor))
                    time.sleep(1)

                    # Search
                    bt_search = driver.find_element_by_id("ctl00__criteria__searchButton")
                    bt_search.send_keys(u'\ue007')
                    print("Clicked search button!")
                    time.sleep(6) # TODO wait for search result to appear
                    # old_page = driver.find_element_by_id('ctl00__results__resultContainer')
                    # WebDriverWait(driver, timeout).until(EC.staleness_of(old_page))

                    try:
                        driver.execute_script("document.getElementById('select-all-reports').click();")
                        print("Search results available")
                        print("Selected all reports!")
                    except Exception as e:
                        print(e)
                        print("No search result")
                        continue # with next contributor

                    # click View
                    view = driver.find_element_by_xpath('//*[@title="View Selected"]')
                    view.send_keys(u'\ue007')
                    print("Clicked View!")

                    # New window should open
                    # https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python
                    WebDriverWait(driver, timeout).until(EC.number_of_windows_to_be(2))

                    current = driver.current_window_handle
                    print("windows:", driver.window_handles)
                    print("current window is", current)
                    newWindow = [window for window in driver.window_handles if window != current][0]
                    driver.switch_to.window(newWindow)
                    print("OPEN NEW WINDOW")

                    try:
                        wait = WebDriverWait(driver, timeout)
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'content')))
                        # driver.switch_to.frame("content")

                        # Select All Reports
                        driver.execute_script("document.getElementById('ctl00_ctlTocDisplay_chkToc').click();")
                        print("Selected all reports 2")
                        # Click view
                        view = driver.find_element_by_name("ctl00$ctlTocDisplay$btnSubmit")
                        print("Clicked View 2")
                        view.send_keys(u'\ue007')
                        # open up the download dialog
                    except Exception as e:
                        print("Can't switch: {}".format(e))
                        driver.close()
                        driver.switch_to.window(current)
                        continue

                    try:
                        # wait for download dialog to open
                        time.sleep(6)
                        print("Download dialog should have opened")

                        # Use pyautogui to save the pdf file
                        # TODO try pywinauto
                        # https://pywinauto.readthedocs.io/en/latest/HowTo.html#how-to-specify-a-usable-application-instance
                        app = Application().connect(title_re="View Downloads")
                        dlg = app.window(title_re="View Downloads")
                        # dlg.print_control_identifiers()
                        # dlg.move_window(500,100) # fixed position
                        dlg.maximize()
                        print("ready to download")
                        # TODO Pyautogui

                        # conflict need version 0.6.6
                        # https://github.com/asweigart/pyautogui/issues/353

                        # https://sqa.stackexchange.com/questions/31436/how-to-save-a-file-by-clicking-on-a-link-in-python-internet-explorer-windows
                        # https://stackoverflow.com/questions/11449179/how-can-i-close-a-specific-window-using-selenium-webdriver-with-java
                        time.sleep(1)
                        gui.moveTo(save_as_x, save_as_y) # position of the 'inverse triangle' button
                        gui.click()
                        gui.press('a') # save as
                        gui.click()
                        # Default file name format: SYMBOL_CONTRIBUTOR
                        gui.typewrite(company + "_" + contributor)
                        time.sleep(1)
                        gui.press('enter')
                        print("started download")

                        try:
                            # this is intended to close the loading window
                            close_unwanted_windows(driver, current)
                            # the new window for reports closes automatically
                        except Exception as e:
                            print("error with close")
                            print(e)
                        print("Switching back to main search page")
                        # this hangs but is a required call somehow
                        driver.switch_to.window(current)
                        print("Gone back to main search page")
                    except Exception as e:
                        print(e)
                        print("Ignore this timeout exception")
                except Exception as e:
                    print("Exception: {}_{}".format(company, contributor))
                    print(e)


            print("PROCESSED {}".format(company))
            print("---------------------------------")
            success.append(company)

            driver.switch_to.default_content()
        except Exception as e:
            print(e)

def handle_cleanup():
    with open(folder + 'success.txt', 'w+') as f:
        f.writelines(success)
    # driver.close()


def main():
    try:
        execute_app()
    finally:
        handle_cleanup()

if __name__=='__main__':
    main()
