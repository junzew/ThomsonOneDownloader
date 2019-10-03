from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

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
driver_location = folder + "IEDriverServer.exe" # needs to be in PATH
login_url = "https://login.ezproxy.library.ubc.ca/login?qurl=http%3a%2f%2fezproxy.library.ubc.ca%2flogin%2fthomsonone"
url = "https://www.thomsonone.com/Workspace/Main.aspx?View=Action%3dOpen&BrandName=www.thomsonone.com&IsSsoLogin=True"
driver = webdriver.Ie(driver_location)
timeout = 5 #seconds
driver.set_page_load_timeout(timeout)
success = []

# UPDATE IF NECESSARY
# position of the 'inverse triangle' button in the download pop up
save_as_x = 1212
save_as_y = 221


def load_symbols():
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
    """Close all windows except keep
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
    symbols = load_symbols()
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
    # print(driver.page_source)
    # with open("page_course.txt", 'w+') as f:
    #     f.write(driver.page_source)

    # Refresh
    # driver.get(url)

    assert "Thomson One" in driver.title

    # bt_company_view = driver.find_element_by_id("td13")
    # print("Company View found, click")
    # bt_company_view.click()

    # foo = input("Please go to Company View->Research, then press Enter")
    # Manually navigate
    # Necessary because may see different thomson one page every time you login
    print("Ready to start jobs")
    # driver.implicitly_wait(10)

    # bt_research = driver.find_element_by_id("tdPages9")
    # print("Research found, click")
    # bt_research.click()
    # driver.implicitly_wait(10)
    # print("wakes up 2")
    # driver.implicitly_wait(10)
    # print("here")
    # time.sleep(10)

    # link_advanced_research = driver.find_element_by_id("ctl00_dd1")
    # print("Advanced Research found, click")
    # link_advanced_research.click()
    # driver.implicitly_wait(10)
    # print("wakes up 2")
    # # driver.implicitly_wait(10)
    # print("here")
    # time.sleep(10)

    companies = ["AAPL-US", "ADBE-US","RDSA-US",]# TODO only exact match
    # companies = symbols
    total_count = len(companies)
    for i, company in enumerate(companies):
        print("{} Processing {}".format(datetime.datetime.now(), company))
        print("Elapsed time: {:.1f} seconds".format(time.time() - start_time))
        print("Progress: {}/{}={:.2f}%".format(i, total_count, i/total_count * 100))
        # for f in ["frameWinF7", "PCPWinF6", "PCPWinF7", "frameWinF6"]:
        try:
            print("try F7")
            driver.switch_to.default_content()
            driver.switch_to.frame("frameWinF7")
            driver.switch_to.frame("PCPWinF7")
            print("F7!")
        except:
            try:
                print("try F6")
                driver.switch_to.default_content()
                driver.switch_to.frame("frameWinF6")
                driver.switch_to.frame("PCPWinF6")
                print("F6!")
            except:
                print("Bad, can't switch into any frame")
                exit()
        try:
            print("trying to find trackInput")
            # print("try switching to frame {}".format(f))
            # driver.switch_to.frame(f)
            # print("switched to frame {}".format(f))
            # time.sleep(1)
            trackinput = driver.find_element_by_id("trackInput")
            # print("find the textbox in frame {}".format(f))
            trackinput.clear()
            print("cleared")
            trackinput.send_keys(company)
            print("entered company name {}".format(company))
            # time.sleep(1)

            go = driver.find_element_by_id("go")
            print("found go")
            go.send_keys(u'\ue007')
            print("clicked go")
            time.sleep(3) #TODO

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
            "WELLS FARGO SECURITIES, LLC"]

            for contributor in contributors:
                try:
                    print("{}_{}".format(company, contributor))
                    print(driver.window_handles)
                    # close_unwanted_windows(driver, driver.current_window_handle)
                    # TODO: close unwanted windows, leave only search window

                    # print(driver.page_source)
                    driver.switch_to.default_content()
                    driver.switch_to.frame("frameWinC1-4054-8135")
                    print("switched to frameWinC1-4054-8135")
                    # siwtch to child frame
                    driver.switch_to.frame("PCPWinC1-4054-8135") #id =3234
                    print("switched to frame PCPWinC1-4054-8135")
                    # time.sleep(1)

                    # range = driver.find_element_by_id("ctl00__criteria_dateRange")
                    driver.execute_script('document.getElementById("ctl00__criteria_dateRange").selectedIndex = 7')

                    driver.execute_script('document.getElementById("ctl00__criteria__fromDate").value = "01/01/15"')
                    print("entered from date")
                    driver.execute_script('document.getElementById("ctl00__criteria__toDate").value = "01/01/18"')
                    # from_date = driver.find_element_by_id("ctl00__criteria__fromDate")
                    # # from_date.send_keys("\b\b\b\b\b\b\b\b")
                    # from_date.clear()
                    # from_date.send_keys("01/01/15")
                    # print("entered from date")
                    #
                    # to_date = driver.find_element_by_id("ctl00__criteria__toDate")
                    # # to_date.send_keys("\b\b\b\b\b\b\b\b")
                    # from_date.clear()
                    # to_date.send_keys("01/01/18")
                    print("entererd to date")

                    ctb_input = driver.find_element_by_id("contributors.searchText")
                    ctb_input.clear() # delete
                    for i in range(4):
                        ctb_input.send_keys('\b\b\b\b\b\b\b')
                    ctb_input.send_keys(contributor)
                    # time.sleep(3)
                    print("entered {}".format(contributor))
                    # choose first by default
                    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, "dijitMenuItem")))
                    # option = driver.find_element_by_xpath("//div/table/tbody/tr/td/b")
                    option = driver.find_element_by_class_name("dijitMenuItem")
                    # Strange offset problem with Dojo combobox
                    h = ActionChains(driver).move_to_element(option).move_by_offset(0,-60)
                    h.click().perform()
                    print("added contibutor {}".format(contributor))

                    # Search
                    bt_search = driver.find_element_by_id("ctl00__criteria__searchButton")
                    bt_search.send_keys(u'\ue007')
                    print("clicked search button!")
                    # web_driver.execute_script("dijit.byId('<id of select field>').set('value','<value you want it set to>')")

                    # Select all
                    # sa = driver.find_element_by_id("select-all-reports")
                    # sa.send_keys(u'\ue007') # doesn't work
                    time.sleep(2) # wait for search result to appear
                    print("search results available")
                    try:
                        driver.execute_script("document.getElementById('select-all-reports').click();")
                    except Exception as e:
                        print(e)
                        print("probably no search result")
                        continue # with next contributor
                    print("selected all")
                    # window.showModalDialog = window.open
                    # click View
                    view = driver.find_element_by_xpath('//*[@title="View Selected"]')
                    view.send_keys(u'\ue007')
                    print("Clicked View!")

                    # https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python
                    # time.sleep(5) # wait for window to open
                    # windows = driver.window_handles
                    # time.sleep(3)
                    # print(driver.window_handles)
                    # time.sleep(3)
                    # print(driver.window_handles)
                    WebDriverWait(driver, timeout).until(EC.number_of_windows_to_be(2))

                    # default_handle = driver.current_window_handle
                    # handles = list(driver.window_handles)
                    # assert len(handles) > 1

                    # handles.remove(default_handle)
                    # assert len(handles) > 0

                    # driver.switch_to.window(handles[0])

                    # current = driver.window_handles[0]
                    current = driver.current_window_handle
                    print("windows:", driver.window_handles)
                    print("current window is", current)
                    newWindow = [window for window in driver.window_handles if window != current][0]
                    driver.switch_to.window(newWindow)
                    # driver.switch_to.window(windows[1])

                    print("open NEW WINDOW")

                    try:
                        wait = WebDriverWait(driver, timeout)
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'content')))
                        # driver.switch_to.frame("content")

                        # Select All Reports
                        driver.execute_script("document.getElementById('ctl00_ctlTocDisplay_chkToc').click();")
                        print("SELECT ALL REPORTS 2")
                        # Click view
                        view = driver.find_element_by_name("ctl00$ctlTocDisplay$btnSubmit")
                        print("Clicked View 2")
                        view.send_keys(u'\ue007')
                    except Exception as e:
                        print("Can't switch: {}".format(e))
                        driver.close()
                        driver.switch_to.window(current)
                        continue

                    # driver_copy = copy.deepcopy(driver)
                    # driver.execute_script("document.getElementById('loading').style.visibility = 'hidden';")
                    # print("Hide it !")

                    # TODO: NEED TO close that annoying loading window
                    try:
                        # wait for download window to open
                        time.sleep(5)
                        print("Download popup should have opened")

                        gui.moveTo(save_as_x, save_as_y) # position of the 'inverse triangle' button
                        gui.click()
                        gui.press('a') # save as
                        gui.click()
                        gui.typewrite(company + "_" + contributor)
                        gui.press('enter')
                        try:
                            # this is intended to close the loading window
                            close_unwanted_windows(driver, current)
                            # the new window for reports closes automatically
                        except Exception as e:
                            print("error with close")
                            print(e)
                        print("switching back to main search page")
                        # this hangs but is a  required call somehow
                        driver.switch_to.window(current)
                        print("go back to main search page")
                    except Exception as e:
                        pass
                        # print(e)
                        # print("Can't switch back to search page")

                    # go back to search window
                    # driver.switch_to.window(windows[0])

                    # time.sleep(10)

                    # do your stuffs
                    # driver.close()
                    # driver.switch_to.window(default_handle)
                    # print("the windows:", driver.window_handles)
                    # print("current:", current)
                    # try:
                    #     driver.switch_to.window(current)
                    #     print("go back to search")
                    # except Exception as e:
                    #     print("Can't switch back to search")
                    # see new_window.html

                    # TODO Pyautogui

                    # close opened window and go back
                    # https://sqa.stackexchange.com/questions/31436/how-to-save-a-file-by-clicking-on-a-link-in-python-internet-explorer-windows
                    # https://stackoverflow.com/questions/11449179/how-can-i-close-a-specific-window-using-selenium-webdriver-with-java
                except Exception as e:
                    print("Exception: {}_{}".format(company, contributor))
                    print(e)


            print("PROCESSED {}".format(company))
            print("---------------------------------")
            success.append(company)

            # TODO: filter results and generate reports
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
