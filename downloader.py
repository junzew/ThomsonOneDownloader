from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time
import datetime
import pandas as pd
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
success = []

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

def execute_app():
    start_time = time.time()
    symbols = load_symbols()
    # Start interacting with the browser
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
    driver.maximize_window()
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

    companies = ["RDSA-US", "AAPL-US", "ADBE-US"]# TODO only exact match
    # companies = symbols
    total_count = len(companies)
    for i, company in enumerate(companies):
        print("{} Processing {}".format(datetime.datetime.now(), company))
        print("Elapsed time: {:.1f} seconds".format(time.time() - start_time))
        print("Progress: {}/{}={:.2f}%".format(i, total_count, i/total_count * 100))
        for f in ["frameWinF7", "PCPWinF6", "PCPWinF7", "frameWinF6"]:
            try:
                # print("try switching to frame {}".format(f))
                driver.switch_to.frame(f)
                # time.sleep(1)
                trackinput = driver.find_element_by_id("trackInput")
                print("find the textbox in frame {}".format(f))
                trackinput.clear()
                print("cleared")
                trackinput.send_keys(company)
                print("entered company name {}".format(company))
                # time.sleep(1)

                go = driver.find_element_by_id("go")
                print("found go")
                go.send_keys(u'\ue007')
                print("clicked go")
                time.sleep(3)

                try:
                    # print(driver.page_source)
                    driver.switch_to.default_content()
                    driver.switch_to.frame("frameWinC1-4054-8135")
                    print("switched f1")
                    # siwtch to child frame
                    driver.switch_to.frame("PCPWinC1-4054-8135") #id =3234
                    print("switched frame")
                    from_date = driver.find_element_by_id("ctl00__criteria__fromDate")
                    from_date.clear()
                    from_date.send_keys("01/01/15")

                    to_date = driver.find_element_by_id("ctl00__criteria__toDate")
                    to_date.clear()
                    to_date.send_keys("01/01/18")
                    print("entererd dates")


                    # Contributor
                    contributors = """BARCLAYS, BMO CAPITAL MARKET, CREDIT SUISSE,
                    DEUTSCHE BANK, EVERCORE ISI,
                    HSBC GLOBAL RESEARCH, JEFFERIES,
                    JPMORGAN, MORGAN STANLEY,
                    RBC CAPITAL MARKETS (CANADA),. UBS RESEARCH,
                    WELLS FARGO SECURITIES, LLC"""

                    ctb_input = driver.find_element_by_id("contributors.searchText")
                    ctb_input.send_keys("BARCLAYS")
                    time.sleep(2)
                    # choose first by default
                    # option = driver.find_element_by_xpath("//div/table/tbody/tr/td/b")
                    option = driver.find_element_by_class_name("dijitMenuItem")
                    # Strange offset problem with Dojo combobox
                    h = ActionChains(driver).move_to_element(option).move_by_offset(0,-60)
                    h.click().perform()

                    # Search
                    bt_search = driver.find_element_by_id("ctl00__criteria__searchButton")
                    bt_search.send_keys(u'\ue007')
# web_driver.execute_script("dijit.byId('<id of select field>').set('value','<value you want it set to>')")

                    # print(driver.page_source)
                    driver.switch_to.default_content()
                    driver.switch_to.frame(f)
                    print("back to f")
                except Exception as e:
                    print(e)
                    pass


                time.sleep(1)
                success.append(company)

                # TODO: filter results and generate reports
                driver.switch_to.default_content()
                break
            except:
                pass

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
