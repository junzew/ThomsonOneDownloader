from selenium import webdriver
import time
# https://stackoverflow.com/questions/48880646/python-selenium-use-a-browser-that-is-already-open-and-logged-in-with-login-cre
# https://stackoverflow.com/questions/24925095/selenium-python-internet-explorer
# https://github.com/SeleniumHQ/selenium/wiki/InternetExplorerDriver#required-configuration
# Enable 64-bit process for Enhanced Protected Mode
# https://stackoverflow.com/questions/40626810/how-to-fix-the-slow-sendkeys-on-ie-11-with-selenium-webdriver-3-0-0
# Use 32 bit version

driver_location = "C:\\Users\\admin\\Documents\\IEDriverServer.exe" # needs to be in PATH
login_url = "https://login.ezproxy.library.ubc.ca/login?qurl=http%3a%2f%2fezproxy.library.ubc.ca%2flogin%2fthomsonone"
url = "https://www.thomsonone.com/Workspace/Main.aspx?View=Action%3dOpen&BrandName=www.thomsonone.com&IsSsoLogin=True"
driver = webdriver.Ie(driver_location)

def execute_app():
    driver.get(login_url)

    foo = input("Please log in, wait for page to load, then press enter")
    # login with CWL
    print("Logged in")

    # refresh
    driver.get(url)
    driver.maximize_window()
    assert "Thomson One" in driver.title

    bt_company_view = driver.find_element_by_id("td13")
    print("Company View found, click")
    bt_company_view.click()
    foo = input("Go to Company View->Research, clear the search box, then Press Enter")
    # time.sleep(20)
    print("ready to start jobs")
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

    # TODO: list of companies should be read from the xlsx file
    companies = ["RDSA-US", "AAPL-US", "ADBE-US"]# TODO only exact match
    for company in companies:
        print("Processing {}".format(company))
        for f in ["frameWinF7", "PCPWinF6", "PCPWinF7", "frameWinF6"]:
            try:
                # print("try switching to frame {}".format(f))
                driver.switch_to.frame(f)
                time.sleep(1)
                trackinput = driver.find_element_by_id("trackInput")
                print("find the textbox in frame {}".format(f))
                trackinput.clear()
                trackinput.send_keys(company)
                print("entered company name {}".format(company))
                time.sleep(1)

                go = driver.find_element_by_id("go")
                print("found go")
                go.send_keys(u'\ue007')
                print("clicked go")
                time.sleep(5)
                # TODO: filter results and generate reports
                driver.switch_to_default_content()
                break
            except:
                pass

def handle_cleanup():
    driver.close()
    pass

def main():
    try:
        execute_app()
    finally:
        handle_cleanup()

if __name__=='__main__':
    main()

# element = driver.find_element_by_name("url")
# print(element)
# login_img.click()
