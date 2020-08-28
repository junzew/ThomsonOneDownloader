# ThomsonOneDownloader


Dowanload pdf reports from the Thomson ONE database from
http://resources.library.ubc.ca/1423

You need a CWL (Campus Wide Login) account to access it.

Thomson One works with IE 10 & 11, so the script only works on Windows.

## To Set Up
* Install git and clone this repo

* Install Python 3.7 and pip:
https://www.python.org/downloads/windows/

* Add Python and Python/Scripts to PATH:
https://datatofish.com/add-python-to-windows-path/

* Install dependencies:

```
pip install selenium
pip install pandas
pip install xlrd
pip install pywinauto==0.6.6
pip install pyautogui
```

* Add IEDriverServer to PATH, and configure IE per:
https://github.com/SeleniumHQ/selenium/wiki/InternetExplorerDriver#required-configuration

## To Run
From cmd, execute `python downloader.py`
Make sure that no other instance of IE is running.

