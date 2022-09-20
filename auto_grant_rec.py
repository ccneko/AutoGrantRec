"""
RGC application grant record auto-filler (Command-line version)
==

This auto-filler script takes in a pre-filled Excel spreadsheet grant record
to fill the grant record section in the RGC grant application. It was written
in Python 3 and first built for a grant application in 2019, and reused in 
2020 and 2021. To ease future work for fellow grant applicating PIs,
the script has been generalized to take in custom User ID, Password, PI name
and input Excel file. A template Excel file goes along with this script.

This is the lightweight command-line (CLI) version.

## Dependencies
### Python3 packages
- `selenium`: webdriver
- `pandas`:   table handler
- `xlrd`:     Excel handler
- `openpyxl`:   Excel handler

Install Python3 package prerequities by
- `python3 -m pip selenium pandas xlrd`

### Chrome Browser
- https://www.google.com/chrome/
- via commmand-line on Linux Ubuntu:
  - `sudo sh -c 'echo "deb http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'`
  - `sudo wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | sudo apt-key add -`
  - `sudo apt update`
  - `sudo apt install google-chrome-stable`
  
### Chrome driver
- Check your Chrome browser version from "Help > About Google Chrome"
  （說明 > 關於 Google Chrome）
- Download chromedriver that matches your OS and Chrome browser version from
  https://chromedriver.chromium.org/downloads
- Unzip the downloaded package and note the path to the chromedriver
  e.g. '/Users/ChanTaiMan/Downloads/chromedriver'

## Usage
- `python3 auto_grant_rec.py -u USER_ID -p PASSWD -n "CHAN, Tai-man"
   -c /path/to/chromedriver -i yourinput.xlsx`

## Remarks
- This script first CLEARS any existing record in the online system before 
  filling the form according to your input file. Please make sure your input 
  file contains all necessary records.
- Add double quotes around PI name to let the argument parser read the whole
  name containing space as one argument.
- The browsing may stuck, e.g. at the proposal menu, in some rare occasions due
  to browser request timing issue. Just rerun the script and this should be
  solved.
- Note that the headless mode skips showing the browser pop-up to free up the
  screen, runs faster, but has a higher chance of Timeout error.

## License (MIT)
Copyright 2021 Claire Chung

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished
to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
"""

__author__ = 'Claire Chung'
__version__ = '1.3'
__license__ = "MIT License"

from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import StaleElementReferenceException, \
    ElementNotInteractableException
import pandas as pd
import datetime
import logging
import argparse

### Set input arguments ###


def fill_rgc():
    parser = argparse.ArgumentParser(description='Parse user ID, password ' +
                                                 'and grant record Excel file '+
                                                 'to the GRF application for '
                                                 'auto form filling.')
    parser.add_argument('-u', '--user_id', metavar='USER_ID', type=str,
                        required=True, help='User ID')
    parser.add_argument('-p', '--pw', metavar='PASSWD', type=str,
                        required=True, help='Password')
    parser.add_argument('-i', '--input', metavar='EXCEL_FILE_PATH', type=str,
                        required=True, help='Input Excel file')
    parser.add_argument('-n', '--pi_name', metavar='PI_NAME', type=str,
                        required=True,
                        help='PI name. Add double quotes, e.g. "Chan, Tai-man"')
    parser.add_argument('-c', '--chromedriver_path',
                        metavar='CHROME_DRIVER_PATH',
                        type=str, default='chromedriver',
                        help='path to chromedriver')
    parser.add_argument('-l', '--log_path', metavar='LOG_FILE_PATH', type=str,
                        default=datetime.datetime.now().strftime(
                                '%Y%m%d-%H%M%S') + '-rgc-grantrec.log',
                        help='path to run log')
    parser.add_argument('--verbose', nargs='?')
    parser.add_argument('--version', '-v', action='version',
                        version='%(prog)s ' + __version__)
    parser.add_argument('--headless', nargs='?', const=True,
                        help='Add this argument to skip showing the browser')
    args = parser.parse_args()


    ### Initialize logger ###
    log_filename = args.log_path
    if not args.verbose:
        logging.basicConfig(filename=log_filename, level=logging.INFO)
    elif args.verbose == 1:
        logging.basicConfig(filename=log_filename, level=logging.DEBUG)
    logFormatter = logging.Formatter("%(asctime)s [%(threadName)-12.12s] " +
                                     "[%(levelname)-5.5s]  %(message)s")
    logger = logging.getLogger()
    consoleHandler = logging.StreamHandler()
    consoleHandler.setFormatter(logFormatter)
    logger.addHandler(consoleHandler)

    logger.debug(args)

    ### Prepare URL ###
    login_url = 'https://cerg1.ugc.edu.hk/cergprod/login.jsp'
    form_url = 'https://cerg1.ugc.edu.hk/cergprod/' + \
               'ControlServlet?FunctionName=UF501&FunctionID=SCRUM501_12' + \
               '&action_type=GOTO&seq=867705'
    # '&seq=' prevents directly entering the page

    ### Prepare data ###
    df = pd.concat(pd.read_excel(args.input,
                                 sheet_name=['On-going', 'Completed',
                                             'Pending']))
    df = df.loc[df['End date'] >= datetime.datetime.strptime(
        str(datetime.datetime.now().year - 4) + '-10-01', '%Y-%m-%d')].\
        reset_index(drop=True)
    df = df.loc[~df['Role'].isna()]

    ### Prepare browser worker ###
    chrome_options = Options()
    # chrome_options.add_argument("--user-data-dir=chrome-data")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=us")
    if args.headless:
        chrome_options.add_argument("--headless")

    chrome_service = fs.Service(executable_path=args.chromedriver_path)
    driver: WebDriver = webdriver.Chrome(options=chrome_options,
                                         service=chrome_service)
    driver.delete_all_cookies()
    wait = WebDriverWait(driver, 10)

    ### Log in ###
    timestamp = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    logger.info(timestamp)

    driver.get(login_url)
    wait.until(EC.element_to_be_clickable((By.NAME, 'submit')))
    logger.info('Login page loaded')

    input_userid = driver.find_element_by_xpath("//input[@maxlength='20']")
    input_userid.send_keys(args.user_id)
    logger.info('User ID filled.')

    pwd_filled = False
    while not pwd_filled:
        try:
            input_pwd = driver.find_element_by_xpath(
                "//input[@type='password']")
            input_pwd.send_keys(args.pw)
            logger.info('User Password input filled.')
            pwd_filled = True
        except StaleElementReferenceException:
            logger.debug(
                "WARNING: Stale Element Reference (Password). Retrying.")
        except ElementNotInteractableException:
            logger.debug(
                "WARNING: ElementNotInteractable (Password). Retrying.")

    driver.find_element_by_name("submit").click()
    logger.info("Login request submitted.")

    ### Select role ###
    wait.until(EC.element_to_be_clickable((By.NAME, "Continue")))
    driver.find_element_by_name("Continue").click()
    logger.info("User role selected.")

    wait.until(EC.element_to_be_clickable(
        (By.LINK_TEXT, "Prepare Proposal / View Internal Comments")))
    logger.info("Project maintenance page loaded.")

    main_window = driver.current_window_handle
    logger.debug(main_window)
    driver.find_element_by_link_text(
        "Prepare Proposal / View Internal Comments").click()
    logger.info("Prepare Proposal clicked.")
    logger.debug(driver.window_handles)
    driver.switch_to.window(driver.window_handles[1])
    wait.until(EC.element_to_be_clickable((By.NAME, "yes"))) # value: "I accept"
    driver.find_element_by_name("yes").click()
    logger.info("Terms accepted.")
    driver.switch_to.window(main_window)
    wait.until(EC.element_to_be_clickable((By.NAME, "ProposalMenu")))
    driver.find_element_by_name("ProposalMenu").click()
    driver.find_element_by_link_text(
        "Grant Record and Related Research Work of Investigator(s)").click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH,"//input[@value=' Add Project / Work " +
         "(GRF/ECS & non-GRF/non-ECS) ']")))

    min_refno_len = 6

    ### Clean all old entries ###
    cleared = False
    while not cleared:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@value=' Add Project / Work " +
                       "(GRF/ECS & non-GRF/non-ECS) ']")))
        buttons = driver.find_elements(By.XPATH, "//input[@type='button']")
        button_value = buttons[0].get_attribute('value').strip()
        if len(button_value) >= min_refno_len \
                and 'Objective' not in button_value \
                and 'Project' not in button_value:
            buttons[0].click()
            wait.until(EC.element_to_be_clickable((By.NAME, 'piName')))
            driver.find_element_by_name('del').click()
            driver.switch_to.alert.accept()
        else:
            buttons = driver.find_elements(By.XPATH, "//input[@type='button']")
            for button in buttons:
                button_value = buttons[0].get_attribute('value').strip()
                assert len(button_value) < min_refno_len \
                       or 'Objective' in button_value \
                       or 'Project' in button_value
            cleared = True
    logger.info('All old entries cleaned to prepare for new input.')

    ### Input record ###
    role_dict = {'PI': 'P', 'PC': 'PC', 'Co-I': 'C', 'Co-PI': 'Co-PI', 0: ''}
    filled_cnt = 0

    for row in df.iterrows():
        timestamp = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        logger.info(timestamp)
        inputdict = {}
        inputdict["NPI"] = args.pi_name
        inputdict["CAP"] = role_dict[row[1]["Role"]]  # value: P/PC/C/Co-PI
        inputdict["FSF"] = ["N", "Y"][row[1]['Funding source'] == "GRF"]  # Y/N
        inputdict["FSR"] = row[1]['Funding source']
        proj_status = {"On-going": "O", "Completed": "Z", "Pending": "U"}
        inputdict["STA"] = proj_status[row[1]["Status"]]
        #inputdict["STA"] = ['Z', 'O'][
        #    row[1]["End date"] >= pd.Timestamp(datetime.date.today())]  # O/Z/U
        try:
            # Prevents adding extra .0 as float due to Excel auto-formatting
            inputdict["RNO"] = str(int(row[1]["Reference number"]))
            logger.warning("Reference number coerced to integer "
                           + inputdict["RNO"] + ". Please check.")
        except ValueError:
            inputdict["RNO"] = str(row[1]["Reference number"])
        if inputdict["RNO"] == 'nan':
            inputdict["RNO"] = ''
        logger.info(inputdict["RNO"])
        inputdict["PTI"] = str(row[1]["Project title"])
        inputdict["FAM"] = str(row[1]["Amount (HK$)"])
        inputdict["RGC"] = str(row[1]["UGC/RGC funding"])  # Y/N
        inputdict["SDA"] = str(row[1]["Start date"].day)
        inputdict["SMO"] = str(row[1]["Start date"].month)
        inputdict["SYR"] = str(row[1]["Start date"].year)
        inputdict["CDA"] = str(row[1]["End date"].day)
        inputdict["CMO"] = str(row[1]["End date"].month)
        inputdict["CYR"] = str(row[1]["End date"].year)
        if inputdict["CAP"] != "C" and inputdict["STA"] != "U":
            inputdict["NHR"] = str(int(row[1]["Number of hours"]))
        else:
            inputdict["NHR"] = 0
        inputdict["OBJ"] = str(row[1]["Project Objectives"])

        # Load the Form
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@value=' Add Project / Work " +
                       "(GRF/ECS & non-GRF/non-ECS) ']")))
        add_proj = driver.find_element_by_xpath(
            "//input[@value=' Add Project / Work " +
            "(GRF/ECS & non-GRF/non-ECS) ']")
        add_proj.click()
        wait.until(EC.element_to_be_clickable((By.NAME, "piName")))  # id same

        # Name of Investigator(s) :
        npi_filled = False
        while not npi_filled:
            try:
                input_pi_name = driver.find_element_by_name("piName")
                input_pi_name.send_keys(inputdict["NPI"])
                assert input_pi_name.get_attribute("value") == inputdict["NPI"]
                npi_filled = True
                logger.info("Name of Investigator(s) filled.")
            except AssertionError:
                logger.debug(
                    "WARNING: Wrong value (Name of Investigator(s)). Retrying")
                input_pi_name.clear()

        # Capacity
        driver.find_element_by_name("capacity").click()
        driver.find_element_by_xpath(
            "//select[@name='capacity']/option[@value='" + str(
                inputdict['CAP']) + "']").click()
        # P/PC/C/Co-PI
        logger.info("(Investigator) Capacity selected.")

        # Funding Sources
        # radio, name=fund_src_flag, id=(fund_src_flag_Y, fund_src_flag_N)
        driver.find_element_by_id(
            "fund_src_flag_" + inputdict["FSF"]).click()  # radio button
        if inputdict['FSF'] == "N":
            driver.find_element_by_name("fund_src").send_keys(inputdict["FSR"])
            assert driver.find_element_by_name("fund_src")\
                       .get_attribute("value") == inputdict["FSR"]
        logger.info("Funding Sources filled.")

        # Status
        driver.find_element_by_name("proj_status").click()
        driver.find_element_by_xpath(
            "//select[@name='proj_status']/option[@value='" + inputdict[
                'STA'] + "']").click()  # id same
        assert Select(driver.find_element_by_name(
            "proj_status")).first_selected_option.get_attribute("value") == \
               inputdict["STA"]
        logger.info("Status filled.")

        # Project Reference No.(if any)
        driver.find_element_by_name("ref_no").send_keys(
            inputdict["RNO"])  # text, no id
        assert driver.find_element_by_name("ref_no").get_attribute("value") == \
               inputdict["RNO"]
        logger.info("Project Reference No. filled.")

        # Project / Work Title
        driver.find_element_by_name("proj_title").send_keys(
            inputdict["PTI"])  # text, no id
        assert driver.find_element_by_name("proj_title").get_attribute("value")\
               == inputdict["PTI"]
        logger.info("Project / Work Title filled.")

        # Funding Amount (HK$) (if not applicable, please input zero)
        try:
            inputdict["FAM"] = str(int(float(inputdict["FAM"])))
        except ValueError:
            inputdict["FAM"] = 0
            assert inputdict["STA"] == 'U'
            logger.warning("0 filled for unknown funding amount.")
        driver.find_element_by_name("fund_amt").send_keys(inputdict["FAM"])
        assert driver.find_element_by_name("fund_amt").get_attribute("value") \
               == inputdict["FAM"] or driver.find_element_by_name("fund_amt")\
                   .get_attribute("value") == '0'
        logger.info("Funding Amount (HK$) filled.")

        # RGC / UGC Funding (radio button)
        if inputdict["FSF"] == "Y":
            assert inputdict["RGC"] == "Y"
        driver.find_element_by_id("ugcfunding_" + inputdict["RGC"]).click()
        logger.info("UGC/RGC funding filled.")

        # Start Date
        driver.find_element_by_name("s_day").click()
        driver.find_element_by_xpath(
            "//select[@name='s_day']/option[@value='" + inputdict[
                'SDA'] + "']").click()
        driver.find_element_by_name("s_month").click()
        driver.find_element_by_xpath(
            "//select[@name='s_month']/option[@value='" + inputdict[
                'SMO'] + "']").click()
        driver.find_element_by_name("s_year").click()
        driver.find_element_by_xpath(
            "//select[@name='s_year']/option[@value='" + inputdict[
                'SYR'] + "']").click()

        # Estimated / Completion Date
        driver.find_element_by_name("c_day").click()
        driver.find_element_by_xpath(
            "//select[@name='c_day']/option[@value=" + inputdict[
                'CDA'] + "]").click()
        driver.find_element_by_name("c_month").click()
        driver.find_element_by_xpath(
            "//select[@name='c_month']/option[@value=" + inputdict[
                'CMO'] + "]").click()
        driver.find_element_by_name("c_year").click()
        cyr_filled = False
        while not cyr_filled:
            try:
                driver.find_element_by_xpath(
                    "//select[@name='c_year']/option[@value=" + inputdict[
                        'CYR'] + "]").click()
                cyr_filled = True
                logger.info('Completion Year filled.')
            except ElementNotInteractableException:
                logger.debug("WARNING: ElementNotInteractable " +
                             "(Completion Year). Retrying.")

        # Number of Hours Per Week Spent by the PI in Each On-going Project*
        text_nhr = driver.find_element_by_name("workHourPer")
        nhr_filled = False
        if text_nhr.is_enabled() and int(float(inputdict["NHR"])) > 0:
            # "Percent of Work Hour Spent should be a positive integer."
            while not nhr_filled:
                try:
                    text_nhr.send_keys(inputdict["NHR"])  # text, d same
                    assert text_nhr.get_attribute("value") == inputdict["NHR"]
                    nhr_filled = True
                    logger.info("Number of Hours filled")
                except AssertionError:
                    text_nhr.clear()
                    logger.debug(
                        "WARNING: Wrong Number of Hours filled. Retrying.")

        # Project / Work Objective
        driver.find_element_by_name("projectObjective").send_keys(
            inputdict["OBJ"])  # textarea

        # Related to the current application
        driver.find_element_by_id("overlap_NA").click()
        # radio, name=overlap, id=(overlap_NA, overlap_RE)

        # Save record
        driver.find_element_by_name("add").click()
        filled_cnt += 1

    timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    logger.info(timestamp)

    logger.info("Record entry complete. A total of " + str(filled_cnt) +
        " entries filled.")

    ### Check and warn for duplicate Ref No ###
    """
    Attention:
    Renewable grants may have same Ref No and different Project Title 
    & Time Period.
    """

    entered = []
    dups = []

    buttons = driver.find_elements(By.XPATH, "//input[@type='button']")

    for button in buttons:
        button_value = button.get_attribute("value").strip()
        if len(
                button_value) >= min_refno_len \
                and "Objective" not in button_value \
                and "Project" not in button_value:
            if button_value in entered:
                dups.append(button_value)
                logger.warning(
                    "WARNING: Potential duplicated entry: " + button_value)
            entered.append(button.get_attribute("value"))

    logger.info("Potential duplicated entries: " + ', '.join(dups))

if __name__ == '__main__':
    fill_rgc()
