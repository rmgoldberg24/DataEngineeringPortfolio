from selenium import webdriver
from selenium.webdriver import ActionChains
import time

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

wait = WebDriverWait(driver, 10)

# Initialize webdriver and navigate to URL
def init(url):
    print('Navigating to {0} . . . '.format(url))
    driver = webdriver.Chrome(keep_alive=True)
    driver.get(url)
    return(driver)

# Login to Cognos
def logIn(driver):
    username = driver.find_element_by_id('sample.username')
    username.click()
    username.send_keys('')
    password = driver.find_element_by_id('sample.password')
    password.click()
    password.send_keys('')
    submit = driver.find_element_by_id('sample.login')
    submit.click()
    return(driver)

# Take Ownership of report (action chain)
def takeOwnershipOptimized(driver, report):
    print(report)
    text = "[data-name^='{0}']".format(report)
    report = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, text)))
    actionChains = ActionChains(driver)
    actionChains.context_click(report).perform()
    time.sleep(1)
    text = "[aria-label^='Take ownership']"
    report = driver.find_elements_by_css_selector(text)
    if report:
        time.sleep(1)
        report[0].click()
        ok = wait.until(EC.presence_of_element_located((By.ID, 'ok')))
        ok.click()
    else:
        print('Moving on . . . ')
        goBack = driver.find_element_by_class_name('commonMenu')
        goBack.click()
    return(driver)

def ownerCapabiltiesOptimized(driver, report):
    print(report)
    text = "[data-name^='{0}']".format(report)
    report = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, text)))
    actionChains = ActionChains(driver)
    actionChains.context_click(report).perform()
    properties = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Properties')]")))
    properties.click()
    reportTab = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabLabel_Report")))
    reportTab.click()
    advanced = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'l_advancedReportProperties')]//div[contains(text(), 'Advanced')]")))
    advanced.click()
    ownerCapabilities = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-value^='runWithOwnerCapabilities']")))
    ownerCapabilities.click()
    backArrow = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='teamFoldersSlideoutContent']//ul[contains(@class, 'breadcrumbPrevious')]")))
    backArrow.click()
    return(driver)

if __name__ == "__main__":
    url = 'COGNOSURL'
    driver = init(url)
    driver = logIn(driver)

    # Open Team Folder
    teamContent = driver.find_element_by_id('com.ibm.bi.contentApps.teamFoldersSlideout')
    teamContent.click()

    # Get List of Reports in Current View
    reportElements = driver.find_elements_by_xpath("//tr[td/div[@title='Report']]")
    reportList= []
    for element in reportElements:
        reportName = element.get_attribute('data-name')
        print(reportName)
        reportList.append(reportName)

    # Take ownership of all reports
    for report in reportList:
        driver = takeOwnershipOptimized(driver, report)

    # Change capability setting for all reports
    for report in reportList:
        driver = ownerCapabiltiesOptimized(driver, report)
