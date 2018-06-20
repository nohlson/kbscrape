from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pickle
import sys

## nateohlson 2018
np = pickle.load(open('np.p', 'rb'))
USER_NAME = np[0]
PASSWORD = np[1]
TIMEOUT = 30


class Kb:

    def __init__(self, number, company, domain, vendor, topic, task_category, task_sub_category, kb_topic, kb_sub_topic,
                 short_description, body):
        self.number = number
        self.company = company
        self.domain = domain
        self.vendor = vendor
        self.topic = topic
        self.task_category = task_category
        self.task_sub_category = task_sub_category
        self.kb_topic = kb_topic
        self.kb_sub_topic = kb_sub_topic
        self.short_description = short_description
        self.body = body

    def print_info(self):
        print(
            'KB NUMBER: {}\n\nCompany: {}\nDomain: {}\nVendor: {}\nTopic: {}\nTask Category: {}\nTask Subcategory: {}\nTopic: {}\nSubtopic: {}\nShort description: {}\n\nBody:\n{}'.format(
                self.number, self.company, self.domain, self.vendor, self.topic, self.task_category,
                self.task_sub_category, self.kb_topic, self.kb_sub_topic, self.short_description,
                self.body.encode('ascii', 'ignore')))

    def write_format_to_file(self, fd):
        fd.write('@@@\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}\n\n{}'.format(self.number, self.company, self.domain,
                                                                        self.vendor, self.topic, self.task_category,
                                                                        self.task_sub_category, self.kb_topic,
                                                                        self.kb_sub_topic, self.short_description,
                                                                        self.body.encode('utf-8')))


def main():
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk',
                           'text/csv/xls/jpeg/octet-stream/png/vnd.openxmlformats-officedoc/pdf/')
    driver = webdriver.Firefox()
    driver.get("https://commandcenter.service-now.com")
    driver.set_page_load_timeout(TIMEOUT)

    login(driver)
    navigate_to_kbs(driver)

    get_all_kbs(driver)

    raw_input("enter...")
    driver.close()


def get_all_kbs(driver):
    total_kbs = 0
    f = open('output.txt', 'a')
    if len(sys.argv) > 1 and sys.argv[1] == 'reset':
        finished_kb_numbers = []
        pickle.dump(finished_kb_numbers, open("finished_kb.p", "wb"))
    else:
        finished_kb_numbers = pickle.load(open("finished_kb.p", "rb"))
    gsft_main = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "gsft_main")))
    driver.switch_to.frame(gsft_main)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[contains(@id, "row_kb_knowledge_")]')))

    first_kb = driver.find_element_by_xpath('//*[contains(@id, "row_kb_knowledge_")]')
    first_kb_link = first_kb.find_element_by_xpath('.//*[@class="linked formlink"]')
    first_kb_link.click()

    next_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[contains(@title, "Next record ")]')))

    while next_button.get_attribute('disabled') != "true":
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kb_knowledge.number"]')))
        number_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.number"]')
        number = number_el.get_attribute('value')

        print("Checking if {} has already been found".format(str(number)))
        already_found = False
        for finished_kb in finished_kb_numbers:
            if str(number) == str(finished_kb):
                already_found = True
                break
        if not already_found:

            print("Getting information from kb")
            company_el = driver.find_element_by_xpath('//*[@id="sys_display.original.kb_knowledge.u_company"]')
            company = company_el.get_attribute('value')

            domain_el = driver.find_element_by_xpath('//*[@id="sys_display.original.kb_knowledge.sys_domain"]')
            domain = domain_el.get_attribute('value')

            vendor_el = driver.find_element_by_xpath('//*[@id="sys_display.original.kb_knowledge.u_vendor"]')
            vendor = vendor_el.get_attribute('value')

            topic_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.topic"]')
            try:
                selected_topic_option = topic_el.find_element_by_xpath('.//*[@selected="SELECTED"]')
                topic = selected_topic_option.get_attribute('value')
            except NoSuchElementException:
                topic = "None"

            task_category_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.u_task_category"]')
            try:
                selected_task_category_option = task_category_el.find_element_by_xpath('.//*[@selected="SELECTED"]')
                task_category = selected_task_category_option.get_attribute('value')
            except NoSuchElementException:
                task_category = "None"

            task_sub_category_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.u_task_subcategory"]')
            try:
                selected_task_sub_category_option = task_sub_category_el.find_element_by_xpath(
                    './/*[@selected="SELECTED"]')
                task_sub_category = selected_task_sub_category_option.get_attribute('value')
            except NoSuchElementException:
                task_sub_category = "None"

            kb_topic_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.u_kb_topic"]')
            try:
                selected_kb_topic_option = kb_topic_el.find_element_by_xpath('.//*[@selected="SELECTED"]')
                kb_topic = selected_kb_topic_option.get_attribute('value')
            except NoSuchElementException:
                kb_topic = "None"

            kb_sub_topic_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.u_kb_subtopic"]')
            kb_sub_topic = kb_sub_topic_el.get_attribute('value')

            short_description_el = driver.find_element_by_xpath('//*[@id="kb_knowledge.short_description"]')
            short_description = short_description_el.get_attribute('value')

            # get body

            text_ifr = driver.find_element_by_id("kb_knowledge.text_ifr")
            driver.switch_to.frame(text_ifr)
            body = driver.page_source
            driver.switch_to_default_content()
            gsft_main = driver.find_element_by_id("gsft_main")
            driver.switch_to.frame(gsft_main)

            new_kb = Kb(number, company, domain, vendor, topic, task_category, task_sub_category, kb_topic,
                        kb_sub_topic, short_description, body)
            total_kbs += 1
            print("Total KBs: {}".format(total_kbs))

            new_kb.write_format_to_file(f)
            finished_kb_numbers.append(str(number))
            pickle.dump(finished_kb_numbers, open("finished_kb.p", "wb"))

        while True:
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(@title, "Next record ")]')))
                next_button.click()
                break
            except KeyboardInterrupt:
                sys.exit(1)
            except:
                print("Failed to click next button retrying...")

        while True:
            try:
                next_button = driver.find_element_by_xpath('//*[contains(@title, "Next record ")]')
                break
            except KeyboardInterrupt:
                sys.exit(1)
            except:
                pass


def navigate_to_kbs(driver):
    # wait for element to exist and then click it. always returning to default content
    edit_kb_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@title="Edit"]')))
    edit_kb_button.click()
    driver.switch_to.default_content()


def login(driver):
    # swtich to login iframe
    frame = driver.find_element_by_name("gsft_main")
    driver.switch_to.frame(frame)

    # enter username
    email_entry_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'user_name')))
    email_entry_element.send_keys(USER_NAME)

    # enter password
    password_entry_element = driver.find_element_by_name("user_password")
    password_entry_element.send_keys(PASSWORD)

    # find login button and click
    login_button = driver.find_element_by_name("not_important")
    login_button.click()

    # return to default content i.e. exit the iframe
    driver.switch_to.default_content()


if __name__ == "__main__":
    main()
