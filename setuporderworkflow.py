from selenium.common import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait
import logging
import os

# importing from another python file
from Orders import SheetOrder

logging.basicConfig(filename='myfile.log', level=logging.ERROR)
logger = logging.getLogger(__name__)


class OrderWorkflow:

    def __init__(self, driver):

        self.driver = driver  # webdriver.WebDriver
        self.action = ActionChains(driver)

        self.level_count_twf = len(so.level_twf)
        print(f'Template workflow level count: {self.level_count_twf}')

        print(f'Order workflow level count: {so.owf_level_count}')

    def setup_for_owf(self):
        try:
            setup_btn = self.driver.find_element(By.ID, "left-tabs-example-tab-UserManagement")
            if setup_btn.is_displayed():
                setup_btn.click()
            sleep(2)
            self.module_order_workflow()
        except NoSuchElementException:
            print("NoSuchElementException:", "Can't find setup button at left so global search will work")
            if NoSuchElementException:
                owf_module_name = "Order Workflow"
                so.search_module(owf_module_name)

    def setup_for_twf(self):
        try:
            setup_btn = self.driver.find_element(By.ID, "left-tabs-example-tab-UserManagement")
            if setup_btn.is_displayed():
                setup_btn.click()
            sleep(2)
            self.module_template_workflow()
        except NoSuchElementException:
            print("NoSuchElementException:", "Can't find setup button at left so global search will work")
            if NoSuchElementException:
                twf_module_name = "Template Workflow"
                so.search_module(twf_module_name)

    def module_template_workflow(self):
        self.driver.find_element(By.XPATH, "//a[@href='/Sheetworkflow']//span[1]").click()
        sleep(10)

    def module_order_workflow(self):
        self.driver.find_element(By.XPATH, "//a[@href='/Orderworkflow']//span[1]").click()
        sleep(10)

    def tab_template_workflow(self):
        self.driver.find_element(By.CSS_SELECTOR, "#rc-tabs-4-tab-2").click()
        sleep(6)

    def template_wf_level(self, level_count, level_name, user_role):

        for r in range(level_count):
            wf_count = str(r)
            wf_box = self.driver.find_element(By.XPATH, "//input[@index='" + wf_count + "']")
            wf_box.clear()
            wf_box.send_keys(level_name[r])
            print(f'level {r}: {level_name[r]}')
            sleep(2)
            # self.driver.find_element(By.XPATH, "//td[@class='kendo-select-col ']//input").click()
            self.assign_user_role_wf(user_role[r])
            if r in range(level_count - 1):
                self.add_step_tw()
            else:
                pass
        self.save_btn()

    def assign_user_role_wf(self, user_role):
        self.driver.find_element(By.XPATH, "//td[text()='" + user_role + "']").click()
        sleep(1)

    def delete_wf_level(self, level_count):
        for d in range(level_count):
            if d == 2:
                sleep(5)
                self.delete_icon()
            else:
                self.delete_icon()
        self.add_step_tw()
        self.delete_icon()

    def add_step_tw(self):
        self.driver.find_element(By.XPATH, "//li[@class='list-inline-item']//button").click()
        sleep(2)
        return

    def save_btn(self):
        self.driver.find_element(By.XPATH, "(//li[@class='list-inline-item']//button)[2]").click()
        try:
            reason = so.audit_reason[2]
            cmts = "workflow saved by Automation"
            so.audit_trail(reason, cmts)
        except ElementNotInteractableException:
            print("Audit trail is not enabled for save btn at template workflow")
        sleep(6)

    def delete_icon(self):
        self.driver.find_element(By.XPATH, "//ul[@class='icon-list-view']//i").click()
        try:
            alert_msg = self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text
            if "Workflow should have atleast one step".upper() == alert_msg:
                print(alert_msg)
                logger.info(f"Alert message: {alert_msg}")
                sleep(5)
        except NoSuchElementException as e:
            logger.warning(f"Failed or pop up not come: {e}")
        try:
            self.driver.find_element(By.XPATH, "//div[@class='react-confirm-alert-button-group']//button[1]").click()
            sleep(2)
        except NoSuchElementException as e:
            logger.warning(f"Failed to find or click on 'Yes' button: {e}")
            pass
        try:
            reason = so.audit_reason[3]
            cmts = "Workflow deleted by Automation"
            so.audit_trail(reason, cmts)
            sleep(2)
        except ElementNotInteractableException:
            print("Audit trail is not enabled for delete btn at template workflow")


if __name__ == '__main__':
    so = SheetOrder()
    so.login()
    ow = OrderWorkflow(so.driver)
    ow.setup_for_owf()
    ow.template_wf_level(ow.level_count_twf, so.level_twf, so.user_role_twf)
    ow.delete_wf_level(ow.level_count_twf)
    ow.template_wf_level(ow.level_count_twf, so.level_twf, so.user_role_twf)
    ow.tab_template_workflow()
    # ow.setup_for_twf()
    ow.template_wf_level(so.owf_level_count, so.level_owf, so.user_role_owf)
    ow.delete_wf_level(so.owf_level_count)
    ow.template_wf_level(so.owf_level_count, so.level_owf, so.user_role_owf)
