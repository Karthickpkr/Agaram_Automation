import json
from selenium import webdriver
from selenium.common import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep


class SheetOrder:

    def __init__(self):
        self.au_dit_reason = None
        self.module_name = None
        self.sheet_ordr_id = None
        self.sheet_order_id_manual = None
        self.sheet_ordr_id_export = None
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.action = ActionChains(self.driver)

        """ To open "Fill_me.json" and read the data as jsonfield,
            load the jsonfield data into self.data

            also can get any data values from the JSON field by using like self.data['FIELD NAME']
        """

        with open("Fill_me.json", 'r') as jsonfield:
            self.data = json.load(jsonfield)

        """ Example: 
                self.data['LOGIN']['username'] is assigned to self.username parameters 
        """

        self.org_id = self.data['LOGIN']['organization_id']
        self.login_url = self.data['LOGIN']['url']
        self.username = self.data['LOGIN']['username']
        self.password = self.data['LOGIN']['password']
        self.fol_name = self.data['FOLDER CREATION AND RENAME']['folder_name']
        self.rename_fol = self.data['FOLDER CREATION AND RENAME']['rename_folder']

        """ Register sheet order parameters """
        self.sheet_order_type = self.data['REGISTER SHEET ORDER FIELDS']['sheet_order_type']
        self.projects = self.data['REGISTER SHEET ORDER FIELDS']['projects']
        self.task = self.data['REGISTER SHEET ORDER FIELDS']['task']
        self.select_sheet_temp = self.data['REGISTER SHEET ORDER FIELDS']['select_sheet_temp']
        self.sample = self.data['REGISTER SHEET ORDER FIELDS']['sample']
        self.keyword = self.data['REGISTER SHEET ORDER FIELDS']['keyword']

        self.sheet_order_id_manual = self.data['SHEET ORDER ID MANUAL PROCESS']['sheet_order_id_manual']

        """ Audit trail """
        self.audit_reason = self.data['AUDIT TRAIL']['audit_trail_reason']
        self.audit_cmts = self.data['AUDIT TRAIL']['audit_trail_cmts']

        """ Data fill sheet order parameters """
        self.cellno_alert = self.data['GENERAL FIELDS SHEET ORDER']['ALERT']['cell_no']
        self.notify_before = self.data['GENERAL FIELDS SHEET ORDER']['ALERT']['notify_before_days']
        self.alert_summary = self.data['GENERAL FIELDS SHEET ORDER']['ALERT']['alert_summary']

        self.cellno_mand_field = self.data['GENERAL FIELDS SHEET ORDER']['MANDATORY FIELD']['cell_no']
        self.mandatory_field = self.data['GENERAL FIELDS SHEET ORDER']['MANDATORY FIELD']['mandatory_value']

        self.cellno_date = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL DATE']['cell_no']
        self.manual_date = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL DATE']['manual_date']

        self.cellno_date_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL DATA AND TIME']['cell_no']
        self.manual_date_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL DATA AND TIME']['manual_date_time']

        self.cellno_manual_field = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL FIELD']['cell_no']
        self.manual_field = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL FIELD']['manual_field']

        self.cellno_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL TIME']['cell_no']
        self.manual_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL TIME']['manual_time']

        self.cellno_multi_combo = self.data['GENERAL FIELDS SHEET ORDER']['MULTISELECT COMBOBOX']['cell_no']

        self.cellno_numeric_field = self.data['GENERAL FIELDS SHEET ORDER']['NUMERIC FIELD']['cell_no']
        self.numeric_field = self.data['GENERAL FIELDS SHEET ORDER']['NUMERIC FIELD']['numeric_field']

        self.cellno_text_wrap = self.data['GENERAL FIELDS SHEET ORDER']['TEXT WRAPPER']['cell_no']
        self.text_wrap = self.data['GENERAL FIELDS SHEET ORDER']['TEXT WRAPPER']['text_wrap']

        self.cellno_esign = self.data['GENERAL FIELDS SHEET ORDER']['E-SIGN']['cell_no']
        self.esign_pwd = self.data['GENERAL FIELDS SHEET ORDER']['E-SIGN']['esign_pwd']
        self.esign_cmts = self.data['GENERAL FIELDS SHEET ORDER']['E-SIGN']['esign_cmts']

        """ Icons """
        self.cellno_hyper_link = self.data['GENERAL FIELDS SHEET ORDER']['ICONS HYPERLINK']['cell_no']
        self.hyper_link = self.data['GENERAL FIELDS SHEET ORDER']['ICONS HYPERLINK']['hyper_link']
        self.attach_path = self.data['GENERAL FIELDS SHEET ORDER']['ICONS ATTACHMENTS']['attach_path']

        """ Order workflow steps"""
        self.work_flow_steps = self.data['SHEET ORDER WORKFLOW']['work_flow_steps']

    def login(self):
        self.driver.get(self.login_url[1])
        self.driver.find_element(By.XPATH, "//input[@id='idUsername']").send_keys(self.org_id[1])
        sleep(3)
        self.driver.find_element(By.ID, "inputText").click()
        sleep(3)
        uname = self.driver.find_element(By.ID, "idUsername")
        uname.send_keys(self.username)
        uname.send_keys(Keys.TAB)
        sleep(3)
        pwd = self.driver.find_element(By.ID, 'idPassword')
        pwd.send_keys(self.password)
        pwd.send_keys(Keys.ENTER)
        sleep(15)
        print(self.driver.title)

    def search_module(self):
        action = ActionChains(self.driver)
        search_box = self.driver.find_element(By.XPATH, "//form[@class='inline-form-search']//input[1]")
        action.click(search_box).send_keys(self.module_name).perform()
        sleep(2)
        module_click = self.driver.find_element(By.XPATH, "//span[@class='list-icon-view']//span")
        action.click(module_click).perform()
        sleep(10)

    def audit_trail(self):
        if self.driver.find_element(By.CSS_SELECTOR, "button#audittrail_closeform").is_enabled():
            sleep(2)
            audit_pwd = self.driver.find_element(By.NAME, "Password")
            audit_pwd.send_keys(self.password)
            sleep(2)
            aud_reason = self.driver.find_element(By.ID, "reasoncode")
            Select(aud_reason).select_by_visible_text(self.au_dit_reason)
            sleep(1)
            audit_cmt = self.driver.find_element(By.NAME, "Comment")
            audit_cmt.send_keys(self.audit_cmts)
            sleep(1)
            submit_btn = self.driver.find_element(By.CSS_SELECTOR, "button#audittrail_submitbtn>span")
            submit_btn.click()

    def orders(self):
        # ele = WebDriverWait(self.driver, 50)
        # ele.until(ec.visibility_of_element_located((By.XPATH,
        #                                   "//div[@id='bodycontent']/div[5]/div[3]/div[3]/div[1]/div[2]")))
        try:
            self.driver.find_element(By.ID, "left-tabs-example-tab-SheetView").click()
            sleep(2)
            self.register_sheet_orders()
        except NoSuchElementException:
            print("NoSuchElementException:", "Can't find Orders button at left so global search will work")
            if NoSuchElementException:
                self.module_name = "Sheet Orders"
                self.search_module()

    def register_sheet_orders(self):
        self.driver.find_element(By.XPATH, "//a[@href='/registertask']//span[1]").click()
        sleep(10)

    def home_folders(self):
        home_icon = self.driver.find_element(By.CSS_SELECTOR, "span#Sheet")
        home_icon.click()
        sleep(7)

    def new_folder(self):
        new_folder = self.driver.find_element(By.XPATH, "//button[contains(@class,'k-button btn')]//span[1]")
        new_folder.click()
        sleep(1)
        folder_name = self.driver.find_element(By.XPATH, "//div[@class='form-group']//input")
        folder_name.send_keys(self.fol_name)
        sleep(1)

        """ Visibility of new folder creation. 
               Site =1 , Only me = 2 , Project team = 3 
            based on selection it can replace 1 or 2 or 3 in the span area
        Start """
        # site = self.driver.find_element(By.XPATH, "(//label[@class='radio']//span)[1]")
        # site.click()
        # only_me = self.driver.find_element(By.XPATH, "(//label[@class='radio']//span)[2]")
        # only_me.click()
        # project_team = self.driver.find_element(By.XPATH, "(//label[@class='radio']//span)[3]")
        # project_team.click()
        """ End """

        add_btn = self.driver.find_element(By.CSS_SELECTOR, "button[value='add']")
        add_btn.click()
        sleep(2)

        try:
            if ("FOLDER NAME ALREADY EXISTS"
                    == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
                labexistcount = self.fol_name + '_' + str(1)
                sleep(1)
                new_folder.click()
                sleep(1)
                folder_name = self.driver.find_element(By.XPATH, "//div[@class='form-group']//input")
                folder_name.send_keys(labexistcount)
                sleep(2)
                add_btn = self.driver.find_element(By.CSS_SELECTOR, "button[value='add']")
                add_btn.click()
                sleep(2)
        except NoSuchElementException:
            print("Folder name exist popup not execute")

    def rename_folder(self):
        action = ActionChains(self.driver)
        sort_by = self.driver.find_element(By.XPATH, "//button[@class='k-button k-button-icon']//span")
        action.click(sort_by).send_keys(Keys.ARROW_DOWN + Keys.ENTER).perform()
        sleep(2)
        select_folder = self.driver.find_element(By.XPATH, "(//span[text()='" + self.fol_name + "'])[2]")
        select_folder.click()
        print(select_folder.text)
        action.context_click(select_folder).perform()
        sleep(2)
        rename_folder = self.driver.find_element(By.XPATH, "(//span[@class='k-link k-menu-link'])[2]")
        rename_folder.click()
        sleep(1)
        rename = self.driver.find_element(By.XPATH, "//div[contains"
                                                    "(@class,'k-content k-window-content')]//input[1]")
        rename.clear()
        rename.send_keys(self.rename_fol)
        sleep(1)
        rename_btn = self.driver.find_element(
            By.XPATH, "//div[@class='k-dialog-buttongroup k-dialog-button-layout-stretched']//button[1]")
        rename_btn.click()
        sleep(2)

    def delete_folder(self):
        action = ActionChains(self.driver)
        select_folder = self.driver.find_element(By.XPATH, "(//span[text()='" + self.rename_fol + "'])[2]")
        select_folder.click()
        action.context_click(select_folder).perform()
        sleep(1)
        del_option = self.driver.find_element(By.XPATH, "(//span[@class='k-link k-menu-link'])[3]")
        del_option.click()
        sleep(1)
        del_btn = self.driver.find_element(By.XPATH, "//button[@value='delete']")
        del_btn.click()
        sleep(2)

    def register_btn(self):
        self.driver.find_element(By.XPATH, "(//button[contains(@class,'k-button ')]//span)[2]").click()
        sleep(2)

    def order_sheet_template(self):
        order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        order_type.send_keys(self.sheet_order_type[0] + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys(self.projects + Keys.ENTER)
        sleep(1)
        task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        task.send_keys(self.task + Keys.ENTER)
        sleep(1)

        try:
            if ("No sheet templates have been mapped with this task".upper()
                    == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
                print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)
        except NoSuchElementException:
            pass

        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)

        if self.select_sheet_temp != "":
            select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
            select_sheet_template.send_keys(self.select_sheet_temp + Keys.ENTER)
            sleep(1)

        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

        if ("PLEASE SELECT THE SHEET"
                == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
            print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)

    def order_research_without_template(self):
        order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        order_type.send_keys(self.sheet_order_type[1] + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys(self.projects + Keys.ENTER)
        sleep(1)
        task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        task.send_keys(self.task + Keys.ENTER)
        sleep(2)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        select_sheet_template.send_keys(self.select_sheet_temp + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def excel_order(self):
        order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        order_type.send_keys(self.sheet_order_type[3] + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys(self.projects + Keys.ENTER)
        sleep(2)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def sheet_validation(self):
        order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        order_type.send_keys(self.sheet_order_type[4] + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys(self.projects + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def save_within(self):
        projects_folder = self.driver.find_element(By.XPATH, "//label[@class='radio']//"
                                                             "span[contains(text(),'Projects')]")
        projects_folder.click()
        sleep(1)

    def register_order_btn(self):
        regorderbtn = self.driver.find_element(By.XPATH, "//button[contains(@class,'btn btn-user')]")
        regorderbtn.click()
        try:
            self.audit_cmts = "Sheet order registered by Automation"
            self.au_dit_reason = self.audit_reason[0]
            self.audit_trail()
        except ElementNotInteractableException:
            print("Audit trail is not enabled for register order button")
        finally:
            sleep(3)

    def process_order_sheet(self):  # process order grid button click

        if self.sheet_order_id_manual == "":
            so_id = self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]")
            print(so_id.text)
            split_id = str(so_id.text).split()
            self.sheet_ordr_id = split_id[0].split(":")[1]
            self.sheet_ordr_id_export = self.sheet_ordr_id
            print(self.sheet_ordr_id)
            order_grid_view = self.driver.find_element(By.XPATH, "//span[text()='" + self.sheet_ordr_id + "']")
            order_grid_view.click()
        else:
            order_grid_view = self.driver.find_element(By.XPATH, "//span[text()='" + self.sheet_order_id_manual + "']")
            order_grid_view.click()
            sleep(1)
        process_order_btn = self.driver.find_element(By.XPATH, "(//span[@class='clsorderopenonlist'])[1]")
        process_order_btn.click()
        try:
            self.audit_cmts = "Process order sheet opened by Automation"
            self.au_dit_reason = self.audit_reason[1]
            self.audit_trail()
        except ElementNotInteractableException:
            print("Audit trail is not enabled for process order sheet")
        finally:
            sleep(10)

    def name_box(self, cellno):
        cell_box = self.driver.find_element(By.XPATH, "//input[@title='Name Box']")
        cell_box.clear()
        cell_box.send_keys(cellno + Keys.ENTER)
        sleep(1)

    def fx(self, formula_box=None):
        fx = self.driver.find_element(By.XPATH, "//div[@class='k-spreadsheet-formula-bar']//div")
        # action = ActionChains(self.driver)
        # action.send_keys(fx + Keys.CONTROL + 'A' + Keys.BACKSPACE + formula_box + Keys.ENTER).perform()
        fx.send_keys(Keys.CONTROL + 'A' + Keys.BACKSPACE)
        fx.send_keys(formula_box + Keys.ENTER)
        sleep(3)

    def data_fill_sheet_order(self):

        """ Add_Resource """
        # add_resource_dd = self.driver.find_element(By.XPATH, "//div[@class='k-button k-spreadsheet-editor-button']//span[1]")
        # add_resource_dd.click()
        # sleep(2)
        # material_type = self.driver.find_element(By.XPATH, "//span[contains(text(),'NA')]")
        # material_type.click()
        # sleep(1)

        # # material_type = WebDriverWait(self.driver, 10).until(
        # #     ec.element_to_be_clickable((By.XPATH, "//span[contains(text(),'NA')]")))
        # # material_type.click()
        # # material_type_dd = WebDriverWait(self.driver, 10).until(
        # #     ec.visibility_of_all_elements_located((By.CSS_SELECTOR, ".k-list[aria-hidden='false'] > ["
        # #                                                             "aria-selected='true']")))
        # # material_type.send_keys(Keys.ARROW_DOWN + Keys.ENTER)
        # # action = ActionChains(self.driver, material_type)
        # # action.click().send_keys(Keys.ARROW_DOWN + Keys.ENTER)
        # # material_type.send_keys(Keys.ENTER)
        # # material_type = Select(self.driver.find_element(By.CSS_SELECTOR, "span[class='k-dropdown-wrap k-state-default k-state-hover'] span[class='k-input']"))
        # # material_type.select_by_visible_text("Standard Type")
        # # material_type.select_by_index(1)
        # sleep(5)

        """ Alert """
        self.name_box(self.cellno_alert)
        alert_dd = self.driver.find_element(By.XPATH, "//span[contains(@class,'k-icon k-icon')]")
        alert_dd.click()
        sleep(2)
        # due_date_icon = self.driver.find_element(By.CSS_SELECTOR, ".k-icon.k-i-calendar[shub-ins='1']")
        due_date_icon = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-i-calendar']")
        due_date_icon.click()
        sleep(1)
        self.driver.find_element(By.LINK_TEXT, "29").click()
        sleep(2)
        try:
            notify_before = self.driver.find_element(By.XPATH, "//span[contains(@class,'k-numeric-wrap k-state-default')]//input[1]")
            # notify_before = self.driver.find_element(By.XPATH, "//span[contains(@class,'k-widget k-numerictextbox')]")
            self.action.click(notify_before).send_keys(self.notify_before + Keys.TAB).perform()
            sleep(1)
            summary = self.driver.find_element(By.XPATH, "//label[text()=' Summary  ']/following::textarea")
            summary.clear()
            sleep(1)
            self.action.click(summary).send_keys(self.alert_summary + Keys.TAB + Keys.ENTER).perform()
            sleep(3)
        except NoSuchElementException or ElementNotInteractableException:
            if NoSuchElementException or ElementNotInteractableException:
                cancel_btn = self.driver.find_element(By.XPATH, "//button[@data-bind='click: cancel']")
                cancel_btn.click()
                sleep(2)

        """ Combobox """
        # combobox_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        # combobox_dd.click()
        # sleep(2)
        # # dd = self.driver.find_element(By.CLASS_NAME, "k-dropdown")
        # # dd.click()
        # # action = ActionChains(self.driver)
        # # action.move_to_element_with_offset(dd, 298, 35).click().send_keys(Keys.ARROW_DOWN + Keys.ENTER).perform()
        # # Select(dd).select_by_visible_text("Python")
        # # sleep(1)
        # ok_btn = self.driver.find_element(By.XPATH, "//button[text()='Ok']")
        # ok_btn.click()
        # sleep(3)

        """ Dynamic Combobox """
        # dynamic_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        # dynamic_dd.click()
        # sleep(2)
        # # dd = self.driver.find_element(By.XPATH, "/html/body/div[213]/div[2]/div[1]/span/span")
        # # dd.click()
        # # sleep(1)
        # # dd.send_keys(Keys.ARROW_DOWN * 4 + Keys.ENTER)
        # # sleep(1)
        # ok_btn = self.driver.find_element(By.XPATH, "//button[text()='Ok']")
        # ok_btn.click()
        # sleep(3)

        """ Mandatory field """
        self.name_box(self.cellno_mand_field)
        self.fx(self.mandatory_field)

        """ Manual date """
        self.name_box(self.cellno_date)
        try:
            manual_date_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
            manual_date_dd.click()
            cal_icon = self.driver.find_element(By.XPATH,
                                                "//span[contains(@class,'k-picker-wrap k-state-default')]//span")
            self.action.click(cal_icon)
            sleep(1)
            self.driver.find_element(By.LINK_TEXT, "24").click()
            sleep(2)
            ok_btn = self.driver.find_element(By.XPATH, "//button[text()='Ok']")
            ok_btn.click()
        except NoSuchElementException:
            print("Xpath of Link text is not working ")
            if NoSuchElementException:
                cal_input = self.driver.find_element(By.XPATH, "//input[@data-role='datepicker']")
                self.action.click(cal_input).send_keys(self.manual_date + Keys.TAB + Keys.ENTER).perform()
        sleep(3)

        """ Manual date & time """
        self.name_box(self.cellno_date_time)
        try:
            manual_date_time_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
            manual_date_time_dd.click()
            cal_time_input = self.driver.find_element(By.XPATH, "//input[@data-role='datetimepicker']")
            self.action.click(cal_time_input).send_keys(self.manual_date_time + Keys.TAB + Keys.ENTER).perform()
        except NoSuchElementException:
            pass
        sleep(3)

        """ Manual field """
        self.name_box(self.cellno_manual_field)
        self.fx(self.manual_field)

        """ Manual time """
        self.name_box(self.cellno_time)
        manual_time_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        manual_time_dd.click()
        sleep(2)
        time = self.driver.find_element(By.XPATH, "//input[@data-role='timepicker']")
        self.action.click(time).send_keys(self.manual_time + Keys.TAB + Keys.ENTER).perform()
        sleep(4)

        """ Multiselect Combobox """
        self.name_box(self.cellno_multi_combo)
        mul_sel_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        mul_sel_dd.click()
        sleep(2)
        dd = self.driver.find_element(By.CSS_SELECTOR, "div#id-selectBox>div")
        dd.click()
        chk_box = self.driver.find_element(By.XPATH, "//div[@id='checkboxes']//input")
        chk_box.click()
        chk_box = self.driver.find_element(By.XPATH, "(//div[@id='checkboxes']//input)[2]")
        chk_box.click()
        sleep(1)
        dd = self.driver.find_element(By.CSS_SELECTOR, "div#id-selectBox>div")
        self.action.click(dd).send_keys(Keys.TAB * 2 + Keys.ENTER).perform()
        sleep(3)

        """ Numeric field """
        self.name_box(self.cellno_numeric_field)
        self.fx(self.numeric_field)

        """ Text wrapper """
        self.name_box(self.cellno_text_wrap)
        self.fx(self.text_wrap)

        """ E-Sign """
        self.name_box(self.cellno_esign)
        esign_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        esign_dd.click()
        sleep(2)
        pwd = self.driver.find_element(By.ID, "focused-input")
        pwd.send_keys(self.esign_pwd)
        # reason = self.driver.find_element(By.XPATH, "//form[@id='SignatureForm']/div[1]/div[1]/div[1]/div[3]/div[2]/span[1]/span[1]/span[1]")
        # action.click(reason).send_keys(Keys.ARROW_DOWN + Keys.ENTER)
        sleep(1)
        cmts = self.driver.find_element(By.XPATH, "//textarea[@class='k-textbox ']")
        cmts.send_keys(self.esign_cmts + Keys.TAB + Keys.ENTER)
        sleep(2)
        fx = self.driver.find_element(By.XPATH, "//div[@class='k-spreadsheet-formula-bar']//div")
        self.action.click(fx).send_keys(Keys.ENTER).perform()
        sleep(3)

        self.hyperlink()
        self.attachments_inside_order()

    def hyperlink(self):
        self.name_box(self.cellno_hyper_link)
        hl_icon = self.driver.find_element(By.XPATH, "(//a[contains(@class,'k-button k-button-icon')]//span)[3]")
        hl_icon.click()
        sleep(2)
        address = self.driver.find_element(By.XPATH, "//div[@class='k-edit-field']//input[1]")
        address.send_keys(self.hyper_link)
        sleep(1)
        ok_btn = self.driver.find_element(By.XPATH, "//button[text()='OK']")
        ok_btn.click()
        sleep(3)

    def export_order(self):  # Export icon
        exp_icon = self.driver.find_element(By.XPATH, "(//button[@class='k-button k-button-icon']//span)[3]")
        exp_icon.click()
        sleep(1)
        filename = self.driver.find_element(By.XPATH, "//div[@class='k-edit-field']//input")
        filename.clear()
        filename.send_keys(self.sheet_ordr_id_export or self.sheet_order_id_manual)
        sleep(1)
        save_btn = self.driver.find_element(By.XPATH, "//button[text()='Save']")
        save_btn.click()
        sleep(5)

    def attachments_inside_order(self):  # Attachments icon
        attachments_btn = self.driver.find_element(By.XPATH, "//button[@title='Attachments']//i[1]")
        attachments_btn.click()
        sleep(2)
        file_input = self.driver.find_element(By.ID, "fileInputobj")
        file_path = self.attach_path  # File import from WINDOWS folder
        file_input.send_keys(file_path)
        sleep(3)
        self.driver.find_element(By.XPATH, "//div[@class='fp-well']//button[1]").click()
        WebDriverWait(self.driver, 50).until(
            ec.presence_of_element_located((By.XPATH, "//i[@class='fas fa-cloud-download-alt']")))
        close_btn = self.driver.find_element(By.XPATH, "//button[text()='CLOSE']")
        close_btn.click()
        sleep(3)

    def save_sheet_order(self):
        save_sheet_order = self.driver.find_element(By.CSS_SELECTOR, "button#SaveSheet>span")
        save_sheet_order.click()
        try:
            self.audit_cmts = "Sheet order saved by Automation"
            self.au_dit_reason = self.audit_reason[2]
            self.audit_trail()
        except ElementNotInteractableException:
            print("Audit trail is not enabled for save sheet order")
        finally:
            sleep(5)
        try:
            if "PLEASE ENTER THE MANDATORY FIELD" == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__"
                                                                                        "']//span[1]").text:
                print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)
                sleep(2)
                self.fx(self.mandatory_field)
                self.save_sheet_order()
        except NoSuchElementException:
            pass

    def order_workflow(self):

        for i in range(int(self.work_flow_steps) + 1):
            try:
                save_btn = self.driver.find_element(By.ID, "SaveSheet")
                save_btn.is_enabled()
                sleep(1)
            except NoSuchElementException:
                print("NoSuchElementException:", "Cannot able to find save button")
            else:
                if save_btn.is_enabled():
                    order_wf_name = self.driver.find_element(By.XPATH, "//label[contains(@class,'my-2 mr-2')]")
                    print(order_wf_name.text)
                    sleep(1)
                    self.so_submit()
                    self.e_sign(str(order_wf_name.text))
                else:
                    self.export_order()
                    com_task = self.driver.find_element(By.ID, "StopSampleLogin")
                    com_task.click()
                    sleep(3)
                    com_task_cmts = "complete task"
                    self.e_sign(com_task_cmts)
                    sleep(1)

                    """ Sheet orders close icon at corner """
                    self.driver.find_element(By.XPATH, "(//button[@class='btn_white_bg']//i)[2]").click()
                    break
            # sleep(6)

    def so_submit(self):
        so_submit_btn = self.driver.find_element(By.XPATH, "//div[@class='form-inline']//button[1]")
        so_submit_btn.click()
        sleep(2)

    def e_sign(self, esign_cmts=None):
        esign_pwd = self.driver.find_element(By.XPATH, "(//input[@name='Password'])[5]")
        esign_pwd.send_keys(self.esign_pwd)
        sleep(1)
        cmts = " successfully submitted"

        if esign_cmts != "complete task":
            esign_cmts_path = self.driver.find_element(By.CSS_SELECTOR, "textarea[type='text']")
            esign_cmts_path.send_keys(esign_cmts, cmts)
        else:
            esign_cmts_path = self.driver.find_element(By.CSS_SELECTOR, "textarea[type='text']")
            esign_cmts_path.send_keys(esign_cmts, cmts)
            print(esign_cmts)
        sleep(2)
        check_box = self.driver.find_element(By.CSS_SELECTOR, "input[value='0']")
        check_box.click()
        sleep(1)
        esign_submit_btn = self.driver.find_element(By.ID, "idSaveSheetAsComment")
        esign_submit_btn.click()
        sleep(6)


class ProtocolOrder(SheetOrder):

    def __init__(self):
        self.protocol_order_id = None
        self.protocol_order_id_manual = None
        self.protocol_order_id_export = None
        self.driver = None
        super().__init__()

        """ Register protocol order parameters """
        self.protocol_order_type = self.data['REGISTER PROTOCOL ORDER FIELDS']['protocol_order_type']
        self.project = self.data['REGISTER PROTOCOL ORDER FIELDS']['projects']
        self.task = self.data['REGISTER PROTOCOL ORDER FIELDS']['task']
        self.select_pro_temp = self.data['REGISTER PROTOCOL ORDER FIELDS']['select_protocol_template']
        self.sample = self.data['REGISTER PROTOCOL ORDER FIELDS']['sample']
        self.keyword = self.data['REGISTER PROTOCOL ORDER FIELDS']['keyword']

        self.protocol_order_id_manual = self.data['PROTOCOL ORDER ID MANUAL PROCESS']['protocol_order_id_manual']

        self.editor_attachments = self.data['EDITORS PROTOCOL ORDER']['ATTACHMENTS']['path']
        self.editor_image = self.data['EDITORS PROTOCOL ORDER']['IMAGE']['path']

    def orders_for_pro(self):
        try:
            self.driver.find_element(By.ID, "left-tabs-example-tab-SheetView").click()
            sleep(2)
            self.register_protocol_orders()
        except NoSuchElementException:
            print("NoSuchElementException:", "Can't find Orders button at left so global search will work")
            if NoSuchElementException:
                self.module_name = "Protocol Orders"
                self.search_module()

    def register_protocol_orders(self):
        self.driver.find_element(By.XPATH, "//a[@href='/Protocolorder']//span[1]").click()
        sleep(10)

    def order_eln_protocol(self):
        proto_order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        proto_order_type.send_keys(self.protocol_order_type[0] + Keys.ENTER)
        sleep(1)
        sel_project = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        sel_project.send_keys(self.project + Keys.ENTER)
        sleep(1)
        sel_task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        sel_task.send_keys(self.task + Keys.ENTER)
        sleep(1)

        try:
            if ("No Protocol templates have been mapped with this task".upper()
                    == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
                print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)
        except NoSuchElementException:
            pass

        sel_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        sel_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)

        if self.select_pro_temp != "":
            sel_pro_temp = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
            sel_pro_temp.send_keys("Default" + Keys.ENTER)
            sleep(1)

        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        super().save_within()
        super().register_order_btn()

    def order_dynamic_protocol(self):
        proto_order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        proto_order_type.send_keys(self.protocol_order_type[1] + Keys.ENTER)
        sleep(1)
        sel_project = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        sel_project.send_keys(self.project + Keys.ENTER)
        sleep(1)
        sel_task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        sel_task.send_keys(self.task + Keys.ENTER)
        sleep(1)
        sel_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        sel_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        sel_proto_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        sel_proto_template.send_keys(self.select_pro_temp + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

        if ("No Protocol selected.".upper()
                == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
            print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)
            sel_proto_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
            sel_proto_template.send_keys("Default template" + Keys.ENTER)
            sleep(1)
            super().register_order_btn()

    def process_order_protocol(self):  # process order grid button click

        if self.protocol_order_id_manual == "":
            so_id = self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]")
            print(so_id.text)
            split_id = str(so_id.text).split()
            self.protocol_order_id = split_id[0].split(":")[1]
            self.protocol_order_id_export = self.protocol_order_id
            print(self.protocol_order_id)
            order_grid_view = self.driver.find_element(By.XPATH, "//span[text()='" + self.protocol_order_id + "']")
            order_grid_view.click()
        else:
            order_grid_view = self.driver.find_element(By.XPATH,
                                                       "//span[text()='" + self.protocol_order_id_manual + "']")
            order_grid_view.click()
            sleep(1)
        process_order_btn = self.driver.find_element(By.XPATH, "(//span[@class='clsorderopenonlist'])[1]")
        process_order_btn.click()
        try:
            self.audit_cmts = "Process order protocol opened by Automation"
            self.audit_trail()
        except ElementNotInteractableException:
            print("Audit trail is not enabled for process_order_protocol")
        finally:
            sleep(10)

    def start_btn(self):
        start = self.driver.find_element(By.CSS_SELECTOR, "button#btn-protocol-start>span")
        start.click()
        sleep(2)

    def editors_protocol_order(self):

        """ Data type - Spreadsheet """
        self.data_fill_sheet_order()

        """ Data type - Attachments """
        file_input = self.driver.find_element(By.XPATH, "//input[@class='input-file']")
        file_path = self.editor_attachments
        file_input.send_keys(file_path)
        WebDriverWait(self.driver, 50).until(
            ec.presence_of_element_located((By.XPATH, "//i[@class='fas fa-cloud-download-alt']")))

        link_file = self.driver.find_element(By.XPATH, "(//input[@class='input-file'])[2]")
        link_file.click()
        sleep(2)
        pro_tab = self.driver.find_element(By.XPATH, "(//div[text()='Protocol'])[2]")
        pro_tab.click()
        sleep(3)
        folders = self.driver.find_element(By.XPATH, "(//span[text()='PKR Protocol'])[2]")
        folders.click()
        sleep(4)
        chk_box = WebDriverWait(self.driver, 50).until(
            ec.presence_of_element_located((By.XPATH, "(//label[text()='B'])[3]/following::input")))
        # chk_box = self.driver.find_element(By.XPATH, "(//label[text()='B'])[3]/following::input")
        self.action.click(chk_box).perform()
        sleep(1)
        insert_link_btn = self.driver.find_element(By.ID, "btn-linkimage-save")
        insert_link_btn.click()
        sleep(2)

        """ Images """
        file_input = self.driver.find_element(By.XPATH, "(//input[@class='input-file'])[3]")
        file_path = self.editor_image
        file_input.send_keys(file_path)
        WebDriverWait(self.driver, 50).until(
            ec.presence_of_element_located((By.XPATH, "//p[@class='card-text middle-textarea']//textarea[1]")))
        img_cmts = self.driver.find_element(By.XPATH, "//p[@class='card-text middle-textarea']//textarea[1]")
        img_cmts.send_keys("Hi")
        sleep(2)

        link_file = self.driver.find_element(By.XPATH, "(//input[@class='input-file'])[4]")
        link_file.click()
        sleep(2)
        pro_tab = self.driver.find_element(By.XPATH, "(//div[text()='Protocol'])[2]")
        pro_tab.click()
        sleep(3)
        folders = self.driver.find_element(By.XPATH, "(//span[text()='PKR Protocol'])[2]")
        folders.click()
        sleep(1)
        chk_box = self.driver.find_element(By.XPATH, "(//label[text()='B'])[3]/following::input")
        chk_box.click()
        sleep(1)
        insert_link_btn = self.driver.find_element(By.ID, "btn-linkimage-save")
        insert_link_btn.click()
        sleep(5)


if __name__ == '__main__':
    so = SheetOrder()
    so.login()
    so.orders()

    # so.home_folders()
    # so.new_folder()
    # so.rename_folder()
    # so.delete_folder()

    so.register_btn()
    so.order_sheet_template()
    so.process_order_sheet()
    # so.attachments_inside_order()
    # so.export_order()
    so.data_fill_sheet_order()
    so.save_sheet_order()
    so.order_workflow()

    # so.register_btn()
    # so.order_research_without_template()
    # so.process_order_sheet()

    # so.register_btn()
    # so.excel_order()
    # so.process_order_sheet()

    # so.register_btn()
    # so.sheet_validation()
    # so.process_order_sheet()

    po = ProtocolOrder()
    SheetOrder.login(po)
    po.orders_for_pro()

    # SheetOrder.register_btn(po)

    # po.order_eln_protocol()
    po.process_order_protocol()
    # po.start_btn()
    # SheetOrder.data_fill_sheet_order(po)
    po.editors_protocol_order()

    # po.order_dynamic_protocol()
    # po.process_order_protocol()
