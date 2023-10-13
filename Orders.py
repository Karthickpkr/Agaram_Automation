import json
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep


class SheetOrder:

    def __init__(self):
        self.order_id = None
        self.order_id_manual = None
        self.export_order_id = None
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

        """ To open data1.json and read the data as jsonfield,
            load the jsonfield data into self.data

            also can get any data values from the JSON field by using like self.data['FIELD NAME']
        """

        with open("Fill_me.json", 'r') as jsonfield:
            self.data = json.load(jsonfield)

        """ Example: self.data['LOGIN']['username'] is assigned to self.username parameters """

        self.login_url = self.data['LOGIN']['url']
        self.username = self.data['LOGIN']['username']
        self.password = self.data['LOGIN']['password']
        self.folder_name = self.data['FOLDER CREATION AND RENAME']['folder_name']
        self.rename_folder = self.data['FOLDER CREATION AND RENAME']['rename_folder']

        """ Register order parameters """
        self.order_type = self.data['ORDER SHEET TEMPLATE']['order_type']
        self.projects = self.data['ORDER SHEET TEMPLATE']['projects']
        self.task = self.data['ORDER SHEET TEMPLATE']['task']
        self.sample = self.data['ORDER SHEET TEMPLATE']['sample']
        self.keyword = self.data['ORDER SHEET TEMPLATE']['keyword']

        self.order_id_manual = self.data['MANUAL ORDER ID PROCESS']['order_id_manual']

        """ Data fill sheet order parameters """
        self.notify_before = self.data['GENERAL FIELDS SHEET ORDER']['ALERT']['notify_before_days']
        self.alert_summary = self.data['GENERAL FIELDS SHEET ORDER']['ALERT']['alert_summary']
        self.mandatory_field = self.data['GENERAL FIELDS SHEET ORDER']['MANDATORY FIELD']['mandatory_value']
        self.manual_date_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL DATA AND TIME']['manual_date_time']
        self.manual_field = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL FIELD']['manual_field']
        self.manual_time = self.data['GENERAL FIELDS SHEET ORDER']['MANUAL TIME']['manual_time']
        self.numeric_field = self.data['GENERAL FIELDS SHEET ORDER']['NUMERIC FIELD']['numeric_field']
        self.text_wrap = self.data['GENERAL FIELDS SHEET ORDER']['TEXT WRAPPER']['text_wrap']
        self.esign_pwd = self.data['GENERAL FIELDS SHEET ORDER']['E-SIGN']['esign_pwd']
        self.esign_cmts = self.data['GENERAL FIELDS SHEET ORDER']['E-SIGN']['esign_cmts']

        """ Icons """
        self.hyper_link = self.data['ICONS']['HYPERLINK']['hyper_link']
        self.attach_path = self.data['ICONS']['ATTACHMENTS']['attach_path']

        """ Order workflow steps"""
        self.work_flow_steps = self.data['ORDER WORKFLOW']['work_flow_steps']

    def login(self):
        self.driver.get(self.login_url)
        uname = self.driver.find_element(By.ID, "idUsername")
        uname.send_keys(self.username)
        uname.send_keys(Keys.TAB)
        pwd = self.driver.find_element(By.ID, 'idPassword')
        pwd.send_keys(self.password)
        sleep(2)
        pwd.send_keys(Keys.ENTER)
        sleep(12)
        print(self.driver.title)

    def orders(self):
        # ele = WebDriverWait(self.driver, 50)
        # ele.until(ec.visibility_of_element_located((By.XPATH,
        #                                             "//div[@id='bodycontent']/div[5]/div[3]/div[3]/div[1]/div[2]")))
        self.driver.find_element(By.ID, "left-tabs-example-tab-SheetView").click()
        sleep(2)

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
        folder_name.send_keys(self.folder_name)
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

        if ("FOLDER NAME ALREADY EXISTS"
                == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text):
            labexistcount = self.folder_name + '_' + str(1)
            sleep(1)
            new_folder.click()
            sleep(1)
            folder_name = self.driver.find_element(By.XPATH, "//div[@class='form-group']//input")
            folder_name.send_keys(labexistcount)
            sleep(2)
            add_btn = self.driver.find_element(By.CSS_SELECTOR, "button[value='add']")
            add_btn.click()
            sleep(2)

    def rename_folder(self):
        action = ActionChains(self.driver)
        select_folder = self.driver.find_element(By.XPATH, "(//span[text()='hoi'])[2]")
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
        rename.send_keys(self.rename_folder)
        sleep(1)
        rename_btn = self.driver.find_element(
            By.XPATH, "//div[@class='k-dialog-buttongroup k-dialog-button-layout-stretched']//button[1]")
        rename_btn.click()
        sleep(2)

    def delete_folder(self):
        action = ActionChains(self.driver)
        select_folder = self.driver.find_element(By.XPATH, "(//span[text()='ELN'])[2]")
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
        order_type = self.driver.find_element(By.XPATH, value="//span[@class='k-searchbar']//input")
        order_type.send_keys(self.order_type + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys(self.projects + Keys.ENTER)
        sleep(1)
        task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        task.send_keys(self.task + Keys.ENTER)
        sleep(1)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        # select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        # select_sheet_template.send_keys("Default" + Keys.ENTER)
        # sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys(self.keyword)
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def order_research_without_template(self):
        order_type = self.driver.find_element(By.XPATH, value="//span[@class='k-searchbar']//input")
        order_type.send_keys("Res" + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys("Order without template" + Keys.ENTER)
        sleep(1)
        task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        task.send_keys("Automation 1" + Keys.ENTER)
        sleep(2)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys("Pyt" + Keys.ENTER)
        sleep(1)
        select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        select_sheet_template.send_keys("Default" + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys("Automation Test 2")
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def excel_order(self):
        order_type = self.driver.find_element(by="xpath", value="//span[@class='k-searchbar']//input")
        order_type.send_keys("Excel" + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys("Excel order" + Keys.ENTER)
        sleep(2)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        select_sample.send_keys("Pyt" + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys("Automation Test 3")
        sleep(1)
        self.save_within()
        self.register_order_btn()

    def sheet_validation(self):
        order_type = self.driver.find_element(by="xpath", value="//span[@class='k-searchbar']//input")
        order_type.send_keys("sheet" + Keys.ENTER)
        sleep(1)
        projects = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        projects.send_keys("Sheet validation" + Keys.ENTER)
        sleep(1)
        keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        keyword.send_keys("Automation Test 4")
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
        sleep(5)

    def process_order(self):  # process order grid button click

        if self.order_id_manual == "":
            so_id = self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]")
            print(so_id.text)
            split_id = str(so_id.text).split()
            self.order_id = split_id[0].split(":")[1]
            self.export_order_id = self.order_id
            print(self.order_id)
            order_grid_view = self.driver.find_element(By.XPATH, "//span[text()='" + self.order_id + "']")
            order_grid_view.click()
        else:
            order_grid_view = self.driver.find_element(By.XPATH, "//span[text()='" + self.order_id_manual + "']")
            order_grid_view.click()
            sleep(1)
        process_order = self.driver.find_element(By.XPATH, "(//span[@class='clsorderopenonlist'])[1]")
        process_order.click()
        sleep(10)

    def name_box(self, cellno=None):
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

        action = ActionChains(self.driver)

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
        so.name_box("A" + str(5))
        alert_dd = self.driver.find_element(By.XPATH, "//span[contains(@class,'k-icon k-icon')]")
        alert_dd.click()
        sleep(2)
        # due_date_icon = self.driver.find_element(By.CSS_SELECTOR, ".k-icon.k-i-calendar[shub-ins='1']")
        due_date_icon = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-i-calendar']")
        due_date_icon.click()
        sleep(1)
        self.driver.find_element(By.LINK_TEXT, "29").click()
        sleep(2)
        notify_before = self.driver.find_element(By.XPATH,
                                                 "//span[contains(@class,'k-numeric-wrap k-state-default')]//input[1]")
        notify_before.send_keys(self.notify_before + Keys.TAB)
        sleep(1)
        summary = self.driver.find_element(By.CSS_SELECTOR, ".k-textbox[data-bind='value: Summary']")
        summary.clear()
        sleep(1)
        summary.send_keys(self.alert_summary + Keys.TAB + Keys.ENTER)
        sleep(3)

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
        so.name_box("A" + str(17))
        so.fx(self.mandatory_field)

        """ Manual date """
        so.name_box("A" + str(19))
        manual_date_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        manual_date_dd.click()
        caldr = self.driver.find_element(By.XPATH, "//span[contains(@class,'k-picker-wrap k-state-default')]//span")
        caldr.click()
        sleep(1)
        self.driver.find_element(By.LINK_TEXT, "29").click()
        sleep(2)
        ok_btn = self.driver.find_element(By.XPATH, "//button[text()='Ok']")
        ok_btn.click()
        sleep(3)

        """ Manual date & time """
        so.name_box("A" + str(21))
        manual_date_time_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        manual_date_time_dd.click()
        calndr = self.driver.find_element(By.XPATH, "//input[@data-role='datetimepicker']")
        action.click(calndr).send_keys(self.manual_date_time + Keys.TAB + Keys.ENTER).perform()
        sleep(3)

        """ Manual field """
        so.name_box("A" + str(23))
        so.fx(self.manual_field)

        """ Manual time """
        so.name_box("A" + str(25))
        manual_time_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        manual_time_dd.click()
        sleep(2)
        time = self.driver.find_element(By.XPATH, "//input[@data-role='timepicker']")
        action.click(time).send_keys(self.manual_time + Keys.TAB + Keys.ENTER).perform()
        sleep(4)

        """ Multiselect Combobox """
        so.name_box("A" + str(27))
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
        action.click(dd).send_keys(Keys.TAB * 2 + Keys.ENTER).perform()
        sleep(3)

        """ Numeric field """
        so.name_box("A" + str(29))
        so.fx(self.numeric_field)

        """ Text wrapper """
        so.name_box("A" + str(41))
        so.fx(self.text_wrap)

        """ E-Sign """
        so.name_box("A" + str(37))
        esign_dd = self.driver.find_element(By.XPATH, "//span[@class='k-icon k-icon k-i-arrow-60-down']")
        esign_dd.click()
        sleep(2)
        pwd = self.driver.find_element(By.ID, "focused-input")
        pwd.send_keys(self.esign_pwd)
        # reason = self.driver.find_element(By.XPATH, "//form[@id='SignatureForm']/div[1]/div[1]/div[1]/div[3]/div[2]/span[1]/span[1]/span[1]")
        # reason.click()
        # reason.send_keys(Keys.ARROW_DOWN + Keys.ENTER)
        sleep(1)
        cmts = self.driver.find_element(By.XPATH, "//textarea[@class='k-textbox ']")
        cmts.send_keys(self.esign_cmts + Keys.TAB + Keys.ENTER)
        sleep(2)
        fx = self.driver.find_element(By.XPATH, "//div[@class='k-spreadsheet-formula-bar']//div")
        action.click(fx).send_keys(Keys.ENTER).perform()
        sleep(2)

        self.hyperlink()
        self.attachments_inside_order()

    def hyperlink(self):
        so.name_box("C" + str(21))
        hl_icon = self.driver.find_element(By.XPATH, "(//a[contains(@class,'k-button k-button-icon')]//span)[3]")
        hl_icon.click()
        address = self.driver.find_element(By.XPATH, "//div[@class='k-edit-field']//input[1]")
        address.send_keys(self.hyper_link)
        sleep(1)
        ok_btn = self.driver.find_element(By.XPATH, "//button[text()='OK']")
        ok_btn.click()
        sleep(3)

    def export_order(self):
        exp_icon = self.driver.find_element(By.XPATH, "(//button[@class='k-button k-button-icon']//span)[3]")
        exp_icon.click()
        sleep(1)
        filename = self.driver.find_element(By.XPATH, "//div[@class='k-edit-field']//input")
        filename.clear()
        filename.send_keys(self.export_order_id or self.order_id_manual)
        sleep(1)
        save_btn = self.driver.find_element(By.XPATH, "//button[text()='Save']")
        save_btn.click()
        sleep(5)

    def attachments_inside_order(self):  # Attachments inside order
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
        sleep(5)

        if "PLEASE ENTER THE MANDATORY FIELD" == self.driver.find_element(By.XPATH, "//div[@id='__react-alert__"
                                                                                    "']//span[1]").text:
            print(self.driver.find_element(By.XPATH, "//div[@id='__react-alert__']//span[1]").text)
            sleep(2)
            so.fx(self.mandatory_field)
            so.save_sheet_order()

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

    def __init__(
            self,
            proordertype,
            project,
            task,
            sample,
            keyword,
            dyaordertype,
    ):
        self.proordertype = None
        self.driver = None
        super().__init__()
        self.proordertype = proordertype
        self.project = project
        self.task = task
        self.sample = sample
        self.keyword = keyword
        self.dyaordertype = dyaordertype

    def register_protocol_orders(self):
        self.driver.find_element(by="xpath", value="//a[@href='/registertask']//span[2]").click()
        sleep(10)

    def order_eln_protocol(self):
        pro_order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        pro_order_type.sendkeys(self.proordertype + Keys.ENTER)
        select_project = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        select_project.send_keys(self.project + Keys.ENTER)
        select_task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        select_task.send_keys(self.task + Keys.ENTER)
        sleep(1)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        # select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        # select_sheet_template.send_keys("Default" + Keys.ENTER)
        # sleep(1)
        enter_keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        enter_keyword.send_keys(self.keyword)
        sleep(1)
        super().save_within()
        super().register_order_btn()

    def order_dynamic_protocol(self):
        pro_order_type = self.driver.find_element(By.XPATH, "//span[@class='k-searchbar']//input")
        pro_order_type.sendkeys(self.dyaordertype + Keys.ENTER)
        select_project = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[2]")
        select_project.send_keys(self.project + Keys.ENTER)
        select_task = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[3]")
        select_task.send_keys(self.task + Keys.ENTER)
        sleep(1)
        select_sample = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[4]")
        select_sample.send_keys(self.sample + Keys.ENTER)
        sleep(1)
        # select_sheet_template = self.driver.find_element(By.XPATH, "(//span[@class='k-searchbar']//input)[5]")
        # select_sheet_template.send_keys("Default" + Keys.ENTER)
        # sleep(1)
        enter_keyword = self.driver.find_element(By.XPATH, "//input[@datatype='Text']")
        enter_keyword.send_keys(self.keyword)
        sleep(1)
        super().save_within()
        super().register_order_btn()

    def main(self):
        self.proordertype = ["ELN"]
        self.project = ["ELN"]
        self.task = ["Automation 3"]
        self.sample = ["selenium"]
        self.keyword = ["Automation"]


if __name__ == '__main__':
    so = SheetOrder()
    so.login()
    so.orders()
    so.register_sheet_orders()

    # so.home_folders()
    # so.new_folder()
    # so.rename_folder()
    # so.delete_folder()

    so.register_btn()
    so.order_sheet_template()

    so.process_order()
    # so.attachments_inside_order()
    # so.export_order()
    so.data_fill_sheet_order()
    so.save_sheet_order()
    so.order_workflow()

    # so.register_btn()
    # so.order_research_without_template()
    # so.process_order()

    # so.register_btn()
    # so.excel_order()
    # so.process_order()

    # so.register_btn()
    # so.sheet_validation()
    # so.process_order()

    # po = ProtocolOrder()
    # po.login()
    # po.register_protocol_orders()
    # po.register_btn()
    # po.order_eln_protocol()
