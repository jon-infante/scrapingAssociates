from selenium import webdriver
from selenium.webdriver import Keys, ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os


EMAIL = "jonathan.infante@revature.com"
PWD = "Kooljonkids1!"

driver = webdriver.Edge('C:\Program Files (x86)/Microsoft/Edge/Application/msedgedriver')
driver.get("https://app.revature.com/caliber/home")

initial_batch = "00001228"


class topQCScores():
    def __init__(self):
        self.names = []
        self.qc_scores = []
        self.trainer_scores = []
        self.trainers = []
        self.iter = 0
        self.cohortFound = False

    def main(self):
        self.wake_up()
        self.searching(initial_batch)
        row_values = [self.names, self.qc_scores, self.trainer_scores, self.trainers]
        self.associatesToDataFrame(row_values)

    def wake_up(self):
        username = WebDriverWait(driver, timeout=10).until(lambda d: d.find_element("id","loginForm:userName-input-id"))
        # username = driver.find_element("id","loginForm:userName-input-id")
        username.clear()
        username.send_keys(EMAIL)

        password = driver.find_element("id","loginForm:input-psw")
        password.clear()
        password.send_keys(PWD)

        driver.find_element("id", "loginForm:login-btn-id").click()

        x = WebDriverWait(driver, timeout=10).until(lambda d: d.find_element("id","batDashDiscourse"))
        caliber = driver.find_element("xpath", ("//*[contains(text(), 'Caliber')]"))
        caliber.click()

        time.sleep(1)
        driver.get("https://app.revature.com/caliber/reports")

        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)
        radio_button = driver.find_element("id","radioPrevious").click()

    def parseBatch(self, batchid):
        nextt = False
        search_element_path = "/html[1]/body[1]/app-root[1]/div[1]/app-reports-container[1]/div[1]/app-reports-toolbar[1]/div[2]/div[1]/div[2]/button[1]"
        search_ele = driver.find_element("xpath", search_element_path)
        if self.iter > 0:
            time.sleep(1)
            search_ele.click()
                
            input_search_bar_path = "/html[1]/body[1]/app-root[1]/div[1]/app-reports-container[1]/div[1]/app-reports-toolbar[1]/div[2]/div[1]/div[2]/div[1]/div[1]/input[1]"
            search = driver.find_element("xpath", input_search_bar_path)
            try:
                search.clear()
                search.send_keys(batchid)
            except:
                pass
            try:
                batch = WebDriverWait(driver, timeout=1).until(lambda d: d.find_element("xpath", (f"//*[contains(text(), {batchid})]")))
                batch.click()
            except:
                self.iter += 1
                new_batch_num = str((int(batchid) + 1))
                batchid = new_batch_num.zfill(8)
                print(batchid)
                nextt = True
        associate_names = []
        associate_qc_scores = []
        associate_trainer_scores = []
        associate_trainers = []

        if not nextt:
            try:
                tbody = WebDriverWait(driver, timeout=10).until(lambda d: d.find_elements(By.CSS_SELECTOR, "tbody"))
            except:
                self.iter += 1
                new_batch_num = str((int(batchid) + 1))
                batchid = new_batch_num.zfill(8)
                self.parseBatch(batchid)

            a = ActionChains(driver)

            tbody_path = "/html[1]/body[1]/app-root[1]/div[1]/app-reports-container[1]/div[2]/app-qc-scores[1]/div[1]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]"
            associatesRemaining = True
            trainer = driver.find_element("xpath", (f"//*[contains(text(), 'Trainer & QC Scores')]"))
            a.move_to_element(trainer).perform()
            trainer = trainer.text

            i = 1
            while (associatesRemaining):
                try:
                    l_name = WebDriverWait(driver, timeout=5).until(lambda d: d.find_element("xpath", f'{tbody_path}/tr[{i}]/td[1]'))
                    a.move_to_element(l_name).perform()
                    a.key_down(Keys.DOWN).send_keys("abc").perform()
                    f_name = WebDriverWait(driver, timeout=5).until(lambda d: d.find_element("xpath", f'{tbody_path}/tr[{i}]/td[2]'))
                    trainer_score = WebDriverWait(driver, timeout=5).until(lambda d: d.find_element("xpath", f'{tbody_path}/tr[{i}]/td[3]/div[1]/div[2]'))
                    qc_score = WebDriverWait(driver, timeout=5).until(lambda d: d.find_element("xpath", f'{tbody_path}/tr[{i}]/td[4]/div[1]/div[2]'))
                    associate_names.append(f_name.text + " " + l_name.text)
                    
                    if qc_score == '': qc_score = "0"
                    if trainer_score == '': trainer_score = "0"
                    
                    associate_trainer_scores.append(trainer_score.text.strip("()"))
                    associate_qc_scores.append(qc_score.text.strip("()"))
                except:
                    associatesRemaining = False
                    print("We have no more names to iterate over")

                i += 1
            

            self.names += associate_names
            self.qc_scores += associate_qc_scores
            self.trainer_scores += associate_trainer_scores
            associate_trainers = [trainer]*(i-2)
            self.trainers += associate_trainers
            print(i)
            print(associate_names)
            print(associate_trainer_scores)
            print(associate_qc_scores)
            print(associate_trainers)
            time.sleep(0.5)
            self.iter += 1


    def searching(self, batchid):
        input_search_bar_path = "/html[1]/body[1]/app-root[1]/div[1]/app-reports-container[1]/div[1]/app-reports-toolbar[1]/div[2]/div[1]/div[2]/div[1]/div[1]/input[1]"
        try: 
            driver.find_element("id", "reportBatchesDropDown").click()
            search = driver.find_element("xpath", input_search_bar_path)
            search.clear()
            search.send_keys(batchid)
            batch = WebDriverWait(driver, timeout=1).until(lambda d: d.find_element("xpath", (f"//*[contains(text(), {batchid})]")))
            batch.click()
        except:
            new_batch_num = str((int(batchid) + 1))
            batchid = new_batch_num.zfill(8)
            self.searching(batchid)
            
        while int(batchid) < 1600:
            self.parseBatch(batchid)
            new_batch_num = str((int(batchid) + 1))
            batchid = new_batch_num.zfill(8)

    def associatesToDataFrame(self, row_values):
        new_row_values = []
        for i in range(len(row_values[0])-1):
            new_row_values.append([row_values[0][i], row_values[1][i], row_values[2][i], row_values[3][i]])


        column_names = ["Name", "QC Score", "Trainer Score", "Trainer/Tech Stack"]
        df = pd.DataFrame(new_row_values, columns=column_names, dtype=object) #Creating the Data Frame

        self.exportToExcel(df)   

    def exportToExcel(self, df):
        df = df.to_excel("output/associateReport.xlsx")   

if __name__ == "__main__":
    qc = topQCScores()
    qc.main()
