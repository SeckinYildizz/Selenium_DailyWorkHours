import time
import pyperclip # type: ignore
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException


user_name = "" # Enter user name
password = "" # Enter password
employee_names = [] # Enter employee names as a list
start_date_year = 2024
start_date_month = 6
start_date_day = 1
the_date = datetime.date(start_date_year, start_date_month, start_date_day)
a_date = the_date
day_numbers = 3
report = [["", ], ]

for _ in range(day_numbers):
    a_day = a_date.strftime("%m/%d/%Y")
    report[0].append(a_day)
    a_date += datetime.timedelta(days=1)

for employee in employee_names:
    an_employee = [employee]
    report.append(an_employee)

# Opens Chrome to execute the code below
driver = webdriver.Chrome()
driver.maximize_window()
# Opens the given website
driver.get("") # I deleted the URL

try:
    # Step 1: Wait for the password input to appear. Then enter email on the first screen and click to the "next" button
    email_field = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "idUserName")))
    email_field.send_keys(user_name)

    # Step 2: Wait for the password input to appear. Enter the password into the password field and click to submit
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "idPassWord")))
    password_field = driver.find_element(By.ID, "idPassWord")
    password_field.send_keys(password)

    # Step 3: Click to login button
    log_in_button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH,
                                "/html/body[@id='body-login']/div[@id='idLogin']/input[@class='login-button']")))
    log_in_button.click()

    # Step 4: Click to menu button
    menu_button = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH,
                            "/html/body[@class='app header-fixed sidebar-fixed sidebar-hidden']/div[@id='root']/"
                            "div[@class='app']/div[@id='id-header']/button[@class='d-md-down-none navbar-toggler']/"
                            "i[@class='oi oi-menu']")))
    menu_button.click()

    # Step 6: Click to IDD operations button
    idd_operations_button = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH,
                            "/html/body/div/div/div[2]/div/nav/ul/li[3]/a")))
    idd_operations_button.click()

    # Step 5: Click to IDD Weekly Time Report button
    idd_weekly_rep_button = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH,
                            "/html/body/div/div/div[2]/div/nav/ul/li[3]/ul/li[9]/a")))
    idd_weekly_rep_button.click()

    # Step 6: Click to search button
    search_button = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH,
                            "/html/body/div/div/div[2]/main/div[3]/div/div/div/div/div/div/div/main/div[1]/button[1]")))
    search_button.click()

    # Step 7: Write a loop that will iterate the steps for all employees
    for emloyee in employee_names:
        # Step 7.1: Enter an employee name
        search_employee = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH,
                            "/html/body/div/div/div[2]/main/div[3]/div/div/div/div/div/div/div/nav/ul/span/div[2]/"
                            "form/div/div[1]/div/div/span[1]/div[2]/input")))
        search_employee.send_keys(Keys.BACKSPACE)
        search_employee.send_keys(emloyee)
        search_employee.send_keys(Keys.RETURN)
        
        a_date = the_date
        a_condition = True

        # Step 7.2: Write a loop to get the work hour for each day
        for _ in range(day_numbers):
            # Step 7.2.1: Clear the start day and then enter preferred start date
            search_sd = driver.find_element(By.XPATH,
                                "/html/body/div/div/div[2]/main/div[3]/div/div/div/div/div/div/div/nav/ul/span/div[2]/"
                                "form/div/div[4]/div/input")
            search_sd.clear()
            pyperclip.copy(a_date.strftime("%m/%d/%Y"))
            search_sd.send_keys(Keys.CONTROL, 'v')            

            # Step 7.2.2: Clear the end day and then enter the same date as start date
            search_ed = driver.find_element(By.XPATH,
                                "/html/body/div/div/div[2]/main/div[3]/div/div/div/div/div/div/div/nav/ul/span/div[2]/"
                                "form/div/div[5]/div/input")
            search_ed.clear()
            pyperclip.copy(a_date.strftime("%m/%d/%Y"))
            search_ed.send_keys(Keys.CONTROL, 'v') # On windows, send_keys method was not able to send the complete date. Hence, a different approach is tried.
    
            # Step 7.2.3: Click to execute the search
            search_execution = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "/html/"
                                    "body[@class='app header-fixed sidebar-fixed']/div[@id='root']/div[@class='app']/"
                                    "div[@class='app-body']/main[@class='main']/div[@class='container-fluid']/"
                                    "div[@id='page8522']/div/div[@class='row']/div[@class='mb-4 col']/"
                                    "div[@class='tab-content']/div[@class='tab-pane active']/div[@class='tab-grid mb-4']/"
                                    "main[@id='main']/div[@class='pb-1 card-header']/"
                                    "button[@class='tlb-button ml-1 btn btn-secondary'][2]")))
            search_execution.click() 
            
            # Step 7.2.3.1: Wait till the table occurs
            if a_condition:
                time.sleep(5)
                a_condition = False
            else:
                time.sleep(2)

            # Step 7.2.4: Get the work hour for the day and add it to a list
            work_hours_element = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div/div[2]/main/div[3]/div/div/div/div/div/div/div/main/div[2]/div[1]/table/tbody/tr/td[19]")))
            work_hours_text = work_hours_element.text 
    
            if len(work_hours_text) == 0:
                work_hours_total = "Empty"
            else:    
                if "hour" in work_hours_text:
                    work_hours = int(work_hours_text[:2].strip())
                else:
                    work_hours = 0    

                # Apply rounding logic
                if "minute" in work_hours_text:
                    work_minutes = int(work_hours_text[work_hours_text.find("min") - 3 : work_hours_text.find("min") - 1].strip())
                
                    if work_minutes <= 7:
                        work_minutes = 0
                    elif work_minutes <= 22:
                        work_minutes = .25
                    elif work_minutes <= 37:
                        work_minutes = .5
                    elif work_minutes <= 52:
                        work_minutes = .75
                    else:
                        work_minutes = 1
                
                else:
                    work_minutes = 0

                work_hours_total = work_hours + work_minutes
                print(work_hours_total)
               
            # Step 7.2.5: Add the work hours to the corresponding part of the report
            report[employee_names.index(emloyee) + 1].append(work_hours_total)

            a_date += datetime.timedelta(days=1)

    output_report = pd.DataFrame(report)
    output_report.to_excel(f"//Downloads/WeeklyHoursReport{the_date}-{the_date + datetime.timedelta(days=day_numbers-1)}.xlsx", index=False, header=False)
    print(output_report)

except WebDriverException as e:
    print("An error occurred while initializing WebDriver:", e)

finally:
    driver.close()
