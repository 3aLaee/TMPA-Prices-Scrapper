import sys
import datetime

# Save the original stdout
original_stdout = sys.stdout

try:
    # Open a file for redirecting output
    with open('terminal_output.txt', 'a') as output_file:
        # Redirect stdout to the file
        sys.stdout = output_file
        print(f"le: {datetime.datetime.now()}")
        

        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import NoSuchElementException
        import time
        import datetime
        import pandas as pd 
        import os
        import csv
        import openpyxl

        # Create instance of Chrome webdriver
        driver = webdriver.Chrome()
        driver.get("https://beta.clickferry.com/fr")

        try:
            # Wait for the page to load
            time.sleep(4)
            today = datetime.date.today()

            # Handle cookies pop-up if it exists
            try:
                cookies_button = driver.find_element(By.XPATH, '//*[@id="cookiescript_accept"]')
                cookies_button.click()
            except:
                pass

            # Click the search button
            search_button = driver.find_element(By.XPATH, '//*[@id="searcher"]/div[1]/button[1]')
            search_button.click()

            # Click the button to open the first field
            first_field_button = driver.find_element(By.XPATH, '//*[@id="route-portal"]/div/div/div[1]')
            first_field_button.click()

            # Enter 'Tanger' in the first field's input
            first_field_input = driver.find_element(By.XPATH, '//*[@id="route-portal"]/div[2]/div[2]/input')
            first_field_input.clear()
            first_field_input.send_keys('Ceuta')
            time.sleep(2)

            # Click on the first option in the dropdown
            first_dropdown_option = driver.find_element(By.XPATH, '//*[@id="route-portal"]/div[2]/ul/li[1]')
            first_dropdown_option.click()
            time.sleep(3)

            # Scroll down using JavaScript
            scroll_amount = 300
            scroll_script = f"window.scrollBy(0, {scroll_amount});"
            driver.execute_script(scroll_script)
            time.sleep(3)

            # Date button
            div_element = driver.find_element(By.CLASS_NAME, "react-datepicker__day--today")
            time.sleep(2)
            div_element.click()
            
            # Vehicule button
            veh_button = driver.find_element(By.XPATH, '//*[@id="vehicle-portal"]/div/div')
            time.sleep(1)
            veh_button.click()
            sec_button = driver.find_element(By.XPATH, '//*[@id="vehicle-portal"]/div[2]/ul/li[2]')
            sec_button.click()
            first_input = driver.find_element(By.XPATH, '//*[@id="vehicle-portal"]/div[2]/div[2]/div[1]/div[1]/input')
            first_input.clear()
            first_input.send_keys('Audi A4')
            time.sleep(2)
            first_option = driver.find_element(By.XPATH, '//*[@id="vehicle-portal"]/div[2]/div[2]/div[1]/div[2]/ul/li[1]')
            first_option.click()
            confirm_button = driver.find_element(By.XPATH, '//*[@id="vehicle-portal"]/div[2]/button')
            confirm_button.click()
            
            

            submit_button = driver.find_element(By.XPATH, '//*[@id="searcher"]/button')
            submit_button.click()
            time.sleep(5)

           

              # Find all company name elements with prices
            company_elements = driver.find_elements(By.CSS_SELECTOR, '.SupplierFilter_option___Hyoa')
            
            # Collect the terminal output
            terminal_output = []
            trajet_element = driver.find_element(By.CSS_SELECTOR, '.SearcherBox_content__BTSJK')
            trajet = trajet_element.text
            print(f"Trajet: {trajet}")
            
            for company_element in company_elements:
                try:
                    # Extract company name
                    company_name = company_element.text
                    

                    # Locate the adjacent double element
                    double_element = company_element.find_element(By.CSS_SELECTOR, '.SupplierFilter_price__xHlJY')

                    # Extract the double value
                    double = double_element.text
                    
                     # Find parent element and then find the trajet element
                    terminal_output.append([f"For: {company_name}", double, f"Trajet: {trajet}"])
                    print(f"Company: {company_name}")

                except NoSuchElementException:
                    
                    
                    print(f"Company: {company_name}")
                    
                    

            
            

        finally:
            # Restore the original stdout
            sys.stdout = original_stdout
            time.sleep(3)
            driver.quit()
            input_file = 'terminal_output.txt'
            output_file = 'rebo.xlsx'

            wb = openpyxl.Workbook()
            ws = wb.worksheets[0]

            with open(input_file, 'r') as data:
                 reader = csv.reader(data, delimiter='\t')
                 for row in reader:
                     ws.append(row)

            wb.save(output_file)

except Exception as e:
    # In case of any exceptions, print the error and traceback
    print(f"An error occurred: {str(e)}")
    import traceback
    traceback.print_exc()