import openpyxl
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def validate_excel_file(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    if not file_path.endswith(".xlsx"):
        raise ValueError("Please provide a valid .xlsx Excel file.")


def setup_driver():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://www.google.com")
    return driver


def get_excel_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    today = datetime.today().strftime('%A')  
    if today in workbook.sheetnames:
        sheet = workbook[today]
        keywords = [
            sheet.cell(row=i, column=1).value
            for i in range(2, sheet.max_row + 1)
            if sheet.cell(row=i, column=1).value is not None  
        ]
        return keywords, sheet, workbook
    else:
        raise Exception(f"No sheet found for {today} in the Excel file.")


def search_keyword(driver, keyword):
    search_box = driver.find_element(By.NAME, "q")
    search_box.clear()
    search_box.send_keys(keyword)
    
    
    suggestions = driver.find_elements(By.CSS_SELECTOR, "ul[role='listbox'] li span")
    
   
    suggestions_list = [suggestion.text for suggestion in suggestions]

    
    longest = max(suggestions_list, key=len) if suggestions_list else None
    shortest = min(suggestions_list, key=len) if suggestions_list else None

    
    print(f"Searching for: {keyword}")
    print(f"Suggestions: {suggestions_list}")
    print(f"Longest: {longest}")
    print(f"Shortest: {shortest}")

    return longest, shortest


def write_results(sheet, workbook, file_path, results):
    for i, (keyword, longest, shortest) in enumerate(results, start=2):
        sheet.cell(row=i, column=2, value=longest)  
        sheet.cell(row=i, column=3, value=shortest)  

    workbook.save(file_path)
    print("Results saved to the Excel file!")


def main():
    file_path = "/home/soomit/Downloads/Admission Matirials/Job CV/4_beats/sample_keywords.xlsx" 
    validate_excel_file(file_path)
    
    
    keywords, sheet, workbook = get_excel_data(file_path)

    if not keywords:
        print("No keywords found for today's sheet.")
    
    
    user_keyword = input("Enter a keyword you want to search: ").strip()
    if user_keyword:
        keywords.append(user_keyword) 

   
    driver = setup_driver()

    results = []
    for keyword in keywords:
       
        longest, shortest = search_keyword(driver, keyword)
        results.append((keyword, longest, shortest))

       
        print(f"Keyword: {keyword}")
        print(f"Longest Suggestion: {longest}")
        print(f"Shortest Suggestion: {shortest}")

    
    write_results(sheet, workbook, file_path, results)

    
    driver.quit()

if __name__ == "__main__":
    main()

