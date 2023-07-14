from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from datetime import datetime
from openpyxl import load_workbook

# Get the day name
dayName = datetime.today().strftime('%A')

workbook = load_workbook('Excel.xlsx')
sheet = workbook[dayName]

driver = webdriver.Chrome()
driver.get("https://www.google.com/")
changeLanguage = driver.find_element(By.CSS_SELECTOR, "#SIvCob > a")
if changeLanguage.text == 'English':    
    changeLanguage.click()
time.sleep(2)

for i in range(10):
    keyword = sheet['C'+str(i+3)].value
    
    searchBar = driver.find_element(By.CSS_SELECTOR, "#APjFqb")
    searchBar.clear()
    searchBar.send_keys(keyword)

    time.sleep(2)

    suggestions = driver.find_elements(By.CSS_SELECTOR, '.erkvQe .wM6W7d span')

    non_empty_suggestions = []
    for suggestion in suggestions:
        text = suggestion.text.strip()
        if text != '':
            non_empty_suggestions.append(suggestion.text)

    if len(non_empty_suggestions) > 0:
        lengths = [len(suggestion) for suggestion in non_empty_suggestions]

        longest_index = lengths.index(max(lengths))
        shortest_index = lengths.index(min(lengths))

        longest_suggestion = non_empty_suggestions[longest_index]
        shortest_suggestion = non_empty_suggestions[shortest_index]

        sheet['D'+str(i+3)].value = longest_suggestion
        sheet['E'+str(i+3)].value = shortest_suggestion
    else:
        print('No non-empty suggestions found.')

workbook.save('your_file.xlsx')
workbook.close()