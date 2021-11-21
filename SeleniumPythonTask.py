from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import Workbook

# Creating driver and opening a web page
service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.maximize_window()
driver.get("https://rpa.hybrydoweit.pl/")

# Finding "Artykuły" section and moving to it
articles_page_button = driver.find_element(By.LINK_TEXT, "Artykuły").click()

# Creating empty lists to collect values and getting number of articles on the page
all_titles_listed = []
all_sectors_listed = []
all_links_listed = []
number = 1
total_number_of_articles = len(driver.find_elements(By.CSS_SELECTOR, "[class='rpa-article-card']"))

# Loop to collect information about articles titles, sectors and links
while number >= 1 and number <= total_number_of_articles:
    all_titles = driver.find_elements(By.CSS_SELECTOR,
                                      f"#articles > div > div > div:nth-child({number}) > article > div > h3")
    all_sectors = driver.find_elements(By.CSS_SELECTOR,
                                       f"#articles > div > div > div:nth-child({number}) > article > div > ul > li")
    all_links = driver.find_elements(By.CSS_SELECTOR, f"#articles > div > div > div:nth-child({number}) > article > a")
    for title in all_titles:
        for sector in all_sectors:
            for link in all_links:
                all_titles_listed.append(title.text)
                all_sectors_listed.append(sector.text)
                all_links_listed.append(link.get_attribute("href"))
    number += 1

# Creating new list of sectors to delete "Branża" and "Dział" words
all_sectors_listed_short = []
for sector_name in all_sectors_listed:
    all_sectors_listed_short.append(sector_name.replace("Branża: ", "").replace("Dział: ", ""))

# Creating list containing information about latest and oldest five articles on the page
articles_indexes = [0, 1, 2, 3, 4, -5, -4, -3, -2, -1]
selected_articles_data = []
for index in articles_indexes:
    selected_articles_data.append([all_titles_listed[index], all_sectors_listed_short[index], all_links_listed[index]][:])

# Creating Excel file
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "List of articles"

# Adding headers to columns
worksheet.append(["Tytuł", "Branża / Dział", "Link"])

# Adding data to the file - information about latest and oldest five articles
for selected_article in selected_articles_data:
    worksheet.append(selected_article)

# Changing columns width in the Excel file
worksheet.column_dimensions['A'].width = 70
worksheet.column_dimensions['B'].width = 15
worksheet.column_dimensions['C'].width = 100

# Saving the Excel file with data
workbook.save("List_of_articles.xlsx")

# Closing the driver
driver.close()
