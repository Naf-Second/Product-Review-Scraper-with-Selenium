from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os


def append_to_excel(filename, new_data, sheet_name='Sheet1'):
    try:
        existing_data = pd.read_excel(filename, sheet_name=sheet_name)
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    except FileNotFoundError:
        updated_data = new_data
    
    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        updated_data.to_excel(writer, index=False, sheet_name=sheet_name)


def uniquenesscheck(filename, new_data, sheet_name='Sheet1'):
    if os.path.exists(filename):
        # File exists, read the existing data
        existing_data = pd.read_excel(filename, sheet_name=sheet_name)
        if 'User ID' in existing_data.columns and not existing_data['User ID'].isnull().all():
            unique_user_ids = existing_data['User ID'].unique()
            unique_user_ids_list = unique_user_ids.tolist()
        else:
            unique_user_ids_list = []
    else:

        existing_data = pd.DataFrame(columns=new_data.columns)
        unique_user_ids_list = []


    if 'User ID' in new_data.columns and not new_data['User ID'].isnull().all():
        new_unique_user_ids = new_data['User ID'].unique()
        

        if any(user_id in unique_user_ids_list for user_id in new_unique_user_ids):
            return False

    return True

def scrape_page(driver):
    names = []
    ratingsall = []
    productids = []
    userid = []
    rev_title = []
    rev_description = []
    # Find the product names and ratings
    
    ratings = driver.find_elements(By.CSS_SELECTOR, "div.rating_image.left")
    products = driver.find_elements(By.CLASS_NAME, "head")
    product_id_elements = driver.find_elements(By.CSS_SELECTOR, "div.module.bb1.mb5.js-review.discussion_content")
    review = driver.find_elements(By.CSS_SELECTOR, "div.js-discussion-edit-toggle")
    user = url.split('/')
    if(user[3]=='profile'):
        user = user[4]
    else:
        user = user[3]
# scraping and appending product name and user name
    
    for product in products:
        names.append(product.text)
        userid.append(user)

# scraping and appending product id

    for result in product_id_elements:
        style = result.get_attribute('id')
        productids.append(style)

# scraping and appending user rating 

    for result in ratings:
        style = result.get_attribute('style')
        background_position = None
        if 'background-position' in style:
            styles = style.split(';')
            for s in styles:
                if 'background-position' in s:
                    background_position = s.split(':')[1].strip()
                    break
        if background_position:
            rating_in_number = ''
            if(len(background_position)>=10):
                
                rating_in_number+=background_position[5] 
                rating_in_number+=background_position[6]
                rating_in_number+=background_position[7]
            else:
                rating_in_number+=background_position[5] 
                rating_in_number+=background_position[6]

            rating_in_number = float(rating_in_number)
            if rating_in_number==250:
                ratingsall.append(5)
            elif rating_in_number==200:
                ratingsall.append(4)
            elif rating_in_number==150:
                ratingsall.append(3)    
            elif rating_in_number==100:
                ratingsall.append(2)
            elif rating_in_number==50:
                ratingsall.append(1)



    for title in review:
        main = title.find_element(By.CSS_SELECTOR, "p.fwb[itemprop='name']")
        rev_title.append(main.text)

    j = 0
    for description in review:
        
        for i in range(len(product_id_elements)):
            temp = product_id_elements[j].get_attribute('id')
            j+=1
            break
        xpath = "//*[@id='"+temp+"']/div[2]/div[2]/div/p[3]"
        main = description.find_element(By.XPATH, xpath)
        rev_description.append(main.text)       
                        

    data = {
        "User ID": userid,
        "Product Name": names,
        "Product ID": productids,
        "Product Rating": ratingsall,
        "User Review Title": rev_title,
        "Review Description": rev_description,
    }

    max_length = max(len(v) for v in data.values())

    for key in data:
        while len(data[key]) < max_length:
            data[key].append(None)

    return pd.DataFrame(data)

# Initialize Chrome webdriver
driver = webdriver.Chrome()

# Open the first page
url = "https://www.beautylish.com/alyssamc"

url+='/products_reviewed'
driver.get(url)


# Collect data from all pages
all_data = pd.DataFrame()

while True:
    time.sleep(2)
    new_data = scrape_page(driver)
    all_data = pd.concat([all_data, new_data], ignore_index=True)
    
    # Check for the 'Next' button and click it if it exists
    try:
        next_button = driver.find_element(By.CSS_SELECTOR, "span.pager_next a")
        next_button.click()
        
        # Wait for the new page to load
        time.sleep(5)
    except Exception as e:
        print("No more pages to scrape.")
        break

# Close the browser
driver.quit()

# Append the data to the Excel file
if(uniquenesscheck("E:/Scraping/beautylish_products2.xlsx", all_data)==True):
    append_to_excel("E:/Scraping/beautylish_products2.xlsx", all_data)
    print("Titles and ratings have been saved to beautylish_products.xlsx")
else:
    print("User Already Exists")


