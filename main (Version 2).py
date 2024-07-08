import os
import math
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC



def append_to_excel(filename, new_data):
    try:
        existing_data = pd.read_excel(filename)
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    except FileNotFoundError:
        updated_data = new_data
    
    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        updated_data.to_excel(writer, index=False)


def uniquenesscheck(filename, new_data):
    if os.path.exists(filename):
        # File exists, read the existing data
        existing_data = pd.read_excel(filename)
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

def scrape_review_links(url):

    urls_array = []

    try:
        driver.get(url)

        wait = WebDriverWait(driver, 2)

        # Scroll down to the end of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)  

        # Wait for the review section to be loaded
        reviews_section = wait.until(EC.presence_of_element_located((By.ID, 'ReviewsAccordion')))

        # Find the review elements
        rev = driver.find_elements(By.CLASS_NAME, "pageWithFooter")
        for loop1 in rev:
            size = loop1.find_element(By.CLASS_NAME, "MuiTypography-root.jss2043.MuiTypography-body1")

            try:
                while(1):
                    button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'MuiButtonBase-root.MuiButton-root.MuiButton-outlined.jss621')))
                    button.click()    
                    time.sleep(2)
            except Exception as e:
                print(f"Error clicking button: ")

            loop2 = loop1.find_elements(By.CLASS_NAME, "pageWithFooter_content")
            for loop3 in loop2:
                time.sleep(2)
                loop4 = loop3.find_elements(By.CLASS_NAME, "jss1915.jss621")

                for links in loop4:
                    link = links.find_element(By.TAG_NAME, 'a')  
                    href = link.get_attribute('href')
                    href += '/products_reviewed'                    
                    urls_array.append(href)
    
    finally:
        driver.quit()
    
    return urls_array    

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
             #   rating_in_number+=background_position[7]
            rating_in_number = float(rating_in_number)
            if rating_in_number==250:
                ratingsall.append(5)
            elif rating_in_number==225:
                ratingsall.append(4.5)
            elif rating_in_number==200:
                ratingsall.append(4)
            elif rating_in_number==175:
                ratingsall.append(3.5)                
            elif rating_in_number==150:
                ratingsall.append(3)
            elif rating_in_number==125:
                ratingsall.append(2.5)                    
            elif rating_in_number==100:
                ratingsall.append(2)
            elif rating_in_number==75:
                ratingsall.append(1.5)                
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


def product_existence_check(driver, url):
    driver.get(url)
    time.sleep(4)
    products = driver.find_elements(By.CSS_SELECTOR, "div.jss2038")
    for product in products:
        
        product_texts = product.find_element(By.CLASS_NAME,"MuiTypography-root.MuiLink-root.MuiLink-underlineHover.jss2039.MuiTypography-colorPrimary")
        text = product_texts.text
        text = text.title()
        text += " "
        product_texts = product.find_element(By.CLASS_NAME,"MuiTypography-root.jss2040.MuiTypography-body1")
        text += product_texts.text
        return text

    


driver = webdriver.Chrome()


# Open the first page
url = "https://www.beautylish.com/s/jouer-cosmetics-deluxe-mini-powder-highlighter-citrine"

product_name = product_existence_check(driver, url)
print(product_name)

df = pd.read_excel('E:/Scraping/beautylish_products.xlsx')
value_counts = df['Product Name'].value_counts()
if product_name:
    filtered_count = value_counts.get(product_name, 0)
    print(f"Count for '{product_name}' in the Excel file: {filtered_count}")

review_links = scrape_review_links(url)
print(review_links)
time.sleep(5)
link_size = len(review_links)
if(link_size==filtered_count):
    print("Product has already been scrapped.")
else:
    for url in review_links:
        try:
            driver = webdriver.Chrome()
            driver.get(url)
            link_size-=1
            print(f'Remaining Profiles:{link_size}')
        except Exception as e:
            print("Invalid link")
            print(url)
            continue    


        # Opening an empty dataframe to store scrapped data
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
        if(uniquenesscheck("E:/Scraping/beautylish_products.xlsx", all_data)==True):
            append_to_excel("E:/Scraping/beautylish_products.xlsx", all_data)
            print("Titles and ratings have been saved to beautylish_products.xlsx")
        else:
            print(url)
            print("User Already Exists")
        



