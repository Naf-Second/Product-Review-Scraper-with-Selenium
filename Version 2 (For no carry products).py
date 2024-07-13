from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
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

def load_unique_user_ids(filename, sheet_name='Sheet1'):
    unique_user_ids_set = set()
    if os.path.exists(filename):
        existing_data = pd.read_excel(filename, sheet_name=sheet_name)
        if 'User ID' in existing_data.columns and not existing_data['User ID'].isnull().all():
            unique_user_ids_set = set(existing_data['User ID'].unique())
    return unique_user_ids_set

def scrape_review_links(url):
    urls_array = []
    i=0
    try:
        driver.get(url)

        wait = WebDriverWait(driver, 2)

        # Scroll down to the end of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(4)  # Allow time for the page to load new content if any


        # Find the review elements
        rev = driver.find_elements(By.CLASS_NAME, "pageWithFooter")
        for loop1 in rev:
            try:
                while(1):
                    button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn.ml10')))
                    button.click()                        
                    time.sleep(4)
                    i+=1
            except Exception as e:
                print(f"Error clicking button: ")

            loop2 = loop1.find_elements(By.CLASS_NAME, "pageWithFooter_content")
            for loop3 in loop2:
                time.sleep(2)
                loop4 = loop3.find_elements(By.CLASS_NAME, "reviewPod")

                for links in loop4:
                    element = links.find_element(By.CLASS_NAME, 'reviewPod_avatar')
                    href = element.get_attribute('href')
                    if href is not None:    
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

    ratings = driver.find_elements(By.CSS_SELECTOR, "div.rating_image.left")
    products = driver.find_elements(By.CLASS_NAME, "head")
    product_id_elements = driver.find_elements(By.CSS_SELECTOR, "div.module.bb1.mb5.js-review.discussion_content")
    review = driver.find_elements(By.CSS_SELECTOR, "div.js-discussion-edit-toggle")
    user = url.split('/')
    if(user[3]=='profile'):
        user = user[4]
    else:
        user = user[3]

    for product in products:
        names.append(product.text)
        userid.append(user)

    for result in product_id_elements:
        style = result.get_attribute('id')
        productids.append(style)

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
            if(len(background_position) >= 10):
                rating_in_number += background_position[5:8]
            else:
                rating_in_number += background_position[5:7]

            rating_in_number = float(rating_in_number)
            if rating_in_number == 250:
                ratingsall.append(5)
            elif rating_in_number == 225:
                ratingsall.append(4.5)
            elif rating_in_number == 200:
                ratingsall.append(4)
            elif rating_in_number == 175:
                ratingsall.append(3.5)                
            elif rating_in_number == 150:
                ratingsall.append(3)
            elif rating_in_number == 125:
                ratingsall.append(2.5)                    
            elif rating_in_number == 100:
                ratingsall.append(2)
            elif rating_in_number == 75:
                ratingsall.append(1.5)                
            elif rating_in_number == 50:
                ratingsall.append(1)

    for title in review:
        main = title.find_element(By.CSS_SELECTOR, "p.fwb[itemprop='name']")
        rev_title.append(main.text)

    j = 0
    for description in review:
        for i in range(len(product_id_elements)):
            temp = product_id_elements[j].get_attribute('id')
            j += 1
            break
        xpath = "//*[@id='" + temp + "']/div[2]/div[2]/div/p[3]"
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
    time.sleep(3)
    products = driver.find_elements(By.CSS_SELECTOR, "div.jss2038")
    for product in products:
        product_texts = product.find_element(By.CLASS_NAME, "fluid.contentRegion")
        text = product_texts.text
        text = text.title()
        text += " "
        product_texts = product.find_element(By.CLASS_NAME, "view_lockup")
        text += product_texts.text
        return text

# Initialize Chrome webdriver
driver = webdriver.Chrome()

# Open the first page
url = "https://www.beautylish.com/p/vaseline-petroleum-jelly"

product_name = product_existence_check(driver, url)
            #fix the path here
try:
    df = pd.read_excel('E:/Scraping/beautylish_products4.xlsx')
except FileNotFoundError:
    data = {'User ID': [], 'Product Name': [], 'Product ID': [], 'Product Rating': [], 'User Review Title': [], 'Review Description': []}
    df = pd.DataFrame(data)
    df.to_excel('E:/Scraping/beautylish_products4.xlsx', index=False)
    

value_counts = df['Product Name'].value_counts()
if product_name:
    filtered_count = value_counts.get(product_name, 0)
    print(f"Count for '{product_name}' in the Excel file: {filtered_count}")
    
filtered_count = 0
review_links = scrape_review_links(url)
print(f'Number of profile links collected: {(len(review_links))}')
time.sleep(2)

#fix the path here
unique_user_ids_set = load_unique_user_ids('E:/Scraping/beautylish_products4.xlsx')
process_links = [link for link in review_links if link.split('/')[-2] not in unique_user_ids_set]
print(f'Lenght of processed_links array: {(len(process_links))}')

process_links = list(set(link for link in review_links if link.split('/')[-2] not in unique_user_ids_set))

link_size = len(review_links)
print(f'Lenght of processed_links array after applying set: {(len(process_links))}')

if link_size - filtered_count <= 7:
    print("Product has already been scraped.")
else:
    reviews_collected = 0
    all_data = pd.DataFrame()
    for url in process_links:
        try:
            driver = webdriver.Chrome()
            driver.get(url)
            link_size -= 1
            reviews_collected += 1
            print(f'Profiles Reviewed: {reviews_collected}')
            print(f'Remaining Profiles: {link_size}')
        except Exception as e:
            print("Invalid link")
            print(url)
            continue

        
        check = False
        while True:
            time.sleep(2)
            new_data = scrape_page(driver)
            if new_data is not False:
                check = True
                all_data = pd.concat([all_data, new_data], ignore_index=True)
                
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "span.pager_next a")
                    next_button.click()
                    time.sleep(3)
                except Exception as e:
                    print("No more pages to scrape.")
                    break
            else:
                break
        driver.quit()

        if check:
            #fix the path here
            append_to_excel("E:/Scraping/beautylish_products4.xlsx", all_data)
            print("Titles and ratings have been saved to beautylish_products4.xlsx")
        else:
            print(url)
        print("User Already Exists")
