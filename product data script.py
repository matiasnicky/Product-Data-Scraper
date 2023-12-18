import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import time

# Set the working directory
os.chdir("YOUR_DIRECTORY_PATH")  # Replace with your working directory path

# Read the Excel file with multiple sheets
excel_file = 'Sample.xlsx'  # Replace with your Excel file's name
sheet_name = 'Sheet 1'      # Replace with your Excel file's name

# Read the selected sheet
try:
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
except Exception as e:
    print(f"Error reading the selected sheet: {e}")
    exit()

# Load the index of the last processed row from a text file (if available)
last_processed_index = 0
if os.path.exists("last_processed_index.txt"):
    with open("last_processed_index.txt", "r") as f:
        last_processed_index = int(f.read())

# List to store scraped data
scraped_data = []

# Function to scrape product details from a URL with timeout and retries
def scrape_product_details(url):
    max_retries = 3   # You can change the value to however you like
    retry_delay = 10  # Retry delay in seconds
    timeout_duration = 30  # When the website does not respond in 30 seconds, it will either stop or retry the scraping process (adjust as needed)

    #IMPORTANT!
    #This is where the Scraping logic starts. You have to inspect your target website and apply accordingly. Every website places the data in different classes and div
    #Generalize the details fetched from the website without revealing specific elements. This will give you the best chance of success
    
    for attempt in range(max_retries):
        try:
            response = requests.get(url, timeout=timeout_duration)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                items = soup.find_all('div', {'class': 'item'})
                for item in items:
                    try:
                        # inside this try loop, the program is trying to find the data from the specified elements For example:
                        type_div = item.find('div', {'style': 'float: left'}) # Here the script tries to get the type of the product by targeting all 'div' that is 'floating on the left'
                        type_info = type_div.text.strip() if type_div else '' # If it finds the text details about the type of the product, it will record it
                    
                        price_div = item.find('div', {'class': 'Price'}) # Same thing is going on here, the target is the price of the product
                        price = price_div.text.strip() if price_div else ''
                        
                        shipping_div = item.find('strong', {'class': 'key'}, text='shipping:')
                        shipping = shipping_div.find_next('span', {'class': 'value'}).text.strip() if shipping_div else ''
                        
                        contact_information_div = item.find('strong', {'class': 'key'}, text='EAN Code:')
                        contact_information = ean_code_div.find_next('span', {'class': 'value'}).text.strip() if contact_information_div else ''
                        
                        scraped_data.append([product_name, product_code, type_info, price, shipping, contact_information]) # the data scraped is recorded to the corresponding columns
                        
                        print(f"Scraped Product:")
                        print(f" Product Name: {product_name}")
                        print(f" Product Code: {product_code}")
                        print(f" Type: {type_info}")
                        print(f" Price: {price}")
                        print(f" Shipping: {shipping}")
                        print(f" Contact Information: {contact_information}")
                        print("=" * 30) # gives spacing to next scraped data
                            
                    except Exception as e:
                        
                        print(f"An error occurred while scraping product details: {e}")
                        break  # Exit the loop on error
                break  # Exit the loop if successful
            else:
                print(f"Failed to fetch data from URL: {url}")
        except requests.exceptions.Timeout:
            print(f"Request timed out for URL: {url}. Retrying in {retry_delay} seconds...")
            time.sleep(retry_delay)
        except Exception as e:
            print(f"An error occurred while processing URL: {url}: {e}")
            break  # Exit the loop on error

# Iterate through each row in the dataframe starting from the last processed index. If you don't want to continue scraping the same file after done, you don't need to worry about this (more explanation on README)
for index in range(last_processed_index, len(df)):
    row = df.iloc[index]
    product_name = row["Product Name"] # In this script, there are two necessary conditions to be fulfilled. The first is that the product name of the scraped product must match the specified product name
    product_code = row["Product Code"] # If the Product name matches the data on the website, it will then scrape the data. 

    # In this sample script, the website search relies on the product code. The reason for this is that there are various names for the same product code.
    # The program is also able to search through different links or different regions as long as the links aren't the same and you can see the pattern
    # In This sample case, the product code is always the same, but the link and the product name are slightly different
    # If the first url fails, it will try to find it on the other links

    if product_name == 'Sample1': # Specify the product name that you require as well as the corresponding link. The link might be the same or different
        url = f'https://example.com/Sample1_search={product_code}'
        print(f'Scraping data for ProductCode: {product_code} using Sample1 link...')
        scrape_product_details(url)
      
    elif product_name == 'Sample2': # Specify other product name that you tolerate
        url = f'https://example.com/Sample2_search={product_code}'
        print(f'Scraping data for ProductCode: {product_code} using Sample2 link...')
        scrape_product_details(url)
      
    else:
        print(f'Skipping ProductCode: {product_code} due to unknown manufacturer.')
    
    # Save the current index to the text file after each successful iteration
    with open("last_processed_index.txt", "w") as f:
        f.write(str(index + 1))  # Save the next index to process

# Create a DataFrame from scraped data
results_df = pd.DataFrame(scraped_data, columns=['Product Name', 'Product Code', 'Type', 'Price', 'Shipping', 'Contact Information'])

# Save the DataFrame to an Excel file
output_file = 'scraped_results.xlsx'
results_df.to_excel(output_file, index=False)
print(f"Scraped data saved to {output_file}")
