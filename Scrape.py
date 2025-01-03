import requests
from bs4 import BeautifulSoup
import pandas as pd
from lxml import html
from time import sleep

def decode_cf_email(cfemail):
    
    email = ""
    r = int(cfemail[:2], 16)
    for i in range(2, len(cfemail), 2):
        email += chr(int(cfemail[i:i+2], 16) ^ r)
    return email



# Function to scrape company data
def scrape_company_data(link):
    while True:  # Repeat until a 200 status code is received
        try:            
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
            }
            response = requests.get(link, headers=headers)
            print(f"HTTP Status Code: {response.status_code}")
            
            if response.status_code == 200:  # Break the loop if successful
                soup = BeautifulSoup(response.content, "html.parser")
                tree = html.fromstring(str(soup))

                # Initialize a dictionary to store company data
                company_data = {}

                # Extract Company Name
                try:
                    company_data["Company Name"] = soup.find("h1", id="title").text.strip()
                except AttributeError:
                    company_data["Company Name"] = "N/A"
                
                # Extract Address
                try:
                    address = soup.find("div", id="contact-details-content").find_all("span")[3].text.strip()
                    company_data["Address"] = address
                except (AttributeError, IndexError):
                    company_data["Address"] = "N/A"

                # Extract Email
                try:
                    email_element = soup.find("a", class_="__cf_email__")
                    if email_element:
                        cfemail = email_element.get("data-cfemail")
                        decoded_email = decode_cf_email(cfemail)
                        company_data["Email"] = decoded_email
                    else:
                        company_data["Email"] = "N/A"
                except Exception as e:
                    company_data["Email"] = "N/A"
                    print(f"Error decoding email: {e}")

                # Extract Activity
                try:
                    activity = soup.find("table", class_="table table-striped").find_all("tr")[11].find_all("td")[1].find_all("span")[1].text.strip()
                    company_data["Activity"] = activity
                except (AttributeError, IndexError):
                    company_data["Activity"] = "N/A"

                # Use XPath to extract Category
                try:
                    category_elements = tree.xpath('/html/body/div[6]/section/div[1]/div[1]/div/div/table/tbody/tr[9]/td[2]')
                    if category_elements:
                        company_data["Category"] = category_elements[0].text.strip()
                    else:
                        company_data["Category"] = "N/A"
                except Exception as e:
                    company_data["Category"] = "N/A"
                    print(f"Error extracting category: {e}")
                try:
                    registration_number =tree.xpath('/html/body/div[6]/section/div[1]/div[1]/div/div/table/tbody/tr[6]/td[2]') 
                    if registration_number:
                        company_data["Registration Number"] = registration_number[0].text.strip()
                    else:
                        company_data["Registration Number"] = "N/A"
                except Exception as e:
                    company_data["Registration Number"] = "N/A"
                # Use XPath to extract Date of Incorporation
                try:
                    date_elements = tree.xpath('/html/body/div[6]/section/div[1]/div[1]/div/div/table/tbody/tr[11]/td[2]')
                    if date_elements:
                        company_data["Date of Incorporation"] = date_elements[0].text.strip()
                    else:
                        company_data["Date of Incorporation"] = "N/A"
                except Exception as e:
                    company_data["Date of Incorporation"] = "N/A"
                    print(f"Error extracting date of incorporation: {e}")

                # Print the extracted data
                print(company_data)
                return company_data

            else:
                print(f"Retrying in 1 seconds... Status code: {response.status_code}")
                sleep(1)  # Wait before retrying

        except Exception as e:
            print(f"Error scraping {link}: {e}. Retrying in 5 seconds...")
            sleep(5)



def get_links(base_url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    }
    response = requests.get(base_url, headers=headers)
    soup = BeautifulSoup(response.content, "html.parser")
    links = []
    
    table = soup.find("table", {"id": "results"})
    if table:
        anchor_tags = table.find_all("a", href=True)  
        for anchor in anchor_tags:
            links.append(anchor['href'])  

    return links

    

links = get_links("https://www.zaubacorp.com/company-by-address/Hyderabad")

company_data_list = []
i = 0
# Loop through each link and scrape the data
for link in links:
    print(i)
    i+=1
    company_data = scrape_company_data(link)
    if company_data:
        company_data_list.append(company_data)
        print(f"Scraped: {company_data['Company Name']}")

# Convert the data to a DataFrame and save it to an Excel file
df = pd.DataFrame(company_data_list)
df.to_excel("company_details.xlsx", index=False)

print("Data successfully saved to company_details.xlsx")