import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random

user_agents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 13_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15'
]


def get_info(lastname, name, school = 'Not Found', title = ''):
    lastname = lastname.title()
    name = name.title()
    header = random.choice(user_agents)
    headers={'User-Agent': header}
    url = f'https://www.effinghamschools.com/directory?utf8=✓&const_search_group_ids=&const_search_role_ids=1&const_search_keyword=&const_search_first_name=&const_search_last_name={lastname}&const_search_location=&const_search_department='
    response = requests.get(url, headers=headers,timeout=10)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find all fsConstituentItem divs
        constituent_items = soup.find_all('div', class_='fsConstituentItem')

        # Iterate through each fsConstituentItem div
        if constituent_items:
            for constituent_item in constituent_items:
                # Find the fsFullName div within each fsConstituentItem div
                full_name_div = constituent_item.find('h3', class_='fsFullName')
                # Extract href and text
                if full_name_div:
                    href = full_name_div.a['href'] if full_name_div.a else None
                    text = full_name_div.text.strip()
                    if text == name:
                        response = requests.get(href, headers=headers,timeout=10)
                        if response.status_code == 200:
                            soup = BeautifulSoup(response.text, 'html.parser')
                            title_div = soup.find('div', class_='fsProfileSectionData fsTitle')
                            if title_div:
                                title_el = title_div.find('div', class_='fsProfileSectionFieldValue')
                                title = title_el.text.strip()

                            campus_div = soup.find('div', class_='fsProfileSectionData fsLocation')
                            if campus_div:
                                campus_el = campus_div.find('div', class_='fsProfileSectionFieldValue')
                                school = campus_el.text.strip()
                        else:
                            print(f"Failed to fetch the URL. Status code: {response.status_code}")
                        break
                    else:
                        school = 'Not Found'
                        title = ''
            
        else:
            school = 'Not Found'
            title = ''
        print('Lastname: ', lastname, 'Fullname: ', name, 'School Name: ', school, 'Title: ', title)
        print('-'*50)

    else:
        print(f"Failed to fetch the URL. Status code: {response.status_code}")
    return school, title




file_path = 'input_data.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Iterate through each row
for index, row in df.iterrows():
    # Check if lastname equals name
    print('Excel lastname: ',row['lastName'], 'fullname : ',row['name'])
    school, title = get_info(row['lastName'], row['name'])
    # Update school and title columns
    df.at[index, 'school'] = school
    df.at[index, 'title'] = title
    if index == 50:
        break
    time.sleep(10)

df['title'] = df['title'].astype('object')

# Save the updated DataFrame to a new Excel file
output_file_path = 'output_data.xlsx'
df.to_excel(output_file_path, index=False)


