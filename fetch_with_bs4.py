# -*- coding: utf-8 -*-
# Importing libraries
import requests
from bs4 import BeautifulSoup

# Set URL of the page where schedule links are listed
uniurl="https://rguk.ru/students/schedule/"                 # URL of the page containing the link to the Excel file

# The following variables can be changed as needed
intstitute='ИХТиПЭ'                                         # Institute to look for
studyform='очно'                                            # Study form to look for
scyear='2курс'                                              # Study year to look for

# Assemble the expected filename based on the variables
filename=f"{intstitute}-{studyform}-{scyear}.xlsx"          # Filename for the downloaded Excel file
# Function to match the href with the expected filename
def match_href(href):
    if href and 'officeapps.live.com' not in href and href.endswith(filename):
        return True
    return False
# Process the webpage
response= requests.get(uniurl)                             # Fetch the webpage
soup= BeautifulSoup(response.content, 'html.parser')       # Parse the webpage content
target_link=soup.find("a", href=match_href)                # Find the link matching the expected filename

# Download the file
if target_link:
    file_url= f'{'/'.join(uniurl.split('/')[:3])}/{target_link['href']}' # This constructs the full URL by taking the base URL and appending the relative path
    file_response= requests.get(file_url)
    with open(filename, 'wb') as file:
        file.write(file_response.content)