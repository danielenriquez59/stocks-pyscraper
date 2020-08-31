# 0. Help Section ############################################################################
""" Info: This code scrapes  marketscreener.com for press releases referred to in 'inputfilename'
          and outputs the results to scraped_data_XXX.xlsx
 Prerequisites for this code to work
 1. python 3
 2. selenium (pip install)
 3. BeautifulSoup (pip install)
 4. pandas (pip install) (or with anaconda)
 5. chromedriver.exe (chromedriver.exe - can be installed from https://chromedriver.storage.googleapis.com/index.html?path=85.0.4183.87/)
 
 Author: Daniel Enriquez
 Release History: 
 00 - 08/2020
"""
############################################################################
#1. Import Section ############################################################
############################################################################
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import csv
import pdb #for debugging
from datetime import datetime
from openpyxl import load_workbook
from re import search


time = datetime.now().strftime("%Y%m%d-%H%M%S")

############################################################################
#2. User Inputs ############################################################
############################################################################
inputfilename = 'stockstoscrape.csv' # csv containing stocks to watch with one header row
outfilename = 'scraped_data_'+time+'.xlsx' #output filename

############################################################################
#3. Import CSV #############################################################
############################################################################
weblinkmain = "https://www.marketscreener.com/" # do not change this

# best to put chromedriver a env var PATH folder. I put mine in 'Scripts'
driver = webdriver.Chrome()

writer = pd.ExcelWriter(outfilename) 
stockstoscrape = open(inputfilename, newline='')
headers = next(stockstoscrape, None)

############################################################################
#6. Define functions #######################################################
############################################################################


############################################################################
#5. Main Scraper ###########################################################
############################################################################

for cnt, symbol in enumerate(stockstoscrape): # search for each stock in csv stockstoscrape
    # intialization of df
    presstitles  = [] # List for PR
    presslinks = [] # hyper links to the PR
    symbols = [] #ticket symbols
    dates = []
    driver.get(weblinkmain) #navigate to landing page
    srch_box = driver.find_element_by_class_name('inputrecherche') #find search box , specific to website
    print('Searching for'+symbol)
    srch_box.send_keys(symbol) #search symbol
    
    wait=WebDriverWait(driver, 15).until( EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'More Instruments')) ) # wait for page to load
    
    linkstring = "//a[contains(@href, 'quote/stock')]"
    firstlink = driver.find_element_by_xpath(linkstring) #will click first link it finds with linkstring
    firstlink.click()
    
    wait=WebDriverWait(driver, 15).until( EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'More about the company')) ) # wait for page to load

    soup = BeautifulSoup(driver.page_source,'lxml')
    
    links = soup.findAll('td', attrs={'class':"newsColCT ptop3 pbottom3 pleft5"}) # extract all row info from PR table
    
    for link in links:     
        hyperlink = link.find('a')['href'] #find hyperlink
        texts = link.text # find text title
        
        date = link.find('a')['title'] #find str containing date
        date.replace('News Stock market from', ' ') #clean up string
        date = search('[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]', date) #clean up string
        
        #save stuff in lists
        presslinks.append(weblinkmain+hyperlink)
        presstitles.append(texts)
        dates.append(date.group(0)) 
        #end links loop
    
    #save stuff in lists
    exec('df'+str(cnt)+' = pd.DataFrame({"Date":dates, "PR Titles":presstitles, "PR Links":presslinks})') # save df
    exec('df'+str(cnt)+'.to_excel(writer, symbol)' ) #write df

    #end WL loop    
    
writer.save()
writer.close()

    
    
    



