from selenium import webdriver
import time
import pandas as pd

class FindingCounties:
    
    def __init__(self, driver_path, website):
        self.driver_path = driver_path
        self.website = website


    def loadingPage(self):
        self.driver = webdriver.Chrome(self.driver_path)
        self.driver.get(self.website)
       
        
    def locatingTable(self, t):
        inputElement = self.driver.find_element_by_name("q")
        inputElement.send_keys(t)
        inputElement.submit()
        time.sleep(10)
        

        table_id_1 = self.driver.find_element_by_id("mainrow")
        table_id_2 = table_id_1.find_element_by_id("map-info")
        table_id_3 = table_id_2.find_element_by_tag_name("table")
        rows = table_id_3.find_elements_by_tag_name("tr")
        
        for row in rows:
            if row.find_elements_by_tag_name("th")[0].text == "County:":
                self.county = row.find_elements_by_tag_name("td")[0].text
        
        return self.county

    
    
    def printingCurrent(self, i):
        print(i, ": ", self.county)
        
        
        
if __name__ == '__main__':        
    
    # You have to install the package "selenium" and a webdriver of your choice
    # The path bewlow is linked to the location, where you saved the webdriver
    driver_path = "***********************"
    
    # Website you want to scrape
    website = "https://www.unitedstateszipcodes.org/"
    
    # List of adresses
    text = ["114 Mt. Auburn St. Cambridge, MA 02138", 
            "110 Championship Way, Ponte Vedra Beach, FL 32082",
            "500 S State St, Ann Arbor, MI 48109", 
            "77 Massachusetts Ave, Cambridge, MA 02139",
            "3911 Figueroa St, Los Angeles, CA 90037"]
    
    # Initializing a list in which the counties are saved
    counties = []
    
    # Creating the "FindingCounties" object
    obj = FindingCounties(driver_path, website)
    
    # Loading the page
    obj.loadingPage()
    
    # Iterating through the adress list
    for i, t in enumerate(text):
        county = obj.locatingTable(t)
        counties.append(county)
        obj.printingCurrent(i)
      
    adress = pd.DataFrame([text, counties], index=["Adresse", "County"]).T