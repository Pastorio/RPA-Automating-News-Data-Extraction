from RPA.Browser.Selenium import Selenium
from RPA.Robocorp.WorkItems import WorkItems
from RPA.Excel.Files import Files

from selenium.common.exceptions import StaleElementReferenceException

from robocorp.tasks import task

import re
from datetime import datetime

class NyTimesScraper:
    def __init__(self):
        self.browser_lib = Selenium()
        self.news_information = []
    
    def open_the_website(self, url: str):
        self.browser_lib.open_browser(url, browser="edge")
        # self.browser_lib.open_available_browser(url)
    
    def close_policy_terms(self):
        button_xpath = "xpath://button[normalize-space()='Continue']"
        # self.browser_lib.wait_until_element_is_enabled(button_xpath)
        self.browser_lib.click_button_when_visible(button_xpath)
        
    def close_cookies(self):
        button_xpath = "xpath://button[normalize-space()='Accept']"
        
        if self.browser_lib.is_element_visible(button_xpath):
            self.browser_lib.click_button(button_xpath)

    def search(self, phrase):
        div_search_xpath = "xpath://div[@class='css-qo6pn ea180rp0']"
        self.browser_lib.wait_until_element_is_visible(div_search_xpath)
        
        button_xpath = "xpath://button[@class='ea180rp1 css-fzvsed']"
        button_search_xpath = "xpath://button[@class='css-tkwi90 e1iflr850']"
        
        if self.browser_lib.is_element_visible(button_search_xpath):
            self.browser_lib.click_button(button_search_xpath)
        elif self.browser_lib.is_element_visible(button_xpath):
            self.browser_lib.click_button(button_xpath)
        
        input_field = "xpath://input[@placeholder='SEARCH']"
        
        self.browser_lib.wait_until_element_is_visible(input_field)
        self.browser_lib.input_text(input_field, phrase)
        self.browser_lib.press_keys(input_field, "ENTER")

    def sort_newest(self):  
        button_xpath = "//select[@class='css-v7it2b']"
        self.browser_lib.click_element_when_visible(button_xpath)
        
        button_xpath = "//option[@value='newest']"
        self.browser_lib.click_element_when_visible(button_xpath)
        
    
    def filter_sections(self, section_list):
        button_xpath = "xpath://div[@data-testid='section']//button[@type='button']"
        self.browser_lib.wait_until_element_is_enabled(button_xpath)
        self.browser_lib.click_button_when_visible(button_xpath)
        
        section_list_lower = [section.lower() for section in section_list]
        
        self.browser_lib.wait_until_element_is_visible("class:css-1qtb2wd")
        elements = self.browser_lib.get_webelements("class:css-1qtb2wd")
        for element in elements:
            text = element.text.lower()
            
            if section_list_lower != []:
                for section in section_list_lower:
                    if section in text:
                        self.browser_lib.driver.execute_script("arguments[0].scrollIntoView();", element)

                        self.browser_lib.wait_until_element_is_visible(element)
                        self.browser_lib.click_element_if_visible(element)

    def calculate_month_diff_from_link(self, url):
            #Get the date from last news
            date_pattern = r'\d{4}/\d{2}/\d{2}'
            match = re.search(date_pattern, url)

            extracted_date_str = match.group()
            extracted_date = datetime.strptime(extracted_date_str, "%Y/%m/%d").date()

            #Compare with actual date
            today_date = datetime.today().date()
            difference_in_months = (today_date.year - extracted_date.year) * 12 + (today_date.month - extracted_date.month)
            
            actual_num_months = difference_in_months+1
            
            return extracted_date_str, actual_num_months
    
    def load_news_in_month_range(self, max_load_pages=100, num_months=1):
        '''Load more news pressing show more button until the number of months was reached'''

        # Wait for the elements with class 'css-1l4w6pd' to be visible
        self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")

        for load in range(max_load_pages):
            try:
                # Get the list of elements with class 'css-1l4w6pd'
                elements = self.browser_lib.get_webelements("class:css-1l4w6pd")

                # Find the last element in the list and wait until it is visible
                last_element = elements[-1]
                self.browser_lib.wait_until_element_is_visible(last_element)

                # Extract the link from the last element
                sub_element = self.browser_lib.get_webelement(locator="css:a", parent=last_element)
                page_link = sub_element.get_attribute("href")

                # Calculate the number of months from the link
                extracted_date_str, actual_num_months = self.calculate_month_diff_from_link(page_link)

                if num_months == 0:
                    num_months = 1

                if num_months >= actual_num_months:
                    # Click the "Show More" button if it is visible
                    button_xpath = "xpath://button[normalize-space()='Show More']"
                    if self.browser_lib.is_element_visible(button_xpath):
                        self.browser_lib.click_button(button_xpath)

                    # Wait for the elements with class 'css-1l4w6pd' to be visible again after clicking
                    self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
                else:
                    break
            except StaleElementReferenceException:
                # If a StaleElementReferenceException occurs, continue the loop without breaking
                # This ensures that the loop continues even if some elements become stale
                continue
            except Exception as e:
                print(f"An error occurred: {str(e)}")
        
    def load_news_in_month_range_old(self, max_load_pages=100, num_months=1):
        '''Load more news pressing show more butter until the number of months was riched'''
        
        self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
        elements = self.browser_lib.get_webelements("class:css-1l4w6pd")
        
        for load in range(max_load_pages):
            try:
                last_element = elements[-1]
                self.browser_lib.wait_until_element_is_visible(last_element)

                sub_element = self.browser_lib.get_webelement(locator="css:a", parent=last_element)
                page_link = sub_element.get_attribute("href")
            except StaleElementReferenceException:
                last_element = self.browser_lib.get_webelement(elements[-1])
                
                sub_element = self.browser_lib.get_webelement(locator="css:a", parent=last_element)
                page_link = sub_element.get_attribute("href")

            try:            
                extracted_date_str, actual_num_months = self.calculate_month_diff_from_link(page_link)
                
                if num_months == 0:
                    num_months = 1
                
                if num_months >= actual_num_months:      
                    button_xpath = "xpath://button[normalize-space()='Show More']"
                    if self.browser_lib.is_element_visible(button_xpath):
                        self.browser_lib.click_button(button_xpath)
                    
                    self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
                    elements = self.browser_lib.get_webelements("class:css-1l4w6pd")
                else:
                    break
            except Exception as e:
                print(f"An error occurred: {str(e)}")
            
    def count_occurrence(self, phrase, text):
        '''Counte the number of times a specific phrase appears in a text'''
        lowercase_phrase = phrase.lower()
        lowercase_text = text.lower()
        
        count = lowercase_text.count(lowercase_phrase)
        
        return count
    
    def get_news_information(self, phrase, num_months=1):
        '''Get all information about news present on the page'''

        self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
        elements = self.browser_lib.get_webelements("class:css-1l4w6pd")
        for element in elements:
            try:
                # Get source news information
                self.browser_lib.wait_until_element_is_visible(element)

                title_element = self.browser_lib.get_webelement(locator="class:css-2fgx4k", parent=element)
                title = title_element.text if title_element else ''

                page_link_element = self.browser_lib.get_webelement(locator="css:a", parent=element)
                page_link = page_link_element.get_attribute("href") if page_link_element else ''
                extracted_date_str, actual_num_months = self.calculate_month_diff_from_link(page_link)

                description_element = self.browser_lib.get_webelement(locator="class:css-16nhkrn", parent=element)
                description = description_element.text if description_element else ''
                
                img_element = self.browser_lib.get_webelement(locator="class:css-rq4mmj", parent=element)
                img_src = img_element.get_attribute("src") if description_element else ''
                
                # Aditional information
                count_title = self.count_occurrence(phrase, title)
                count_description = self.count_occurrence(phrase, description)

                title_contains_money = any(self.count_occurrence(keyword, title) for keyword in ['$','dollar','USD'])
                description_contains_money = any(self.count_occurrence(keyword, description) for keyword in ['$','dollar','USD'])

                # Verify if the news are on the month range
                if num_months == 0:
                    num_months = 1
                if num_months < actual_num_months:
                    break
                
                # Save news information not saved yet
                title_already_saved = any(item.get("title") == title for item in self.news_information)
                
                if not title_already_saved:
                    self.news_information.append({
                        'title': title,
                        'date': extracted_date_str,
                        'description': description,
                        'count_title': count_title,
                        'count_description': count_description,
                        'title_contains_money': title_contains_money,
                        'description_contains_money': description_contains_money,
                        'img_src': img_src
                    })
                # print(title_already_saved, "len - ", len(self.news_information), "title - ", title)
                
            except Exception as e:
                print(f"An error occurred: {str(e)}")
    
    def save_news_image(self, path):
        
        url_query = self.browser_lib.get_location()
        i = 0
        for actual_news in self.news_information:
            img_src = actual_news['img_src']
            self.browser_lib.go_to(img_src)
            
            img_path = f"{path}/img_{str(i)}.png"
            self.browser_lib.capture_element_screenshot("tag:img",img_path)
            
            actual_news['img_name'] = f"img_{str(i)}.png"
            i += 1
            
        self.browser_lib.go_to(url_query)
    
    def save_information_in_excel(self, path):
        lib = Files()
        
        lib.create_workbook(path, fmt="xlsx")
        lib.create_worksheet(name="News", content=self.news_information, header=True)
        lib.save_workbook()
        
        lib.close_workbook()
        
    def store_screenshot(self, filename):
        self.browser_lib.screenshot(filename=filename)
    
    def finish_execution(self):
        self.browser_lib.close_all_browsers()
    
    def __del__(self):
        self.browser_lib.close_all_browsers()
        
        

@task
def main():
    try:
        wi = WorkItems()
        wi.get_input_work_item()
        search_phrase = wi.get_work_item_variable("search_phrase")
        sections = wi.get_work_item_variable("sections")
        months = wi.get_work_item_variable("months")
        
        scraper = NyTimesScraper()
        scraper.open_the_website(url="https://www.nytimes.com/")
        
        scraper.close_policy_terms()
        scraper.close_cookies()
        
        scraper.search(phrase=search_phrase)
        scraper.close_cookies()
        
        scraper.sort_newest()

        scraper.filter_sections(section_list=sections)
        
        scraper.load_news_in_month_range(num_months=months)
        
        scraper.get_news_information(phrase=search_phrase, num_months=months)
        
        scraper.save_news_image(path=f"output")
        
        scraper.save_information_in_excel(path="output/my_new_excel.xlsx")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        scraper.finish_execution()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()