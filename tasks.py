from RPA.Browser.Selenium import Selenium
from RPA.Robocorp.WorkItems import WorkItems
from RPA.Excel.Files import Files

from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException

from robocorp.tasks import task
from robot.api import logger

import re
from datetime import datetime

class NyTimesScraper:
    def __init__(self):
        self.browser_lib = Selenium()
        self.news_information = []
    
    def open_the_website(self, url: str):
        try:
            # self.browser_lib.open_browser(url, browser="edge")
            self.browser_lib.open_available_browser(url)
        except Exception as err:
            logger.warn(f"Failed to open the website, error: {err:}")
    
    def close_policy_terms(self):
        button_xpath = "xpath://button[normalize-space()='Continue']"
        try:
            self.browser_lib.wait_until_element_is_enabled(button_xpath)
            self.browser_lib.click_button(button_xpath)
        except Exception as err:
            logger.warn(f"Failed to close policy terms, error: {err:}")

    def close_cookies(self):
        button_xpath = "xpath://button[normalize-space()='Accept']"
        try:
            self.browser_lib.wait_until_element_is_enabled(button_xpath)
            self.browser_lib.click_button(button_xpath)
        except Exception as err:
            logger.warn(f"Failed to close policy terms, error: {err:}")

    def search(self, phrase: str):
        try:
            div_search_xpath = "xpath://div[@class='css-qo6pn ea180rp0']"
            self.browser_lib.wait_until_element_is_visible(div_search_xpath)
            
            button_xpath = "xpath://button[@class='ea180rp1 css-fzvsed']"
            button_search_xpath = "xpath://button[@class='css-tkwi90 e1iflr850']"
            
            if self.browser_lib.is_element_visible(button_search_xpath):
                self.browser_lib.click_button(button_search_xpath)
            elif self.browser_lib.is_element_visible(button_xpath):
                self.browser_lib.click_button(button_xpath)
        except Exception as err:
            logger.warn(f"Failed click on search button, error: {err:}")
            
            raise ValueError("An error occurred during click on search button.") from err
            
        try:
            input_field = "xpath://input[@placeholder='SEARCH']"
            
            self.browser_lib.wait_until_element_is_visible(input_field)
            self.browser_lib.input_text(input_field, phrase)
            self.browser_lib.press_keys(input_field, "ENTER")
        except Exception as err:
            logger.warn(f"Failed to search the phrase, error: {err:}")
            
            raise ValueError("An error occurred during search the phrase.") from err

    def sort_news_by_newest(self):  
        try:
            button_xpath = "//select[@class='css-v7it2b']"
            self.browser_lib.click_element_when_visible(button_xpath)
            
            button_xpath = "//option[@value='newest']"
            self.browser_lib.click_element_when_visible(button_xpath)
        except Exception as err:
            logger.warn(f"Failed to sort by newst, error: {err:}")
    
    def filter_sections(self, section_list: list):
        try:
            button_xpath = "xpath://div[@data-testid='section']//button[@type='button']"
            self.browser_lib.wait_until_element_is_enabled(button_xpath)
            self.browser_lib.click_button_when_visible(button_xpath)
        except Exception as err:
            logger.warn(f"Failed to open section options, error: {err:}")
        
        section_list_lower = [section.lower() for section in section_list]
        
        try:
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
        except Exception as err:
            logger.warn(f"Failed to select sections, error: {err:}")

    def calculate_month_diff_from_link(self, url: str):
        '''Calculate the difference in months from news to actual date'''
        date_pattern = r'\d{4}/\d{2}/\d{2}'
        match = re.search(date_pattern, url)

        extracted_date_str = match.group()
        extracted_date = datetime.strptime(extracted_date_str, "%Y/%m/%d").date()

        today_date = datetime.today().date()
        difference_in_months = (today_date.year - extracted_date.year) * 12 + (today_date.month - extracted_date.month)
        
        actual_num_months = difference_in_months+1
        
        return extracted_date_str, actual_num_months
    
    def get_element_safe(self, locator, parent=None):
        '''Handling for stale element reference or no such element exception'''
        try:
            return self.browser_lib.get_webelement(locator=locator, parent=parent)
        except (NoSuchElementException, StaleElementReferenceException):
            logger.warn(f"Element not found or stale element reference. Retrying")
            
            return self.browser_lib.get_webelement(locator=locator, parent=parent)    

    def load_news_in_month_range(self, max_load_pages=100, num_months=1):
        '''Load more news pressing show more button until the number of months was reached'''

        self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")

        for load in range(max_load_pages):
            try:
                elements = self.browser_lib.get_webelements("class:css-1l4w6pd")

                last_element = elements[-1]
                self.browser_lib.wait_until_element_is_visible(last_element)

                sub_element = self.get_element_safe(locator="css:a", parent=last_element)
                page_link = sub_element.get_attribute("href")

                extracted_date_str, actual_num_months = self.calculate_month_diff_from_link(page_link)

                if num_months == 0:
                    num_months = 1

                if num_months >= actual_num_months:
                    button_xpath = "xpath://button[normalize-space()='Show More']"
                    if self.browser_lib.is_element_visible(button_xpath):
                        self.browser_lib.click_button(button_xpath)

                    self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
                else:
                    break
            except (NoSuchElementException, StaleElementReferenceException) as err:
                logger.warn(f"Element not found: {str(err)}")
            except Exception as err:
                logger.warn(f"Failed to load past news: {str(err)}")
                
        
            
    def count_phrase_in_text(self, phrase: str, text: str) -> int:
        '''Counte the number of times a specific phrase appears in a text'''
        lowercase_phrase = phrase.lower()
        lowercase_text = text.lower()
        
        count = lowercase_text.count(lowercase_phrase)
        
        return count
    

    def get_news_information(self, phrase: str, num_months=1):
        self.browser_lib.wait_until_element_is_visible("class:css-1l4w6pd")
        elements = self.browser_lib.get_webelements("class:css-1l4w6pd")
        
        for element in elements:
            try:
                self.browser_lib.driver.execute_script("arguments[0].scrollIntoView();", element)
                
                # Get source news information
                title_element = self.get_element_safe(locator="class:css-2fgx4k", parent=element)
                title = title_element.text if title_element else ''

                page_link_element = self.get_element_safe(locator="css:a", parent=element)
                page_link = page_link_element.get_attribute("href") if page_link_element else ''
                extracted_date_str, actual_num_months = self.calculate_month_diff_from_link(page_link)

                description_element = self.get_element_safe(locator="class:css-16nhkrn", parent=element)
                description = description_element.text if description_element else ''
                
                img_element = self.get_element_safe(locator="class:css-rq4mmj", parent=element)
                img_src = img_element.get_attribute("src") if img_element else ''
                
                # Additional information
                count_title = self.count_phrase_in_text(phrase, title)
                count_description = self.count_phrase_in_text(phrase, description)

                title_contains_money = any(self.count_phrase_in_text(keyword, title) for keyword in ['$','dollar','USD'])
                description_contains_money = any(self.count_phrase_in_text(keyword, description) for keyword in ['$','dollar','USD'])

                # Verify if the news is in the month range
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
                    
            except (NoSuchElementException, StaleElementReferenceException) as err:
                logger.warn(f"Element not found: {str(err)}")
            except Exception as err:
                logger.warn(f"Failed to get news: {str(err)}")
                
    
    def save_news_image(self, path: str):
        url_query = self.browser_lib.get_location()
        i = 0
        try:
            for actual_news in self.news_information:
                img_src = actual_news['img_src']
                self.browser_lib.go_to(img_src)
                
                img_path = f"{path}/img_{str(i)}.png"
                self.browser_lib.capture_element_screenshot("tag:img",img_path)
                
                actual_news['img_name'] = f"img_{str(i)}.png"
                i += 1
                
            self.browser_lib.go_to(url_query)
        except Exception as err:
                logger.warn(f"Failed to save images: {str(err)}") 
                
    def save_information_in_excel(self, path: str):
        try:
            lib = Files()
            
            lib.create_workbook(path, fmt="xlsx")
            lib.create_worksheet(name="News", content=self.news_information, header=True)
            lib.save_workbook()
            
            lib.close_workbook()
        except Exception as err:
                logger.warn(f"Failed to save excel informations: {str(err)}") 
        
    def store_screenshot(self, path):
        self.browser_lib.screenshot(filename=path)
    
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
        
        scraper.sort_news_by_newest()

        scraper.filter_sections(section_list=sections)
        
        scraper.load_news_in_month_range(num_months=months)
        
        scraper.get_news_information(phrase=search_phrase, num_months=months)
        
        scraper.save_news_image(path=f"output")
        
        scraper.save_information_in_excel(path="output/news_information.xlsx")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        scraper.finish_execution()
        

if __name__ == "__main__":
    main()