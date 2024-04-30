from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Files import Files
import time
import re
import datetime
from pytz import timezone
import os
import requests
from robocorp import workitems
import json
from dateutil.relativedelta import relativedelta

@task
def automation_robot():

    browser.configure(
	    headless=True
        )
    #cloud_search_phrase = handle_item()
    #cloud_search_phrase = "money"
    #cloud_number_of_months = "6"
    cloud_search_phrase, cloud_number_of_months = handle_item()
    open_website()
    there_are_items_to_process = insert_query(cloud_search_phrase)
    if there_are_items_to_process:
        select_newest_news()
        limit_date = get_limit_date(int(cloud_number_of_months))
        list_of_articles = get_list_of_articles(cloud_number_of_months, limit_date)
        excel = create_excel_file()
        excel_insert_rows(excel, list_of_articles,cloud_search_phrase,limit_date)

def handle_item():
    item = workitems.inputs.current
    payload = item.payload
    search_phrase = payload["SearchPhrase"]
    number_of_months = payload["NumberOfMonths"]
    print("Received payload:", search_phrase)
    return search_phrase, number_of_months

def load_page():
    try:
        browser.goto("https://www.aljazeera.com")
    except:
        # Wait for more 20 seconds to see if browser can load properly and try to interact with it
        time.sleep(20)
        page = browser.page()
        element = page.locator(".site-header__search-trigger")
        if not element.is_visible:
            raise ValueError("The page could not load properly.")

def open_website():
    # Try to open and load the browser. The maximum number of retry attempts is set to 3.
    max_retries = 3
    retries = 0
    
    while retries < max_retries:
        try:
            load_page()
            break
        except Exception as e:
            print(f"An error occurred: {e}")
            retries += 1
            if retries < max_retries:
                print(f"Retrying... Attempt {retries + 1}")
                time.sleep(5)  # Wait for a short time before retrying
            else:
                print("Maximum retries reached. Exiting.")
                break

def insert_query(cloud_search_phrase):
    page = browser.page()
    page.click(".site-header__search-trigger")
    page.fill(".search-bar__input", cloud_search_phrase)
    page.keyboard.press("Enter")
    #Wait a little for the page to load
    time.sleep(5)
    no_results = page.is_visible('text="About 0 results"', timeout=6000)
    if no_results:
        print("About 0 results were found for search phrase: " + cloud_search_phrase)
        return False

    print("Successfully searched for: " + cloud_search_phrase)
    return True

def select_newest_news():
    page = browser.page()
    page.select_option("#search-sort-option", "Date")

def get_list_of_articles(cloud_number_of_months, limit_date):

    keep_expanding = True
    print("Getting the list of articles...")

    while keep_expanding:
        page = browser.page()
        locator = page.locator(".search-results").inner_html()
        list_of_articles = locator.split('<article class="gc u-clickable-card gc--type-customsearch#result gc--list gc--with-image">')[1:]
        keep_expanding = expand_list_if_needed(list_of_articles,limit_date)
    return list_of_articles

def create_excel_file():
    print("Creating excel file...")

    # Create an instance of the RPA.Excel.Files library
    excel = Files()

    # Define the file path for the Excel file
    file_path = 'output/news_report.xlsx'

    # Create a new workbook at the specified file path
    excel.create_workbook(file_path)

    # Define the headers for the worksheet
    headers = ["Title", "Date", "Description", "Picture Filename", "Count of Phrases", "Contains Money Amount"]

    # Append the headers row directly as a list of lists
    excel.append_rows_to_worksheet([headers])

    # Save the workbook
    excel.save_workbook(file_path)

    # Return the Excel instance
    return excel

def calculate_date_from_description(date):
    format = "%d %b %Y"

    now_qatar = get_current_date_qatar()

    if "minute" in date or "hour" in date:
        return now_qatar.strftime(format) 
    else:
        # Define the number of days to subtract
        days_to_subtract = re.search(r'\d+', date)

        if days_to_subtract:
            days_to_subtract = int(days_to_subtract.group())
            # Subtract the number of days from now_qatar
            subtracted_date = now_qatar - datetime.timedelta(days=days_to_subtract)

            formatted_date = subtracted_date.strftime(format)
            # print(formatted_date)  # Optional: Uncomment if you want to print the formatted date
            return formatted_date
        else:
            # Default value if no match is found
            raise ValueError("No number of days found in the date string")

def extract_article_details(pattern, int_match_group, article, value_to_look_for):
        match = re.search(pattern, article)
        # If a match is found, extract the title
        if match:
            return match.group(int_match_group)
        else:
            return None

def download_image(image_url, file_name):

    img_data = requests.get(image_url).content
    with open(os.path.join("output", file_name), 'wb') as handler:
        handler.write(img_data)

    print("Image downloaded successfully.")

def extract_article_details(pattern, int_match_group, article, value_to_look_for):
        match = re.search(pattern, article)
        # If a match is found, extract the title
        if match:
            return match.group(int_match_group)
        else:
            return None

def count_substring_occurrences(search_phrase,title,description):
    full_text = "{} {}".format(title, description)

    full_text = full_text.lower()  # Convert text to lowercase
    search_phrase = search_phrase.lower()  # Convert substring to lowercase
    return full_text.count(search_phrase)

def extract_date(description):
    pattern = r'\b\d+\s*(?:hour|minute|day)s?\s*ago\b'
    pattern_Two = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2}, \d{4}'

    match = re.search(pattern, description)
    if match:
        date = match.group(0)
    else:
        match = re.search(pattern_Two, description)
        if match:
            date = match.group(0)
            date_obj = datetime.datetime.strptime(date, '%b %d, %Y')
            return date_obj.strftime('%d %b %Y')
        else:
            print("Date was not found in the description...")
            return None

    return calculate_date_from_description(date)

def contains_money_amount(description):
    pattern1 = r'\b\d+\s*(?:dollars?|usd)\b'
    pattern2 = r'\$\d+.*'

    if re.search(pattern1, description, re.IGNORECASE) or re.search(pattern2, description):
        return True
    return False

def excel_insert_rows(excel, list_of_articles,cloud_search_phrase, limit_date):
    for index, article in enumerate(list_of_articles):

        print(f"Processing article number {index+1}")
        article = article.replace("\xad", "")

        # TITLE
        # Define the pattern to match the title within <span> tags
        pattern = r'<span>(.*?)</span>'
        title = extract_article_details(pattern,1,article,"title")

        # DESCRIPTION
        pattern = r'<p>(.*?)</p>'
        description = extract_article_details(pattern,1,article,"description")
        # If a match is found, extract the title
        description = re.sub(r'<.*?>|&\w+;', '', description)

        # DATE
        date = extract_date(description)

        if not is_article_date_bigger_than_limit_date(date, limit_date):
            break        
        
        #SRC
        pattern = r'(?:src=")([^"]+)"'
        image_url = extract_article_details(pattern,1,article,"source")

        #ALT
        pattern = r'alt="([^"]+)"'
        file_name = extract_article_details(pattern,1,article,"alt")
        if file_name:
            file_name = re.sub(r'[^a-zA-Z0-9\s]', '', file_name)
            if len(file_name) > 100:
                words = file_name.split()
                smaller_file_name = ' '.join(words[:3])
                file_name = smaller_file_name
            file_name += ".jpg"
        else:
            file_name = f"Article {index+1}.jpg"

        download_image(image_url, file_name)
        count_of_phrases = count_substring_occurrences(cloud_search_phrase,title,description)
        has_money_amount = contains_money_amount(description)
        
        # Create a row containing URL, title, date, and description
        row = [title,date,description,file_name,count_of_phrases,has_money_amount]

        excel.append_rows_to_worksheet([row])

        excel.save_workbook("output/excel.xlsx")

def is_article_date_bigger_than_limit_date(date_from_article, limit_date):
    #Convert date from article to datetime
    date_object_article = datetime.datetime.strptime(date_from_article, "%d %b %Y")

    if date_object_article>limit_date:
        return True
    else:
        return False

def expand_list_if_needed(list_of_articles, limit_date):
    #Get the last article from list
    last_article = list_of_articles[-1]

    # DESCRIPTION
    pattern = r'<p>(.*?)</p>'
    description = extract_article_details(pattern,1,last_article,"description")
    # If a match is found, extract the title
    description = re.sub(r'<.*?>|&\w+;', '', description)

    # DATE
    date_from_last_article = extract_date(description)

    should_click_show_more = is_article_date_bigger_than_limit_date(date_from_last_article, limit_date)

    if should_click_show_more:
        print("Checking if 'Show more' button is available...")
        page = browser.page()

        try:
            # Wait for the 'Show more' button to become visible
            page.wait_for_selector('text="Show more"', timeout=5000)
            
            # If the button becomes visible within the timeout period, click it
            print("'Show more' button found. Clicking...")
            page.click('text="Show more"')
            print("Clicked 'Show more'.")
            return True
        except:
            # If the button does not become visible within the timeout period, print a message and return False
            print("No more articles to display. The page has reached its limit.")
            return False
    else:
        print("No need to click 'Show more'.")
        return False


def get_current_date_qatar():

    # Current time in UTC
    now_utc = datetime.datetime.now(timezone('UTC'))

    # Convert to Asia/Qatar time zone
    now_qatar = now_utc.astimezone(timezone('Asia/Qatar')).replace(tzinfo=None)
    return now_qatar     

def get_limit_date(cloud_number_of_months):
    # Current time in UTC
    now_utc = datetime.datetime.now(timezone('UTC'))

    # Convert to Asia/Qatar time zone
    now_qatar = now_utc.astimezone(timezone('Asia/Qatar'))

    if cloud_number_of_months > 1:
        #days_to_subtract = (cloud_number_of_months - 1) * 30
        #now_qatar = now_qatar - datetime.timedelta(days=days_to_subtract)
        now_qatar = now_qatar - relativedelta(months=cloud_number_of_months-1)
    
    return now_qatar.replace(day=1).replace(tzinfo=None)