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
from dateutil.relativedelta import relativedelta

@task
def automation_robot():

    # Configure the browser to run in headless mode
    browser.configure(
        headless=True
    )

    #Testing items
    #cloud_search_phrase = "Miley"
    #cloud_number_of_months = "12"

    # Retrieve the search phrase and number of months from the work item payload
    cloud_search_phrase, cloud_number_of_months = handle_item()

    # Open the website
    open_website()

    # Check if there are items to process
    there_are_items_to_process = insert_query(cloud_search_phrase)
    if there_are_items_to_process:

        # Select the option to sort the news articles by date
        select_newest_news()

        # Calculate the limit date based on the number of months provided
        limit_date = get_limit_date(int(cloud_number_of_months))

        # Retrieve a list of articles from the website within the specified time frame
        list_of_articles = get_list_of_articles(cloud_number_of_months, limit_date)

        # Creates the excel used by the process
        excel = create_excel_file()

        # Insert rows into the Excel file containing information about the articles
        excel_insert_rows(excel, list_of_articles,cloud_search_phrase,limit_date)

def handle_item():
    # Retrieve the current work item
    item = workitems.inputs.current

    # Extract the payload from the work item
    payload = item.payload

    # Extract the search phrase from the payload
    search_phrase = payload["SearchPhrase"]

    # Extract the number of months from the payload
    number_of_months = payload["NumberOfMonths"]

    # Return the search phrase and number of months as a tuple
    return search_phrase, number_of_months

def load_page():
    try:
        # Attempt to navigate to the specified URL
        browser.goto("https://www.aljazeera.com")
    except:
        # If an error occurs during navigation, wait for 20 seconds to see if the browser can load properly
        time.sleep(20)
        # Get the current page object
        page = browser.page()
        # Locate the search trigger element on the page header
        element = page.locator(".site-header__search-trigger")
        # If the search trigger element is not visible, raise a ValueError indicating that the page could not load properly
        if not element.is_visible:
            raise ValueError("The page could not load properly.")

def open_website():
    # Try to open and load the browser. The maximum number of retry attempts is set to 3.
    max_retries = 3  # Maximum number of retry attempts
    retries = 0  # Initialize the retry counter
    
    # Loop until the maximum number of retry attempts is reached
    while retries < max_retries:
        try:
            # Attempt to load the page
            load_page()
            # If successful, exit the loop
            break
        except Exception as e:
            # If an error occurs during page loading
            print(f"An error occurred: {e}")  # Print the error message
            retries += 1  # Increment the retry counter
            if retries < max_retries:
                # If there are remaining retry attempts, print a message indicating a retry is being attempted
                print(f"Retrying... Attempt {retries + 1}")
                time.sleep(5)  # Wait for a short time before retrying
            else:
                # If maximum retries reached, print a message and exit the loop
                print("Maximum retries reached. Exiting.")
                break

def insert_query(cloud_search_phrase):
    # Get the current page
    page = browser.page()
    
    # Click on the search trigger element in the site header
    page.click(".site-header__search-trigger")
    
    # Fill in the search bar input field with the provided search phrase
    page.fill(".search-bar__input", cloud_search_phrase)
    
    # Simulate pressing the Enter key to perform the search
    page.keyboard.press("Enter")
    
    # Wait for a short time for the page to load after performing the search
    time.sleep(5)
    
    # Check if there are no search results (indicated by the absence of a specific text element)
    no_results = page.is_visible('text="About 0 results"', timeout=6000)
    if no_results:
        # If no results are found, print a message indicating this
        print("About 0 results were found for search phrase: " + cloud_search_phrase)
        return False  # Return False to indicate no results found
    
    # If search is successful (i.e., results are found), print a success message
    print("Successfully searched for: " + cloud_search_phrase)
    return True  # Return True to indicate successful search

def select_newest_news():
    page = browser.page()

    # Select the option to order news articles by date
    page.select_option("#search-sort-option", "Date")

def get_list_of_articles(cloud_number_of_months, limit_date):
    # Initialize a variable to control the expansion of the article list
    keep_expanding = True
    
    print("Getting the list of articles...")

    # Continue expanding the list of articles until it's no longer needed
    while keep_expanding:

        page = browser.page()
        
        # Get the HTML content of the element containing the search results
        locator = page.locator(".search-results").inner_html()
        
        # Split the HTML content to extract individual articles
        list_of_articles = locator.split('<article class="gc u-clickable-card gc--type-customsearch#result gc--list gc--with-image">')[1:]
        
        # Check if further expansion is needed based on the limit date
        keep_expanding = expand_list_if_needed(list_of_articles, limit_date)
    
    # Return the list of articles
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
    # Define the date format
    format = "%d %b %Y"

    # Get the current date and time in Qatar timezone
    now_qatar = get_current_date_qatar()

    # Check if the description indicates a recent time (minutes or hours)
    if "minute" in date or "hour" in date:
        # Return the current date and time in the specified format
        return now_qatar.strftime(format) 
    else:
        # Extract the number of days from the description
        days_to_subtract = re.search(r'\d+', date)

        if days_to_subtract:
            # Convert the extracted days to an integer
            days_to_subtract = int(days_to_subtract.group())
            # Subtract the extracted number of days from the current date and time
            subtracted_date = now_qatar - datetime.timedelta(days=days_to_subtract)

            # Format the subtracted date and time according to the specified format
            formatted_date = subtracted_date.strftime(format)
            # Return the formatted date and time
            return formatted_date
        else:
            # Raise a ValueError if no number of days is found in the date string
            raise ValueError("No number of days found in the date string")

def extract_article_details(pattern, int_match_group, article):
        match = re.search(pattern, article)
        if match:
            return match.group(int_match_group)
        else:
            return None

def download_image(image_url, file_name):
    try:
        # Send a GET request to the specified image URL and retrieve the image data
        img_data = requests.get(image_url).content
        
        # Open a file in binary write mode in the 'output' directory and save the image data
        with open(os.path.join("output", file_name), 'wb') as handler:
            handler.write(img_data)
    except Exception as e:
        # If an error occurs during the image downloading process, raise a custom error message
        raise Exception(f"Error downloading image from URL '{image_url}': {e}")

def count_substring_occurrences(search_phrase, title, description):
    # Combine the title and description into a single string
    full_text = "{} {}".format(title, description)
    
    # Convert the full text and search phrase to lowercase for case-insensitive matching
    full_text = full_text.lower()
    search_phrase = search_phrase.lower()
    
    # Count the number of occurrences of the search phrase in the full text
    return full_text.count(search_phrase)

def extract_date(description):
    # Define regex patterns to match date formats
    pattern = r'\b\d+\s*(?:hour|minute|day)s?\s*ago\b'
    pattern_Two = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2}, \d{4}'

    # Attempt to match the first pattern in the description
    match = re.search(pattern, description)
    if match:
        # If a match is found, extract the date
        date = match.group(0)
    else:
        # If the first pattern does not match, try the second pattern
        match = re.search(pattern_Two, description)
        if match:
            # If a match is found with the second pattern, extract the date
            date = match.group(0)
            # Convert the extracted date string to a datetime object
            date_obj = datetime.datetime.strptime(date, '%b %d, %Y')
            # Format the datetime object as 'dd mmm YYYY' (e.g., '25 Apr 2024')
            return date_obj.strftime('%d %b %Y')
        else:
            # If neither pattern matches, print a message and return None
            print("Date was not found in the description...")
            return None

    # If a date is extracted, pass it to the calculate_date_from_description function for further processing
    return calculate_date_from_description(date)

def contains_money_amount(description):
    pattern1 = r'\b\d+\s*(?:dollars?|usd)\b'
    pattern2 = r'\$\d+.*'

    #Check if description has money amounts
    if re.search(pattern1, description, re.IGNORECASE) or re.search(pattern2, description):
        return True
    return False

def excel_insert_rows(excel, list_of_articles, cloud_search_phrase, limit_date):
# Iterate over the list of articles
    for index, article in enumerate(list_of_articles):

        try:
            # Print the processing status of the current article
            print(f"Processing article number {index+1}")

            # Replace special characters in the article
            article = article.replace("\xad", "")

            # TITLE
            # Define the pattern to match the title within <span> tags
            pattern = r'<span>(.*?)</span>'
            # Extract the title from the article using the defined pattern
            title = extract_article_details(pattern, 1, article)

            # DESCRIPTION
            pattern = r'<p>(.*?)</p>'
            # Extract the description from the article using the defined pattern
            description = extract_article_details(pattern, 1, article)
            # If HTML tags are present, remove them from the description
            description = re.sub(r'<.*?>|&\w+;', '', description)

            # DATE
            # Extract the date from the description
            date = extract_date(description)
            # Check if the article date is greater than the limit date
            if not is_article_date_bigger_than_limit_date(date, limit_date):
                # If the article date is not greater than the limit date, break the loop
                break

            # SRC
            pattern = r'(?:src=")([^"]+)"'
            # Extract the image URL from the article using the defined pattern
            image_url = extract_article_details(pattern, 1, article)

            # ALT
            pattern = r'alt="([^"]+)"'
            # Extract the alt text from the article using the defined pattern
            file_name = extract_article_details(pattern, 1, article)
            if file_name:
                # Remove special characters from the file name
                file_name = re.sub(r'[^a-zA-Z0-9\s]', '', file_name)
                # Truncate the file name if its length exceeds 100 characters
                if len(file_name) > 100:
                    words = file_name.split()
                    smaller_file_name = ' '.join(words[:3])
                    file_name = smaller_file_name
                # Append ".jpg" extension to the file name
                file_name += ".jpg"
            else:
                # If no alt text is found, use a default file name
                file_name = f"Article {index+1}.jpg"

            # Download the image associated with the article
            download_image(image_url, file_name)

            # Count the occurrences of the search phrase in the title and description
            count_of_phrases = count_substring_occurrences(cloud_search_phrase, title, description)
            # Check if the description contains a money amount
            has_money_amount = contains_money_amount(description)
            
            # Create a row containing title, date, description, image file name, count of phrases, and money amount flag
            row = [title, date, description, file_name, count_of_phrases, has_money_amount]

            # Append the row to the Excel worksheet
            excel.append_rows_to_worksheet([row])

            # Save the workbook
            excel.save_workbook("output/news_report.xlsx")

        except Exception as e:
            print(f"An error occurred for article number {index+1}: {e}")


def is_article_date_bigger_than_limit_date(date_from_article, limit_date):
    #Convert date from article to datetime
    date_object_article = datetime.datetime.strptime(date_from_article, "%d %b %Y")

    if date_object_article>limit_date:
        return True
    else:
        return False

def expand_list_if_needed(list_of_articles, limit_date):
    # Get the last article from the list
    last_article = list_of_articles[-1]

    # DESCRIPTION
    # Define the pattern to extract the description from the article
    pattern = r'<p>(.*?)</p>'
    # Extract the description from the last article using the defined pattern
    description = extract_article_details(pattern, 1, last_article)
    # If HTML tags are present, remove them from the description
    description = re.sub(r'<.*?>|&\w+;', '', description)

    # DATE
    # Extract the date from the description of the last article
    date_from_last_article = extract_date(description)

    # Check if the date of the last article is greater than the limit date
    should_click_show_more = is_article_date_bigger_than_limit_date(date_from_last_article, limit_date)

    if should_click_show_more:
        # If the date of the last article is greater than the limit date, attempt to click the 'Show more' button
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
        # If the date of the last article is not greater than the limit date, return False
        print("No need to click 'Show more'.")
        return False

def get_current_date_qatar():

    # Current time in UTC
    now_utc = datetime.datetime.now(timezone('UTC'))

    # Convert to Asia/Qatar time zone
    now_qatar = now_utc.astimezone(timezone('Asia/Qatar')).replace(tzinfo=None)
    return now_qatar     

def get_limit_date(cloud_number_of_months):
    # Get the current time in UTC
    now_utc = datetime.datetime.now(timezone('UTC'))

    # Convert the current time to the Asia/Qatar time zone
    now_qatar = now_utc.astimezone(timezone('Asia/Qatar'))

    if cloud_number_of_months > 1:
        # If the specified number of months is greater than 1, subtract the corresponding number of months
        now_qatar = now_qatar - relativedelta(months=cloud_number_of_months - 1)
    
    # Set the day of the month to 1 (e.g., first day of the month) and remove the time zone information
    return now_qatar.replace(day=1).replace(tzinfo=None)
