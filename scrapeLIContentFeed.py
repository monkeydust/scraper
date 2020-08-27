#! Python3
import time
import bs4, sys, os
import argparse
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from msedge.selenium_tools import Edge, EdgeOptions
import smtplib
from email.message import EmailMessage

# Instantiate the parser
parser = argparse.ArgumentParser(
    description="Personal CLI for fetching posts from LinkedIn feed."
)

# Required keyoword arguments
parser.add_argument(
    "--keywords", nargs="+", help="Keywords to search feed for.", required=True
)

# Optional depth argument
parser.add_argument(
    "--depth", type=int, help="Number of times PAGE_DOWN should be executed."
)

# Optional order argument
parser.add_argument(
    "--order", help="If set to 'date' will sort by recent results instead of relevant (default) "
)

# Set file name for job
def set_file_name():
    searchName= get_search_str().replace(" ", "").replace("%20","")
    file_name = f"Feed-{searchName}-{date.today().strftime('%d-%m-%Y')}.txt"
    return file_name


# Launch Microsoft Edge (Chromium)
options = EdgeOptions()
options.use_chromium = True
#options.binary_location = "C:\\Drivers\\msedgedriver.exe"
options.add_argument("headless")
browser = Edge(options = options)

def get_search_str():
    """Composes URL embeddable string (search query) from keywords"""

    searchString = ""
    default_search_string = "Fixed Income Automation"
    keywords = parser.parse_args().keywords or []
    
    # Get content keywords from command line if provided and form url search string
    if len(keywords) > 0:
        for keyword in keywords:
            searchString = searchString + "%20" + keyword

        return searchString

    return default_search_string


def get_depth():
    """Depth number determines how many times we should perform PAGE_DOWN (to load new posts)"""

    default_depth = 2
    depth = parser.parse_args().depth

    return depth or default_depth


def get_text(element):
    """Custom utility error-safe wrapper around BeautifulSoups's get_text() method
       Returns stripped string or None if the element does not exist."""

    if not element:
        return None

    return element.get_text().strip()


def save_post(text):
    """Saves post to a text file"""
    file_name = set_file_name() 
    posts_separator = "-" * 60 + "\n"

    # Explicitly set encoding (otherwise we encounter an issue on Windows)
    with open(file_name, "a+", encoding="utf-8") as file:
        file.write(posts_separator)
        file.write(text + "\n")
        file.close()


def login_linkedin():
    browser.get(
        "https://www.linkedin.com/login?fromSignIn=true&trk=guest_homepage-basic_nav-header-signin"
    )

    try:
        userElem = browser.find_element_by_id("username")
        userElem.send_keys("mailsamd@gmail.com")
        pwdElem = browser.find_element_by_id("password")
        pwdElem.send_keys("gBrom275!!")
        ### Alternative: loginElem = browser.find_element_by_xpath('//*[@id="app__container"]/main/div[2]/form/div[3]')
        loginElem = browser.find_element_by_xpath('//button[text()="Sign in"]').click()
    except:
        print("Error Logging in")


def fetch_posts():
    """Fetches posts from LinkedIn feed and saves them to a text file"""
    login_linkedin()
    searchString = get_search_str()
    
    # Check if sort order overide to sort by recent
    order = parser.parse_args().order or []
    print(order)
    
    # Create urlString with search
    if order == 'date':
        urlString = f"https://www.linkedin.com/search/results/content/?facetSortBy=date_posted&keywords=\"{searchString[3:]}\"&origin=SORT_RESULTS"
    else:
        urlString = f"https://www.linkedin.com/search/results/content/?keywords=\"{searchString[3:]}\"&origin=GLOBAL_SEARCH_HEADER"
    
    # Now call LinkedIn with our search string
    browser.get(urlString)
   
    # Scroll <depth> times to the end of the browser viewport to make LinkedIn load new posts
    for i in range(1, get_depth()):
        page = browser.find_element_by_tag_name("html")
        page.send_keys(Keys.END)

        # Wait for LinkedIn to load posts (wait number is in seconds)
        time.sleep(5)

    bsScrape = bs4.BeautifulSoup(browser.page_source, "html.parser")
    posts = bsScrape.select("li.search-content__result")

    return posts


def scrape_posts(posts):
    """Extracts data from each post, composes the final result string and saves it to the file"""
    fullResults=""
    for post in posts:
        post_author_name = get_text(post.select_one("span.feed-shared-actor__name"))
        post_author_description = get_text(
            post.select_one("span.feed-shared-actor__description")
        )
        post_age = get_text(
            post.select_one(
                "span.feed-shared-actor__sub-description span.visually-hidden"
            )
        )
        post_text = get_text(post.select_one('span[class="break-words"]'))

        result = (
            f"{post_author_name}, {post_author_description}, {post_age}\n\n{post_text}"
        )

        save_post(result)


    
def sendViaEmail():
    # Create the email
    msg = EmailMessage()
    msg['Subject'] = f'{file_name[:-15]}'
    msg['From'] = 'auto@fwd.ac'
    msg['To'] = 'jklondon@gmail.com'

   
    # Connect to SMTP Server
    smtpObj = smtplib.SMTP('mail.gandi.net',587)
    print(type(smtpObj))
    smtpObj.ehlo()

    # Start encryption and provide login credentials
    smtpObj.starttls()
    smtpObj.login('auto@fwd.ac', 'Ravi1234')
    
    # Open the file. parse it and attach to the body of the message

    with open(file_name, "r", encoding="utf8") as filename:
        text = filename.read()
        msg.add_header('Conent-Type','text/plain')
        msg.set_payload(text.encode())
    
    # Send email
    print('About to send email')
    smtpObj.send_message(msg)
    print('Email sent')

    # Quit the SMTP 
    smtpObj.quit()

def checkFileExists(file_name):
        # If file exists delete it
    if os.path.exists(file_name):
        os.remove(file_name) #this deletes the file
        print("The search file existed, deleting it") 
    else:
        print("The search file does not exist, will be created") #add this to prevent errors

## Main program flow

# Set the file name to be used 
file_name = set_file_name()

# Check if exists already - if so delete and create new one
checkFileExists(file_name)

# Get posts
posts = fetch_posts()
scrape_posts(posts)

# Close browser
browser.close()

# Send email
sendViaEmail()

# Exit program
sys.exit('Job Complete')
