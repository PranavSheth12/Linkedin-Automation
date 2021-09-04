# importing modules
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
import os
from csv import writer
import xlrd


# getting current working directory
cwd = os.getcwd()


# getting location of respective url
loc = cwd + r'\configurations\Book.xlsx'
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

# open xlsx workbook
wb = xlrd.open_workbook(loc)
worksheet = wb.sheet_by_index(0)
worksheet.cell_value(0, 0)

# getting number of requests 
no_of_req = open(cwd + r"\configurations\number_of_requests.txt", "r")
no_of_req = int(no_of_req.read())


# initialize empty list of urls
target_links = []

for i in range(worksheet.nrows):
    a = (worksheet.cell_value(i, 0))
    target_links.append(a)

# initialise empty list of profile urls
profile_urls = []


# define function to write results back to the existing csv file
def append_list_as_row(file_name, list_of_elem):
    # Open file in append mode
    with open(file_name, 'a+', newline='') as write_obj:
        # Create a writer object from csv module
        csv_writer = writer(write_obj)
        # Add contents of list as last row in the csv file
        csv_writer.writerow(list_of_elem)


# opening message text file which contains our personal customizable message
msg = open(cwd + r"\configurations\Message.txt", "r")
client_msg = msg.read()

# OPENING TIMING FILE
timing = open(cwd + r"\configurations\time.txt", 'r')
timing = timing.read()

# opening login credentials text file which contains login information
un = open(cwd + r"\configurations\login.txt", "r")
user = un.readlines()  # it will return a list that contains every line as individual string

# initialize empty list of comments
comments_url = []

# call browser
browser = webdriver.Chrome(executable_path=cwd + r'\configurations\chromedriver.exe')
# action = ActionChains(browser)
# headers = {"User-Agent": "Mozilla/5.0 (X11; U; Linux i686) Gecko/20071127 Firefox/2.0.0.11"}

# getting linkedin login page
browser.get("https://www.linkedin.com/login")
username = browser.find_element_by_id('username')
username.send_keys(user[0])
password = browser.find_element_by_id('password')
password.send_keys(user[1])
time.sleep(1)


time.sleep(3)

for link in target_links:

    # looping through 1 to 100 pages
    for i in range(2, 101):


        browser.get(link + '&page=' + str(i))
        time.sleep(2)

        # decreasing size of our browser so every 10 profile on 1 page can be fit into screen
        browser.execute_script("document.body.style.zoom='25%'")

        # finding all the containers
        profile_containers = browser.find_elements_by_class_name("reusable-search__result-container")

        # initialising profile links list
        profile_links = []

        # looping through each container
        for container in profile_containers:
            try:
                profile1 = container.find_element_by_class_name('entity-result__content')
                profile_url1 = profile1.find_element_by_class_name('app-aware-link')
                profile_url = profile_url1.get_attribute('href')
                print(profile_url)

                profile_links.append(profile_url)
            except:
                continue

        # getting browser to our normal version
        browser.execute_script("document.body.style.zoom='100%'")

        # looping through each link in profile_link list
        for links in profile_links:
            time.sleep(float(timing))
            if no_of_req == 0 :
                browser.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                time.sleep(2)
                current_connections = browser.find_element_by_class_name("mn-connections__header").text
                current_connections = current_connections.split(" ")[0].replace(',', '')
                current_connections = int(current_connections)
                last_count = current_connections
                time.sleep(30)
                while no_of_req==0:
                    try:
                        browser.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                        time.sleep(2)
                        current_connections = browser.find_element_by_class_name("mn-connections__header").text
                        current_connections = current_connections.split(" ")[0].replace(',', '')
                        current_connections = int(current_connections)
                        if last_count < current_connections:
                            no_of_req = abs(current_connections - last_count)
                            continue
                        last_count = current_connections
                        # no_of_req = checkout(browser, last_count)
                        time.sleep(30)
                    except:
                        pass


            try:
                # getting every 10 links on every page one by one
                try:
                    browser.get(links)
                    time.sleep(3)
                    name = browser.find_element_by_css_selector('.t-24').text
                    print(name)
                    print(no_of_req)
                    first_name = name.split()[0]
                except:
                    continue

                # clicking connect button

                try:
                    connect3 = browser.find_element_by_css_selector(".artdeco-button--primary .artdeco-button__text")

                    connect3.click()
                except:
                    print('connect button not found')
                    continue

                # note_button click to write our personal invitation

                try:
                    note_button = browser.find_element_by_xpath('/html/body/div[3]/div/div/div[3]/button[1]')

                    note_button.click()
                except:
                    print('Request already sent!')
                    continue

                # typing our personalized message

                try:
                    text = browser.find_element_by_xpath('/html/body/div[3]/div/div/div[2]/div[1]/textarea')
                    message = 'Hello ' + first_name + '!' + '\n' + client_msg
                    text.send_keys(message)
                    time.sleep(5)
                except:
                    continue

                # clicking submit button

                try:
                    submit = browser.find_element_by_xpath('/html/body/div[3]/div/div/div[3]/button[2]')
                    submit.click()
                    no_of_req = no_of_req-1
                except:
                    continue

                # getting a list that contains profile_url and our message we sent just

                content = [links, message]

                # writing back our result to existing csv file

                append_list_as_row(cwd + r'\configuration\Sent_Requests.csv', content)
            except:
                print("Couldn't retrived url")
