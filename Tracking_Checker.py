from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import json
import time
import datetime
import os
import re
import collections
import requests
from datetime import datetime
import traceback
import pyexcel



def Set_Chrome():
    path_to_dir = "C:/CustonChromeProfiles"
    profilefolder = "PythonScraper"
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--user-data-dir=' + path_to_dir)
    chrome_options.add_argument('--profile-directory=' + profilefolder)

    print("\nSetting Chrome Options.")
    driver = webdriver.Chrome(chrome_options=chrome_options)
    return driver


def add_to_records(single_record, list_of_details):
    """Append record of a single order to list of records """
    print("\nRunning Function: \"add_to_records\" ")
    try:
        list_of_details.append(single_record)
        print("record added")
        # print("\nCurrent record for all accounts so far")
        print(json.dumps(single_record, indent=2))
    except:
        print("record not added")
    return list_of_details


def gmail_login(driver):
    print("Logging into Gmail")
    driver.get("https://accounts.google.com/signin/v2/identifier?&service=mail")

    page_title = driver.title
    if "random_email@gmail.com.com" not in page_title:
        driver.refresh()

    # find the email input box
    compose_page_xpath = "//div[@class]//*[text()[contains(.,'COMPOSE')]]"
    try:
        compose_page = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, compose_page_xpath)))

    except:
        print("No Compose window")
        print("Have to login first")

        try:
            inputEmail = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierId']")))
            inputEmail.clear()
            inputEmail.send_keys("random_email@gmail.com")
        except Exception as exc:
            print("Error with input Email. ")
            print(exc)

        # Click the next button
        try:
            next_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']/content/span")))
            next_button.click()
        except Exception as exc:
            print("Error with next Button.")
            print(exc)

        # password field will appear.
        try:
            inputPassword = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='password']//input")))
            inputPassword.clear()
            inputPassword.send_keys("randompassword123")
        except Exception as exc:
            print("Error with input password.")
            print(exc)

        # Click login button
        try:
            loginButton_2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']/content/span")))
            time.sleep(1)
            loginButton_2.click()
            time.sleep(2)
        except Exception as exc:
            print("Error with login button.")
            print(exc)

        driver.get('https://mail.google.com/mail/u/0/h/1t9zc9mtg2ong/?zy=e&f=1')
        time.sleep(2)
        driver.refresh()
    return


def post_tracking_information(jsondata):
    url_endpoint = "http://192.168.1.7/api/odata/ShipmentTrackings"

    data = None
    try:
        print(jsondata)
        jsonstr = json.dumps(jsondata)
        print(jsonstr)

        headers = {'Content-type': 'application/json'}
        response = requests.post(url_endpoint, data=jsonstr, headers=headers)

        response.raise_for_status()

        data = response.json()
        print("Data posted.")
    except Exception as e:
        print(" Tracking post error " + str(e))

    return data


def post_tracking_logs(jsondata):
    url_endpoint = "http://192.168.1.7/api/odata/ShipmentTrackingLogs"

    data = None
    try:
        print(jsondata)
        jsonstr = json.dumps(jsondata)
        print(jsonstr)

        headers = {'Content-type': 'application/json'}
        response = requests.post(url_endpoint, data=jsonstr, headers=headers)

        response.raise_for_status()

        data = response.json()

    except Exception as e:
        print(" Tracking post error " + str(e))

    return data


def send_tracking_info_database(tracking_number, carrier_name, status, order_number, shipping_address, referenceID):
    current_date = (datetime.utcnow()).isoformat() + "Z"
    batch = str(current_date)
    check_data = None
    tracking_post_data = {
        "TrackingNumber": tracking_number,
        "CarrierName": carrier_name,
        "CreatedDate": current_date,
        "ShipmentDate": None,
        "Username": "PythonBot",
        "Source": "Python Script",
        "Status": status,
        "ServiceName": None,
        "OrderNumber": order_number,
        "ReferenceId": referenceID,
        "Note": None,
        "MessageId": "",
        "BatchId": batch,

        "ShipFromAddress": None,
        "ShipToAddress": shipping_address,

        "OrderId": None
    }

    if len(tracking_number) != 0:
        print("Checking for existing record")
        check_url = f"http://192.168.1.7/api/odata/ShipmentTrackings?$filter=TrackingNumber eq '{tracking_number}'"
        check_response = ""
        try:
            check_response = requests.get(check_url)
            print(check_response)
        except:
            print("Check failed.")
            check_response = json.dumps(check_response)

        try:
            data = check_response.json()
            check_data = (data['value'])
        except:
            check_data = []

        found = next((i for i in check_data if i['OrderNumber'] == order_number), None)
        # Changed from i['OrderNumber'] == order_number

        if not found:
            print("Existing record for this tracking number with order information not found in the database.\n Posting new record")
            post_tracking_information(tracking_post_data)
        else:
            print("\nRecord(s) for this tracking number exists.")  # Unless its checked in. In that case the tracking cycle has been completed

            index = 1
            add_log = True
            add_address_entry = True

            for entry in check_data:
                print(f"\nChecking record {index}")
                if entry['Status'] != "CheckedIn":
                    print("Item has not been checked in for current record.")

                else:
                    print("Item has been checked in for current record. ")
                    add_log = False

                index = index + 1

            for entry in check_data:
                print(f"\nChecking record {index}")
                if entry['ShipToAddress'] == "":
                    print("Shipping address is blank for current record.")

                else:
                    print("Shipping address is filled for current record. ")
                    add_address_entry = False

                index = index + 1

            if add_address_entry == True:
                print("Adding record with shipping address")
                post_tracking_information(tracking_post_data)

            if add_log == True:
                print("\n All records checked. No \'checked-in\' status available for tracking number. Adding log.")
                log_post_data = {
                    "ShipmentTrackingId": found['Id'],
                    "TrackingNumber": tracking_number,
                    "InputType": status,
                    "UserId": "api_PythonBot",
                    "Source": "pythonscript"
                }

                print("Tracking record exists. Adding log")
                post_tracking_logs(log_post_data)

            else:
                print("Tracking number has check in record. Tracking Cycle complete. Will not add log")
    else:
        print("Tracking number is not defined.")


    return check_data


def process_main_search():
    # ---------Paramaters------------------------------------------------------#

    search_term_NE = '"Tracking" from:info@newegg.com subject:Tracking newer_than:15d'
    search_term_NEB = '"Tracking" from:info@neweggbusiness.com subject:Tracking newer_than:15d'
    search_term_BB = '"Tracking" from:BestBuyInfo@emailinfo.bestbuy.com subject:"has shipped" newer_than:15d'
    search_term_AMZ = '{"Your Amazon.com order"  "Your Prime order"} "Track" subject:has shipped newer_than:15d'
    search_term_WAL = '"Tracking" from:help@walmart.com subject:Shipped newer_than:15d'
    search_term_EBAY = '"Thank you for shopping on eBay! Your order has shipped" subject:ORDER SHIPPED newer_than:20d'
    search_term_STAPLES = '"Items in your Staples Order" "TRACK" newer_than:15d'
    search_term_DELL = '"Dell Order Has Shipped for Dell Purchase ID" newer_than:15d'
    search_term_OD = '"Shipment Confirmation #" "Office Depot order has shipped" newer_than:15d'
    search_term_RAKUTEN = '"Your order with Rakuten.com has shipped" newer_than:20d'
    search_term_ANTOnline = '"from ANTOnline" "ORDER SHIPPED" newer_than:10d'
    search_term_target = 'subject: Good news! Your order "is on its way" "target" newer_than:30d'
    search_term_target2 = '"Good news! An item has shipped from your order #" "target" newer_than:30d'


    search_list = [search_term_WAL,search_term_BB,  search_term_NE, search_term_NEB, search_term_AMZ,
                   search_term_EBAY, search_term_STAPLES, search_term_DELL,
                   search_term_OD, search_term_RAKUTEN, search_term_ANTOnline, search_term_target, search_term_target2
                   ]

    all_order_records_list = []

    driver = Set_Chrome()
    print(driver.get_window_size())

    print("Getting search results from gmail")
    gmail_login(driver)
    # driver.get(search_string_url)
    driver.get("https://mail.google.com/mail/u/0/h/1t9zc9mtg2ong/?zy=e&f=1")

    for search_term in search_list:
        print("Finding Search Box")
        search_box_xpath = "(//input[@title='Search'])"
        try:
            search_box = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.XPATH, search_box_xpath)))
            print("Search box found")
            time.sleep(1)
            search_box.clear()
            search_box.click()
            search_box.send_keys(search_term)
        except:
            print("search box failed")

        print("Clicking on search button")
        search_button_xpath = "(//input[@value='Search Mail'])"
        try:
            email_submit_button = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.XPATH, search_button_xpath)))
            print("Submit button found")
            time.sleep(1)
            email_submit_button.click()
        except:
            print("Search button failed")

        search_string_url = driver.current_url
        if "newegg" in search_term:
            process_Newegg_pages(driver, search_string_url, all_order_records_list)
        elif "bestbuy" in search_term:
            process_Bestbuy_pages(driver, search_string_url, all_order_records_list)
        elif "Amazon" in search_term:
            process_Amazon_pages(driver, search_string_url, all_order_records_list)
        elif "walmart" in search_term:
            process_Walmart_pages(driver, search_string_url, all_order_records_list)
        elif "eBay" in search_term:
            process_ebay_pages(driver, search_string_url, all_order_records_list)
        elif "Staples " in search_term:
            process_Staples_pages(driver, search_string_url, all_order_records_list)
        elif "Dell " in search_term:
            process_Dell_pages(driver, search_string_url, all_order_records_list)
        elif "Office Depot" in search_term:
            process_OD_pages(driver, search_string_url, all_order_records_list)
        elif "neweggbusiness" in search_term:
            process_Newegg_pages(driver, search_string_url, all_order_records_list)
        elif "Rakuten" in search_term:
            process_Rakuten_pages(driver, search_string_url, all_order_records_list)
        elif "ANTOnline" in search_term:
            process_ANTOnline_pages(driver, search_string_url, all_order_records_list)
        elif "target" in search_term:
            process_Target_pages(driver, search_string_url, all_order_records_list)

    print("Saving results")
    time_run = time.strftime("%Y-%b-%d__%H_%M_%S", time.localtime())
    try:
        dest_file_name = "C:/saveToFile/package_status/" + "package_status_" + time_run + ".xlsx"
        pyexcel.save_as(records=all_order_records_list, dest_file_name=dest_file_name)
        print(f"Results saved to{dest_file_name}.")
    except:
        print("Save directory not present. Making directory")
        os.makedirs('C:/saveToFile/package_status/')
        try:
            dest_file_name = "C:/saveToFile/package_status/" + "package_status_" + time_run + ".xlsx"
            pyexcel.save_as(records=all_order_records_list, dest_file_name=dest_file_name)
            print(f"Results saved to{dest_file_name}.")
        except:
            print("Save failed")

    driver.quit()


def process_Target_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = re.compile(r"Order\s*#\s+([\s\S]*?)\s")
        try:
            order_number = (order_number_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Order Number: {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Tracking\s*#\s*([\d\w]+)")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number: {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Delivers\s+to:[\s\S]*?,\s+(.+)")
        shipping_address = "Nothing"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Shipping Address: {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ")
            next_page = False

        email_index = email_index + 1


def process_ANTOnline_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    order_number = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Reference ID")
        referenceID = ""
        referenceID_regex = re.compile(r"Order\s*Number:\s*([\s\S]*?)\s")
        try:
            referenceID = (referenceID_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"ReferenceID: {referenceID}")
        except:
            print("Getting ReferenceID fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Tracking:\s*([\d\w]+)")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number: {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Ship\s*To\s*.+\s*([\s\S]*?)Payment\s")
        shipping_address = "Null"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Shipping Address: {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ")
            next_page = False

        email_index = email_index + 1


def process_Rakuten_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = re.compile(r"Order\s*#:([\s\S]*?)\s")
        try:
            order_number = (order_number_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Order Number: {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Tracking\s*Number\s*([\d\w]+)")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number: {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Ship\s*To\s*.+\s*([\s\S]*?)Payment\s")
        shipping_address = "Nothing"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Shipping Address: {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ")
            next_page = False

        email_index = email_index + 1


def process_OD_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = re.compile(r"Order\s{0,}Number:\s{0,}([\d-]+)")
        try:
            order_number = (order_number_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Order Number {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Tracking\s{0,}number:\s{0,}([\d\w]+)")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Shipping\s{0,}to:.+\s{0,}([\s\S]*?)Account")
        shipping_address = "Nothing"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Shipping Address {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_Walmart_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = "Order\s+number:\s+(\d{7}-\d{6})"
        try:
            order_number = (re.findall(order_number_regex, message_body_field_message))[0]
            time.sleep(1)
            print(f"Order Number {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = "tracking\s+number\s+(.+)"
        try:
            tracking_number_found = ((re.findall(tracking_number_regex, message_body_field_message)))
            tracking_number_list = tracking_number_found
            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = "Shipped\s+by\s+.+\s+.+\s+([\s\S]*?\d+)\s+\w+\s+tracking"
        shipping_address = "Nothing"
        try:
            shipping_address = ((re.findall(shipping_text_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Shipping Address {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_Dell_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = re.compile(r"Order\s+number\s+(\d+)")
        try:
            order_number = (order_number_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Order Number {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Carrier\s{0,}tracking\s{0,}(.+)")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Shipping\s{0,}.+\s{0,}([\s\S]*?)Carrier")
        shipping_address = "Nothing"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            time.sleep(1)
            print(f"Shipping Address {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_Staples_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Order Number")
        order_number = ""
        order_number_xpath = "/html/body/table[3]/tbody/tr/td[2]/table[1]/tbody/tr/td[2]/table[3]/tbody/tr/td/h2/font/b"
        order_number_regex = re.compile(r"Order\s{0,}(\d+)")
        try:
            order_number_text = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, order_number_xpath))).text
            order_number = order_number_regex.search(order_number_text).group(1)
            time.sleep(1)
            print(f"Order Number: {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text")

        print("Getting Tracking Number")
        tracking_number_list = []
        # tracking_number_regex = re.compile(r"Tracking\s{0,}#:\s+(.+)")
        tracking_number_regex = re.compile(r"Tracking\s{0,}#:\s+([\s\S]*?)TRACK")
        try:
            tracking_number_found = (tracking_number_regex.search(message_body_field_message).group(1)).strip()
            tracking_number_list = tracking_number_found.split('\n')
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting shipping address")
        shipping_text_regex = re.compile(r"Ship\s{0,}.+\s{0,}([\s\S]*?)Carrier")
        shipping_address = "Nothing"
        try:
            shipping_address = (shipping_text_regex.search(message_body_field_message).group(1)).strip()
            print(f"Shipping Address {shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_Amazon_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing..."

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = "Order\s#([\w\d-]+)"
        try:
            order_number = (re.findall(order_number_regex, message_body_field_message))[0]
            time.sleep(1)
            print(f"Order Number: {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Shipping address")
        shipping_address = ""
        shipping_address_regex1 = "SHIP\s+TO\s+.+(\s+.+)"
        shipping_address_regex2 = "Order\s#[\w\d-]+\s+.+\s(.+\s.+\s+.+)"
        try:
            shipping_address = ((re.findall(shipping_address_regex1, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Shipping Address:\n{shipping_address}")
        except:
            print("Getting shipping address fail")

        if len(shipping_address) < 2:
            print("Alternate amazon tracking email.")
            try:
                shipping_address = ((re.findall(shipping_address_regex2, message_body_field_message))[0]).strip()
                time.sleep(1)
                print(f"Shipping Address:\n{shipping_address}")
            except:
                print("Getting shipping address fail")

        # About to click to new page. Need to get current page
        print("Getting current email position")
        current_email_position = driver.current_url

        print("Getting Tracking Number.\nFinding tracking button.")
        amazon_tracking_button_xpath = "(//*[text()[contains(.,'Track')]]|//*[contains(@alt,'Track')])"
        try:
            amazon_tracking_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, amazon_tracking_button_xpath)))
            print("Tracking button found")
            time.sleep(1)
            print("Clicking email")
            amazon_tracking_button.click()
        except Exception as exc:
            print("Clicking tracking button failed.")
            print(exc)

        # On New Webpaage. Switch to it.
        try:
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(2)
            driver.refresh()
        except:
            print("Tab switch failed")

        print("Getting tracking number from page.")
        tracking_number_xpath = '//*[@id="carrierRelatedInfo-container"]/div/span/a'
        tracking_number_regex = 'Tracking\s+ID\s+([\w\d]+)'
        tracking_number_list = []
        try:
            tracking_number_found = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, tracking_number_xpath))).text
            tracking_number = (re.findall(tracking_number_regex, tracking_number_found))[0]
            tracking_number_list = tracking_number.split(',')
            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting shipping page text")
        shipping_page_xpath = "//span[@id='primaryStatus']"
        try:
            shipping_page_text = driver.find_element_by_xpath(shipping_page_xpath).text
        except:
            print("Copy Page fail")
            shipping_page_text = "Nothing..."

        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)
        # close new tab that contains new webpage and switch back to original tab

        driver.close()
        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Going back to Gmail")
        # driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_Bestbuy_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page.....")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing...:-("

        print("Working on body text.....")
        print("Getting Order Number.....")
        order_number = ""
        order_number_regex = "ORDER\s#\s+([\w\d-]+)"
        try:
            order_number = (re.findall(order_number_regex, message_body_field_message))[0]
            time.sleep(1)
            print(f"Order Number {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number.....")
        tracking_number_list = []
        tracking_number_regex = "TRACKING\s#\s+([\w\d]+)"
        try:
            tracking_number_found = ((re.findall(tracking_number_regex, message_body_field_message))[0]).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number: {tracking_number_list}")
        except:
            print("Getting tracking number fail")
            tracking_number_list[0] = " "

        print("Getting shipping address....")
        shipping_text_regex = "ORDER\s+SHIPPED\s+ON:\s+\w{3}\s+\d\d/\d\d\s+.+([\s\S]*?\w\w\s+\d{5})"
        shipping_address = "Nothing"
        try:
            shipping_address = ((re.findall(shipping_text_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Shipping Address:\n{shipping_address}")
        except Exception as exc:
            print("Getting shipping address fail")
            print(exc)

        print("Extra step for bestbuy: Need to check if full tracking number is given.")

        tracking_number_length = len(tracking_number_list[0])
        if tracking_number_length < 6:
            print("Bestbuy email does not show full tracking number (USPS). Need to navigate to bestbuy site to get full tracking number.")
            bestbuy_tracking_button_xpath = "/html/body/table[3]/tbody/tr/td[2]/table[1]/tbody/tr/td[2]/table[4]/tbody/tr/td/table[1]/tbody/tr[5]/td/div/div/table/tbody/tr/td/div/table/tbody/tr[3]/td/table/tbody/tr/td/div[2]/table/tbody/tr[1]/td/a"
            try:
                bestbuy_tracking_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, bestbuy_tracking_button_xpath)))
                time.sleep(1)
                print("Bestbuy tracking button found.")
                bestbuy_tracking_button.click()
                time.sleep(1)
            except:
                print("Bestbuy tracking button not present.")

            try:
                driver.switch_to.window(driver.window_handles[1])
                time.sleep(2)
                driver.refresh()
            except:
                print("Tab switch failed")
                traceback.print_exc()

            bb_tracking_xpath = "//div[@class='tracker-numContainer']//a"
            bb_tracking_regex = "(.+)\s+Opens"
            try:
                bb_tracking_number1 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, bb_tracking_xpath))).text
                bb_tracking_number = ((re.findall(bb_tracking_regex, bb_tracking_number1))[0]).strip()
                time.sleep(1)
                print("Bestbuy tracking found.")
                tracking_number_list = bb_tracking_number.split(',')
                print(f"Tracking Number: {tracking_number_list}")
            except Exception as exc:
                print("Tracking number not found on bestbuy page")
                print(exc)

            driver.close()

            try:
                driver.switch_to.window(driver.window_handles[0])
            except:
                print("Tab switch failed")
        else:
            print("Full tracking number given. Continuing.")

        try:
            print("Getting current email position")
            current_email_position = driver.current_url
        except:
            traceback.print_exc()
            current_email_position = " "

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, older_button_xpath)))
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1

        # --------------------------------------#-----------------------------------------#------------------------------------#---------------------------------------------------#


def process_Newegg_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing..."

        print("Working on body text")
        print("Getting Order Number")
        order_number = ""
        order_number_regex = "Order\s+Number:\s+(.+)"
        try:
            order_number = (re.findall(order_number_regex, message_body_field_message))[0]
            time.sleep(1)
            print(f"Order Number {order_number}")
        except:
            print("Getting order number fail")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = re.compile(r"Tracking\s+Number:\s+([\s\S]*?)(\(|Shipping)")
        try:
            tracking_number_found_1 = tracking_number_regex.search(message_body_field_message)
            tracking_number_found = (tracking_number_found_1.group(1)).strip()

            tracking_number_list = tracking_number_found.split(',')

            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        # --------------------------------------#-----------------------------------------#------------------------------------#---------------------------------------------------#

        print("Getting shipping address process")

        order_confirmation_search_term = "subject:order confirmation " + order_number
        search_box_xpath = "//input[@title='Search']"
        search_button_xpath = "(//input[@value='Search Mail'])"
        try:
            input_search = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, search_box_xpath)))
            input_search.clear()
            input_search.send_keys(order_confirmation_search_term)

            search_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, search_button_xpath)))
            search_button.click()
        except:
            print("Search button error")

        email_content_link_xpath = "(//td//a[@href]//span[@class])"
        try:
            email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath)))
            print("Email link element found")
            time.sleep(1)
            print("Clicking email")
            email_content_link_element.click()
        except Exception as exc:
            print("Going to email body fail")
            print(exc)

        print("Getting text of order confirmation")
        try:
            message_body_field_message_shipping = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message_shipping = "Nothing..."

        # shipping_text_regex = "Shipping\s+Information\s+([\s\S]*?)Newegg.com"  #Imcludes Name
        shipping_text_regex = "Shipping\s+Information\s+.+(\s.+\s.+\s.+)"
        shipping_address = "Nothing"
        try:
            shipping_address = ((re.findall(shipping_text_regex, message_body_field_message_shipping))[0]).strip()
            time.sleep(1)
            print(f"Shipping Address:\n{shipping_address}")
        except:
            print("Getting shipping address fail")

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def process_ebay_pages(driver, current_page_url, all_order_records_list):
    # current_page = driver.current_url
    referenceID = ""
    next_page = True
    email_index = 1

    email_content_link_xpath = "(//td//a[@href]//span[@class])"
    try:
        email_content_link_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, email_content_link_xpath + "[" + str(email_index) + "]")))
        print("Email link element found")
        time.sleep(1)
        print("Clicking email")
        email_content_link_element.click()
    except Exception as exc:
        print("Going to email body fail")
        print(exc)

    while next_page:
        print("\n****************Going to email number " + str(email_index) + " ****************")

        try:
            driver.switch_to.window(driver.window_handles[0])
        except:
            print("Tab switch failed")

        print("Getting Ship date")
        ship_date_xpath = "(//a[@name]/..//td[@align='right'])[1]"
        ship_date_regex = "(\d{1,2}\/\d{1,2}\/\d{2,4})"
        try:
            ship_date = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ship_date_xpath))).text
            print(f"Ship date: {ship_date}")
        except:
            ship_date = " "

        print("Getting Text of the page")
        try:
            message_body_field_message = driver.find_element_by_xpath("//body//div[@class='msg']").text
        except:
            print("Copy Page fail")
            message_body_field_message = "Nothing..."

        print("Working on body text....")
        print("Getting transaction ID")
        transaction_id_xpath = "//*[text()[contains(.,'Track')]]"
        transaction_id_regex = "transid%\d+D([\s\S]*?)%"
        transaction_id = ""
        try:
            transaction_id_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, transaction_id_xpath)))
            transaction_id_tag = (transaction_id_element.get_attribute('href')).strip()
            transaction_id = ((re.findall(transaction_id_regex, transaction_id_tag))[0]).strip()
            print(f"Transaction ID: {transaction_id}")
        except:
            print("Can't get transaction ID")

        print("Getting item ID")
        item_id = ""
        item_id_regex = "Item\s+ID:\s+(\d+)"
        try:
            item_id = ((re.findall(item_id_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Item ID: {item_id}")
        except:
            print("Getting item ID fail")

        print("Getting Order Number")
        order_number = item_id + "-" + transaction_id
        print(f"Order Number: {order_number}")

        print("Constructing Reference ID")
        print("Getting name")
        name = ""
        name_regex = "order\s+is\s+[\s\S]*?,\s+(.+?)!"
        try:
            name = ((re.findall(name_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Item ID: {name}")
        except:
            print("Getting name fail")

        print("Getting qty")
        qty = ""
        qty_regex = "Quantity:\s+(\d+)"
        try:
            qty = ((re.findall(qty_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Item ID: {qty}")
        except:
            print("Getting name fail")

        print("getting amount paid")
        paid = ""
        paid_regex = "Paid:[\s\S]*?([\d.,]+)"
        try:
            paid = ((re.findall(paid_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Item ID: {paid}")
        except:
            print("Getting amount paid fail")

        referenceID = name + "-" + item_id + "-" + qty + "-" + paid
        print(f"Reference ID: {referenceID}")

        print("Getting Tracking Number")
        tracking_number_list = []
        tracking_number_regex = "The\s+tracking\s+number\s+is\s+([\w\d]+)"
        try:
            tracking_number_found = ((re.findall(tracking_number_regex, message_body_field_message))[0]).strip()
            tracking_number_list = tracking_number_found.split(',')
            time.sleep(1)
            print(f"Tracking Number {tracking_number_list}")
        except:
            print("Getting tracking number fail")

        # --------------------------------------#-----------------------------------------#------------------------------------#---------------------------------------------------#

        print("Getting shipping address ")
        shipping_text_regex = "shipped\s+to\s+([\s\S]*?).\s+The\s+tracking"
        shipping_address = "Nothing"
        try:
            shipping_address = ((re.findall(shipping_text_regex, message_body_field_message))[0]).strip()
            time.sleep(1)
            print(f"Shipping Address:\n{shipping_address}")
        except:
            print("Getting shipping address fail")

        print("Getting current email position")
        current_email_position = driver.current_url

        shipping_page_text = ""
        check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID)

        print("Going back to Gmail")
        driver.get(current_email_position)

        print("Going to next email")
        older_button_xpath = "(//*[text()[contains(.,'Older')]])[1]/.."
        try:
            older_button = driver.find_element_by_xpath(older_button_xpath)
            time.sleep(1)
            print("Older button found.")
            older_button.click()
            next_page = True
        except:
            print("Older button not present. ..")
            next_page = False

        email_index = email_index + 1


def check_carrier_name_and_add_record(tracking_number_list, driver, order_number, shipping_address, all_order_records_list, ship_date, shipping_page_text, referenceID):
    print("checking carrier and adding to record")

    print("Checking for Carrier Name")
    tracking_site = ""
    status = "Unknown"
    delivery_date = "Unknown"
    for tracking_number in tracking_number_list:
        tracking_number = tracking_number.strip()

        if "TBA" in tracking_number:
            carrier_name = "Amazon Logistics"
            print(f"{tracking_number}: Amazon Logistics")

            print("Getting shipment delivery date (Amazon)")
            shiment_status_regex = "^(\w+)\s+"
            delivery_date_regex = "^\w+\s+(.+)"
            try:
                delivery_date = ((re.findall(delivery_date_regex, shipping_page_text))[0]).strip()
                print(f"{tracking_number} delivery date: {delivery_date}")
            except:
                delivery_date = " "

            print("Getting shipment status")
            try:
                status = ((re.findall(shiment_status_regex, shipping_page_text))[0]).strip()
                print(f"{tracking_number} status: {status}")
            except:
                status = "Created"

            if "delivered" in status.lower():
                delivery_date = delivery_date

            if "today" in delivery_date.lower() and "delivered" in status.lower():
                delivery_date = str(datetime.today().strftime('%Y-%m-%d'))

            if "arriving" in status.lower():
                delivery_date = "not yet delivered"
                print(f"{tracking_number} updated delivery date: {delivery_date}")

        elif "1Z" in tracking_number:
            carrier_name = "UPS"
            print(f"{tracking_number}: UPS")
            tracking_site = 'https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums='
            print("Getting status from tracking site")
            status_url = tracking_site + tracking_number
            status_xpath = "//a[@id='tt_spStatus']"
            status_xpath = '//*[@id="stApp_txtPackageStatus"]'
            driver.get(status_url)
            try:
                status = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, status_xpath))).text
                print(f"{tracking_number} status: {status}")
            except:
                status = "Created"

            print("Getting delivery date")
            # delivery_date_xpath = "//*[@id='fontControl']/div[3]/div/div/div[1]/div[1]/div[2]/div/div[1]/div[1]/p[2]"
            delivery_date_xpath = '//*[@id="stApp_deliveredDate"]'
            delivery_date_regex = '(\d{1,2}\/\d{1,2}\/\d{2,4})'
            try:
                delivery_info = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, delivery_date_xpath))).text
                delivery_date = ((re.findall(delivery_date_regex, delivery_info))[0]).strip()
                print(f"{tracking_number} delivery date: {delivery_date}")
            except:
                delivery_date = " "

            if status.lower() == "delivered":
                delivery_date = delivery_date
            else:
                delivery_date = " "
                print(f"{tracking_number} updated delivery date: {delivery_date}")


        elif "1LS" in tracking_number:
            carrier_name = "Lasership"
            print(f"{tracking_number}: Lasership")
            tracking_site = 'http://www.lasership.com/track/'
            print("Getting status from tracking site")
            status_url = tracking_site + tracking_number
            status_xpath = "//div[@class='row track-banner']//*[@class='event-banner-on']"
            driver.get(status_url)
            try:
                status = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, status_xpath))).text
                print(f"{tracking_number} status: {status}")
            except:
                status = "Created"

            print("Getting delivery date")
            delivery_date_xpath = "//*[@class='event-title-green']"
            delivery_date_regex = '(\d{1,2}\/\d{1,2}\/\d{2,4})'
            try:
                delivery_info = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, delivery_date_xpath))).text
                delivery_date = ((re.findall(delivery_date_regex, delivery_info))[0]).strip()
                print(f"{tracking_number} delivery date: {delivery_date}")
            except:
                delivery_date = " "

            if status.lower() == "delivered":
                delivery_date = delivery_date
            else:
                delivery_date = " "
                print(f"{tracking_number} updated delivery date: {delivery_date}")


        elif len(tracking_number) >= 22:
            carrier_name = "USPS"
            print(f"{tracking_number}: USPS")
            tracking_site = 'https://tools.usps.com/go/TrackConfirmAction?tLabels='
            print("Getting status from tracking site")
            status_url = tracking_site + tracking_number
            status_xpath = "//div[@class='delivery_status']//strong"
            driver.get(status_url)
            try:
                status = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, status_xpath))).text
                print(f"{tracking_number} status: {status}")
            except:
                status = "Created"

            print("Getting delivery date")
            delivery_date_xpath = "(//div[@class='delivery_status']//div[@class='status_feed']/p)[1]"
            delivery_date_regex = '(\w+\s+\d+,\s+\d+)'
            try:
                delivery_info = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, delivery_date_xpath))).text
                delivery_date = ((re.findall(delivery_date_regex, delivery_info))[0]).strip()
                print(f"{tracking_number} delivery date: {delivery_date}")
            except:
                delivery_date = " "

            if status.lower() == "delivered":
                delivery_date = delivery_date
            else:
                delivery_date = " "
                print(f"{tracking_number} updated delivery date: {delivery_date}")

        elif len(tracking_number) == 0:
            carrier_name = "Unknown or tracking number not displayed"
            print(f"{tracking_number}: Unknown")



        else:
            carrier_name = "FedEx"
            print(f"{tracking_number}: FedEx")
            tracking_site = 'https://www.fedex.com/apps/fedextrack/?tracknumbers='
            print("Getting status from tracking site")
            status_url = tracking_site + tracking_number
            status_xpath = "//*[contains(@class,'key_status')]"
            driver.get(status_url)
            try:
                status = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, status_xpath))).text
                print(f"{tracking_number} status: {status}")
            except:
                status = "Created"

            print("Getting delivery date")
            delivery_date_xpath = "//*[contains(@class,'date dest')]"
            delivery_date_regex = '(\d{1,2}\/\d{1,2}\/\d{2,4})'
            try:
                delivery_info = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, delivery_date_xpath))).text
                delivery_date = ((re.findall(delivery_date_regex, delivery_info))[0]).strip()
                print(f"{tracking_number} delivery date: {delivery_date}")
            except:
                delivery_date = " "

            if status.lower() == "delivered":
                delivery_date = delivery_date
            else:
                delivery_date = " "
                print(f"{tracking_number} updated delivery date: {delivery_date}")

        current_order_record = collections.OrderedDict()
        current_order_record['Order Number'] = order_number
        current_order_record['Tracking Number'] = tracking_number
        current_order_record['Carrier Name'] = carrier_name
        current_order_record['Shipping Address'] = shipping_address
        current_order_record['Status'] = status
        current_order_record['Ship Date'] = ship_date
        current_order_record['Delivery Date'] = delivery_date

        info_database = send_tracking_info_database(tracking_number, carrier_name, status, order_number, shipping_address, referenceID)

        # Check the number of entries in ShipmentTrackings for this tracking number.
        add_to_sheet = True
        print("Checking if item has been picked up")
        for entry in info_database:
            if entry['Status'] == "CheckedIn":
                print("Item has check-in scan. Record will not be added to the sheet. ")
                add_to_sheet = False
            else:
                pass

        if add_to_sheet == True:
            print(f"No pickup scan for items for tracking number \"{tracking_number}\" ")
            print("Adding current tracking record to sheet")
            add_to_records(current_order_record, all_order_records_list)

    return


if __name__ == '__main__':
    process_main_search()
