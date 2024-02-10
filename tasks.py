import os
import shutil

from robocorp.tasks import task
from robocorp import browser
from time import sleep
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():    
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    page = open_robot_order_website()
    close_annoying_modal(page)
    data = get_orders()
    for row in data:
        fill_form_and_submit_order(row, page)
        take_screenshot_and_save_as_pdf(row['Order number'], page)
        page.click("xpath=//button[@id='order-another']") # click the order another button
        close_annoying_modal(page)
    archive_receipts()

def open_robot_order_website():
    """
    Opens the robot order website. 
    Returns a page object to be passed to other functions.
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    return browser.page()

def get_orders(): 
    """
    Downloads the orders excel file using requests.
    Reads it in as a table. 
    Returns the table.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = Tables()
    data = table.read_table_from_csv("orders.csv")
    return data

def close_annoying_modal(page): 
    """
    Click 'OK' on the modal that pops up when visiting the URL.
    """
    page.click("text=OK")

def check_for_submit_error(): 
    """
    Checks for an order submit error and deals with it if necessary.
    """
    page = browser.page()
    submit_error = page.is_visible("xpath=//div[@class='alert alert-danger']")
    while submit_error: 
        page.click("xpath=//button[@id='order']")
        submit_error = page.is_visible("xpath=//div[@class='alert alert-danger']")
    
def fill_form_and_submit_order(row, page): 
    """
    Given a row, fills out the form on the site. 
    Presses submit button. 
    Checks to see if the submit button failed. 
    """
    page.select_option("#head", str(row["Head"]))
    page.click(f"#id-body-{str(row['Body'])}")
    page.fill("xpath=//input[@placeholder='Enter the part number for the legs']", str(row["Legs"]))
    page.fill("xpath=//input[@placeholder='Shipping address']", str(row["Address"]))
    page.click("xpath=//button[@id='order']")
    check_for_submit_error()

def take_screenshot_and_save_as_pdf(order_number, page): 
    """
    Creates a PDF with the order number in the name. 
    Returns a path to the PDF.
    """
    screenshot_path = f"output/receipts/order_{order_number}.png"
    pdf_path = f"output/receipts/order_{order_number}.pdf"
    page.screenshot(path=screenshot_path)
    pdf = PDF()
    receipt_html = page.locator("xpath=//div[@id= 'receipt']").inner_html()
    pdf.html_to_pdf(receipt_html, pdf_path)
    pdf.open_pdf(pdf_path)
    pdf.add_files_to_pdf(files = [screenshot_path], target_document = pdf_path, append=True)
    pdf.save_pdf(output_path=f"output/receipts/order_{order_number}.pdf")
    os.remove(screenshot_path) # clean up screenshot

def archive_receipts(): 
    """
    Takes the PDFs in the receipts folder and zips them together.
    """
    lib = Archive()
    lib.archive_folder_with_zip('./output/receipts', 'output/receipts.zip')
    shutil.rmtree('./output/receipts') # clean up receipts after they have been zipped