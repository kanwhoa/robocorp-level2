from robocorp.tasks import task
from robocorp import browser
from robocorp import vault

# Very simple, prefer httpx and save to a temp dr?
from RPA.HTTP import HTTP
# Prefer win32com. This library uses openpyxl, which has cserious bug esp dealing with macros
from RPA.Excel.Files import Files
from RPA.Tables import Tables, Table, Row
from RPA.PDF import PDF
from RPA.Archive import Archive

import glob
import os
from typing import Dict, Any, List

# General TODOs:
# 1. logging
# 2. Common timeouts
# 3. Generalise the application in a library to deal with custom controls and access methods
# 4. Use Chrome not Chromium
# 5. Remove browser scrolling

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    config = setup()

    # Create the receipt store.
    # TODO: error handling. Put this in a startup
    receipts_path = r'output/receipts'
    os.makedirs(receipts_path, exist_ok=True)
    files = glob.glob(receipts_path + '/*')
    for f in files:
        os.remove(f)

    # Log into the site so that we can submit the orders later.
    portal_front_page = log_in(config)

    # Get the orders as a table.
    orders = get_orders(config)
    filenames = place_orders(config, orders)

    # Zip the filenames
    archive_receipts(receipts_path, "output/receipts.zip")


def setup() -> Dict[str, Any]:
    """Convert this into an common initialse function/library
    TODO
    1. Temp dirs
    2. Environment configuration/checks (browser profile etc)
    3. Debug mode to slow down. Assume always slow for now
    4. screenshot directory which is permament
    5. The running environment (dev/test/prod)
    6. Cleanup of the output directory?
    7. Maximise the browser
    8. Set the default playwright timeout
    9. Global debug flag
    """

    browser.configure(
        slowmo=100,
        screenshot="only-on-failure",
    )

    # Config properties
    return {
        "base_url": "https://robotsparebinindustries.com",
        "env": "dev",
        "default_timeout": 5000
    }


def log_in(config: dict[str, Any]) -> browser.Page:
    """Fills in the login form and clicks the 'Log in' button
    TODO:
    1. clean environment at the run start - e.g. browsers
    2. Credentials into a vault
    3. Checks to make sure the page loads correctly.
    """
    browser.goto(config['base_url'])
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

    # Wait for the post login page to ensure all good.
    page.wait_for_selector(r'xpath=//input[@id="firstname"]', timeout=5000) #, state="visible")
    # TODO. should really wait until the modal has disappeared.... wait for element with stae or negative xpath

    return page


def get_orders(config: Dict[str, Any]) -> Table:
    """TODO
    1. Convert this to an httpx function to use parallelism for multiple downloads.
    2. Put all downloads into a run specifc temp dir - this drops them in the current run dir, don't want confusion between clients/runs
    3. Retries & errors
    4. Convert the Tables to Pandas
    5. Could split this into a download and then extract function, but in this case, it doesn't make sense."""
    http = HTTP()
    order_filename = "orders.csv"
    http.download(url="{base_url}/{filename}".format(base_url = config["base_url"], filename = order_filename), overwrite=True)

    tables = Tables()
    return tables.read_table_from_csv(order_filename)


def place_orders(config: Dict[str, Any], orders: Table) -> List[str]:
    """Iterate over the orders and place them into the form"""

    filenames = []
    for order in orders:
        # TODO: log which order we are processing
        # This retry is very bad, better to have some form of limit and fail the work item
        filename = place_order(config, order)
        filenames.append(filename)

    # Return all the order filenames to package up
    return filenames


def place_order(config: Dict[str, Any], order: Row) -> bool:
    """Place the order.
    TODO:
    1. proper error handling
    2. Confirm we are on the right page before starting"""

    page = browser.page()
    # Navigate to the order page. The order button doesn't work from a existing order, click home to reset. Doesn't matter if clicked multiple times.
    page.click(r'xpath=//a[text()="Home"]', timeout=5000)
    page.click(r'xpath=//a[text()="Order your robot!"]', timeout=5000)
    page.click(r'xpath=//div[@class="modal"]//button[text()="OK"]', timeout=5000)

    # Fill the form, cheap and cheerful way
    page.select_option(r'xpath=//select[@id="head"]', value=order["Head"], timeout=5000)
    # TODO: obvious code injection to fix
    page.set_checked(r'xpath=//input[@id="id-body-{num}"]'.format(num=order["Body"]), checked=True, timeout=5000)
    page.fill(r'xpath=//label[text()="3. Legs:"]/following-sibling::input', value=order["Legs"], timeout=5000)
    page.fill(r'xpath=//input[@id="address"]', value=order["Address"], timeout=5000)

    # Preview... this doesn't seem to fail, but real should check for that. Also don't need, the images is on the order page
    # which is better to pick up from as we know the order...
    page.click(r'xpath=//button[@id="preview"]', timeout=5000)

    # Make sure the submit is good
    while True:
        # Click submit
        page.click(r'xpath=//button[@id="order"]', timeout=5000)

        # Wait for a success or failure condition (both at the same time). Once detected, see what happened
        success_xpath=r'//div[@id="receipt" and contains(@class, "alert-success")]/p[contains(@class, "badge-success")]'
        failed_xpath=r'//div[contains(@class, "alert-danger")]'
        page.wait_for_selector(r'xpath={success}|{failure}'.format(success=success_xpath, failure=failed_xpath))

        # If we got an error, just reclick the submit button until it works. In reality, this would throw errors/retry the whole form etc.
        # Should also do a proper else check.
        if page.is_visible(r'xpath=' + success_xpath, timeout=100):
            break
        # If not success, loop and retry. Eugh.

    # Layout like this for now, but structure better for prod.
    # TODO: check the order number isn't blank and is of the right format.

    order_number = page.text_content(r'xpath=//div[@id="receipt"]//p[1]', timeout=5000)
    
    # Not put this in a seperate function because it would mean passing a lot of state between the two. 
    # A production version would handle this nicely. Also, would get the locator then base further queries from there.
    # Save the resulting form
    receipt_filename = store_receipt_as_pdf(config, order_number)
    screenshot_filename = screenshot_robot(config, order_number)
    embed_screenshot_to_receipt(screenshot_filename, receipt_filename)
    return receipt_filename


def store_receipt_as_pdf(config: Dict[str, Any], order_number: str) -> str:
    """Store the receipt, but badly, many assumptions not passed in arguments
    TODO:
    1. loses all style information from the HTML. Look at the page.pdf() function."""
    page = browser.page()
    order_html = page.locator(r'xpath=//div[@id="receipt"]').inner_html()

    pdf = PDF()
    filename = "output/receipts/{order_number}.pdf".format(order_number=order_number)
    pdf.html_to_pdf(order_html, filename)
    return filename


def screenshot_robot(config: Dict[str, Any], order_number) -> str:
    """Take a screenshot of the robot
    TODO: big hack, this just hopes that the image is small enough to fit in the page bounding box"""
    page = browser.page()
    page.evaluate(r'window.scrollTo({top: 0, left: 0, behavior: "instant"});')
    pic_locator = page.locator(r'xpath=//div[@id="robot-preview-image"]')
    pic_coords = pic_locator.bounding_box(timeout=5000)

    filename = "output/receipts/{order_number}_ss.png".format(order_number=order_number)
    #page.screenshot(path=filename, type="png", full_page=False, clip=pic_coords, timeout=5000)
    page.screenshot(path=filename, type="png", full_page=True, clip=pic_coords, timeout=5000)
    return filename


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Merge the two PDFs
    TODO:
    Better aporoach here is to save all the files then to the merge in one shot at the end, handles io failures better."""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)


def archive_receipts(receipts_path: str, archive_path: str):
    """Archive a directory into zip file
    TODO:
    1. remove paths from the zip file"""
    archive = Archive()

    archive.archive_folder_with_zip(receipts_path, archive_path, exclude="*.png")