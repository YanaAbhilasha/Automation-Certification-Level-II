from robocorp.tasks import task
from robocorp import browser
from PIL import Image
from RPA.Excel.Files import Files as ExcelFiles
from RPA.HTTP import HTTP as Http
from RPA.PDF import PDF as Pdf
from RPA.Tables import Tables as Tables
import time
from RPA.Archive import Archive as ZipArchive


@task
def order_robots_from_robot_spare_bin():
    """
    Automates the process of ordering robots from RobotSpareBin Industries Inc.
    - Saves the order confirmation as a PDF.
    - Captures and resizes a screenshot of the ordered robot.
    - Embeds the screenshot into the PDF receipt.
    - Archives the receipts and images into a ZIP file.
    """
    browser.configure(slowmo=50)
    open_robot_order_website()
    close_annoying_modal()
    download_excel_file()
    process_orders()
    archive_receipts()


def open_robot_order_website():
    # Navigate to the RobotSpareBin Industries order page
    browser.goto('https://robotsparebinindustries.com/#/robot-order')


def close_annoying_modal():
    # Close any pop-up modal on the page
    browser.page().click("text=OK")


def download_excel_file():
    # Download the orders CSV file from the specified URL
    http = Http()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def process_orders():
    """
    Read the order data from the downloaded CSV file and process each order.
    """
    orders = Tables().read_table_from_csv(
        "orders.csv",
        columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    for order in orders:
        fill_the_form(order)


def fill_the_form(order):
    """
    Fill in the order form on the website using data from the CSV file.
    - Select the appropriate head and body.
    - Enter the legs part number and shipping address.
    - Submit the form and handle the receipt.
    """
    page = browser.page()

    def select_head_option(head_id):
        # Map head_id to the corresponding robot head option
        return {
            '1': 'Roll-a-thor head',
            '2': 'Peanut crusher head',
            '3': 'D.A.V.E head',
            '4': 'Andy Roid head',
            '5': 'Spanner mate head',
            '6': 'Drillbit 2000 head'
        }.get(head_id, '')

    def select_body_option(body_id):
        # Select the body part based on the body_id
        return f"#id-body-{body_id}"

    # Fill out the order form with the selected head, body, legs, and address
    page.select_option('#head', select_head_option(order["Head"]))
    page.click(select_body_option(order['Body']))
    page.fill(
        "input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", str(order["Address"]))
    time.sleep(1)

    # Submit the order and handle the response
    while True:
        page.click("css=#order")
        if next_order_button := page.query_selector("css=#order-another"):
            # If the order was successful, store the receipt and screenshot
            store_receipt_as_pdf(order["Order number"])
            screenshot = screenshot_robot(order["Order number"])
            embed_screenshot_to_receipt(screenshot, order["Order number"])
            page.click("css=#order-another")
            print("Order successful")
            close_annoying_modal()
            break
        else:
            print("Order failed, retrying...")


def store_receipt_as_pdf(order_number):
    """
    Save the order receipt as a PDF file.
    """
    order_html = browser.page().locator("#receipt").inner_html()
    pdf = Pdf()
    path = f"output/{order_number}.pdf"
    pdf.html_to_pdf(order_html, path)
    return path


def screenshot_robot(order_number):
    """
    Capture a screenshot of the robot preview and resize it.
    """
    page = browser.page()
    element = page.query_selector("#robot-preview-image")
    path = f"output/{order_number}.png"
    element.screenshot(path=path)

    # Resize the image to a fixed width, maintaining the aspect ratio
    image = Image.open(path)
    resized_image = image.resize(
        (500, int(image.height * (500 / image.width))))
    resized_image.save(path)

    return path


def embed_screenshot_to_receipt(screenshot, order_number):
    """
    Embed the robot screenshot into the corresponding PDF receipt.
    """
    pdf = Pdf()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=f"output/{order_number}.pdf",
        append=True
    )


def archive_receipts():
    """
    Archive all receipts and screenshots into a single ZIP file.
    """
    archive = ZipArchive()
    archive.archive_folder_with_zip('output/', 'merged.zip')
