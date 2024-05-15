from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.FileSystem import FileSystem
from RPA.Tables import Tables
import shutil
from pathlib import Path
import shutil

browser = Selenium()
pdf = PDF()


@task
def order_robots_from_RobotSpareBin():
    create_new_directory()
    open_robot_order_website()
    log_in()
    download_excel_file()
    get_orders_and_fill_form()
    # Archive orders folder as zip
    shutil.make_archive("output/orders", 'zip', "output/orders")
    # Remove unneccessary folder
    fs = FileSystem()
    fs.remove_directory(path="output/pictures", recursive=True)
    fs.remove_directory(path="output/orders", recursive=True)


def open_robot_order_website():
    browser.open_available_browser(
        'https://robotsparebinindustries.com/', headless=False, maximized=True)


def log_in():
    browser.input_text(locator="id:username", text="maria")
    browser.input_text(locator="id:password", text="thoushallnotpass")
    browser.click_button_when_visible("//button[contains(text(),'Log in')]")
    # Go to order robot page
    browser.click_element_when_visible(
        "//a[contains(text(),'Order your robot!')]")


def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", overwrite=True, target_file="output")


def close_annoying_modal():
    # Close the popup modal
    browser.click_button_when_visible("//button[contains(text(),'OK')]")


def get_orders_and_fill_form():
    tables = Tables()
    table = tables.read_table_from_csv(path="output/orders.csv", header=True)
    for row in table:
        close_annoying_modal()
        fill_the_form(row)
        store_receipt_as_pdf(row)
        # Go to order another robot
        browser.click_button_when_visible("id:order-another")


def fill_the_form(row):
    browser.select_from_list_by_value(
        "id:head", str(row["Head"]))
    body_value = row["Body"]
    browser.click_element_when_visible(f"//input[@value='{body_value}']")
    browser.input_text("//input[@type='number']", row["Legs"])
    browser.input_text("id:address", row["Address"])
    browser.click_button_when_visible("//button[contains(text(),'Order')]")
    check_order_success()


def check_order_success():
    is_order_success = browser.is_element_visible(
        "//div[@class='alert alert-success']")
    while is_order_success == False:
        browser.click_button_when_visible("//button[contains(text(),'Order')]")
        is_order_success = browser.is_element_visible(
            "//div[@class='alert alert-success']")


def store_receipt_as_pdf(row):
    pdf = PDF()
    # Get order
    browser.wait_until_element_is_visible("id:receipt", timeout=15)
    recipent_html = browser.get_element_attribute(
        "id:receipt", attribute="outerHTML")
    order_number = row["Order number"]
    pdf.html_to_pdf(recipent_html, f"output/orders/order_{order_number}.pdf")
    # Get robot image
    browser.wait_until_element_is_visible("id:robot-preview-image", timeout=15)
    browser.wait_until_element_is_visible(
        "//div[@id='robot-preview-image']//img[@alt='Head']", timeout=15)
    browser.wait_until_element_is_visible(
        "//div[@id='robot-preview-image']//img[@alt='Body']", timeout=15)
    browser.wait_until_element_is_visible(
        "//div[@id='robot-preview-image']//img[@alt='Legs']", timeout=15)
    browser.screenshot("id:robot-preview-image",
                       f"output/pictures/robot_{order_number}.png")
    # Merge PDF and image
    list_of_files = [
        f"output/orders/order_{order_number}.pdf",
        f"output/pictures/robot_{order_number}.png:align=center",
    ]

    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f"output/orders/order_{order_number}.pdf"
    )


def create_new_directory():
    directory = "output/orders"
    directory_path = Path(directory)

    # Check if the directory exists
    if directory_path.exists() and directory_path.is_dir():
        # Remove the existing directory
        shutil.rmtree(directory)
        print(f"Directory '{directory}' was removed.")

    # Create the new directory
    directory_path.mkdir()
    print(f"Directory '{directory}' was created.")


def log_out():
    """Presses the 'Log out' button"""
    browser.click_element_when_visible("//button[contains(text(),'Log out')]")
