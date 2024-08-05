import logging
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Robocorp.Vault import Vault
from RPA.Robocorp.Storage import Storage

@task #Means that this function is the entry point
def robot_spare_bin_python():
    browser.configure(slowmo=100)
    
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    logout()

def open_the_intranet_website():
    storage = Storage()

    url = storage.get_text_asset("RobotSpareBin_URL")

    browser.goto(url)

def log_in():

    _secret = Vault().get_secret("RobotSpareBin_Credential")

    USER_NAME = _secret["user"]
    PASSWORD = _secret["password"]

    page = browser.page()
    page.fill("#username", USER_NAME)
    page.fill("#password", PASSWORD)
    page.click("button:text('Log in')")

def fill_and_submit_sales_form(sales_rep):
    page = browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def download_excel_file():
    http = HTTP()
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook

    for row in worksheet:
        fill_and_submit_sales_form(row)
        

def collect_results():
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def logout():
    page = browser.page()
    page.click("text=Log out")

def export_as_pdf():
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

