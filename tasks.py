from robocorp.tasks import task
import robocorp.browser as browser
import robocorp.http as http
from RPA.Tables import Tables   # Tables모듈 안 Tables클래스 사용
import time
from RPA.PDF import PDF # PDF모듈 안 PDF클래스 사용
from RPA.Archive import Archive

#? 전역변수(상수) 설정
OUTPUT_PATH = "output"

@task
def order_robots_from_RobotSpareBin():
    """
    - Orders robots from RobotSpareBin Industries Inc.
    - Saves the order HTML receipt as a PDF file.
    - Saves the screenshot of the ordered robot.
    - Embeds the screenshot of the robot to the PDF receipt.
    - Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    orders = get_orders()
    for row in orders:  # row: dict
        close_annoying_modal()
        fill_the_form(row)
        download_and_store_the_order_receipt_as_pdf(row)
        order_another_robot()
    archive_receipts()
    #? Teardown: it will be automatically closed when the task finishes


def browser_setting():
    browser.configure(

    )

def open_robot_order_website():
    #* 1. 브라우저 초기 설정
    browser_setting()
    #* 2. 브라우저 실행
    browser.goto(url="https://robotsparebinindustries.com/#/robot-order")

def download_the_orders_file():
    download_url = "https://robotsparebinindustries.com/orders.csv"
    http.download(
        url=download_url, path=OUTPUT_PATH, overwrite=True
    )

def get_orders():
    """
    Download the orders file, read it as a table, and return the result
    """
    download_the_orders_file()

    table = Tables()    # 인스턴스화
    orders = table.read_table_from_csv(path=OUTPUT_PATH + "/" + "orders.csv")   # @keyword("Read table from CSV")
    return orders

def close_annoying_modal():
    page = browser.page()
    page.click(selector="button:text('OK')")

def fill_the_form(row: dict):
    page = browser.page()
    page.select_option(selector="#head", value=row["Head"])
    
    group_name = "body"
    value = row["Body"]
    #! parsing selector -> xpath:(X), xpath=(O)
    radio_xpath = (
            #// f"xpath=//input[@type='radio' and @name='{group_name}' and "
            f"//input[@type='radio' and @name='{group_name}' and "
            f"(@value='{value}' or @id='{value}')]"
        )
    page.click(selector=radio_xpath)

    legs_xpath_expression = "//input[@class='form-control' and  @type='number']"
    page.fill(selector=legs_xpath_expression, value=row["Legs"])

    page.fill(selector="#address", value=row["Address"])

    preview_the_robot()
    submit_the_order_until_success()

def preview_the_robot():
    page = browser.page()
    page.click(selector="#preview")

def submit_the_order_and_check():
    page = browser.page()
    page.click(selector="#order")

    element_exist = page.is_visible(selector="#receipt")
    return element_exist

def submit_the_order_until_success():
    RETRY_AMOUNT = 10
    RETRY_INTERVAL = 0.5

    run_count = 0
    while True:
        run_count += 1
        if RETRY_AMOUNT < run_count:
            raise "run over!"
        
        result_flag = submit_the_order_and_check()

        if result_flag == True:
            break

        time.sleep(RETRY_INTERVAL)

def store_receipt_as_pdf(order_number):
    '''
    - make automatically directory
    - overwrite = True
    '''
    page = browser.page()
    #! return -> None
    #// order_receipt_html = page.get_attribute(selector="#receipt", name="outerHTML")
    locator = page.locator(selector="#receipt")
    order_receipt_html = locator.inner_html()
    # Ensure a unique name
    file_name = "0" + order_number + "_order_receipt.pdf"
    dir_name = "order_receipt"
    file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name
    pdf = PDF()
    #! 메서드 인식이 되지 않는다!
    pdf.html_to_pdf(order_receipt_html, file_path)  # make automatically "order_receipt" dir
    return file_path

def screenshot_robot(order_number):
    '''
    - make automatically directory
    - overwrite = True
    '''
    # Ensure a unique name
    file_name = "0" + order_number + "_robot_preview_image.png"
    dir_name = "robot_preview_image"
    file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name

    page = browser.page()
    element = page.query_selector(selector="#robot-preview-image")
    element.screenshot(path=file_path)   # make automatically "robot_preview_image" dir
    return file_path

def embed_screenshot_to_receipt(order_number, screenshot, pdf_file):
    '''
    - make automatically directory
    - overwrite = True

    .. PDF 문서에 쓴 후에는 닫는 것이 좋습니다.
        -> 스크린샷을 PDF에 추가하는 기능은 PDF 문서 열기 및 닫기도 처리합니다.
    '''
    # Ensure a unique name
    file_name = "0" + order_number + "_result.pdf"
    dir_name = "result"
    file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name

    pdf = PDF()
    #! 메서드 인식이 되지 않는다!
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=file_path
    )   # make automatically "result" dir

def download_and_store_the_order_receipt_as_pdf(row):
    order_number = row["Order number"]
    pdf_file = store_receipt_as_pdf(order_number)
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(order_number, screenshot, pdf_file)

def order_another_robot():
    page = browser.page()
    page.click(selector="#order-another")

def archive_receipts():
    '''
    Create a ZIP file of receipt PDF files
    '''
    folder = OUTPUT_PATH + "/" + "result"
    file_name = "result.zip"
    archive_name = OUTPUT_PATH + "/" + file_name

    #! 인스턴스화 하지않으면 발생하는 에러
    # missing 1 required positional argument: 'self'
    lib = Archive()
    lib.archive_folder_with_zip(folder=folder, archive_name=archive_name)