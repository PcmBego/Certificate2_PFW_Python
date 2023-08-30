from robocorp.tasks import task
import robocorp.browser as browser
import robocorp.http as http
from RPA.Tables import Tables   # Tables모듈 안 Tables클래스 사용
import time
from RPA.PDF import PDF # PDF모듈 안 PDF클래스 사용
from RPA.Archive import Archive
from RPA.Assistant import Assistant
import shutil
import os
import getpass

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
    url = user_input_task()
    if url != "Fail":
        open_robot_order_website()
        orders = get_orders()
        for row in orders:  # row: dict
            close_annoying_modal()
            fill_the_form(row)
            download_and_store_the_order_receipt_as_pdf(row)
            order_another_robot()
        archive_receipts()
        #? Teardown: it will be automatically closed when the task finishes
        copy_output_dir_to_local()


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
    file_name = "orders.csv"
    src_path = os.path.join(OUTPUT_PATH, file_name)
    #// orders = table.read_table_from_csv(path=OUTPUT_PATH + "/" + "orders.csv")   # @keyword("Read table from CSV")
    orders = table.read_table_from_csv(path=src_path)   # @keyword("Read table from CSV")
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
    #// file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name
    file_path = os.path.join(OUTPUT_PATH, dir_name, file_name)
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
    # file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name
    file_path = os.path.join(OUTPUT_PATH, dir_name, file_name)

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
    #//file_path = OUTPUT_PATH + "/" + dir_name + "/" + file_name
    file_path = os.path.join(OUTPUT_PATH, dir_name, file_name)

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
    #// folder = OUTPUT_PATH + "/" + "result"
    folder = os.path.join(OUTPUT_PATH, "result")
    file_name = "result.zip"
    #// archive_name = OUTPUT_PATH + "/" + file_name
    archive_name = os.path.join(OUTPUT_PATH, file_name)

    #! 인스턴스화 하지않으면 발생하는 에러
    # missing 1 required positional argument: 'self'
    lib = Archive()
    lib.archive_folder_with_zip(folder=folder, archive_name=archive_name)

def user_input_task():
    assistant = Assistant()
    #// assistant.add_icon("warning")
    assistant.add_heading("Input from user")
    assistant.add_text_input("text_input", label="URL", placeholder="Please enter URL")
    assistant.add_submit_buttons(buttons="Submit, Cancel", default="Submit")
    assistant.add_text("https://robotsparebinindustries.com/#/robot-order", size="small")
    result = assistant.run_dialog()
    try:
        url = result.text_input # == result["text_input"]
    except  AttributeError:
        url = "Fail"

    return url

def copy_output_dir_to_local():
    """
    .. Assistant 
        로컬 PC를 사용하는 Assistant도 Temp경로로 가상의 환경을 만들어 실행 후 제거하는것으로 보인다
        (Control room - Cloude Worker 방법과 유사)
        때문에 output폴더에 존재하는 결과물이 유지되지 않는다 -> 웹상(Control room)의 artifact로는 존재

        편집기를 통한 프로젝트를 직접 실행하는 환경에서는 불필요
    """
    src_path = OUTPUT_PATH
    #* 1. C:\Users\min88
    user_path = os.path.expanduser('~')
    #* 2. Certificate2_PFW
    current_path = os.getcwd()
    temp_list = current_path.split('\\')
    prject_name = temp_list[-1]
    #* 3. C:\Users\min88\Desktop\Certificate2_PFW_output
    dst_path = os.path.join(user_path, "Desktop", prject_name + "_output")
    # os.makedirs() 기능 포함
    shutil.copytree(src_path, dst_path, dirs_exist_ok=True)