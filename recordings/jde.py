import re
from playwright.sync_api import Page, expect

def test_example(page: Page) -> None:
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("广州中国政府报表").click()
    page.get_by_role("link", name="利润表").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("CapsLock")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).fill("TKHXLS001")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("Enter")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("CapsLock")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("checkbox", name="选择所有提取的记录").check()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("checkbox", name="数据选择").check()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("a").filter(has_text="提交(U)").click()

    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#RightOperand1").select_option("Literal")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#LITtf").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#LITtf").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#LITtf").fill("02501")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()

    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+0)").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("input[name=\"E3\"]").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("input[name=\"E3\"]").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("input[name=\"E3\"]").fill("20")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("input[name=\"E4\"]").fill("26")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("input[name=\"E28\"]").fill("10")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+0)").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("a").filter(has_text="屏幕(F)").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_text("已提交作业").click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="作业", exact=True).click()
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="作业", exact=True).press("CapsLock")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="作业", exact=True).fill("R55GZ001_TKHXLS001")
    page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="作业", exact=True).press("Enter")
    with page.expect_download() as download1_info:
        page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("link", name="R55GZ001_TKHXLS001").click()
    download1 = download1_info.value
    page.get_by_title("主页").click()