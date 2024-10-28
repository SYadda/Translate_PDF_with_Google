import os
import time
from pywinauto import Application
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

folder_path = r"C:\Users\hang\OneDrive\Introduction to Control and Machine Learning\Lecture notes 2024ws"  # 替换为包含pdf文件的文件夹路径

def google_translate_pdf(folder):
    # 获取文件夹中的所有pdf文件
    pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
    if len(pdf_files) == 0:
        return

    # Edge浏览器的驱动路径
    edge_driver_path = 'C:\\Program Files\\msedgedriver.exe'  # 替换为你Edge驱动的实际路径

    # Edge浏览器的选项
    options = Options()
    options.add_argument('--start-maximized')  # 启动时最大化窗口
    service = Service(edge_driver_path)

    # 启动Edge浏览器
    driver = webdriver.Edge(service=service, options=options)



    # 隐藏webdriver属性
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
        get: () => undefined
        })
    """
    })



    try:
        # 遍历所有pdf文件
        for pdf_file in pdf_files:
            # Step 2: 等待200毫秒
            time.sleep(0.2)

            # Step 3: 打开Google翻译文档上传页面
            driver.get("https://translate.google.com/?hl=zh-CN&sl=en&tl=zh-CN&op=docs")

            # Step 4: 等待100毫秒
            time.sleep(0.1)

            # Step 4: 点击接受所有的cookies
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button'))
            )
            accept_button.click()

            # Step 5: 点击浏览文件按钮
            upload_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div/div[1]/div[2]/div[2]/div/div/button'))
            )
            upload_button.click()

            # Step 6: 等待200毫秒
            time.sleep(0.2)

            # # Step 7: 在弹出的窗口中选择文件
            # file_path = os.path.join(folder, pdf_file)
            # file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            # file_input.send_keys(file_path)

            # Step 7: 在弹出的窗口中选择文件
            file_path = os.path.join(folder, pdf_file)
            # 启动应用程序
            app = Application().connect(title_re=".*打开.*")
            # 选择文件选择窗口
            file_dialog = app.window(title_re=".*打开.*")
            # 输入文件路径
            file_dialog.Edit.set_text(file_path)
            time.sleep(0.2)
            # 点击打开按钮
            file_dialog.打开.click()


            # Step 8: 等待200毫秒
            time.sleep(0.2)

            # Step 9: 点击翻译按钮
            translate_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/div/button'))
            )
            translate_button.click()

            # 【目前的问题】selenium打开的浏览器，被谷歌识别为“通过软件自动控制”的浏览器，因此拒绝服务。需要想个办法骗过它。

            # Step 10: 等待1000毫秒
            time.sleep(1)

            # Step 11: 循环等待翻译结果出现
            while True:
                # Step 11(a): 等待500毫秒
                time.sleep(0.5)

                # Step 11(b): 检查翻译结果是否可下载
                try:
                    download_button = driver.find_element(By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/button[1]')
                    download_button.click()
                    break  # 如果找到了下载按钮，退出循环
                except:
                    continue  # 否则继续循环

            # Step 12: 等待1000毫秒
            time.sleep(1)

            # Step 13: 点击清除按钮，准备处理下一个文件
            clear_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[1]/div[3]/div/span/button'))
            )
            clear_button.click()

    finally:
        # 关闭浏览器
        driver.quit()

# 示例使用
google_translate_pdf(folder_path)
