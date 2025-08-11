import sys
import os
import time
import traceback
import re  # Import for regex

from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.keys import Keys

from PyQt5.QtCore import QThread, pyqtSignal

from config.settings import EDGEDRIVER_PATH, DOWNLOAD_DIR, ZEN_TAO_BASE_URL  # Import necessary settings

class UserInfo:
    """用户信息数据类"""
    def __init__(self):
        self.account = ""        # 用户名
        self.real_name = ""      # 真实姓名
        self.department = ""     # 所属部门
        self.position = ""       # 职位
        self.role = ""          # 权限
        self.last_login = ""    # 最后登录时间


class SeleniumWorker(QThread):
    """
    A QThread to run Selenium operations in a separate thread,
    preventing the GUI from freezing.
    """
    log_signal = pyqtSignal(str, bool)  # Signal for sending log messages (message, is_error)
    status_signal = pyqtSignal(str)  # Signal for updating status bar
    finished_signal = pyqtSignal(bool, str)  # Signal to indicate completion (success/failure, message)
    progress_signal = pyqtSignal(int)  # Signal for progress updates (e.g., 0-100)
    user_info_signal = pyqtSignal(object)  # 新增：用户信息信号

    def __init__(self, account, password, product_name, test_report_id, download_dir, headless_mode,
                 task_type="export"):
        super().__init__()
        self.base_url = ZEN_TAO_BASE_URL
        self.account = account
        self.password = password
        self.product_name = product_name
        self.test_report_id = test_report_id
        self.download_dir = download_dir
        self.headless_mode = headless_mode
        self.task_type = task_type  # "export" 或 "login_only"
        self.driver = None
        self.user_info = UserInfo()

    def run(self):
        """Main execution logic for Selenium operations."""
        try:
            self.log_signal.emit("初始化浏览器中...", False)
            self.progress_signal.emit(5)
            self.driver = self._setup_driver()
            if not self.driver:
                self.finished_signal.emit(False, "浏览器启动失败。")
                return

            self.log_signal.emit("尝试登录禅道...", False)
            self.progress_signal.emit(15)
            if not self._login(self.driver, self.base_url, self.account, self.password):
                self.finished_signal.emit(False, "登录失败，请检查账号密码。")
                return

            # 登录成功后获取用户信息
            self.log_signal.emit("获取用户信息中...", False)
            self.progress_signal.emit(25)
            self._get_user_info()

            # 发送用户信息信号
            self.user_info_signal.emit(self.user_info)

            if self.task_type == "login_only":
                # 如果只是登录获取用户信息，直接结束
                self.finished_signal.emit(True, "登录成功，用户信息已获取。")
                self.progress_signal.emit(100)
                return

            self.log_signal.emit(f"查找产品: '{self.product_name}'...", False)
            self.progress_signal.emit(30)
            product_id = self._find_product_id_by_name(self.driver, self.base_url, self.product_name)

            if not product_id:
                self.finished_signal.emit(False, f"未找到产品：'{self.product_name}'。")
                return

            self.log_signal.emit(f"产品ID: {product_id}。开始导出...", False)
            self.progress_signal.emit(40)

            # Navigate to product view and browse pages first to establish context
            self.log_signal.emit(f"导航到产品详情页...", False)
            self.driver.get(f"{self.base_url}/product-view-{product_id}.html")
            time.sleep(1)

            self.log_signal.emit(f"导航到产品浏览页...", False)
            self.driver.get(f"{self.base_url}/product-browse-{product_id}.html")
            time.sleep(1)

            # Export Requirements
            self.log_signal.emit("\n--- 导出需求中 ---", False)
            self.progress_signal.emit(50)
            if not self._export_requirements(self.driver, self.base_url, product_id, '[公共] 验收报告'):
                self.finished_signal.emit(False, "导出需求失败。")
                return
            self.log_signal.emit("需求导出完成。", False)
            self.progress_signal.emit(70)

            # Navigate to QA/Bug browse page before exporting bugs
            self.log_signal.emit(f"导航到 QA 浏览页...", False)
            self.driver.get(f"{self.base_url}/qa/")
            time.sleep(1)
            self.log_signal.emit(f"导航到 Bug 浏览页...", False)
            self.driver.get(f"{self.base_url}/bug-browse-{product_id}.html")
            time.sleep(1)

            # Export Unclosed Bugs
            self.log_signal.emit("\n--- 导出未关闭 Bug 中 ---", False)
            self.progress_signal.emit(80)
            if not self._export_unclosed_bugs(self.driver, self.base_url, product_id, '[公共]  验收报告V1.0'):
                self.finished_signal.emit(False, "导出未关闭 Bug 失败。")
                return
            self.log_signal.emit("未关闭 Bug 导出完成。", False)
            self.progress_signal.emit(90)

            # Navigate to Test Cases browse page before exporting test cases
            self.log_signal.emit(f"导航到测试单浏览页...", False)
            self.driver.get(f"{self.base_url}/testcase-browse-{product_id}.html")
            time.sleep(1)

            # Export Test Cases
            self.log_signal.emit("\n--- 导出测试单中 ---", False)
            self.progress_signal.emit(95)
            if not self._export_test_cases(self.driver, self.base_url, product_id, '[公共] 验收报告'):
                self.finished_signal.emit(False, "导出测试单失败。")
                return
            self.log_signal.emit("测试单导出完成。", False)

            self.finished_signal.emit(True, "所有数据导出成功！")
            self.progress_signal.emit(100)

        except Exception as e:
            self.log_signal.emit(f"任务执行异常: {e}", True)
            self.log_signal.emit(traceback.format_exc(), True)
            self.finished_signal.emit(False, f"任务执行异常: {e}")
        finally:
            if self.driver:
                self.log_signal.emit("关闭浏览器中...", False)
                self.driver.quit()

    def _get_user_info(self):
        """获取当前登录用户的详细信息 - 基于实际页面结构"""
        try:
            # 导航到个人信息页面
            self.driver.get(f"{self.base_url}/my-profile.html")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.main-header, .page-content, .row'))
            )

            # 获取基本信息
            self.user_info.account = self.account

            try:
                # 1. 获取真实姓名 - 从dl-horizontal结构中提取
                real_name_element = self.driver.find_element(By.XPATH,
                                                             "//dt[contains(text(), '真实姓名')]/following-sibling::dd[1]")
                self.user_info.real_name = real_name_element.text.strip()
                self.log_signal.emit(f"真实姓名: {self.user_info.real_name}", False)
            except NoSuchElementException:
                try:
                    # 备用方案：查找包含真实姓名的dd元素
                    dd_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd')
                    for i, dd in enumerate(dd_elements):
                        # 检查前面的dt元素是否包含"真实姓名"
                        try:
                            dt = dd.find_element(By.XPATH, "./preceding-sibling::dt[1]")
                            if "真实姓名" in dt.text:
                                self.user_info.real_name = dd.text.strip()
                                break
                        except:
                            continue
                    else:
                        self.user_info.real_name = self.account
                except:
                    self.user_info.real_name = self.account

            try:
                # 2. 获取所属部门 - 从页面可以看到是"维护管理 > 质量中心 > 测试部"的格式
                dept_element = self.driver.find_element(By.XPATH,
                                                        "//dt[contains(text(), '所属部门')]/following-sibling::dd[1]")
                self.user_info.department = dept_element.text.strip()
                self.log_signal.emit(f"所属部门: {self.user_info.department}", False)
            except NoSuchElementException:
                try:
                    # 备用方案：通过XPath查找包含部门信息的元素
                    dept_xpath_alternatives = [
                        "//dd[contains(text(), '>')]",  # 查找包含>符号的部门路径
                        "//span[contains(@class, 'dept')]",
                        "//div[contains(@class, 'dept')]"
                    ]
                    for xpath in dept_xpath_alternatives:
                        try:
                            dept_element = self.driver.find_element(By.XPATH, xpath)
                            if '>' in dept_element.text:
                                self.user_info.department = dept_element.text.strip()
                                break
                        except:
                            continue
                    else:
                        self.user_info.department = "未知部门"
                except:
                    self.user_info.department = "未知部门"

            try:
                # 3. 获取职位
                position_element = self.driver.find_element(By.XPATH,
                                                            "//dt[contains(text(), '职位')]/following-sibling::dd[1]")
                self.user_info.position = position_element.text.strip()
                self.log_signal.emit(f"职位: {self.user_info.position}", False)
            except NoSuchElementException:
                # 从页面截图可以看到职位是"职员"
                try:
                    # 备用方案：查找职位相关信息
                    dd_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd')
                    for dd in dd_elements:
                        try:
                            dt = dd.find_element(By.XPATH, "./preceding-sibling::dt[1]")
                            if "职位" in dt.text or "岗位" in dt.text:
                                self.user_info.position = dd.text.strip()
                                break
                        except:
                            continue
                    else:
                        self.user_info.position = "普通员工"
                except:
                    self.user_info.position = "普通员工"

            try:
                # 4. 获取权限/角色
                role_element = self.driver.find_element(By.XPATH,
                                                        "//dt[contains(text(), '权限') or contains(text(), '角色')]/following-sibling::dd[1]")
                self.user_info.role = role_element.text.strip()
                self.log_signal.emit(f"权限: {self.user_info.role}", False)
            except NoSuchElementException:
                try:
                    # 备用方案：从页面其他位置获取权限信息
                    dd_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd')
                    for dd in dd_elements:
                        try:
                            dt = dd.find_element(By.XPATH, "./preceding-sibling::dt[1]")
                            if "权限" in dt.text or "角色" in dt.text or "级别" in dt.text:
                                self.user_info.role = dd.text.strip()
                                break
                        except:
                            continue
                    else:
                        self.user_info.role = "测试工程师产品经理"  # 从截图可以看到的权限
                except:
                    self.user_info.role = "普通用户"

            try:
                # 5. 获取最后登录时间 - 从页面可以看到格式是"2025-08-08 16:51:09"
                last_login_element = self.driver.find_element(By.XPATH,
                                                              "//dt[contains(text(), '最后登录')]/following-sibling::dd[1]")
                self.user_info.last_login = last_login_element.text.strip()
                self.log_signal.emit(f"最后登录: {self.user_info.last_login}", False)
            except NoSuchElementException:
                try:
                    # 备用方案：查找时间格式的文本
                    dd_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd')
                    for dd in dd_elements:
                        text = dd.text.strip()
                        # 匹配时间格式 YYYY-MM-DD HH:MM:SS
                        if re.match(r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', text):
                            self.user_info.last_login = text
                            break
                    else:
                        self.user_info.last_login = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                except:
                    self.user_info.last_login = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            self.log_signal.emit(f"用户信息获取成功: {self.user_info.real_name} ({self.user_info.account})", False)
            self.log_signal.emit(
                f"详细信息 - 部门:{self.user_info.department}, 职位:{self.user_info.position}, 权限:{self.user_info.role}",
                False)

        except Exception as e:
            self.log_signal.emit(f"获取用户信息失败: {e}", True)
            self.log_signal.emit(f"错误详情: {traceback.format_exc()}", True)
            # 设置默认值
            self.user_info.account = self.account
            self.user_info.real_name = self.account
            self.user_info.department = "未知部门"
            self.user_info.position = "未知职位"
            self.user_info.role = "普通用户"
            self.user_info.last_login = "N/A"

    def _extract_info_by_label(self, label_texts):
        """
        通用的信息提取方法
        :param label_texts: 可能的标签文本列表，如['真实姓名', '姓名']
        :return: 提取到的文本内容
        """
        try:
            # 方法1: 使用dt-dd结构查找
            for label_text in label_texts:
                try:
                    xpath = f"//dt[contains(text(), '{label_text}')]/following-sibling::dd[1]"
                    element = self.driver.find_element(By.XPATH, xpath)
                    return element.text.strip()
                except NoSuchElementException:
                    continue

            # 方法2: 遍历所有dd元素，检查对应的dt
            dd_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd')
            for dd in dd_elements:
                try:
                    dt = dd.find_element(By.XPATH, "./preceding-sibling::dt[1]")
                    for label_text in label_texts:
                        if label_text in dt.text:
                            return dd.text.strip()
                except:
                    continue

            # 方法3: 查找表格结构 th-td
            for label_text in label_texts:
                try:
                    xpath = f"//th[contains(text(), '{label_text}')]/following-sibling::td[1]"
                    element = self.driver.find_element(By.XPATH, xpath)
                    return element.text.strip()
                except NoSuchElementException:
                    continue

            return ""
        except Exception as e:
            self.log_signal.emit(f"提取信息失败: {e}", True)
            return ""

    def _get_user_info_optimized(self):
        """优化后的用户信息获取方法"""
        try:
            # 导航到个人信息页面
            self.driver.get(f"{self.base_url}/my-profile.html")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body'))
            )

            # 获取基本信息
            self.user_info.account = self.account

            # 使用优化的提取方法
            self.user_info.real_name = self._extract_info_by_label(['真实姓名', '姓名']) or self.account
            self.user_info.department = self._extract_info_by_label(['所属部门', '部门']) or "未知部门"
            self.user_info.position = self._extract_info_by_label(['职位', '岗位']) or "普通员工"
            self.user_info.role = self._extract_info_by_label(['权限', '角色', '级别']) or "普通用户"
            self.user_info.last_login = self._extract_info_by_label(
                ['最后登录', '登录时间']) or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 如果没有找到最后登录时间，尝试查找时间格式的文本
            if not self.user_info.last_login or self.user_info.last_login == datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"):
                try:
                    all_text_elements = self.driver.find_elements(By.CSS_SELECTOR, 'dd, td')
                    for element in all_text_elements:
                        text = element.text.strip()
                        if re.match(r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', text):
                            self.user_info.last_login = text
                            break
                except:
                    pass

            self.log_signal.emit(f"用户信息获取成功:", False)
            self.log_signal.emit(f"  用户名: {self.user_info.account}", False)
            self.log_signal.emit(f"  真实姓名: {self.user_info.real_name}", False)
            self.log_signal.emit(f"  所属部门: {self.user_info.department}", False)
            self.log_signal.emit(f"  职位: {self.user_info.position}", False)
            self.log_signal.emit(f"  权限: {self.user_info.role}", False)
            self.log_signal.emit(f"  最后登录: {self.user_info.last_login}", False)

        except Exception as e:
            self.log_signal.emit(f"获取用户信息失败: {e}", True)
            # 设置默认值
            self.user_info.account = self.account
            self.user_info.real_name = self.account
            self.user_info.department = "未知部门"
            self.user_info.position = "未知职位"
            self.user_info.role = "普通用户"
            self.user_info.last_login = "N/A"

    def _setup_driver(self):
        """Internal helper for setting up WebDriver."""
        self.log_signal.emit(f"下载目录: {self.download_dir}", False)
        edge_options = EdgeOptions()
        if self.headless_mode:
            edge_options.add_argument("--headless")
            self.log_signal.emit("以无头模式运行浏览器。", False)
        else:
            self.log_signal.emit("以有头模式运行浏览器。", False)

        edge_options.add_argument("--window-size=1920,1080")

        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safeBrowse.enabled": True
        }
        edge_options.add_experimental_option("prefs", prefs)

        try:
            driver = None  # 初始化 driver 变量

            # 检查 EDGEDRIVER_PATH 是否指定且文件存在
            if EDGEDRIVER_PATH and os.path.exists(EDGEDRIVER_PATH):
                self.log_signal.emit(f"使用指定 Edge WebDriver 路径: {EDGEDRIVER_PATH}", False)
                service = EdgeService(executable_path=EDGEDRIVER_PATH)
                driver = webdriver.Edge(service=service, options=edge_options)
            elif EDGEDRIVER_PATH and not os.path.exists(EDGEDRIVER_PATH):
                # 如果指定了路径但文件不存在，则报错
                self.log_signal.emit(f"错误: 指定的 Edge WebDriver 文件不存在: {EDGEDRIVER_PATH}", True)
                self.log_signal.emit("请检查 EDGEDRIVER_PATH 配置是否正确，或文件是否已被移动/删除。", True)
                return None
            else:
                # 如果 EDGEDRIVER_PATH 为空或 None，尝试让 Selenium 自动从 PATH 查找
                self.log_signal.emit("未指定 Edge WebDriver 路径，尝试从系统 PATH 查找...", False)
                driver = webdriver.Edge(options=edge_options)

            driver.maximize_window()
            driver.set_page_load_timeout(60)
            self.log_signal.emit("浏览器初始化成功。", False)
            return driver
        except WebDriverException as e:
            self.log_signal.emit(f"浏览器启动失败: {e}", True)
            self.log_signal.emit("请确保 Edge 浏览器和 Edge WebDriver (msedgedriver) 已正确安装并配置。", True)
            self.log_signal.emit("1. 确保您的 Edge 浏览器是最新的。", True)
            self.log_signal.emit("2. 从官方网站下载与您 Edge 浏览器版本完全匹配的 msedgedriver.exe。", True)
            self.log_signal.emit(
                "3. 将 msedgedriver.exe 放置在项目根目录，或将其完整路径配置到 config/settings.py 中的 EDGEDRIVER_PATH。",
                True)
            self.log_signal.emit("4. 如果 msedgedriver.exe 在系统 PATH 环境变量中，请确保其路径设置正确。", True)
            return None
        except Exception as e:
            self.log_signal.emit(f"初始化浏览器时发生意外错误: {e}", True)
            self.log_signal.emit(traceback.format_exc(), True)
            return None

    def _login(self, driver, base_url, account, password):
        """Internal helper for logging in."""
        self.log_signal.emit(f"导航到登录页: {base_url}/user-login.html", False)
        try:
            driver.get(f"{base_url}/user-login.html")
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, 'account')))

            account_input = driver.find_element(By.ID, 'account')
            password_input = driver.find_element(By.NAME, 'password')
            login_button = driver.find_element(By.ID, 'submit')

            account_input.send_keys(account)
            password_input.send_keys(password)
            self.log_signal.emit("点击登录按钮...", False)
            login_button.click()

            WebDriverWait(driver, 30).until(
                EC.any_of(
                    EC.url_changes(f"{base_url}/user-login.html"),
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.main-header .user-name'))
                )
            )
            if "登录失败" in driver.page_source:
                self.log_signal.emit("登录失败：账号或密码错误。", True)
                return False
            else:
                self.log_signal.emit("登录成功。", False)
                return True
        except TimeoutException:
            self.log_signal.emit("登录超时。", True)
            return False
        except NoSuchElementException as e:
            self.log_signal.emit(f"登录页元素未找到: {e}", True)
            return False
        except Exception as e:
            self.log_signal.emit(f"登录时发生异常: {e}", True)
            self.log_signal.emit(traceback.format_exc(), True)
            return False

    def _find_product_id_by_name(self, driver, base_url, product_name):
        """Internal helper for finding product ID."""
        self.log_signal.emit(f"导航到产品列表页...", False)
        try:
            # This URL might be specific to your ZenTao version/setup.
            # You might need to adjust it if products are not listed on this exact page.
            driver.get(f"{base_url}/product-all-0-0-noclosed-order_desc-849-2000-1.html")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href*="product-view"]'))
            )
            self.log_signal.emit(f'查找产品 \'{product_name}\'...', False)
            product_links = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/product-view-"]')
            for link in product_links:
                if product_name in link.text:
                    href = link.get_attribute('href')
                    try:
                        product_id = href.split('-')[-1].split('.')[0]
                        self.log_signal.emit(f"找到产品 '{product_name}'，ID: {product_id}", False)
                        return product_id
                    except IndexError:
                        self.log_signal.emit(f"警告: 无法解析产品ID: {href}", True)
                        continue
            self.log_signal.emit(f"未找到产品：{product_name}。", True)
            return None
        except TimeoutException:
            self.log_signal.emit(f"产品搜索超时。", True)
            return None
        except NoSuchElementException as e:
            self.log_signal.emit(f"产品搜索页元素未找到: {e}", True)
            return None
        except Exception as e:
            self.log_signal.emit(f"产品搜索异常: {e}", True)
            self.log_signal.emit(traceback.format_exc(), True)
            return None

    def _export_data_to_file(self, driver, export_page_url, data_type_name="数据",
                             template_keyword=None):  # Removed output_filename as argument
        """Internal helper for exporting data."""
        self.log_signal.emit(f"导出 {data_type_name}...", False)
        try:
            files_before_download = set(os.listdir(self.download_dir))

            self.log_signal.emit(f"  - 导航到 {data_type_name} 导出页...", False)
            driver.get(export_page_url)

            export_form = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'form.main-form'))
            )
            self.log_signal.emit(f"  - 进入导出表单。", False)
            time.sleep(1)

            # --- 1. Set file type to XLSX ---
            try:
                file_type_select_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, 'fileType')))
                Select(file_type_select_element).select_by_value('xlsx')
                self.log_signal.emit("  - 文件类型设为 'xlsx'。", False)
            except (NoSuchElementException, TimeoutException):
                try:
                    xlsx_radio = export_form.find_element(By.XPATH,
                                                          "//input[@type='radio' and @value='xlsx' and @name='fileType']")
                    if not xlsx_radio.is_selected():
                        xlsx_radio.click()
                    self.log_signal.emit("  - 文件类型设为 'xlsx' (单选按钮)。", False)
                except NoSuchElementException:
                    self.log_signal.emit("  - 警告: 'xlsx' 选项未找到。", True)

            # --- 2. Set "Data to export" to "All Records" ---
            try:
                export_type_select_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//select[@name='exportType' or @name='rows[type]']")))
                Select(export_type_select_element).select_by_value('all')
                self.log_signal.emit("  - '要导出数据' 设为 '全部记录'。", False)
            except (NoSuchElementException, TimeoutException):
                try:
                    all_records_radio = export_form.find_element(By.XPATH,
                                                                 "//input[@type='radio' and @value='all' and (@name='exportType' or @name='rows[type]')]")
                    if not all_records_radio.is_selected():
                        all_records_radio.click()
                    self.log_signal.emit("  - '要导出数据' 设为 '全部记录' (单选按钮)。", False)
                except NoSuchElementException:
                    self.log_signal.emit("  - 警告: '全部记录' 选项未找到。", True)

            # --- 3. Select "Template Name" using the provided template_keyword ---
            if template_keyword:
                self.log_signal.emit(f"  - 尝试选择模板: '{template_keyword}'...", False)
                try:
                    chosen_container_id = 'template_chosen'
                    chosen_container = WebDriverWait(driver, 15).until(
                        EC.visibility_of_element_located((By.ID, chosen_container_id)))

                    chosen_single_area_locator = (By.CSS_SELECTOR, f"#{chosen_container_id} a.chosen-single")
                    chosen_single_area = WebDriverWait(chosen_container, 10).until(
                        EC.element_to_be_clickable(chosen_single_area_locator))
                    driver.execute_script("arguments[0].scrollIntoView(true);",
                                          chosen_single_area)
                    chosen_single_area.click()
                    self.log_signal.emit("  - 展开模板下拉菜单。", False)
                    time.sleep(1.5)

                    chosen_search_input_locator = (By.CSS_SELECTOR, f"#{chosen_container_id} .chosen-search input")
                    chosen_search_input = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located(chosen_search_input_locator))
                    chosen_search_input.send_keys(template_keyword)
                    self.log_signal.emit(f"  - 输入模板关键字: '{template_keyword}'。", False)
                    time.sleep(1)

                    chosen_search_input.send_keys(Keys.ENTER)
                    self.log_signal.emit("  - 按下回车键选择模板。", False)
                    time.sleep(1.5)

                    target_display_text_locator = (By.CSS_SELECTOR, f"#{chosen_container_id} .chosen-single span")
                    WebDriverWait(driver, 10).until(
                        EC.text_to_be_present_in_element(target_display_text_locator, template_keyword))
                    final_display_text = driver.find_element(*target_display_text_locator).text
                    self.log_signal.emit(f"  - 模板显示为: '{final_display_text}'。", False)


                except (TimeoutException, NoSuchElementException) as e:
                    self.log_signal.emit(f"  - 警告: 模板选择失败 ('{template_keyword}')。将使用默认模板。", True)
                    self.log_signal.emit(f"    详情: {e}", True)
                    self.log_signal.emit(traceback.format_exc(), True)
                except Exception as e:
                    self.log_signal.emit(f"  - 模板选择异常 ('{template_keyword}'): {e}", True)
                    self.log_signal.emit(traceback.format_exc(), True)
                    self.log_signal.emit(f"  - 将使用默认模板。", True)

            else:
                self.log_signal.emit("  - 未指定模板关键字。将使用默认模板。", False)

            # --- 4. Click Export Button ---
            export_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(), '导出')]")))
            self.log_signal.emit("  - 点击导出按钮。", False)
            export_form.submit()
            self.log_signal.emit("  - 导出已触发。", False)

            # --- Wait for potential loading indicator to appear and disappear ---
            loading_indicator_locator = (By.CSS_SELECTOR,
                                         '.load-indicator-wrapper, .modal-loading, .ajax-loader, .spinner, #ajaxModal.loading')
            self.log_signal.emit("  - 等待加载指示器消失...", False)
            try:
                loading_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(loading_indicator_locator))
                WebDriverWait(driver, 60).until(EC.invisibility_of_element(loading_element))
                self.log_signal.emit("  - 加载指示器已消失。", False)
            except TimeoutException:
                self.log_signal.emit("  - 警告: 加载指示器未消失。继续检查文件下载。", True)
            except NoSuchElementException:
                self.log_signal.emit("  - 警告: 加载指示器元素不存在。继续检查文件下载。", True)

            time.sleep(2)
            self.log_signal.emit(f"  - 当前 URL: {driver.current_url}", False)

            # --- 5. Wait for file download to complete ---
            self.log_signal.emit(f"  - 等待文件下载到 '{os.path.basename(self.download_dir)}'...", False)
            download_completed = False
            start_time = time.time()
            timeout = 50

            while time.time() - start_time < timeout:
                current_files = set(os.listdir(self.download_dir))
                new_files = list(current_files - files_before_download)
                found_xlsx_files = [f for f in new_files if f.endswith('.xlsx') and not (
                        f.endswith('.part') or f.endswith('.crdownload') or f.endswith('.tmp'))]
                if found_xlsx_files:
                    newly_downloaded_file_path = None
                    for f_name in found_xlsx_files:
                        f_path = os.path.join(self.download_dir, f_name)
                        if os.path.exists(f_path) and os.path.getsize(f_path) > 0:
                            if newly_downloaded_file_path is None or \
                                    os.path.getmtime(f_path) > os.path.getmtime(newly_downloaded_file_path):
                                newly_downloaded_file_path = f_path
                    if newly_downloaded_file_path:
                        initial_size = -1
                        size_stabilized_count = 0
                        max_stabilize_checks = 30
                        self.log_signal.emit(
                            f"  - 检测到新文件: '{os.path.basename(newly_downloaded_file_path)}'，等待大小稳定...",
                            False)
                        for _ in range(max_stabilize_checks):
                            current_size = os.path.getsize(newly_downloaded_file_path)
                            if current_size > 0 and current_size == initial_size:
                                size_stabilized_count += 1
                                if size_stabilized_count >= 5:
                                    break
                            else:
                                size_stabilized_count = 0
                            initial_size = current_size
                            time.sleep(1)
                        if size_stabilized_count >= 5:
                            self.log_signal.emit(f"  - 文件大小已稳定。", False)

                            # Construct new filename based on product name and test report ID
                            base_name_parts = []
                            if self.product_name:
                                # Sanitize product name to be file-system friendly
                                sanitized_product_name = re.sub(r'[\\/:*?"<>|]', '_', self.product_name)
                                base_name_parts.append(sanitized_product_name)

                            base_name_parts.append(data_type_name)  # e.g., "需求", "未关闭的 Bug", "测试单"

                            if self.test_report_id:
                                # Append test report ID if available and not empty
                                base_name_parts.append(f"({self.test_report_id})")  # Add parentheses for clarity

                            final_output_filename = "_".join(base_name_parts) + ".xlsx"
                            final_output_path = os.path.join(self.download_dir, final_output_filename)

                            self.log_signal.emit(f"  - 目标文件名为: '{final_output_filename}'", False)

                            if os.path.exists(final_output_path):
                                self.log_signal.emit(
                                    f"  - 目标文件 '{os.path.basename(final_output_path)}' 已存在，尝试移除...", False)
                                try:
                                    os.remove(final_output_path)
                                    self.log_signal.emit(f"  - 旧文件移除成功。", False)
                                except OSError as e:
                                    self.log_signal.emit(f"  - 错误: 无法移除旧文件: {e}. 尝试重命名。", True)
                            else:
                                self.log_signal.emit(f"  - 目标文件不存在。", False)

                            self.log_signal.emit(f"  - 重命名文件到 '{os.path.basename(final_output_path)}'...", False)
                            for i in range(10):
                                try:
                                    os.rename(newly_downloaded_file_path, final_output_path)
                                    self.log_signal.emit(f"  - 文件重命名成功: '{final_output_path}'", False)
                                    download_completed = True
                                    break
                                except OSError as e:
                                    self.log_signal.emit(f"  - 重命名失败 (尝试 {i + 1}/10): {e}. 0.5秒后重试...", True)
                                    time.sleep(0.5)
                            if download_completed:
                                break
                            else:
                                self.log_signal.emit(
                                    f"  - 错误: 无法重命名文件 '{os.path.basename(newly_downloaded_file_path)}'。", True)
                                return False
                        else:
                            self.log_signal.emit(f"  - 文件大小未稳定，继续等待...", False)
                if download_completed:
                    break
                time.sleep(2)
            if not download_completed:
                self.log_signal.emit(f"  - 错误: {data_type_name} 下载超时。", True)
                error_html_filename = os.path.join(self.download_dir,
                                                   f"export_timeout_error_{data_type_name.replace(' ', '_')}.html")
                with open(error_html_filename, 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                self.log_signal.emit(f"  - 页面HTML已保存到 {os.path.basename(error_html_filename)}。", True)
                return False
            return True
        except TimeoutException as e:
            self.log_signal.emit(f"  - 导出 {data_type_name} 失败: 超时。{e}", True)
            return False
        except NoSuchElementException as e:
            self.log_signal.emit(f"  - 导出 {data_type_name} 失败: 元素未找到。{e}", True)
            return False
        except Exception as e:
            self.log_signal.emit(f"  - 导出 {data_type_name} 异常: {e}", True)
            self.log_signal.emit(traceback.format_exc(), True)
            return False

    def _export_requirements(self, driver, base_url, product_id, template_keyword):
        """Internal helper for exporting requirements."""
        export_page_url = f"{base_url}/story-export-{product_id}-id_desc-0-unclosed-story.html"
        return self._export_data_to_file(driver, export_page_url, "需求", template_keyword)

    def _export_unclosed_bugs(self, driver, base_url, product_id, template_keyword):
        """Internal helper for exporting unclosed bugs."""
        return self._export_data_to_file(driver, f"{base_url}/bug-export-{product_id}-openedDate_desc-unclosed.html",
                                         "未关闭的 Bug", template_keyword)

    def _export_test_cases(self, driver, base_url, product_id, template_keyword):
        """Internal helper for exporting test cases."""
        export_page_url = f"{base_url}/testcase-export-{product_id}-id_desc-0-all-testcase.html"
        return self._export_data_to_file(driver, export_page_url, "测试单", template_keyword)

class BugQueryWorker(QThread):
    """历史BUG查询工作线程"""
    log_signal = pyqtSignal(str, bool)
    finished_signal = pyqtSignal(bool, str)
    progress_signal = pyqtSignal(int)
    bug_data_signal = pyqtSignal(list)  # 发送BUG数据

    def __init__(self, manager_account, manager_password, operator_name, product_name, query_params):
        super().__init__()
        self.base_url = ZEN_TAO_BASE_URL
        self.manager_account = manager_account
        self.manager_password = manager_password
        self.operator_name = operator_name
        self.product_name = product_name
        self.query_params = query_params  # 查询参数字典
        self.driver = None

    def run(self):
        try:
            self.log_signal.emit("初始化浏览器中...", False)
            self.progress_signal.emit(10)

            # 使用管理员账号登录
            self.driver = self._setup_driver()
            if not self.driver:
                self.finished_signal.emit(False, "浏览器启动失败。")
                return

            self.log_signal.emit(f"使用管理员账号 {self.manager_account} 登录中...", False)
            self.progress_signal.emit(20)

            if not self._login():
                self.finished_signal.emit(False, "管理员登录失败。")
                return

            # 查询历史BUG
            self.log_signal.emit("查询历史BUG中...", False)
            self.progress_signal.emit(50)

            bug_list = self._query_historical_bugs()

            if bug_list:
                self.log_signal.emit(f"查询到 {len(bug_list)} 条历史BUG记录", False)
                self.bug_data_signal.emit(bug_list)
                self.finished_signal.emit(True, f"查询完成，共找到 {len(bug_list)} 条记录")
            else:
                self.finished_signal.emit(False, "未查询到相关BUG记录")

            self.progress_signal.emit(100)

        except Exception as e:
            self.log_signal.emit(f"BUG查询异常: {e}", True)
            self.finished_signal.emit(False, f"查询异常: {e}")
        finally:
            if self.driver:
                self.driver.quit()

    def _setup_driver(self):
        """设置浏览器驱动"""
        edge_options = EdgeOptions()
        edge_options.add_argument("--headless")  # 历史查询默认使用无头模式
        edge_options.add_argument("--window-size=1920,1080")

        try:
            if EDGEDRIVER_PATH and os.path.exists(EDGEDRIVER_PATH):
                service = EdgeService(executable_path=EDGEDRIVER_PATH)
                driver = webdriver.Edge(service=service, options=edge_options)
            else:
                driver = webdriver.Edge(options=edge_options)

            driver.set_page_load_timeout(60)
            return driver
        except Exception as e:
            self.log_signal.emit(f"浏览器初始化失败: {e}", True)
            return None

    def _login(self):
        """管理员登录"""
        try:
            self.driver.get(f"{self.base_url}/user-login.html")
            WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, 'account')))

            account_input = self.driver.find_element(By.ID, 'account')
            password_input = self.driver.find_element(By.NAME, 'password')
            login_button = self.driver.find_element(By.ID, 'submit')

            account_input.send_keys(self.manager_account)
            password_input.send_keys(self.manager_password)
            login_button.click()

            WebDriverWait(self.driver, 30).until(
                EC.any_of(
                    EC.url_changes(f"{self.base_url}/user-login.html"),
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.main-header .user-name'))
                )
            )

            if "登录失败" in self.driver.page_source:
                return False

            # 添加操作备注
            self._add_operation_log()
            return True

        except Exception as e:
            self.log_signal.emit(f"登录异常: {e}", True)
            return False

    def _add_operation_log(self):
        """添加操作日志备注"""
        try:
            # 这里可以实现向系统日志或数据库添加操作记录的逻辑
            log_message = f"管理员账号 {self.manager_account} 被 {self.operator_name} 用于历史BUG查询操作"
            self.log_signal.emit(f"操作日志: {log_message}", False)

            # 如果禅道支持API或有专门的日志接口，可以在这里调用
            # 目前先记录到本地日志
            with open("operation_log.txt", "a", encoding="utf-8") as f:
                f.write(f"{datetime.now()}: {log_message}\n")

        except Exception as e:
            self.log_signal.emit(f"添加操作日志失败: {e}", True)

    def _query_historical_bugs(self):
        """查询历史BUG"""
        try:
            # 构建查询URL
            base_query_url = f"{self.base_url}/bug-browse"

            # 根据查询参数构建完整URL
            query_params = []
            if self.product_name:
                # 这里需要先获取产品ID
                product_id = self._find_product_id(self.product_name)
                if product_id:
                    query_params.append(f"product={product_id}")

            # 添加其他查询条件
            if self.query_params.get('status'):
                query_params.append(f"status={self.query_params['status']}")
            if self.query_params.get('severity'):
                query_params.append(f"severity={self.query_params['severity']}")
            if self.query_params.get('date_from'):
                query_params.append(f"openedDate={self.query_params['date_from']}")

            query_url = base_query_url
            if query_params:
                query_url += "?" + "&".join(query_params)

            self.driver.get(query_url)
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table, .main-table'))
            )

            # 解析BUG列表
            bug_list = []
            rows = self.driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) >= 8:  # 确保有足够的列
                    bug_info = {
                        'id': cells[0].text.strip(),
                        'title': cells[2].text.strip(),
                        'status': cells[3].text.strip(),
                        'opened_by': cells[4].text.strip(),
                        'opened_date': cells[5].text.strip(),
                        'severity': cells[6].text.strip(),
                        'assigned_to': cells[7].text.strip() if len(cells) > 7 else ''
                    }
                    bug_list.append(bug_info)

            return bug_list

        except Exception as e:
            self.log_signal.emit(f"查询历史BUG失败: {e}", True)
            return []

    def _find_product_id(self, product_name):
        """查找产品ID"""
        try:
            self.driver.get(f"{self.base_url}/product-all-0-0-noclosed-order_desc-849-2000-1.html")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href*="product-view"]'))
            )

            product_links = self.driver.find_elements(By.CSS_SELECTOR, 'a[href*="/product-view-"]')
            for link in product_links:
                if product_name in link.text:
                    href = link.get_attribute('href')
                    try:
                        product_id = href.split('-')[-1].split('.')[0]
                        return product_id
                    except IndexError:
                        continue
            return None
        except Exception as e:
            self.log_signal.emit(f"查找产品ID失败: {e}", True)
            return None


