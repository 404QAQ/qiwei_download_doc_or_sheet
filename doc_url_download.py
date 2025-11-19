#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import time
import json
import re
import shutil
import logging
from pathlib import Path
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------------- 配置区 ----------------
# 根目录配置 - 会遍历这个目录下的所有子目录
ROOT_DIRECTORY = os.path.abspath("")  # 修改为你的根目录

# Cookie 配置
cookie_file = "cookies.json"  # 放在脚本同目录下

# 浏览器配置
USE_REAL_PROFILE = False
CHROME_PROFILE_PATH = ""
PROFILE_NAME = "Default"

headless = False

# 超时设置
PAGE_LOAD_TIMEOUT = 30
WAIT_TIMEOUT = 15
DOWNLOAD_TIMEOUT = 120

# 基本等待时间
MENU_WAIT = 2
CLICK_WAIT = 1
PAGE_STABLE_WAIT = 3

# 下载记录文件
DOWNLOAD_LOG_FILE = "downloaded_files.txt"

# ---------------------------------------------------------

# 配置日志格式
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


def safe_filename(name: str):
    """清理非法文件名字符"""
    name = re.sub(r'[\\/:"*?<>|]+', "_", str(name).strip())
    return name or "unnamed"


def guess_ext_from_url(url: str):
    url = str(url).lower()
    if "sheet" in url:
        return "xlsx"
    if "doc" in url:
        return "docx"
    return "bin"


def format_time(seconds):
    """格式化时间"""
    if seconds < 60:
        return f"{int(seconds)}秒"
    elif seconds < 3600:
        return f"{int(seconds // 60)}分{int(seconds % 60)}秒"
    else:
        return f"{int(seconds // 3600)}小时{int((seconds % 3600) // 60)}分"


def print_progress_bar(current, total, prefix='', length=50):
    """打印进度条"""
    percent = current / total
    filled = int(length * percent)
    bar = '█' * filled + '░' * (length - filled)
    logging.info(f"{prefix} [{bar}] {current}/{total} ({percent*100:.1f}%)")


def log_downloaded_file(filepath, filename):
    """记录已下载的文件到txt"""
    try:
        log_path = os.path.join(ROOT_DIRECTORY, DOWNLOAD_LOG_FILE)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {filepath} | {filename}\n")
        
        logging.debug(f"📝 已记录到下载日志: {filename}")
    except Exception as e:
        logging.warning(f"⚠️  写入下载日志失败: {e}")


def setup_browser(download_path, use_profile=False, profile_path="", profile_name="Default"):
    """设置浏览器"""
    options = uc.ChromeOptions()
    
    # 基础设置
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--lang=zh-CN")
    
    # 下载设置 - 动态设置下载目录
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.notifications": 2,
        "intl.accept_languages": "zh-CN,zh;q=0.9",
    }
    options.add_experimental_option("prefs", prefs)
    
    # 使用真实浏览器 Profile
    if use_profile and profile_path:
        logging.info(f"使用真实浏览器 Profile: {profile_path}/{profile_name}")
        options.add_argument(f"--user-data-dir={profile_path}")
        options.add_argument(f"--profile-directory={profile_name}")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-software-rasterizer")
    
    try:
        logging.info("🚀 正在启动 Chrome 浏览器...")
        driver = uc.Chrome(
            options=options,
            version_main=None,
            headless=False,
            use_subprocess=False,
            log_level=3,
        )
        
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        driver.set_script_timeout(30)
        
        logging.info("✅ 浏览器启动成功")
        return driver
        
    except Exception as e:
        logging.error(f"❌ 启动浏览器失败: {e}")
        raise


def load_cookies_from_file(path: str) -> list:
    """从 JSON 文件读取 cookie"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            
            if isinstance(data, list):
                return data
            
            if isinstance(data, dict):
                cookies = []
                for name, value in data.items():
                    cookies.append({
                        "name": name,
                        "value": value,
                        "domain": ".weixin.qq.com",
                        "path": "/",
                        "secure": True,
                        "httpOnly": False
                    })
                return cookies
    except Exception as e:
        logging.warning(f"⚠️  读取 cookie_file 失败: {e}")
    return []


def add_cookies(driver, cookies_list, domain="doc.weixin.qq.com"):
    """注入 cookie"""
    if not cookies_list:
        return
    
    driver.get(f"https://{domain}")
    time.sleep(2)
    
    success_count = 0
    for cookie in cookies_list:
        try:
            if "name" not in cookie or "value" not in cookie:
                continue
            
            cookie_dict = {
                "name": cookie["name"],
                "value": cookie["value"],
                "domain": cookie.get("domain", domain),
                "path": cookie.get("path", "/"),
                "secure": cookie.get("secure", True),
            }
            
            if "httpOnly" in cookie:
                cookie_dict["httpOnly"] = cookie["httpOnly"]
            if "sameSite" in cookie:
                cookie_dict["sameSite"] = cookie["sameSite"]
            if "expiry" in cookie:
                cookie_dict["expiry"] = cookie["expiry"]
            
            driver.add_cookie(cookie_dict)
            success_count += 1
        except Exception as e:
            logging.debug(f"注入 cookie {cookie.get('name')} 失败: {e}")
    
    logging.info(f"✅ 成功注入 {success_count}/{len(cookies_list)} 个 cookie")
    time.sleep(2)


def wait_for_new_download(before_files, download_folder, timeout=DOWNLOAD_TIMEOUT):
    """等待下载完成 - 改进版本"""
    start = time.time()
    logging.info(f"⏳ 开始等待下载 (超时: {timeout}秒)")
    logging.info(f"   下载目录: {download_folder}")
    logging.info(f"   下载前文件数: {len(before_files)}")
    
    # 定期检查
    check_interval = 5  # 每5秒报告一次进度
    last_check_time = start
    
    while time.time() - start < timeout:
        current_time = time.time()
        
        # 获取当前文件列表
        try:
            now_files = {p.name for p in Path(download_folder).iterdir() if p.is_file()}
        except Exception as e:
            logging.warning(f"⚠️  读取目录失败: {e}")
            time.sleep(1)
            continue
        
        new_files = now_files - before_files
        
        # 定期输出进度信息
        if current_time - last_check_time >= check_interval:
            elapsed = int(current_time - start)
            logging.info(f"   [{elapsed}s] 当前文件数: {len(now_files)}, 新增: {len(new_files)}")
            if new_files:
                logging.info(f"   新文件列表: {list(new_files)}")
            last_check_time = current_time
        
        if new_files:
            # 排除临时文件
            valid_files = [f for f in new_files 
                          if not f.endswith('.crdownload') 
                          and not f.endswith('.tmp')
                          and not f.startswith('.')
                          and not f.startswith('~')]
            
            if not valid_files:
                logging.debug("   只有临时文件，继续等待...")
                time.sleep(1)
                continue
            
            # 选择最新的文件
            candidate = Path(download_folder) / sorted(valid_files)[0]
            logging.info(f"   🔍 检测到新文件: {candidate.name}")
            
            # 等待文件稳定
            stable_checks = 0
            stable_needed = 3  # 需要3次检查都稳定
            last_size = -1
            
            for check_num in range(10):  # 最多检查10次
                if not candidate.exists():
                    logging.warning(f"   ⚠️  文件消失: {candidate.name}")
                    break
                
                try:
                    current_size = candidate.stat().st_size
                    
                    if current_size == last_size and current_size > 0:
                        stable_checks += 1
                        logging.info(f"   ✓ 文件稳定检查: {stable_checks}/{stable_needed} ({current_size:,} bytes)")
                        
                        if stable_checks >= stable_needed:
                            file_size_mb = current_size / (1024 * 1024)
                            logging.info(f"✅ 文件下载完成: {candidate.name} ({file_size_mb:.2f} MB)")
                            return str(candidate)
                    else:
                        if last_size != -1:
                            logging.info(f"   ⏬ 文件大小变化: {last_size:,} → {current_size:,} bytes")
                        stable_checks = 0
                        last_size = current_size
                    
                    time.sleep(1)
                    
                except Exception as e:
                    logging.warning(f"   ⚠️  检查文件时出错: {e}")
                    time.sleep(1)
        
        time.sleep(1)
    
    # 超时后做最后检查
    logging.warning("⚠️  下载等待超时，做最后检查...")
    try:
        final_files = {p.name for p in Path(download_folder).iterdir() if p.is_file()}
        final_new = final_files - before_files
        if final_new:
            # 返回最新的非临时文件
            valid = [Path(download_folder) / f for f in final_new 
                    if not f.endswith(('.crdownload', '.tmp')) and not f.startswith('.')]
            if valid:
                latest = max(valid, key=lambda f: f.stat().st_mtime)
                logging.info(f"✅ 找到文件: {latest.name}")
                return str(latest)
    except Exception as e:
        logging.error(f"❌ 最后检查失败: {e}")
    
    return None


def update_download_directory(driver, new_download_path):
    """动态更新浏览器下载目录"""
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": new_download_path
        })
        logging.info(f"📁 下载目录已更新为: {new_download_path}")
    except Exception as e:
        logging.warning(f"⚠️  更新下载目录失败: {e}")


def save_debug(driver, prefix, current_dir=None):
    """保存调试信息"""
    debug_dir = os.path.join(current_dir if current_dir else "debug", "debug")
    os.makedirs(debug_dir, exist_ok=True)
    
    ts = int(time.time())
    html_path = os.path.join(debug_dir, f"{prefix}_{ts}.html")
    png_path = os.path.join(debug_dir, f"{prefix}_{ts}.png")
    
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(png_path)
        logging.info(f"💾 已保存调试文件：{html_path}, {png_path}")
    except Exception as e:
        logging.warning(f"⚠️  保存调试文件失败: {e}")


def check_file_exists(directory, filename, url):
    """检查文件是否已存在"""
    # 获取期望的文件扩展名
    ext = guess_ext_from_url(url)
    full_filename = f"{safe_filename(filename)}.{ext}"
    file_path = Path(directory) / full_filename
    
    if file_path.exists():
        logging.info(f"⏭️  文件已存在，跳过: {full_filename}")
        return True
    
    # 检查是否有带编号的版本
    i = 1
    while True:
        alt_filename = f"{safe_filename(filename)}({i}).{ext}"
        alt_path = Path(directory) / alt_filename
        if alt_path.exists():
            logging.info(f"⏭️  文件已存在，跳过: {alt_filename}")
            return True
        if not alt_path.exists():
            break
        i += 1
    
    return False


def click_export_and_download(driver, name, url, idx, total, download_dir, before_files):
    """点击导出并下载 - 改进版本"""
    
    url_l = url.lower()
    is_sheet = "sheet" in url_l
    is_doc = "doc" in url_l and "sheet" not in url_l
    
    doc_type = "表格" if is_sheet else "文档" if is_doc else "文件"
    
    try:
        # 1. 点击菜单
        logging.info(f"🔍 [{idx}/{total}] 查找菜单按钮...")
        menu = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable((By.ID, "main-menu-file"))
        )
        menu.click()
        logging.info(f"✅ 菜单按钮已点击")
        time.sleep(MENU_WAIT)
        
        # 2. 点击导出
        logging.info(f"🔍 查找导出按钮...")
        if is_sheet:
            export_xpaths = [
                "//li[contains(@class,'mainmenu-submenu-exportAs') and contains(normalize-space(.),'导出')]",
                "//li[contains(@class,'mainmenu-submenu') and contains(normalize-space(.),'导出')]",
            ]
        else:
            export_xpaths = [
                "//li[contains(@class,'mainmenu-submenu-export-as') and contains(normalize-space(.),'导出')]",
                "//li[contains(@class,'mainmenu-submenu') and contains(normalize-space(.),'导出')]",
            ]
        
        export_li = None
        for xpath_idx, xpath in enumerate(export_xpaths, 1):
            try:
                logging.debug(f"   尝试 XPath {xpath_idx}/{len(export_xpaths)}")
                export_li = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                logging.info(f"✅ 找到导出按钮（XPath {xpath_idx}）")
                break
            except TimeoutException:
                continue
        
        if not export_li:
            return None, "未找到导出按钮"
        
        export_li.click()
        logging.info(f"✅ 导出按钮已点击")
        time.sleep(CLICK_WAIT)
        
        # 3. 选择导出类型
        logging.info(f"🔍 查找导出类型选项（{doc_type}）...")
        if is_sheet:
            export_type_xpaths = [
                "//li[contains(@class,'mainmenu-item-export-local') and contains(normalize-space(.),'本地')]",
                "//li[contains(@class,'export-local') and contains(normalize-space(.),'本地')]",
            ]
        elif is_doc:
            export_type_xpaths = [
                "//li[contains(@class,'mainmenu-item-export-as-docx') and contains(normalize-space(.),'本地')]",
                "//li[contains(@class,'export-as-docx') and contains(normalize-space(.),'本地')]",
            ]
        else:
            export_type_xpaths = [
                "//*[contains(normalize-space(.),'本地') and (self::li or self::button)]",
            ]
        
        target = None
        for xpath_idx, xpath in enumerate(export_type_xpaths, 1):
            try:
                logging.debug(f"   尝试导出类型 XPath {xpath_idx}/{len(export_type_xpaths)}")
                target = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                logging.info(f"✅ 找到导出类型选项（XPath {xpath_idx}）")
                break
            except TimeoutException:
                continue
        
        if not target:
            return None, "未找到导出类型选项"
        
        # 点击前记录文件列表
        before_click_files = {p.name for p in Path(download_dir).iterdir() if p.is_file()}
        logging.info(f"📊 点击前文件数: {len(before_click_files)}")
        
        target.click()
        logging.info(f"✅ 导出类型已选择，开始下载...")
        
        # 点击后立即检查是否有弹窗
        time.sleep(2)
        
        # 检查是否有确认按钮或其他弹窗
        try:
            confirm_buttons = driver.find_elements(By.XPATH, 
                "//button[contains(normalize-space(.),'确定') or "
                "contains(normalize-space(.),'确认') or "
                "contains(normalize-space(.),'下载')]"
            )
            if confirm_buttons:
                logging.info(f"🔍 发现 {len(confirm_buttons)} 个确认按钮")
                for btn in confirm_buttons:
                    if btn.is_displayed():
                        btn.click()
                        logging.info(f"✅ 点击了确认按钮")
                        time.sleep(1)
                        break
        except Exception as e:
            logging.debug(f"检查确认按钮时出错: {e}")
        
        # 再次记录文件列表用于比较
        after_click_files = {p.name for p in Path(download_dir).iterdir() if p.is_file()}
        new_immediate = after_click_files - before_click_files
        if new_immediate:
            logging.info(f"⚡ 点击后立即出现新文件: {list(new_immediate)}")
        
        # 等待下载
        downloaded = wait_for_new_download(before_files, download_dir, timeout=DOWNLOAD_TIMEOUT)
        
        return downloaded, "成功" if downloaded else "下载超时"
        
    except TimeoutException as e:
        logging.warning(f"⚠️  等待页面元素超时: {e}")
        return None, "元素超时"
    except Exception as e:
        logging.warning(f"⚠️  自动点击导出失败: {e}")
        return None, f"点击失败: {str(e)}"


def process_directory(directory_path, driver, dir_idx, total_dirs):
    """处理单个目录 - 修复返回值问题"""
    json_file = os.path.join(directory_path, "data.json")
    
    if not os.path.exists(json_file):
        logging.warning(f"⚠️  目录 {directory_path} 中没有 data.json")
        return 0, 0, 0, "无data.json文件"
    
    dir_name = os.path.basename(directory_path)
    logging.info(f"\n{'='*80}")
    logging.info(f"📂 [{dir_idx}/{total_dirs}] 处理目录: {dir_name}")
    logging.info(f"   路径: {directory_path}")
    logging.info(f"{'='*80}")
    
    # 读取 JSON
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        logging.error(f"❌ 无法读取 JSON 文件 {json_file}: {e}")
        return 0, 0, 0, "JSON读取失败"
    
    infos = data.get("body", {}).get("file_list", [])
    if not infos:
        logging.warning("⚠️  JSON 中未找到 file_list，跳过此目录")
        return 0, 0, 0, "无file_list数据"
    
    logging.info(f"📊 本目录共有 {len(infos)} 个文档待处理")
    
    # 更新下载目录为当前目录
    download_dir = directory_path
    update_download_directory(driver, download_dir)
    
    # 主循环
    success_count = 0
    failed_count = 0
    skipped_count = 0
    failed_details = []
    
    start_time = time.time()
    
    for idx, info in enumerate(infos, start=1):
        name_raw = info.get("name", f"doc_{idx}")
        name = safe_filename(name_raw)
        url = info.get("doc_url")
        
        # 显示进度
        logging.info(f"\n{'─'*80}")
        print_progress_bar(idx - 1, len(infos), prefix=f'📈 目录进度')
        logging.info(f"📄 [{idx}/{len(infos)}] 正在处理: {name}")
        logging.info(f"   URL: {url[:80]}..." if len(url) > 80 else f"   URL: {url}")
        
        if not url:
            logging.warning(f"⚠️  未提供 doc_url，跳过")
            failed_count += 1
            failed_details.append((name, "无URL"))
            continue
        
        # 检查文件是否已存在
        if check_file_exists(download_dir, name, url):
            skipped_count += 1
            continue
        
        # 打开页面
        try:
            logging.info(f"🌐 打开页面...")
            driver.get(url)
            time.sleep(PAGE_STABLE_WAIT)
            logging.info(f"✅ 页面加载完成")
            
        except Exception as e:
            logging.warning(f"❌ 打开页面异常: {e}")
            save_debug(driver, f"{idx}_{name}_open_err", download_dir)
            failed_count += 1
            failed_details.append((name, "打开失败"))
            continue
        
        # 在点击下载前记录文件列表
        before_files = {p.name for p in Path(download_dir).iterdir() if p.is_file()}
        
        # 点击导出并下载
        downloaded, status = click_export_and_download(
            driver, name, url, idx, len(infos), download_dir, before_files
        )
        
        if not downloaded:
            logging.warning(f"❌ 下载失败: {status}")
            save_debug(driver, f"{idx}_{name}_{status}", download_dir)
            failed_count += 1
            failed_details.append((name, status))
            continue
        
        # 重命名文件
        src = Path(downloaded)
        ext = src.suffix if src.suffix else f".{guess_ext_from_url(url)}"
        
        # 目标文件名（不带后缀）
        dest = Path(download_dir) / f"{name}{ext}"
        
        # 如果下载的文件已经是正确的名字，就不需要重命名
        if src == dest:
            file_size_mb = dest.stat().st_size / (1024 * 1024)
            logging.info(f"✅ 下载完成: {dest.name} ({file_size_mb:.2f} MB)")
            
            # 记录到下载日志
            rel_path = os.path.relpath(str(dest), ROOT_DIRECTORY)
            log_downloaded_file(rel_path, dest.name)
            
            success_count += 1
        else:
            # 需要重命名，检查目标文件是否存在
            if dest.exists():
                i = 1
                while True:
                    alt = Path(download_dir) / f"{name}({i}){ext}"
                    if not alt.exists():
                        dest = alt
                        break
                    i += 1
            
            try:
                shutil.move(str(src), str(dest))
                file_size_mb = dest.stat().st_size / (1024 * 1024)
                logging.info(f"✅ 下载完成: {dest.name} ({file_size_mb:.2f} MB)")
                
                # 记录到下载日志
                rel_path = os.path.relpath(str(dest), ROOT_DIRECTORY)
                log_downloaded_file(rel_path, dest.name)
                
                success_count += 1
            except Exception as e:
                logging.warning(f"❌ 重命名失败: {e}")
                failed_count += 1
                failed_details.append((name, "重命名失败"))
                continue
        
        # 简单的间隔
        if idx < len(infos):
            time.sleep(2)
    
    # 计算耗时
    elapsed_time = time.time() - start_time
    
    # 输出当前目录的处理结果
    logging.info(f"\n{'='*80}")
    logging.info(f"📊 目录 [{dir_name}] 处理完成")
    logging.info(f"{'='*80}")
    logging.info(f"✅ 成功: {success_count}")
    logging.info(f"⏭️  跳过: {skipped_count}")
    logging.info(f"❌ 失败: {failed_count}")
    logging.info(f"⏱️  耗时: {format_time(elapsed_time)}")
    
    if failed_details:
        logging.info(f"\n失败详情:")
        for name, reason in failed_details:
            logging.info(f"  ❌ {name}: {reason}")
    
    return success_count, failed_count, skipped_count, "完成"


def main():
    start_time = time.time()
    
    logging.info("=" * 80)
    logging.info("🚀 企业微信文档批量下载工具（目录遍历版）")
    logging.info("=" * 80)
    logging.info(f"📁 根目录: {ROOT_DIRECTORY}")
    logging.info(f"⏰ 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("=" * 80)
    
    if not os.path.exists(ROOT_DIRECTORY):
        logging.error(f"❌ 根目录不存在: {ROOT_DIRECTORY}")
        return
    
    # 查找所有包含data.json的目录
    directories_with_data = []
    logging.info("🔍 正在扫描目录...")
    for root, dirs, files in os.walk(ROOT_DIRECTORY):
        if "data.json" in files:
            directories_with_data.append(root)
            rel_path = os.path.relpath(root, ROOT_DIRECTORY)
            logging.info(f"   ✓ 找到: {rel_path}")
    
    if not directories_with_data:
        logging.warning("⚠️  未找到包含data.json的目录")
        return
    
    # 按路径深度排序，确保先处理父目录
    directories_with_data.sort(key=lambda x: x.count(os.sep))
    
    logging.info(f"\n✅ 找到 {len(directories_with_data)} 个包含data.json的目录:")
    for i, directory in enumerate(directories_with_data, 1):
        rel_path = os.path.relpath(directory, ROOT_DIRECTORY)
        logging.info(f"   {i}. {rel_path}")
    
    # 启动浏览器
    logging.info(f"\n{'='*80}")
    logging.info("🌐 正在启动浏览器...")
    logging.info("=" * 80)
    driver = setup_browser(
        ROOT_DIRECTORY,  # 初始下载目录设为根目录
        use_profile=USE_REAL_PROFILE,
        profile_path=CHROME_PROFILE_PATH,
        profile_name=PROFILE_NAME
    )
    
    # 如果不使用 profile，则需要加载 cookie
    if not USE_REAL_PROFILE:
        cookies_list = load_cookies_from_file(cookie_file) if os.path.exists(cookie_file) else []
        if cookies_list:
            logging.info(f"🔐 正在注入 {len(cookies_list)} 个 cookie...")
            add_cookies(driver, cookies_list)
        else:
            logging.warning("⚠️  未提供 cookie 且未使用 profile，可能需要手动登录")
            logging.info("👉 请在打开的浏览器中登录，然后按回车继续...")
            input()
    else:
        logging.info("✅ 使用真实浏览器 Profile，已自动登录")
    
    # 统计信息
    total_success = 0
    total_failed = 0
    total_skipped = 0
    directory_results = []
    
    # 处理每个目录
    for idx, directory in enumerate(directories_with_data, 1):
        try:
            # 调用 process_directory，确保总是返回4个值
            result = process_directory(directory, driver, idx, len(directories_with_data))
            
            # 检查返回值
            if result is None or len(result) != 4:
                logging.error(f"❌ process_directory 返回值异常: {result}")
                success, failed, skipped, status = 0, 0, 0, "返回值异常"
            else:
                success, failed, skipped, status = result
            
            total_success += success
            total_failed += failed
            total_skipped += skipped
            
            directory_results.append({
                'directory': os.path.relpath(directory, ROOT_DIRECTORY),
                'success': success,
                'failed': failed,
                'skipped': skipped,
                'status': status
            })
            
            # 显示总体进度
            logging.info(f"\n{'='*80}")
            print_progress_bar(idx, len(directories_with_data), prefix='🎯 总体进度')
            logging.info(f"📊 累计统计: 成功 {total_success} | 跳过 {total_skipped} | 失败 {total_failed}")
            logging.info(f"{'='*80}")
            
            # 目录间休息
            if idx < len(directories_with_data):
                logging.info(f"\n⏸️  休息 3 秒后处理下一个目录...")
                time.sleep(3)
                
        except Exception as e:
            logging.error(f"❌ 处理目录 {directory} 时发生异常: {e}")
            import traceback
            logging.error(traceback.format_exc())
            
            directory_results.append({
                'directory': os.path.relpath(directory, ROOT_DIRECTORY),
                'success': 0,
                'failed': 0,
                'skipped': 0,
                'status': f"异常: {str(e)}"
            })
    
    driver.quit()
    logging.info("\n🔒 浏览器已关闭")
    
    # 计算总耗时
    total_time = time.time() - start_time
    
    # 最终统计
    logging.info(f"\n{'='*80}")
    logging.info("🎉 全部处理完成！")
    logging.info("=" * 80)
    logging.info(f"⏰ 结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info(f"⏱️  总耗时: {format_time(total_time)}")
    logging.info(f"\n📊 最终统计:")
    logging.info(f"   ✅ 总成功: {total_success}")
    logging.info(f"   ⏭️  总跳过: {total_skipped}")
    logging.info(f"   ❌ 总失败: {total_failed}")
    logging.info(f"   📁 处理目录: {len(directories_with_data)}")
    
    if total_success + total_failed > 0:
        success_rate = (total_success / (total_success + total_failed)) * 100
        logging.info(f"   📈 成功率: {success_rate:.1f}%")
    
    # 详细结果表
    logging.info(f"\n{'='*80}")
    logging.info("📋 各目录处理结果:")
    logging.info("=" * 80)
    logging.info(f"{'目录':<40} {'成功':>8} {'跳过':>8} {'失败':>8} {'状态':<15}")
    logging.info("-" * 80)
    
    for result in directory_results:
        dir_name = result['directory']
        if len(dir_name) > 38:
            dir_name = "..." + dir_name[-35:]
        
        success = result['success']
        skipped = result['skipped']
        failed = result['failed']
        status = result['status']
        
        # 根据状态选择图标
        if status == "完成" and failed == 0:
            icon = "✅"
        elif status == "完成" and failed > 0:
            icon = "⚠️"
        else:
            icon = "❌"
        
        logging.info(f"{dir_name:<40} {success:>8} {skipped:>8} {failed:>8} {icon} {status:<15}")
    
    logging.info("=" * 80)
    
    # 保存结果到JSON文件
    try:
        result_file = os.path.join(ROOT_DIRECTORY, f"download_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(result_file, "w", encoding="utf-8") as f:
            json.dump({
                "start_time": datetime.fromtimestamp(start_time).strftime('%Y-%m-%d %H:%M:%S'),
                "end_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "total_time_seconds": int(total_time),
                "total_success": total_success,
                "total_failed": total_failed,
                "total_skipped": total_skipped,
                "total_directories": len(directories_with_data),
                "directory_results": directory_results
            }, f, ensure_ascii=False, indent=2)
        logging.info(f"\n💾 结果已保存到: {result_file}")
    except Exception as e:
        logging.warning(f"⚠️  保存结果文件失败: {e}")
    
    # 显示下载日志文件位置
    log_file_path = os.path.join(ROOT_DIRECTORY, DOWNLOAD_LOG_FILE)
    if os.path.exists(log_file_path):
        logging.info(f"📝 下载文件日志: {log_file_path}")
    
    logging.info("\n✨ 程序执行完毕！")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info("\n\n⚠️  用户中断程序")
    except Exception as e:
        logging.error(f"\n\n❌ 程序异常退出: {e}")
        import traceback
        logging.error(traceback.format_exc())
    finally:
        logging.info("\n👋 再见！")
