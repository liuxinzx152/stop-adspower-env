"""
AdsPower 环境通用自动化脚本
功能：QuickQ连接 → API启动 → 连接CDP → 点击未读 → 最大化窗口 → 截图
支持多环境：#17、#18 等，通过参数传入环境ID和CDP端口

使用方法:
    python adspower_env11_full_auto_copy.py --env-id k19g237w --port 49907
    python adspower_env11_full_auto_copy.py -e k19g237w -p 49907
    python adspower_env11_full_auto_copy.py -e k19g237w              # 通过API启动并自动获取端口
"""

import requests
import time
import subprocess
import sys
import ctypes
import re
import argparse
import threading
from ctypes import wintypes
from playwright.sync_api import sync_playwright

# ============ 全局硬性超时机制 ============
SCRIPT_START_TIME = time.time()
MAX_SCRIPT_RUNTIME = 240  # 240秒硬性超时

def check_global_timeout():
    """检查脚本是否运行超过最大时间"""
    elapsed = time.time() - SCRIPT_START_TIME
    if elapsed > MAX_SCRIPT_RUNTIME:
        print(f"\n[Error] 脚本运行超过 {MAX_SCRIPT_RUNTIME} 秒，强制退出")
        sys.exit(1)
    return elapsed

# ============ 命令行参数解析 ============
parser = argparse.ArgumentParser(description='AdsPower 环境通用自动化脚本')
parser.add_argument('-e', '--env-id', type=str, required=True,
                    help='AdsPower 环境ID (如: k19g237w, kqqvpqj)')
parser.add_argument('-p', '--port', type=int, default=None,
                    help='CDP 调试端口 (如: 49907)。如果不指定，将通过API启动环境后自动获取')
parser.add_argument('--skip-quickq', action='store_true',
                    help='跳过 QuickQ 连接步骤')
parser.add_argument('--skip-adspower-launch', action='store_true',
                    help='跳过 AdsPower 客户端启动')
args = parser.parse_args()

# ============ 配置 ============
ADSPOWER_PATH = r"C:\Program Files\AdsPower Global\AdsPower Global.exe"
API_BASE = "http://local.adspower.net:50325"
ENV_ID = args.env_id  # 从命令行参数获取环境ID
debug_port = args.port  # 从命令行参数获取端口（可能为None）

# ============ Step 1: QuickQ 连接完成 ============
print("=" * 50)
print("Step 1: QuickQ 连接完成")
print("=" * 50)

def find_window_by_title_contains(title_contains):
    """根据标题关键词查找窗口句柄"""
    found = []
    def enum_callback(hwnd, extra):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                window_title = buffer.value
                if title_contains.lower() in window_title.lower():
                    found.append(hwnd)
        return True
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    callback = EnumWindowsProc(enum_callback)
    ctypes.windll.user32.EnumWindows(callback, 0)
    return found[0] if found else None

def get_window_rect(hwnd):
    """获取窗口位置和大小"""
    rect = wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top

def click_at(x, y):
    """在屏幕指定位置点击"""
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
    time.sleep(0.1)
    ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放

def check_internet_connection(timeout=120):
    """检测是否能访问外网，确认 QuickQ 已连接"""
    import socket
    start_time = time.time()
    check_urls = [
        ("8.8.8.8", 53),      # Google DNS
        ("1.1.1.1", 53),      # Cloudflare DNS
        ("223.5.5.5", 53),    # 阿里 DNS
    ]
    
    print("[Check] 正在检测网络连接...")
    while time.time() - start_time < timeout:
        for host, port in check_urls:
            try:
                socket.setdefaulttimeout(3)
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                result = sock.connect_ex((host, port))
                sock.close()
                if result == 0:
                    print(f"[OK] 网络已连接 (通过 {host}:{port})")
                    return True
            except:
                pass
        time.sleep(2)
    return False

def get_public_ip():
    """获取当前公网 IP，用于确认 VPN 是否生效"""
    try:
        response = requests.get("https://api.ipify.org?format=json", timeout=10)
        if response.status_code == 200:
            return response.json().get("ip")
    except:
        pass
    try:
        response = requests.get("https://httpbin.org/ip", timeout=10)
        if response.status_code == 200:
            return response.json().get("origin", "").split(",")[0].strip()
    except:
        pass
    return None

def check_vpn_connected(timeout=120):
    """检测 VPN 是否真正连接（通过外网 IP 判断）"""
    print("[Check] 检测 VPN 连接状态...")
    start_time = time.time()
    
    # 先获取初始 IP（如果可用）
    initial_ip = None
    try:
        initial_ip = get_public_ip()
        print(f"[Info] 初始 IP: {initial_ip}")
    except:
        pass
    
    while time.time() - start_time < timeout:
        try:
            ip = get_public_ip()
            if ip:
                # 简单判断：如果能获取到 IP 且不是常见的中国 IP 段，认为 VPN 已连接
                # 更可靠的方式是检查 IP 是否变化或与预期地区匹配
                print(f"[OK] 当前公网 IP: {ip}")
                return ip
        except Exception as e:
            print(f"[Wait] 获取 IP 失败: {e}")
        time.sleep(3)
    
    return None

if not args.skip_quickq:
    # 查找 QuickQ 窗口
    print("[Action] 查找 QuickQ 窗口...")
    quickq_hwnd = None
    for _ in range(10):  # 尝试10次
        quickq_hwnd = find_window_by_title_contains("QuickQ")
        if quickq_hwnd:
            break
        time.sleep(1)

    if not quickq_hwnd:
        print("[Error] 未找到 QuickQ 窗口，请确保 QuickQ 已运行")
        sys.exit(1)

    print(f"[OK] 找到 QuickQ 窗口")

    # 将 QuickQ 窗口带到前台
    ctypes.windll.user32.SetForegroundWindow(quickq_hwnd)
    time.sleep(0.5)

    # 获取窗口位置，计算中央电源按钮坐标（更精确的中心位置）
    x, y, width, height = get_window_rect(quickq_hwnd)
    center_x = x + width // 2
    center_y = y + height // 2 - 20  # 电源按钮在窗口正中央

    print(f"[Action] 点击 QuickQ 电源按钮 ({center_x}, {center_y})")
    click_at(center_x, center_y)

    # 等待并检测 QuickQ 连接状态
    print("[Wait] 等待 QuickQ 连接...")
    print("[Info] 给 QuickQ 15秒时间建立连接...")
    time.sleep(15)  # 增加等待时间到15秒

    max_wait = 120  # 增加最大等待时间到120秒
    connected = False
    public_ip = None
    stable_count = 0
    last_ip = None

    for i in range(max_wait):
        # 检查全局超时
        check_global_timeout()
        
        # 检测 VPN 是否真正连接（通过外网 IP）
        current_ip = get_public_ip()
        if current_ip:
            # 检查 IP 是否稳定（连续2次相同）
            if current_ip == last_ip:
                stable_count += 1
                if stable_count >= 2:
                    public_ip = current_ip
                    connected = True
                    print(f"[OK] QuickQ 连接稳定，公网 IP: {public_ip}")
                    break
            else:
                stable_count = 0
                last_ip = current_ip
                print(f"[Wait] QuickQ 连接中，当前 IP: {current_ip} ({i+1}/{max_wait})")
        else:
            print(f"[Wait] QuickQ 连接中，等待获取 IP... ({i+1}/{max_wait})")
        time.sleep(1)

    if not connected:
        print("[Error] 120秒内未能建立稳定的 VPN 连接")
        sys.exit(1)

    print(f"[OK] 电源按钮应显示绿色指示灯")
    print(f"[OK] 公网 IP: {public_ip}")

else:
    print("[Skip] 跳过 QuickQ 连接步骤")
    public_ip = None

# ============ Step 2: AdsPower 启动成功 ============
print("\n" + "=" * 50)
print("Step 2: AdsPower 启动成功")
print("=" * 50)

if not args.skip_adspower_launch:
    try:
        subprocess.Popen([ADSPOWER_PATH], shell=True)
        print("[OK] AdsPower 已启动")
        time.sleep(3)  # 等待 AdsPower 启动
    except Exception as e:
        print(f"[Warning] 启动 AdsPower 可能失败: {e}")
else:
    print("[Skip] 跳过 AdsPower 启动")

# ============ Step 2.1: 最大化 AdsPower 窗口 ============
print("\n" + "=" * 50)
print("Step 2.1: 最大化 AdsPower 窗口")
print("=" * 50)

def maximize_adspower_window():
    """最大化 AdsPower 主窗口 - 使用 win32gui 更可靠"""
    try:
        import win32gui
        import win32con
        
        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                # 检测 AdsPower 主窗口（标题包含 AdsPower 但不包含 SunBrowser/WhatsApp）
                if "AdsPower" in window_title and "SunBrowser" not in window_title and "WhatsApp" not in window_title:
                    extra.append((hwnd, window_title))
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            hwnd, title = windows[0]
            # 先恢复窗口，再激活，最后最大化
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            print(f"[OK] 已最大化: {title}")
            return True
        else:
            print("[Warning] 未找到 AdsPower 主窗口")
            return False
    except ImportError:
        print("[Warning] 未安装 pywin32，回退到 ctypes 方法")
        # 回退到 ctypes 方法
        def enum_windows_callback(hwnd, extra):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                    window_title = buffer.value
                    if "AdsPower" in window_title and "SunBrowser" not in window_title and "WhatsApp" not in window_title:
                        ctypes.windll.user32.ShowWindow(hwnd, win32con.SW_RESTORE)
                        ctypes.windll.user32.SetForegroundWindow(hwnd)
                        time.sleep(0.3)
                        ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE = 3
                        print(f"[OK] 已最大化: {window_title}")
                        return False
            return True
        
        EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        callback = EnumWindowsProc(enum_windows_callback)
        ctypes.windll.user32.EnumWindows(callback, 0)
        return True

time.sleep(1)
maximize_adspower_window()
time.sleep(1)

# ============ Step 2.2: 检测并关闭广告 ============
print("\n" + "=" * 50)
print("Step 2.2: 检测广告弹窗，点击右上角 X 关闭")
print("=" * 50)

# 广告弹窗关闭按钮坐标（右上角 X）
AD_CLOSE_X = 1234
AD_CLOSE_Y = 158

print(f"[Action] 点击广告关闭按钮: ({AD_CLOSE_X}, {AD_CLOSE_Y})")
ctypes.windll.user32.SetCursorPos(AD_CLOSE_X, AD_CLOSE_Y)
ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
time.sleep(0.1)
ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放
print("[OK] 广告关闭点击完成")
time.sleep(1)

# ============ Step 2.3: 点击环境管理 ============
print("\n" + "=" * 50)
print("Step 2.3: 点击'环境管理'按钮 (85, 210)")
print("=" * 50)

x, y = 85, 210
print(f"[Action] 点击坐标: ({x}, {y})")
ctypes.windll.user32.SetCursorPos(x, y)
ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
time.sleep(0.1)
ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放
print("[OK] 环境管理点击完成")

# 等待 API 服务就绪（最多2分钟）
print("[Check] 等待 AdsPower API 服务就绪（最多2分钟）...")
start_time = time.time()
api_ready = False
while time.time() - start_time < 120:  # 2分钟超时
    # 检查全局超时
    check_global_timeout()
    
    try:
        response = requests.get(f"{API_BASE}/status", timeout=5)
        if response.status_code == 200:
            print("[OK] API 服务已就绪")
            api_ready = True
            break
    except:
        pass
    elapsed = int(time.time() - start_time)
    if elapsed % 5 == 0:  # 每5秒输出一次
        print(f"[Wait] API 服务启动中... ({elapsed}/120)")
    time.sleep(1)

if not api_ready:
    print("[Error] 2分钟内 API 服务未就绪")
    sys.exit(1)

# ============ Step 3: 环境启动成功 ============
print("\n" + "=" * 50)
print(f"Step 3: 环境 {ENV_ID} 启动成功")
print("=" * 50)

# 如果通过命令行传入了端口，跳过 API 启动流程
if args.port:
    print(f"[Info] 使用指定的 CDP 端口: {debug_port}")
    print(f"[Info] 跳过 API 启动流程（假设环境已在运行）")
else:
    # 等待 API 服务就绪
    def wait_for_api_ready(timeout=60):
        """等待 AdsPower API 服务就绪"""
        print("[Check] 等待 AdsPower API 服务就绪...")
        start_time = time.time()
        while time.time() - start_time < timeout:
            # 检查全局超时
            check_global_timeout()
            
            try:
                response = requests.get(f"{API_BASE}/status", timeout=5)
                if response.status_code == 200:
                    print("[OK] API 服务已就绪")
                    return True
            except:
                pass
            print(f"[Wait] API 服务启动中... ({int(time.time() - start_time)}/{timeout})")
            time.sleep(2)
        return False

    if not wait_for_api_ready(timeout=60):
        print("[Error] API 服务未就绪")
        sys.exit(1)

    url = f"{API_BASE}/api/v1/browser/start?user_id={ENV_ID}&open_tabs=[]&headless=false"
    print(f"[Request] {url}")

    # 重试机制
    max_retries = 3
    for attempt in range(max_retries):
        try:
            print(f"[Attempt {attempt + 1}/{max_retries}] 正在启动环境...")
            response = requests.get(url, timeout=60)
            result = response.json()
            print(f"[Response] {result}")
            
            if result.get("code") == 0:
                debug_port = result["data"]["debug_port"]
                ws_url = result["data"]["ws"]["puppeteer"]
                print(f"[OK] 环境启动成功")
                print(f"    Debug Port: {debug_port}")
                print(f"    WebSocket: {ws_url}")
                
                # 保存端口到文件，供其他脚本使用
                with open(f"adspower_port_{ENV_ID}.txt", "w", encoding="utf-8") as f:
                    f.write(debug_port)
                print(f"[Save] Debug Port 已保存到 adspower_port_{ENV_ID}.txt")
                
                break
            else:
                print(f"[Error] 启动失败: {result.get('msg')}")
                if attempt < max_retries - 1:
                    print(f"[Wait] 5秒后重试...")
                    time.sleep(5)
                else:
                    sys.exit(1)
        except Exception as e:
            print(f"[Error] API 调用失败: {e}")
            if attempt < max_retries - 1:
                print(f"[Wait] 5秒后重试...")
                time.sleep(5)
            else:
                sys.exit(1)

# ============ Step 4: CDP 连接成功，找到并激活 WhatsApp 页面 ============
print("\n" + "=" * 50)
print("Step 4: CDP 连接成功，找到并激活 WhatsApp 页面")
print("=" * 50)

p = sync_playwright().start()
browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{debug_port}")
context = browser.contexts[0] if browser.contexts else browser.new_context()

print(f"[OK] 已连接到 CDP (端口 {debug_port})")
print(f"    Contexts: {len(browser.contexts)}")
print(f"    Pages: {len(context.pages)}")

# 找到 WhatsApp 页面
whatsapp_page = None
for page in context.pages:
    if "web.whatsapp.com" in page.url:
        whatsapp_page = page
        break

if not whatsapp_page:
    print("[Info] 未找到 WhatsApp 页面，正在打开新页面...")
    whatsapp_page = context.new_page()
    whatsapp_page.goto("https://web.whatsapp.com")
    print("[OK] 已打开 WhatsApp Web")
    print("[Wait] 等待 10 秒让页面加载...")
    time.sleep(10)
    
if not whatsapp_page:
    print("[Error] 无法创建或找到页面")
    browser.close()
    p.stop()
    sys.exit(1)

print(f"[OK] 页面已就绪")
print(f"    URL: {whatsapp_page.url}")
try:
    print(f"    Title: {whatsapp_page.title()}")
except:
    print(f"    Title: (无法获取)")

# 激活页面
whatsapp_page.bring_to_front()
print("[OK] WhatsApp 页面已激活 (bring_to_front)")
time.sleep(1)

# ============ Step 5: 等待浏览器完全启动 ============
print("\n" + "=" * 50)
print("Step 5: 等待浏览器完全启动")
print("=" * 50)
print("[Wait] 等待 5 秒让浏览器完全启动...")
time.sleep(5)

# ============ Step 6: 最大化浏览器窗口 ============
print("\n" + "=" * 50)
print("Step 6: 最大化浏览器窗口")
print("=" * 50)

SW_MAXIMIZE = 3

def find_window_by_title(title_substring):
    found_hwnd = []
    def enum_windows_callback(hwnd, extra):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                window_title = buffer.value
                if title_substring.lower() in window_title.lower():
                    found_hwnd.append((hwnd, window_title))
        return True
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    callback = EnumWindowsProc(enum_windows_callback)
    ctypes.windll.user32.EnumWindows(callback, 0)
    return found_hwnd

# 查找并最大化浏览器窗口（不限于 WhatsApp）
browser_windows = find_window_by_title("SunBrowser") or find_window_by_title("Chrome") or find_window_by_title("WhatsApp")
if browser_windows:
    for hwnd, title in browser_windows:
        ctypes.windll.user32.ShowWindow(hwnd, SW_MAXIMIZE)
        print(f"[OK] 已最大化窗口: {title}")
else:
    print("[Warning] 未找到浏览器窗口")

# 同时设置视口大小
try:
    whatsapp_page.set_viewport_size({"width": 1920, "height": 1080})
    print("[OK] 视口已设置为 1920x1080")
except:
    pass

time.sleep(1)

# ============ Step 7: 点击"未读"按钮（30秒超时） ============
print("\n" + "=" * 50)
print("Step 7: 点击'未读'按钮（30秒超时）")
print("=" * 50)

unread_clicked = False
start_time = time.time()
timeout = 30  # 30秒超时

while time.time() - start_time < timeout and not unread_clicked:
    try:
        # 策略1: 通过 button 标签遍历查找（extract_chats_run.py 用的方式）
        buttons = whatsapp_page.query_selector_all('button')
        for btn in buttons:
            try:
                if not btn.is_visible():
                    continue
                text = btn.inner_text() or ''
                if re.search(r'(未读|Unread)', text, re.IGNORECASE):
                    btn.click()
                    print("[OK] 已点击'未读'按钮（通过 button 标签）")
                    unread_clicked = True
                    break
            except:
                continue
        
        if unread_clicked:
            break
        
        # 策略2: 通过 :has-text 查找（备用）
        unread_button = whatsapp_page.query_selector('button:has-text("未读")')
        if unread_button:
            unread_button.click()
            print("[OK] 已点击'未读'按钮（通过 :has-text）")
            unread_clicked = True
            break
        
        # 策略3: 通过 role="button" 和文本查找
        buttons = whatsapp_page.query_selector_all('[role="button"]')
        for btn in buttons:
            try:
                text = btn.inner_text() or ''
                if '未读' in text:
                    btn.click()
                    print("[OK] 已点击'未读'按钮（通过 role=button）")
                    unread_clicked = True
                    break
            except:
                continue
        
        if unread_clicked:
            break
        
        if not unread_clicked:
            print("[Wait] 未找到未读按钮，1秒后重试...")
            time.sleep(1)
    except Exception as e:
        print(f"[Warning] 点击未读按钮出错: {e}")
        time.sleep(1)

if not unread_clicked:
    print("[Warning] 30秒内未能点击未读按钮，进入后备方案")

# ============ Step 8: 截图判断状态 ============
print("\n" + "=" * 50)
print("Step 8: 截图判断状态")
print("=" * 50)

try:
    whatsapp_page.screenshot(path="whatsapp_status_check.png")
    print("[OK] 状态截图已保存: whatsapp_status_check.png")
except Exception as e:
    print(f"[Warning] 截图失败: {e}")

time.sleep(2)

# ============ Step 9: 后备坐标点击（如果Step 7失败时执行） ============
print("\n" + "=" * 50)
print("Step 9: 后备坐标点击（如果 Step 7 失败时执行）")
print("=" * 50)

if not unread_clicked:
    try:
        print("[Fallback] 尝试使用元素动态位置点击未读按钮...")
        
        # 方法1: 通过 aria-label 查找
        unread_btn = whatsapp_page.query_selector('[aria-label*="未读"], [aria-label*="Unread"]')
        
        # 方法2: 如果找不到，通过文本匹配查找
        if not unread_btn:
            buttons = whatsapp_page.query_selector_all('button')
            for btn in buttons:
                try:
                    text = btn.inner_text() or ''
                    if '未读' in text or 'Unread' in text:
                        unread_btn = btn
                        break
                except:
                    continue
        
        # 后备方案1: 找父级 BUTTON 并点击
        if unread_btn:
            try:
                # 尝试找父级 BUTTON
                parent_btn = unread_btn.evaluate_handle('''el => {
                    let parent = el.parentElement;
                    while (parent) {
                        if (parent.tagName === 'BUTTON') return parent;
                        parent = parent.parentElement;
                    }
                    return el;  // 如果没找到，返回自己
                }''')
                
                if parent_btn:
                    box = parent_btn.bounding_box()
                    if box:
                        click_x = box['x'] + box['width'] / 2
                        click_y = box['y'] + box['height'] / 2
                        print(f"  [Info] 父级按钮位置: ({click_x:.0f}, {click_y:.0f})")
                        parent_btn.click()
                        print("  [OK] 第1次点击完成")
                        time.sleep(0.3)
                        parent_btn.click()
                        print("  [OK] 第2次点击完成")
                        time.sleep(2)
                    else:
                        raise Exception("无法获取按钮位置")
                else:
                    raise Exception("未找到父级按钮")
            except Exception as e1:
                print(f"  [Warn] 父级按钮点击失败: {e1}")
                # 后备方案2: 终极后备 - 使用 ctypes 屏幕坐标点击
                print("  [Fallback2] 使用 ctypes 屏幕坐标 (195, 250) 点击")
                
                # 获取浏览器窗口位置
                def find_whatsapp_window():
                    found = []
                    def enum_callback(hwnd, extra):
                        if ctypes.windll.user32.IsWindowVisible(hwnd):
                            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                            if length > 0:
                                buffer = ctypes.create_unicode_buffer(length + 1)
                                ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                                if "WhatsApp" in buffer.value and "SunBrowser" in buffer.value:
                                    rect = wintypes.RECT()
                                    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                                    found.append((rect.left, rect.top))
                        return True
                    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                    callback = EnumWindowsProc(enum_callback)
                    ctypes.windll.user32.EnumWindows(callback, 0)
                    return found[0] if found else (0, 0)
                
                win_x, win_y = find_whatsapp_window()
                click_x = win_x + 195
                click_y = win_y + 250
                
                print(f"    [Info] 窗口位置: ({win_x}, {win_y})")
                print(f"    [Info] 屏幕点击坐标: ({click_x}, {click_y})")
                
                # 移动鼠标并点击两次
                ctypes.windll.user32.SetCursorPos(click_x, click_y)
                time.sleep(0.3)
                ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
                time.sleep(0.1)
                ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放
                print("    [OK] 第1次点击完成")
                
                time.sleep(0.3)
                
                ctypes.windll.user32.SetCursorPos(click_x, click_y)
                time.sleep(0.3)
                ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
                time.sleep(0.1)
                ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放
                print("    [OK] 第2次点击完成")
                time.sleep(2)
        else:
            print("  [Warn] 未找到未读按钮元素，直接使用终极后备方案")
            # 后备方案2: 终极后备 - 使用 ctypes 屏幕坐标点击
            print("  [Fallback2] 使用 ctypes 屏幕坐标 (195, 250) 点击")
            
            def find_whatsapp_window():
                found = []
                def enum_callback(hwnd, extra):
                    if ctypes.windll.user32.IsWindowVisible(hwnd):
                        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                        if length > 0:
                            buffer = ctypes.create_unicode_buffer(length + 1)
                            ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                            if "WhatsApp" in buffer.value and "SunBrowser" in buffer.value:
                                rect = wintypes.RECT()
                                ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                                found.append((rect.left, rect.top))
                    return True
                EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                callback = EnumWindowsProc(enum_callback)
                ctypes.windll.user32.EnumWindows(callback, 0)
                return found[0] if found else (0, 0)
            
            win_x, win_y = find_whatsapp_window()
            click_x = win_x + 195
            click_y = win_y + 250
            
            print(f"    [Info] 窗口位置: ({win_x}, {win_y})")
            print(f"    [Info] 屏幕点击坐标: ({click_x}, {click_y})")
            
            ctypes.windll.user32.SetCursorPos(click_x, click_y)
            time.sleep(0.3)
            ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
            time.sleep(0.1)
            ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
            print("    [OK] 第1次点击完成")
            
            time.sleep(0.3)
            
            ctypes.windll.user32.SetCursorPos(click_x, click_y)
            time.sleep(0.3)
            ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
            time.sleep(0.1)
            ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
            print("    [OK] 第2次点击完成")
            time.sleep(2)
    except Exception as e:
        print(f"[Warning] 所有后备点击方案都失败: {e}")
else:
    print("[Skip] Step 7 已成功，跳过后备点击")

# ============ Step 10: 最终截图保存 ============
print("\n" + "=" * 50)
print("Step 10: 最终截图保存")
print("=" * 50)

try:
    whatsapp_page.screenshot(path="whatsapp_final.png")
    print("[OK] 最终截图已保存: whatsapp_final.png")
except Exception as e:
    print(f"[Warning] 截图失败: {e}")

# ============ 完成 ============
print("\n" + "=" * 50)
print("所有步骤已完成!")
print("=" * 50)
print(f"环境 ID: {ENV_ID}")
print(f"调试端口: {debug_port}")
print(f"CDP地址: http://127.0.0.1:{debug_port}")

browser.close()
p.stop()
