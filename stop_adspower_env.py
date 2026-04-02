# -*- coding: utf-8 -*-
"""
关闭 AdsPower 环境 (通用版本)
用法: python stop_adspower_env.py [环境ID]
示例: 
  python stop_adspower_env.py k19g237w    # 关闭环境 #18
  python stop_adspower_env.py k19g22i7    # 关闭环境 #17
  python stop_adspower_env.py             # 默认关闭环境 #11
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import requests
import subprocess
import os
import time
import argparse
import ctypes
from ctypes import wintypes

# 解析命令行参数
parser = argparse.ArgumentParser(description='关闭 AdsPower 环境')
parser.add_argument('env_id', nargs='?', default='kqqvpqj', 
                    help='环境 ID (默认: kqqvpqj 即环境#11)')
args = parser.parse_args()

ENV_ID = args.env_id
API_BASE = "http://local.adspower.net:50325"
PORT_FILE = f"adspower_port_{ENV_ID}.txt"  # 按环境ID命名的端口文件

print("=" * 60)
print(f"关闭 AdsPower环境 #{ENV_ID}")
print("=" * 60)

# 读取当前端口
port = None
if os.path.exists(PORT_FILE):
    try:
        with open(PORT_FILE, "r", encoding='utf-8') as f:
            port = f.read().strip()
            print(f"[Info] 从文件读取端口: {port}")
    except:
        pass

if not port:
    # 尝试读取旧格式的端口文件
    OLD_PORT_FILE = "adspower_port.txt"
    if os.path.exists(OLD_PORT_FILE):
        try:
            with open(OLD_PORT_FILE, "r", encoding='utf-8') as f:
                port = f.read().strip()
                print(f"[Info] 从旧文件读取端口: {port}")
        except:
            pass

if not port:
    port = "49907"  # 默认端口
    print(f"[Info] 使用默认端口: {port}")

# 方法1: 通过 CDP 关闭浏览器
print(f"\n[方法1] 通过 CDP 关闭浏览器 (端口 {port})...")
try:
    from playwright.sync_api import sync_playwright
    
    p = sync_playwright().start()
    browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
    
    print("  [OK] 已连接到 CDP")
    
    # 先处理所有页面的 beforeunload 弹窗
    context = browser.contexts[0] if browser.contexts else None
    if context:
        for page in context.pages:
            try:
                # 移除 beforeunload 监听器，防止弹窗
                page.evaluate("""
                    window.onbeforeunload = null;
                    window.addEventListener('beforeunload', function(e) {
                        e.preventDefault();
                        e.returnValue = '';
                    });
                """)
                print(f"  [OK] 已处理页面弹窗: {page.url[:50]}...")
            except:
                pass
    
    # 设置 dialog 处理（如果还有弹窗出现）
    if context:
        for page in context.pages:
            try:
                page.on("dialog", lambda dialog: dialog.accept())
            except:
                pass
    
    print(f"  [Info] 关闭浏览器...")
    
    # 强制关闭浏览器（不触发 beforeunload）
    browser.close()
    print("  [OK] 浏览器已关闭")
    p.stop()
    
except Exception as e:
    print(f"  [Warn] CDP 关闭失败: {e}")

# 方法2: API 关闭
print(f"\n[方法2] 通过 API 关闭环境 #{ENV_ID}...")
try:
    url = f"{API_BASE}/api/v1/browser/stop?user_id={ENV_ID}"
    response = requests.get(url, timeout=30)
    result = response.json()
    
    if result.get("code") == 0:
        print(f"  [OK] 环境 #{ENV_ID} 已通过 API 关闭")
    else:
        print(f"  [Info] API 返回: {result.get('msg')}")
except Exception as e:
    print(f"  [Warn] API 关闭失败: {e}")

# 方法3: 强制关闭进程
print("\n[方法3] 强制关闭相关进程...")
processes = [
    ('SunBrowser.exe', 'SunBrowser 浏览器'),
    ('chromedriver.exe', 'ChromeDriver'),
    # 注意: 不杀死 QuickQ.exe，通过点击电源按钮正常断开
]

for proc_name, desc in processes:
    try:
        result = subprocess.run(
            ['taskkill', '/F', '/IM', proc_name],
            capture_output=True,
            check=False
        )
        if result.returncode == 0:
            print(f"  [OK] {desc} ({proc_name}) 已终止")
        else:
            print(f"  [Info] {desc} 未运行或已关闭")
    except Exception as e:
        print(f"  [Error] 关闭 {proc_name} 失败: {e}")

# 清理端口文件
port_files_to_remove = [
    PORT_FILE,  # adspower_port_{ENV_ID}.txt
    "adspower_port.txt",  # 旧格式
]

for pf in port_files_to_remove:
    if os.path.exists(pf):
        try:
            os.remove(pf)
            print(f"\n[OK] 端口文件 {pf} 已删除")
        except:
            pass

# 方法4: 点击 QuickQ 电源按钮断开 VPN
print("\n[方法4] 点击 QuickQ 电源按钮断开 VPN...")
try:
    def find_window_by_title_contains(title_contains):
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
        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        return rect.left, rect.top, rect.right, rect.bottom
    
    # 查找 QuickQ 窗口
    quickq_hwnd = find_window_by_title_contains("QuickQ")
    if quickq_hwnd:
        # 将窗口带到前台
        ctypes.windll.user32.SetForegroundWindow(quickq_hwnd)
        time.sleep(0.5)
        
        # 获取窗口位置和大小
        left, top, right, bottom = get_window_rect(quickq_hwnd)
        width = right - left
        height = bottom - top
        
        # 计算电源按钮中心坐标
        center_x = left + width // 2
        center_y = top + height // 2 - 20
        
        print(f"  [Info] QuickQ 窗口位置: ({left}, {top}), 大小: {width}x{height}")
        print(f"  [Action] 点击电源按钮: ({center_x}, {center_y})")
        
        # 移动鼠标并点击
        ctypes.windll.user32.SetCursorPos(center_x, center_y)
        time.sleep(0.3)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # 按下
        time.sleep(0.1)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # 释放
        
        print("  [OK] 已点击 QuickQ 电源按钮")
        print("  [Wait] 等待 3 秒让 VPN 断开...")
        time.sleep(3)
    else:
        print("  [Info] 未找到 QuickQ 窗口")
except Exception as e:
    print(f"  [Warn] 点击 QuickQ 电源按钮失败: {e}")

# 方法5: 关闭 AdsPower 窗口
print("\n[方法5] 关闭 AdsPower 窗口...")
try:
    def find_window_by_title_contains(title_contains):
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
    
    # 查找并关闭 AdsPower 窗口
    adspower_hwnd = find_window_by_title_contains("AdsPower")
    if adspower_hwnd:
        ctypes.windll.user32.PostMessageW(adspower_hwnd, 0x10, 0, 0)  # WM_CLOSE
        print("  [OK] AdsPower 窗口已关闭")
        time.sleep(1)
    else:
        print("  [Info] 未找到 AdsPower 窗口")
except Exception as e:
    print(f"  [Warn] 关闭 AdsPower 窗口失败: {e}")

print("\n" + "=" * 60)
print("✅ 关闭完成！")
print("=" * 60)
print(f"\n环境 #{ENV_ID} 已关闭")
print("可以重新启动了。")
