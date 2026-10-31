#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智慧牧业数字化大屏系统 · 云端高可用防休眠与自动自愈唤醒探测器
功能：
1. 轻量检测 Streamlit Cloud 容器健康端点及主页响应
2. 识别休眠特征（This app has gone to sleep）
3. 联动无头浏览器自动模拟点击官方唤醒按钮，实现零人工介入全自动唤醒
"""

import sys
import time
import urllib.request
import urllib.error

STREAMLIT_URL = "https://smartpastoral.streamlit.app/"
HEALTH_URL = "https://smartpastoral.streamlit.app/_stcore/health"
CUSTOM_DOMAIN_URL = "https://smartpastoral.com/"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

def check_http_status():
    """轻量 HTTP 探测"""
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 正在发起云端健康与主页探测...")
    is_sleeping = False
    
    # 1. 探测健康接口
    try:
        req = urllib.request.Request(HEALTH_URL, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read().decode('utf-8', errors='ignore')
            print(f"  -> 健康端点响应状态码: {resp.status}, 内容: {content.strip()[:60]}")
    except Exception as e:
        print(f"  -> 健康端点探测提示: {e}")

    # 2. 探测主页是否返回休眠页面特征
    try:
        req = urllib.request.Request(STREAMLIT_URL, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=20) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
            if "This app has gone to sleep" in html or "get this app back up" in html:
                print("  [!] 警告：检测到 Streamlit 云端容器已进入休眠状态！")
                is_sleeping = True
            else:
                print("  [✓] 正常：页面未出现休眠提示，容器处于活跃状态。")
    except urllib.error.HTTPError as e:
        print(f"  [!] 主页请求 HTTP 异常: {e.code} {e.reason}")
        if e.code in (404, 503, 500):
            is_sleeping = True
    except Exception as e:
        print(f"  [!] 主页请求异常: {e}")
        is_sleeping = True

    return is_sleeping

def trigger_playwright_wakeup():
    """使用 Playwright 无头浏览器触发自动唤醒"""
    try:
        import asyncio
        from playwright.async_api import async_playwright
    except ImportError:
        print("  [*] 当前运行环境未安装 playwright，已跳过无头浏览器唤醒模拟。")
        return False

    async def _async_wake():
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 启动 Playwright 无头浏览器执行自动点击唤醒...")
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1280, "height": 800}
            )
            try:
                await page.goto(STREAMLIT_URL, timeout=60000)
                await page.wait_for_timeout(4000)
                
                # 寻找官方唤醒按钮
                btn = page.locator("button:has-text('Yes, get this app back up')").or_(
                      page.locator("button:has-text('wake')")).or_(
                      page.locator("[data-testid='wakeup-button']"))
                
                if await btn.count() > 0:
                    print("  [✓] 成功定位到官方休眠唤醒按钮！正在自动模拟点击...")
                    await btn.first.click()
                    print("  [✓] 点击完成，正在等待容器拉起与初始化（25秒）...")
                    await page.wait_for_timeout(25000)
                    print("  [✓] 自动唤醒流程执行完毕！")
                    return True
                else:
                    print("  [*] 未发现唤醒按钮，应用可能已在唤醒中或已正常展示。")
                    return True
            except Exception as e:
                print(f"  [!] Playwright 唤醒过程异常: {e}")
                return False
            finally:
                await browser.close()

    return asyncio.run(_async_wake())

def main():
    print("=" * 65)
    print("智慧牧业云端大屏 · 全天候防休眠巡检与自动唤醒自愈程序")
    print(f"目标系统: {STREAMLIT_URL} / {CUSTOM_DOMAIN_URL}")
    print("=" * 65)
    
    sleeping = check_http_status()
    if sleeping:
        print("\n检测到系统休眠或服务未就绪，开始执行自愈唤醒...")
        success = trigger_playwright_wakeup()
        if not success:
            print("若在本地调试且未安装 playwright，建议在终端运行: pip install playwright && playwright install chromium")
    else:
        print("\n[✓] 检查完毕：云端系统处于健康活跃状态，保活心跳已记录。")
    print("=" * 65)

if __name__ == "__main__":
    main()
