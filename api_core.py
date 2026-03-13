"""
核心 API 模块 - 供 GUI 和 CLI 共用
"""
import requests
import json
import os
import uuid
import hashlib
import math
import asyncio
import websockets
import urllib.request
from datetime import datetime
from urllib.parse import quote_plus

QUERY_URL = "https://sucaiwang-api-elb.zhishangsoft.com/api/video/query"
_SIGN_POSITIONS = [2, 4, 5, 7, 11, 14, 15, 18, 22, 23, 26, 28, 31, 33, 35, 36]
TOKEN_FILE = os.path.join(os.path.dirname(__file__), ".token_cache")

# 复用 TCP 连接，大幅减少握手延迟
_session = requests.Session()


def _build_sign(token: str, timestamp: int, request_id: str) -> str:
    parts = token.split(".")
    payload_str = parts[1] if len(parts) > 1 else ""
    extracted = "".join(
        payload_str[i] if 0 <= i < len(payload_str) else ""
        for i in _SIGN_POSITIONS
    )
    params = {"requestId": request_id, "timestamp": timestamp}
    sorted_str = "&".join(
        f"{k}={params[k]}"
        for k in sorted(params.keys(), key=lambda x: x.lower())
    )
    raw = f"{sorted_str}&{extracted}"
    return hashlib.md5(raw.encode()).hexdigest()


def make_headers(token: str) -> dict:
    ts = math.floor(datetime.now().timestamp())
    req_id = str(uuid.uuid4())
    sign = _build_sign(token, ts, req_id)
    return {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/x-www-form-urlencoded",
        "token": token,
        "isInner": "0",
        "Access-Control-Allow-Private-Network": "True",
        "requestId": req_id,
        "timestamp": str(ts),
        "sign": sign,
    }


def encode_form(data: dict) -> str:
    parts = []
    for k, v in data.items():
        parts.append(f"{quote_plus(str(k))}={quote_plus(str(v))}")
    return "&".join(parts)


def load_cached_token():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return f.read().strip()
    return None


def save_token(token: str):
    with open(TOKEN_FILE, "w") as f:
        f.write(token)


def _ensure_chrome_debug():
    """确保 Chrome 或 Edge 以调试模式运行，如果没有就自动启动"""
    import subprocess, platform, time
    try:
        urllib.request.urlopen("http://localhost:9222/json", timeout=2)
        return  # 已经在运行
    except Exception:
        pass

    sys_name = platform.system()
    browser_path = None

    if sys_name == "Windows":
        candidates = [
            # Edge
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            # Chrome
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        browser_path = next((p for p in candidates if os.path.exists(p)), None)
        if browser_path:
            subprocess.Popen([browser_path, "--remote-debugging-port=9222",
                              "https://sucaiwang.zhishangsoft.com"])
    elif sys_name == "Darwin":
        candidates = [
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
        browser_path = next((p for p in candidates if os.path.exists(p)), None)
        if browser_path:
            subprocess.Popen([browser_path, "--remote-debugging-port=9222",
                              "https://sucaiwang.zhishangsoft.com"])

    if not browser_path:
        raise RuntimeError("未找到 Chrome 或 Edge 浏览器，请手动安装后重试")

    # 等待浏览器启动
    for _ in range(20):
        time.sleep(0.5)
        try:
            urllib.request.urlopen("http://localhost:9222/json", timeout=1)
            return
        except Exception:
            pass

    raise RuntimeError("浏览器启动超时，请手动打开浏览器并登录 sucaiwang.zhishangsoft.com")


def get_token_from_browser() -> str:
    _ensure_chrome_debug()  # 自动启动 Chrome 调试模式

    async def _fetch():
        import time
        # 等待 sucaiwang 标签页出现（最多 30 秒，等用户登录）
        for _ in range(60):
            tabs = json.loads(urllib.request.urlopen("http://localhost:9222/json", timeout=5).read())
            tab = next((t for t in tabs if "sucaiwang.zhishangsoft.com" in t.get("url", "")), None)
            if tab:
                break
            await asyncio.sleep(0.5)
        if not tab:
            raise RuntimeError("未找到 sucaiwang 标签页，请在 Chrome 中打开并登录")
        ws_url = tab["webSocketDebuggerUrl"]
        async with websockets.connect(ws_url) as ws:
            cmd = json.dumps({"id": 1, "method": "Runtime.evaluate",
                              "params": {"expression": "localStorage.getItem('material_token')"}})
            await ws.send(cmd)
            resp = json.loads(await ws.recv())
            return resp.get("result", {}).get("result", {}).get("value", "")

    loop = asyncio.new_event_loop()
    try:
        token = loop.run_until_complete(_fetch())
    finally:
        loop.close()
    if not token or not token.startswith("eyJ"):
        raise RuntimeError("读取 token 失败，请确认已在 Chrome 中登录")
    return token


def is_token_valid(token: str) -> bool:
    try:
        payload = {"length": 1, "start": 0, "name": "test", "searchType": 1,
                   "videoType": 0, "commentType": 0, "dyStatCostType": 0, "isPubArea": 0}
        resp = _session.post(QUERY_URL, data=encode_form(payload),
                             headers=make_headers(token), timeout=10)
        return resp.status_code == 200 and resp.json().get("code") == 1
    except Exception:
        return False


def get_token() -> str:
    token = load_cached_token()
    if token and is_token_valid(token):
        return token
    token = get_token_from_browser()
    save_token(token)
    return token


def search_video(token: str, name: str):
    payload = {
        "length": 60, "start": 0, "videoType": 0, "order": 1,
        "sortField": 1, "sortOrder": 1, "name": name,
        "commentType": 0, "dyStatCostType": 0, "isPubArea": 0,
        "searchType": 1, "linkIdJson": "[]", "isUseCache": "",
    }
    try:
        resp = _session.post(QUERY_URL, data=encode_form(payload),
                             headers=make_headers(token), timeout=15)
        data = resp.json()

        if not data.get("success"):
            err = data.get("info", "API 返回失败")
            if "token" in err.lower() or "登录" in err:
                return False, "", {"token_expired": True, "error": err}
            return False, "", {"error": err}

        items = data.get("data", {}).get("list", [])
        total = data.get("data", {}).get("total", 0)

        if not items:
            return False, "", {}

        # 收集所有返回的视频（API 按名称搜索，返回的都是相关结果）
        all_ids = [str(v.get("videoId", "")) for v in items]
        all_raws = items
        # 精确匹配的放前面
        matched = [v for v in items if v.get("name", "").strip() == name.strip()]

        if matched:
            v = matched[0]
            return True, ",".join(all_ids), {
                "videoState": v.get("videoState", ""),
                "sumStatCost": v.get("sumStatCost", 0),
                "sumPayOrderAmount": v.get("sumPayOrderAmount", 0),
                "sumRoi": v.get("sumRoi", 0),
                "match_count": len(items),
                "all_ids": all_ids,
                "all_raws": all_raws,
            }
        else:
            first_name = items[0].get("name", "") if items else ""
            return False, "", {"info": f"返回{total}条无精确匹配，第一条: {first_name}"}
    except Exception as e:
        return False, "", {"error": str(e)}


def fetch_video_objects_by_ids(token: str, video_ids: list) -> list:
    results = []
    chunk_size = 60
    for i in range(0, len(video_ids), chunk_size):
        chunk = video_ids[i:i + chunk_size]
        ids_str = ",".join(chunk)
        payload = {
            "length": chunk_size, "start": 0, "searchVideoIds": ids_str,
            "videoType": 0, "commentType": 0, "dyStatCostType": 0,
            "isPubArea": 0, "searchType": 1, "linkIdJson": "[]", "isUseCache": "",
        }
        resp = _session.post(QUERY_URL, data=encode_form(payload),
                             headers=make_headers(token), timeout=15)
        data = resp.json()
        if data.get("success"):
            items = data.get("data", {}).get("list", [])
            for obj in items:
                obj["id"] = obj.get("videoId", 0)
                obj["title"] = obj.get("name", "")
            results.extend(items)
    return results


def set_workbench_via_cdp(video_objects: list) -> str:
    async def _set():
        tabs = json.loads(urllib.request.urlopen("http://localhost:9222/json", timeout=5).read())
        tab = next((t for t in tabs if "sucaiwang.zhishangsoft.com" in t.get("url", "")), None)
        if not tab:
            raise RuntimeError("未找到 sucaiwang 标签页")
        ws_url = tab["webSocketDebuggerUrl"]
        async with websockets.connect(ws_url, max_size=50 * 1024 * 1024) as ws:
            js_data = json.dumps(video_objects)
            code = f"""(()=>{{
  const data = {js_data};
  localStorage.setItem('workbenchList', JSON.stringify(data));
  const app = document.querySelector('#app').__vue__;
  if (app && app.$store) app.$store.commit('app/SET_WORKBENCH_LIST', data);
  return data.length;
}})()"""
            await ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate",
                                      "params": {"expression": code, "returnByValue": True}}))
            for _ in range(10):
                msg = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
                if msg.get("id") == 1:
                    return str(msg["result"]["result"].get("value", "0"))
        return "0"

    loop = asyncio.new_event_loop()
    try:
        return loop.run_until_complete(_set())
    finally:
        loop.close()
