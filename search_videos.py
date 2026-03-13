"""
视频素材批量查询脚本
- 从 Chrome 浏览器自动读取 token（无需手动操作）
- 正确生成 sign 签名（逆向自前端 JS）
- 从 Excel 读取素材名称，逐行调用 API 查询
- 将结果（找到/未找到）回写到 Excel
"""
import requests
import pandas as pd
import json
import time
import os
import uuid
import hashlib
import math
from datetime import datetime

# ========== 配置 ==========
QUERY_URL = "https://sucaiwang-api-elb.zhishangsoft.com/api/video/query"
EXCEL_FILE = "/Users/msc/rag/全域数据_素材分析_视频_2026-03-01 00_00_00-2026-03-10 23_59_59-7615898179875143721 (1).xlsx"
NAME_COLUMN = "素材名称"
REQUEST_INTERVAL = 0.5
BATCH_SIZE = 60
TOKEN_FILE = os.path.join(os.path.dirname(__file__), ".token_cache")
# ==========================

# sign 算法中提取 token payload 的字符位置（逆向自前端 JS）
_SIGN_POSITIONS = [2, 4, 5, 7, 11, 14, 15, 18, 22, 23, 26, 28, 31, 33, 35, 36]


def _build_sign(token: str, timestamp: int, request_id: str) -> str:
    """
    复现前端 sign 生成逻辑：
      s = token.split('.')[1] 中特定位置的字符拼接
      params = {timestamp, requestId} 按 key 排序后 join '&'
      sign = MD5(params + '&' + s)
    """
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


def _make_headers(token: str) -> dict:
    """生成一次请求所需的完整 headers"""
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


def _encode_form(data: dict) -> str:
    """将 dict 编码为 x-www-form-urlencoded 字符串（与前端 transformRequest 一致）"""
    from urllib.parse import quote_plus
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


def get_token_from_browser() -> str:
    """通过 CDP 连接已运行的 Chrome，从 localStorage 读取 material_token"""
    import asyncio
    import websockets

    async def _fetch():
        import urllib.request
        tabs = json.loads(urllib.request.urlopen("http://localhost:9222/json", timeout=5).read())
        tab = next((t for t in tabs if "sucaiwang.zhishangsoft.com" in t.get("url", "")), None)
        if not tab:
            raise RuntimeError("未找到 sucaiwang.zhishangsoft.com 标签页，请先在 Chrome 中打开并登录")
        ws_url = tab["webSocketDebuggerUrl"]
        async with websockets.connect(ws_url) as ws:
            cmd = json.dumps({"id": 1, "method": "Runtime.evaluate",
                              "params": {"expression": "localStorage.getItem('material_token')"}})
            await ws.send(cmd)
            resp = json.loads(await ws.recv())
            return resp.get("result", {}).get("result", {}).get("value", "")

    token = asyncio.run(_fetch())
    if not token or not token.startswith("eyJ"):
        raise RuntimeError("从浏览器读取 token 失败，请确认已在 Chrome 中登录")
    print("从浏览器读取 token 成功")
    return token


def _is_token_valid(token: str) -> bool:
    """发一个测试请求验证 token 是否有效"""
    try:
        payload = {"length": 1, "start": 0, "name": "test", "searchType": 1,
                   "videoType": 0, "commentType": 0, "dyStatCostType": 0, "isPubArea": 0}
        resp = requests.post(QUERY_URL, data=_encode_form(payload),
                             headers=_make_headers(token), timeout=10)
        return resp.status_code == 200 and resp.json().get("code") == 1
    except Exception:
        return False


def get_token() -> str:
    """优先使用缓存 token，失效则从浏览器重新读取"""
    token = load_cached_token()
    if token and _is_token_valid(token):
        print("使用缓存 token")
        return token
    if token:
        print("缓存 token 已失效，从浏览器重新读取...")
    token = get_token_from_browser()
    save_token(token)
    return token


def search_video(token: str, name: str):
    """
    按素材名称搜索视频
    返回: (found: bool, video_id: str, extra: dict)
    """
    payload = {
        "length": BATCH_SIZE,
        "start": 0,
        "videoType": 0,
        "order": 1,
        "sortField": 1,
        "sortOrder": 1,
        "name": name,
        "commentType": 0,
        "dyStatCostType": 0,
        "isPubArea": 0,
        "searchType": 1,
        "linkIdJson": "[]",
        "isUseCache": "",
    }
    try:
        resp = requests.post(QUERY_URL, data=_encode_form(payload),
                             headers=_make_headers(token), timeout=15)
        data = resp.json()

        if not data.get("success"):
            err = data.get("info", "API 返回失败")
            # token 失效特征
            if "token" in err.lower() or "登录" in err:
                return False, "", {"token_expired": True, "error": err}
            return False, "", {"error": err}

        items = data.get("data", {}).get("list", [])
        total = data.get("data", {}).get("total", 0)

        matched = [v for v in items if v.get("name", "").strip() == name.strip()]
        if matched:
            v = matched[0]
            return True, str(v.get("videoId", "")), {
                "videoState": v.get("videoState", ""),
                "sumStatCost": v.get("sumStatCost", 0),
                "sumPayOrderAmount": v.get("sumPayOrderAmount", 0),
                "sumRoi": v.get("sumRoi", 0),
            }
        elif total > 0:
            first_name = items[0].get("name", "") if items else ""
            return False, "", {"info": f"返回{total}条但无精确匹配，第一条: {first_name}"}
        else:
            return False, "", {}

    except Exception as e:
        return False, "", {"error": str(e)}


def main():
    print("=" * 50)
    print("视频素材批量查询脚本")
    print("=" * 50)

    token = get_token()
    print(f"Token 就绪（前20字符）: {token[:20]}...\n")

    if not os.path.exists(EXCEL_FILE):
        print(f"Excel 文件不存在: {EXCEL_FILE}")
        return

    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    print(f"已加载 Excel，共 {len(df)} 行")

    if NAME_COLUMN not in df.columns:
        print(f"找不到列 '{NAME_COLUMN}'，现有列: {df.columns.tolist()}")
        return

    for col in ["查询状态", "视频ID", "查询备注", "查询时间"]:
        if col not in df.columns:
            df[col] = ""

    found_count = not_found_count = skip_count = error_count = 0

    for idx, row in df.iterrows():
        name = row[NAME_COLUMN]

        if pd.isna(name) or str(name).strip() == "":
            skip_count += 1
            continue

        name = str(name).strip()

        if str(row.get("查询状态", "")).strip() == "找到":
            skip_count += 1
            continue

        print(f"[{idx+1}/{len(df)}] {name}", end=" ... ", flush=True)

        found, video_id, extra = search_video(token, name)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # token 失效时自动刷新
        if extra.get("token_expired"):
            print("token 失效，刷新...", end=" ", flush=True)
            token = get_token_from_browser()
            save_token(token)
            found, video_id, extra = search_video(token, name)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if found:
            df.at[idx, "查询状态"] = "找到"
            df.at[idx, "视频ID"] = video_id
            df.at[idx, "查询备注"] = f"state={extra.get('videoState','')} cost={extra.get('sumStatCost',0)} roi={extra.get('sumRoi',0)}"
            df.at[idx, "查询时间"] = now
            found_count += 1
            print(f"找到 (ID: {video_id})")
        elif "error" in extra:
            df.at[idx, "查询状态"] = "错误"
            df.at[idx, "查询备注"] = extra["error"]
            df.at[idx, "查询时间"] = now
            error_count += 1
            print(f"错误: {extra['error']}")
        else:
            df.at[idx, "查询状态"] = "未找到"
            df.at[idx, "查询备注"] = extra.get("info", "")
            df.at[idx, "查询时间"] = now
            not_found_count += 1
            print("未找到")

        if (idx + 1) % 10 == 0:
            df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
            print("  [进度已保存]")

        time.sleep(REQUEST_INTERVAL)

    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

    print("\n" + "=" * 50)
    print("查询完成！")
    print(f"  找到:   {found_count} 条")
    print(f"  未找到: {not_found_count} 条")
    print(f"  错误:   {error_count} 条")
    print(f"  跳过:   {skip_count} 条")
    print(f"结果已保存到: {EXCEL_FILE}")
    print("=" * 50)


if __name__ == "__main__":
    main()
