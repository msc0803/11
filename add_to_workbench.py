"""
批量加入工作台脚本
- 从 Excel 读取已查询到的视频 ID
- 通过 CDP 直接写入浏览器 localStorage.workbenchList
- 每批最多 200 个，自动分批提示
"""
import json
import asyncio
import websockets
import urllib.request
import pandas as pd
import math

# ========== 配置 ==========
EXCEL_FILE = "/Users/msc/rag/全域数据_素材分析_视频_2026-03-01 00_00_00-2026-03-10 23_59_59-7615898179875143721 (1).xlsx"
BATCH_SIZE = 200  # 工作台最多容纳数量
# ==========================


def get_video_ids_from_excel():
    """从 Excel 读取所有查询状态为'找到'的视频 ID"""
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    if "查询状态" not in df.columns or "视频ID" not in df.columns:
        raise RuntimeError("Excel 中找不到'查询状态'或'视频ID'列，请先运行 search_videos.py")

    found = df[df["查询状态"] == "找到"]["视频ID"].dropna()
    ids = [str(int(float(v))) for v in found]
    print(f"Excel 中找到 {len(ids)} 条有效视频 ID")
    return ids


async def get_browser_tab():
    tabs = json.loads(urllib.request.urlopen("http://localhost:9222/json", timeout=5).read())
    tab = next((t for t in tabs if "sucaiwang.zhishangsoft.com" in t.get("url", "")), None)
    if not tab:
        raise RuntimeError("未找到 sucaiwang.zhishangsoft.com 标签页，请先在 Chrome 中打开并登录")
    return tab


async def get_current_workbench(ws) -> list:
    """读取当前工作台列表"""
    cmd = json.dumps({"id": 10, "method": "Runtime.evaluate", "params": {
        "expression": "localStorage.getItem('workbenchList') || '[]'",
        "returnByValue": True
    }})
    await ws.send(cmd)
    for _ in range(10):
        msg = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
        if msg.get("id") == 10:
            val = msg["result"]["result"].get("value", "[]")
            return json.loads(val)
    return []


async def fetch_video_objects(ws, token: str, video_ids: list, make_headers_fn) -> list:
    """通过 API 获取视频完整对象（工作台需要完整对象，不只是 ID）"""
    from urllib.parse import quote_plus
    import requests as req_lib

    QUERY_URL = "https://sucaiwang-api-elb.zhishangsoft.com/api/video/query"
    results = []
    chunk_size = 60

    for i in range(0, len(video_ids), chunk_size):
        chunk = video_ids[i:i + chunk_size]
        ids_str = ",".join(chunk)
        payload_dict = {
            "length": chunk_size,
            "start": 0,
            "searchVideoIds": ids_str,
            "videoType": 0,
            "commentType": 0,
            "dyStatCostType": 0,
            "isPubArea": 0,
            "searchType": 1,
            "linkIdJson": "[]",
            "isUseCache": "",
        }
        body = "&".join(f"{quote_plus(k)}={quote_plus(str(v))}" for k, v in payload_dict.items())
        headers = make_headers_fn(token)
        resp = req_lib.post(QUERY_URL, data=body, headers=headers, timeout=15)
        data = resp.json()
        if data.get("success"):
            items = data.get("data", {}).get("list", [])
            results.extend(items)
            print(f"  获取视频对象 {i+1}~{i+len(chunk)}：{len(items)} 条")
        else:
            print(f"  获取失败: {data.get('info')}")

    return results


async def set_workbench(ws, video_objects: list):
    """将视频对象列表写入 localStorage 并同步到 Vuex store"""
    js_data = json.dumps(video_objects)
    code = f"""(()=>{{
  const data = {js_data};
  localStorage.setItem('workbenchList', JSON.stringify(data));
  // 同步到 Vuex store
  const app = document.querySelector('#app').__vue__;
  if (app && app.$store) {{
    app.$store.commit('app/SET_WORKBENCH_LIST', data);
  }}
  return 'OK: ' + data.length + ' items';
}})()"""
    cmd = json.dumps({"id": 20, "method": "Runtime.evaluate", "params": {
        "expression": code, "returnByValue": True
    }})
    await ws.send(cmd)
    for _ in range(10):
        msg = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
        if msg.get("id") == 20:
            return msg["result"]["result"].get("value", "")
    return "unknown"


async def main():
    print("=" * 50)
    print("批量加入工作台")
    print("=" * 50)

    # 读取视频 ID
    video_ids = get_video_ids_from_excel()
    if not video_ids:
        print("没有可添加的视频 ID")
        return

    total_batches = math.ceil(len(video_ids) / BATCH_SIZE)
    print(f"共 {len(video_ids)} 个 ID，分 {total_batches} 批（每批最多 {BATCH_SIZE} 个）\n")

    tab = await get_browser_tab()
    ws_url = tab["webSocketDebuggerUrl"]

    async with websockets.connect(ws_url, max_size=50 * 1024 * 1024) as ws:
        # 读取当前工作台
        current = await get_current_workbench(ws)
        current_ids = {str(v.get("videoId", "")) for v in current}
        print(f"当前工作台已有 {len(current)} 个视频")

        # 过滤掉已在工作台的
        new_ids = [vid for vid in video_ids if vid not in current_ids]
        print(f"去重后新增 {len(new_ids)} 个\n")

        if not new_ids:
            print("所有视频已在工作台中，无需添加")
            return

        # 获取 token（从 localStorage）
        cmd = json.dumps({"id": 5, "method": "Runtime.evaluate", "params": {
            "expression": "localStorage.getItem('material_token')", "returnByValue": True
        }})
        await ws.send(cmd)
        token = ""
        for _ in range(10):
            msg = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if msg.get("id") == 5:
                token = msg["result"]["result"].get("value", "")
                break

        if not token:
            raise RuntimeError("无法读取 token")

        # 构建 make_headers 函数
        import sys
        sys.path.insert(0, "/Users/msc/rag/api_search")
        from search_videos import _make_headers

        # 分批处理
        for batch_idx in range(total_batches):
            batch_ids = new_ids[batch_idx * BATCH_SIZE:(batch_idx + 1) * BATCH_SIZE]
            print(f"--- 第 {batch_idx + 1}/{total_batches} 批（{len(batch_ids)} 个）---")

            if batch_idx > 0:
                input(f"\n工作台已加载第 {batch_idx} 批，请先处理推送，然后按回车继续下一批...")
                # 清空工作台准备下一批
                await set_workbench(ws, [])

            print("正在从 API 获取视频完整数据...")
            video_objects = await fetch_video_objects(ws, token, batch_ids, _make_headers)
            print(f"获取到 {len(video_objects)} 个视频对象")

            # 补全 id 和 title 字段（工作台必须）
            for obj in video_objects:
                obj["id"] = obj.get("videoId", 0)
                obj["title"] = obj.get("name", "")

            result = await set_workbench(ws, video_objects)
            print(f"写入工作台结果: {result}")
            print(f"✓ 第 {batch_idx + 1} 批已加入工作台，请在浏览器中查看并推送！")

    print("\n全部批次处理完成！")


if __name__ == "__main__":
    import sys
    # 可传入批次号参数，如: python3 add_to_workbench.py 1
    # 不传则从第1批开始，每批结束后自动等待用户在终端按回车
    asyncio.run(main())
