# 视频素材批量查询工具

批量查询视频素材库，支持将找到的视频一键加入工作台推送。

## 功能

- 导入 Excel，批量按素材名称查询视频
- 实时显示查询进度、找到/未找到/错误统计
- 表格支持按状态筛选、关键词搜索、复制单元格
- 找到的视频可批量加入工作台（分批推送，每批 200 个）
- 结果自动保存到独立目录，不覆盖原文件
- 支持多 Sheet Excel，自动检测是否有表头

## 使用前提

Chrome 浏览器需以调试模式启动（程序通过此方式读取登录 Token）：

**Mac：**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
```

**Windows：**
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
```

> 建议创建快捷方式，在目标路径末尾加上 `--remote-debugging-port=9222`

Chrome 启动后，打开并登录 `sucaiwang.zhishangsoft.com`，然后再运行本工具。

## Mac 运行

**安装依赖：**
```bash
pip3 install pyqt6 pandas openpyxl requests websockets
```

**启动：**
```bash
python3 gui_app.py
```

## Windows 打包（生成 exe）

1. 安装 [Python 3.10](https://www.python.org/downloads/release/python-31011/)，安装时勾选 `Add Python to PATH`
2. 将以下文件拷贝到同一文件夹：
   - `gui_app.py`
   - `api_core.py`
   - `build_win.bat`
3. 右键 `build_win.bat` → 以管理员身份运行
4. 等待 3-5 分钟，打包完成后 exe 在 `dist\视频查询工具.exe`

## 使用流程

1. 点击 **打开 Excel**，选择素材名称文件
2. 选择工作表和搜索列（默认第一列）
3. 点击 **开始查询**，实时查看进度
4. 查询完成后，点击 **加入工作台** 批量推送
5. 每批 200 个，推送完一批后确认继续下一批

## 文件说明

| 文件 | 说明 |
|------|------|
| `gui_app.py` | 主程序（GUI） |
| `api_core.py` | API 核心模块（签名、Token、查询） |
| `search_videos.py` | 命令行批量查询脚本 |
| `add_to_workbench.py` | 命令行工作台脚本 |
| `build_win.bat` | Windows 一键打包脚本 |
