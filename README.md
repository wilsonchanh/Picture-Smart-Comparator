图纸找茬智能核对神器 (Drawing Smart Comparator)

```markdown
# 🚀 展览图纸智能找茬神器 (Drawing Smart Comparator)

## 📖 项目简介
在项目对接与图纸多轮修改中，肉眼核对变动细节耗时且极易遗漏。本项目利用本地开源 OCR 大模型结合 OpenCV 计算机视觉与相对坐标系算法，实现新旧图纸（或 PPT）的毫秒级比对，精准框注增、删、改的位置。

## ✨ 核心功能
- **纯本地离线引擎**：内置开源顶流 RapidOCR 引擎，无网络限制，图纸数据绝对本地保密。
- **相对坐标系套准**：无视新旧图纸分辨率差异，基于 0%~100% 的相对坐标系定位，防止排版错位误报。
- **长短句降维智脑**：
  - 🔬 **精确匹配模式**：内置正则表达式，强行剥离标点与空格。对价格、数字等短字符执行 100% 严格比对；对长句容忍极轻微的识别噪点。修改必抓，排版无视！
  - 🎯 **模糊匹配模式**：容忍大段重排或一定程度的错别字，适合粗略核对。

## 🛠️ 环境配置
本项目依赖视觉与离线 AI 库，一键安装指令（推荐使用清华源加速）：
```bash
pip install opencv-python numpy pywin32 rapidocr_onnxruntime -i [https://pypi.tuna.tsinghua.edu.cn/simple](https://pypi.tuna.tsinghua.edu.cn/simple)
注：如运行报错 DLL load failed，请安装 Microsoft Visual C++ Redistributable 运行库。

🚀 安装与右键菜单注入
为支持多张图纸的快速选中比对，需注入“发送到”菜单：

新建 install_comparator.py 并运行以下代码：

Python
import os, sys, win32com.client
pythonw_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
tool_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "图纸找茬双模式版.pyw")

sendto_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'SendTo')
shortcut_path = os.path.join(sendto_dir, "智能找茬比对图纸.lnk")

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(shortcut_path)
shortcut.TargetPath = pythonw_exe
shortcut.Arguments = f'"{tool_path}"'
shortcut.IconLocation = pythonw_exe + ",0" 
shortcut.Save()
print("✅ 图纸找茬功能已成功注入右键'发送到'菜单！")
💡 使用说明
将新、老版本的图纸（或 PPT 文件）放在一起。

框选需要对比的多个文件 -> 右键 -> 发送到 -> 智能找茬比对图纸。

在弹出的可视化面板中选择“精确匹配”或“模糊匹配” -> 立即启动。

目录下将生成 diff_result_xxx.png，红框为修改/删除，黄框为原图缺失/新增。
