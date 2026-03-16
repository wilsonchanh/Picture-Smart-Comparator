import cv2
import numpy as np
import math
import sys
import os
import glob
import win32com.client
import copy
import difflib
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading

engine = None

def get_ocr_result(image_path):
    global engine
    try:
        if engine is None:
            print(">>> 正在唤醒本地离线 OCR 模型，请稍候...")
            from rapidocr_onnxruntime import RapidOCR
            engine = RapidOCR()
            
        result, _ = engine(image_path)
        res_list = []
        if result:
            for box, text, score in result:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                left, right = min(xs), max(xs)
                top, bottom = min(ys), max(ys)
                res_list.append({
                    'words': text,
                    'location': {
                        'left': left, 'top': top, 
                        'width': right - left, 'height': bottom - top
                    }
                })
        return res_list
    except Exception as e:
        print(f"本地 OCR 识别报错: {e}")
        return []

def cv_imread(file_path):
    return cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), cv2.IMREAD_COLOR)

def cv_imwrite(file_path, img):
    cv2.imencode('.png', img)[1].tofile(file_path)

# ================= 核心：精准过滤，严格保留大小写 =================
def normalize_text(text, mode="fuzzy"):
    cleaned = re.sub(r'[^\w\u4e00-\u9fa5]', '', text)
    if mode == "exact":
        # 精确模式：绝对保留原始大小写形态
        return cleaned
    else:
        # 模糊模式：全部转小写，忽略大小写差异
        return cleaned.lower()
# ==============================================================

def convert_ppt_to_images(ppt_path, output_dir):
    img_paths = []
    try:
        import pythoncom
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        ppt_abs_path = os.path.abspath(ppt_path)
        deck = powerpoint.Presentations.Open(ppt_abs_path, ReadOnly=True, WithWindow=False)
        
        slide_w = deck.PageSetup.SlideWidth
        slide_h = deck.PageSetup.SlideHeight
        ratio = slide_w / slide_h if slide_h else 16/9
        
        export_h = 2160
        export_w = int(export_h * ratio)

        for i in range(1, deck.Slides.Count + 1):
            out_name = f"temp_slide_{i}_{os.path.basename(ppt_path)}.png"
            out_abs_path = os.path.abspath(os.path.join(output_dir, out_name))
            deck.Slides(i).Export(out_abs_path, "PNG", export_w, export_h)
            img_paths.append(out_abs_path)
            
        deck.Close()
        powerpoint.Quit()
        return img_paths
    except Exception as e:
        try: powerpoint.Quit()
        except: pass
        return []
    finally:
        try: pythoncom.CoUninitialize()
        except: pass

def get_median_height(res_list):
    if not res_list: return 40
    heights = [item['location']['height'] for item in res_list if 'location' in item]
    return np.median(heights) if heights else 40

def group_texts(res_list, merge_radius):
    clusters = []
    for item in res_list:
        loc = item.get('location', {})
        if not loc: continue
        clusters.append({
            'words': item.get('words', ''),
            'left': loc['left'], 'top': loc['top'],
            'right': loc['left'] + loc['width'], 'bottom': loc['top'] + loc['height'],
            'cx': loc['left'] + loc['width'] / 2, 'cy': loc['top'] + loc['height'] / 2,
            'items': [item]
        })

    changed = True
    while changed:
        changed = False
        for i in range(len(clusters)):
            for j in range(i + 1, len(clusters)):
                c1, c2 = clusters[i], clusters[j]
                dist = math.hypot(c1['cx'] - c2['cx'], c1['cy'] - c2['cy'])
                if dist < merge_radius:
                    c1['left'], c1['top'] = min(c1['left'], c2['left']), min(c1['top'], c2['top'])
                    c1['right'], c1['bottom'] = max(c1['right'], c2['right']), max(c1['bottom'], c2['bottom'])
                    c1['cx'], c1['cy'] = (c1['left'] + c1['right']) / 2, (c1['top'] + c1['bottom']) / 2
                    c1['items'].extend(c2['items'])
                    c1['items'].sort(key=lambda x: x['location']['top'])
                    c1['words'] = "".join([x['words'] for x in c1['items']])
                    del clusters[j]
                    changed = True
                    break 
            if changed: break
    return clusters

def calculate_similarity_for_pairing(resA, resB):
    # 配对找兄弟时，为了防止因为大小写变了而找不到对象，强制用模糊模式对比
    setA = set([normalize_text(item.get('words', ''), "fuzzy") for item in resA])
    setB = set([normalize_text(item.get('words', ''), "fuzzy") for item in resB])
    if not setA and not setB: return 0
    return len(setA & setB) / len(setA | setB)

def auto_compare(input_data, mode):
    print(f"当前比对模式：{'[精确找茬 - 严格区分大小写]' if mode == 'exact' else '[模糊找茬]'}")

    if isinstance(input_data, list):
        output_dir = os.path.dirname(input_data[0]) if input_data else ""
        raw_files = input_data
    else:
        if not os.path.exists(input_data):
            print("[错误] 选定的路径不存在！")
            return
        output_dir = input_data
        exts = ['*.png', '*.jpg', '*.jpeg', '*.bmp', '*.ppt', '*.pptx']
        raw_files = []
        for ext in exts:
            raw_files.extend(glob.glob(os.path.join(input_data, ext)))
            raw_files.extend(glob.glob(os.path.join(input_data, ext.upper())))
    
    unique_files = {}
    for f in raw_files:
        norm_name = os.path.normcase(os.path.abspath(f))
        if "diff_result" not in norm_name and "~$" not in norm_name and "temp_slide" not in norm_name:
            unique_files[norm_name] = f
    files = list(unique_files.values())

    ppt_files = [f for f in files if f.lower().endswith(('.ppt', '.pptx'))]
    all_image_paths = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]

    if ppt_files:
        print(">>> 正在以原生比例无损导出 PPT 高清图纸...")
    for ppt in ppt_files:
        all_image_paths.extend(convert_ppt_to_images(ppt, output_dir))

    if len(all_image_paths) < 2:
        print("[错误] 检测到的有效图纸页面不足 2 张，无法比对。")
        return

    print("\n>>> 正在动用本地离线引擎识别文字...")
    ocr_results = {}
    for path in all_image_paths:
        print(f" -> 扫描: {os.path.basename(path)}")
        ocr_results[path] = get_ocr_result(path)

    print("\n>>> 正在智能匹配新旧图纸...")
    all_pairs = []
    
    pngs = [f for f in all_image_paths if f.lower().endswith('.png')]
    jpgs = [f for f in all_image_paths if f.lower().endswith(('.jpg', '.jpeg'))]

    if len(pngs) > 0 and len(jpgs) > 0:
        for pathA in pngs:
            for pathB in jpgs:
                score = calculate_similarity_for_pairing(ocr_results[pathA], ocr_results[pathB])
                if score > 0.05:
                    all_pairs.append((score, pathA, pathB))
    else:
        for i in range(len(all_image_paths)):
            for j in range(i + 1, len(all_image_paths)):
                pathA = all_image_paths[i]
                pathB = all_image_paths[j]
                score = calculate_similarity_for_pairing(ocr_results[pathA], ocr_results[pathB])
                if score > 0.05:
                    all_pairs.append((score, pathA, pathB))
                
    all_pairs.sort(key=lambda x: x[0], reverse=True)
    
    matched_pairs = []
    used_images = set()
    
    for score, pathA, pathB in all_pairs:
        if pathA not in used_images and pathB not in used_images:
            if os.path.basename(pathA) > os.path.basename(pathB):
                pathA, pathB = pathB, pathA
            matched_pairs.append((pathA, pathB, ocr_results[pathA], ocr_results[pathB]))
            used_images.add(pathA)
            used_images.add(pathB)
            print(f" [配对成功] '{os.path.basename(pathA)}' <-> '{os.path.basename(pathB)}'")

    for idx, (imgA_path, imgB_path, resA_raw, resB_raw) in enumerate(matched_pairs):
        print(f"\n>>> 正在生成第 {idx+1} 组比对结果...")
        imgA_visual = cv_imread(imgA_path)
        imgB_visual = cv_imread(imgB_path)
        
        if imgA_visual is None or imgB_visual is None:
            continue

        hA, wA = imgA_visual.shape[:2]
        hB, wB = imgB_visual.shape[:2]
        
        h_median_A = get_median_height(resA_raw)
        h_median_B = get_median_height(resB_raw)
        merge_radius_A = h_median_A * 1.5
        merge_radius_B = h_median_B * 1.5

        clustersA = group_texts(resA_raw, merge_radius_A)
        clustersB = group_texts(resB_raw, merge_radius_B)
        
        candidate_matches = []
        max_dist_norm = 0.15 
        
        for i_B, cB in enumerate(clustersB):
            norm_B = normalize_text(cB['words'], mode)
            if not norm_B: continue
            
            cx_B_norm = cB['cx'] / wB
            cy_B_norm = cB['cy'] / hB
                
            for i_A, cA in enumerate(clustersA):
                norm_A = normalize_text(cA['words'], mode)
                if not norm_A: continue
                
                cx_A_norm = cA['cx'] / wA
                cy_A_norm = cA['cy'] / hA
                
                similarity = difflib.SequenceMatcher(None, norm_A, norm_B).ratio()
                is_match = False
                
                # ================= 终极重拳出击 =================
                if mode == "exact":
                    # 彻底取消长句容错，只认 100% 绝对死理！App 和 APP 就是不一样！
                    is_match = (norm_A == norm_B)
                else:
                    is_match = (similarity >= 0.75)
                # ================================================
                
                if is_match:
                    dist_norm = math.hypot(cx_B_norm - cx_A_norm, cy_B_norm - cy_A_norm)
                    if dist_norm < max_dist_norm:
                        candidate_matches.append((similarity, -dist_norm, i_A, i_B))
                    
        candidate_matches.sort(key=lambda x: (x[0], x[1]), reverse=True)
        matched_A_indices, matched_B_indices = set(), set()
        
        for sim, neg_dist, i_A, i_B in candidate_matches:
            if i_A not in matched_A_indices and i_B not in matched_B_indices:
                matched_A_indices.add(i_A)
                matched_B_indices.add(i_B)

        target_h = max(hA, hB)
        
        scale_A = target_h / hA
        imgA_final = cv2.resize(imgA_visual, (int(wA * scale_A), target_h))

        scale_B = target_h / hB
        imgB_final = cv2.resize(imgB_visual, (int(wB * scale_B), target_h))

        diff_count_B = diff_count_A = 0
        
        for i_B, cB in enumerate(clustersB):
            if i_B not in matched_B_indices and normalize_text(cB['words'], mode):
                left = int(cB['left'] * scale_B)
                top = int(cB['top'] * scale_B)
                right = int(cB['right'] * scale_B)
                bottom = int(cB['bottom'] * scale_B)
                cv2.rectangle(imgB_final, (left, top), (right, bottom), (0, 0, 255), 4)
                diff_count_B += 1

        for i_A, cA in enumerate(clustersA):
            if i_A not in matched_A_indices and normalize_text(cA['words'], mode):
                left = int(cA['left'] * scale_A)
                top = int(cA['top'] * scale_A)
                right = int(cA['right'] * scale_A)
                bottom = int(cA['bottom'] * scale_A)
                cv2.rectangle(imgA_final, (left, top), (right, bottom), (0, 215, 255), 4)
                diff_count_A += 1

        print(f" - 第 {idx+1} 组完成 (左缺失: {diff_count_A}处, 右修改: {diff_count_B}处)")

        combined_img = cv2.hconcat([imgA_final, imgB_final])
        output_name = f"diff_result_{idx+1}_{'精确' if mode == 'exact' else '模糊'}.png"
        output_path = os.path.join(output_dir, output_name)
        cv_imwrite(output_path, combined_img)

    for temp_img in all_image_paths:
        if "temp_slide" in os.path.basename(temp_img) and os.path.exists(temp_img):
            os.remove(temp_img)
            
    print(f"\n🎉 恭喜！所有配对比对完成！请前往文件夹查看结果图。")

# ================= GUI 可视化界面 =================
class RedirectText(object):
    def __init__(self, text_ctrl):
        self.output = text_ctrl
    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)
    def flush(self): pass

def start_gui():
    global selected_files_list
    root = tk.Tk()
    root.title("展览图纸本地核对神器 (大小写绝对敏感版)")
    root.geometry("650x550")
    root.configure(bg="#f0f0f0")

    frame_top = tk.Frame(root, bg="#f0f0f0")
    frame_top.pack(fill="x", padx=15, pady=15)

    tk.Label(frame_top, text="处理目标 (文件夹 或 右键选中的图纸):", bg="#f0f0f0", font=("微软雅黑", 10)).pack(anchor="w", pady=(0, 5))
    
    path_entry = tk.Entry(frame_top, font=("微软雅黑", 10))
    path_entry.pack(side="left", fill="x", expand=True, ipady=4)

    selected_files_list = []
    if len(sys.argv) > 1:
        if os.path.isdir(sys.argv[1]):
            path_entry.insert(0, sys.argv[1])
        else:
            selected_files_list = sys.argv[1:]
            path_entry.insert(0, f"✅ 准备就绪！已获取您选中的 {len(selected_files_list)} 个文件")
            path_entry.config(state="readonly", fg="green")
    else:
        path_entry.insert(0, os.path.join(os.path.expanduser("~"), "Desktop", "PPT"))

    def browse_folder():
        global selected_files_list
        folder = filedialog.askdirectory()
        if folder:
            path_entry.config(state="normal", fg="black")
            path_entry.delete(0, tk.END)
            path_entry.insert(0, folder)
            selected_files_list = [] 

    btn_browse = tk.Button(frame_top, text="选文件夹", font=("微软雅黑", 9), command=browse_folder)
    btn_browse.pack(side="right", padx=(10, 0), ipadx=5)

    frame_mode = tk.LabelFrame(root, text=" 选择比对严苛度 ", bg="#f0f0f0", font=("微软雅黑", 10, "bold"), fg="#333333")
    frame_mode.pack(fill="x", padx=15, pady=(0, 15))

    compare_mode = tk.StringVar(value="exact")

    rb_exact = ttk.Radiobutton(frame_mode, text="精确匹配：极度严苛，只要有一个字母大小写不同，立刻标红！", variable=compare_mode, value="exact")
    rb_exact.pack(anchor="w", padx=10, pady=8)

    rb_fuzzy = ttk.Radiobutton(frame_mode, text="模糊匹配：忽略大小写差异，允许大段落错别字时的粗略核对。", variable=compare_mode, value="fuzzy")
    rb_fuzzy.pack(anchor="w", padx=10, pady=8)

    def run_compare():
        if selected_files_list:
            target_data = selected_files_list
        else:
            target_data = path_entry.get()
            if not os.path.exists(target_data):
                messagebox.showerror("错误", "路径不存在！")
                return
                
        mode = compare_mode.get() 
        btn_start.config(state="disabled", text="智脑全速核对中...")
        log_area.delete(1.0, tk.END)
        
        def target_func():
            try:
                auto_compare(target_data, mode)
            except Exception as e:
                print(f"\n发生严重错误: {e}")
            finally:
                btn_start.config(state="normal", text="立即启动本地核对")
        
        threading.Thread(target=target_func, daemon=True).start()

    btn_start = tk.Button(root, text="立即启动本地核对", bg="#FFD700", fg="black", font=("微软雅黑", 12, "bold"), command=run_compare)
    btn_start.pack(pady=5, ipadx=20, ipady=5)

    tk.Label(root, text="执行日志:", bg="#f0f0f0", font=("微软雅黑", 10)).pack(anchor="w", padx=15)
    log_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Consolas", 9), bg="#1e1e1e", fg="#00ff00")
    log_area.pack(padx=15, pady=(0, 15), fill="both", expand=True)

    sys.stdout = RedirectText(log_area)
    sys.stderr = RedirectText(log_area)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
