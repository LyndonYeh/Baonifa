import re
import os
import pandas as pd
from openpyxl import load_workbook

# 設定檔案路徑與名稱
def define_paths(folder_path, patch_file_name, output_excel_name):
    patch_file = os.path.join(folder_path, patch_file_name)
    excel_file = os.path.join(folder_path, output_excel_name)
    return patch_file, excel_file

# 過濾非法字符
def remove_illegal_characters(text):
    """移除非法字符，避免 openpyxl 的 IllegalCharacterError"""
    return re.sub(r"[\x00-\x1F]", "", text)

# 讀取 .patch 檔案內容
def parse_patch_to_excel(patch_file, excel_file):
    with open(patch_file, 'r', encoding='utf-8', errors='ignore') as file:
        patch_content = file.read()
        
    # 解析 commit 訊息
    commit_messages = {}
    current_commit = None
    current_message = ""
    for line in patch_content.splitlines():
        if line.startswith("commit "):
            current_commit = line.split()[1]
        elif current_commit and line.startswith("    ") and (line.strip().startswith("fix:") or line.strip().startswith("style:")  or line.strip().startswith("feat:") or line.strip().startswith("fix") or line.strip().startswith("Issue")or line.strip().startswith("Revert")or line.strip().startswith("chore")or line.strip().startswith("refactor")):
            commit_messages[current_commit] = line.strip()
        #elif current_commit and line.startswith("    "):
        #    commit_messages[current_commit] = line.strip()

    # 模組代號對應表
    module_mapping = {
        "wob": "工作管理",
        "cus": "客戶管理",
        "pro": "商品管理",
        "mkt": "行銷管理",
        "rpt": "營運管理",
        "gen": "訊息溝通",
        "adm": "系統管理",
        "common": "共用設定",
        "com": "共用設定",
        "wkf": "審核相關",
        "database": "DB連線管理",
        "external": "外部電文相關",
        "soap": "電文相關",
        "rest": "電文相關",
        "mock": "電文虛擬資料",
        "util": "程式工具"
    }

    # 解析檔案變更
    pattern = re.compile(r'diff --git a/(.+) b/(.+)')
    rows = []
    for match in pattern.finditer(patch_content):
        file_path = match.group(2)
        commit_hash = None
        commit_msg = ""
        
        # 找到該程式最近的 commit message
        for line in reversed(patch_content[:match.start()].splitlines()):
            if line.startswith("commit "):
                commit_hash = line.split()[1]
                commit_msg = commit_messages.get(commit_hash, "")
                break
        
        # 取得模組代號
        module_match = re.search(r'/tbm/([^/]+)/', file_path)
        modal_match = re.search(r'/modal/([^/]+)/', file_path)
        static_match = re.search(r'/static/model/([^/]+)/', file_path)
        service_match = re.search(r'/service/impl/([^/]+)/', file_path)
        templates_match = re.search(r'/templates/([^/]+)/', file_path)
        
        module_code = module_match.group(1) if module_match else (
            modal_match.group(1) if modal_match else (
                static_match.group(1) if static_match else (
                    service_match.group(1) if service_match else (
                        templates_match.group(1) if templates_match else ""
                    )
                )
            )
        )
        
        # 如果模組代號為空或是 'include'，則改為 'com'
        if not module_code or module_code.lower() == 'include':
            module_code = 'com'
        
        # 取得模組名稱
        module_name = module_mapping.get(module_code, "")
        
        directory, filename = file_path.rsplit("/", 1) if "/" in file_path else ("", file_path)
        rows.append(["", module_code, module_name, commit_msg, "AP", directory, filename])

    # 轉換為 DataFrame 並移除非法字符
    df = pd.DataFrame(rows, columns=["過版日期", "模組代號", "模組名稱", "過版內容說明", "屬性", "程式路徑", "檔名"])
    for col in df.columns:
        df[col] = df[col].apply(lambda x: remove_illegal_characters(x) if isinstance(x, str) else x)

    # 存為 Excel
    df.to_excel(excel_file, index=False, engine='openpyxl')

    # 調整欄寬
    wb = load_workbook(excel_file)
    ws = wb.active
    column_widths = {"D": 90, "F": 90, "G": 40, "E": 10}  # 指定欄寬
    default_width = 14
    
    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = column_widths.get(col_letter, default_width)
    
    wb.save(excel_file)

# 直接執行
if __name__ == "__main__":
    parse_patch_to_excel("diff.patch", "程式修改清單.xlsx")
