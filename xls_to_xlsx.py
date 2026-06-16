import os
from pathlib import Path
import pandas as pd

def batch_convert_xls_to_xlsx_xlrd(search_dir):
    base_dir = Path(search_dir)
    # 查找所有 .xls 文件
    xls_files = [f for f in base_dir.glob("*.xls") if f.suffix == '.xls']
    
    if not xls_files:
        print(f"未在目录 '{search_dir}' 下找到任何 .xls 文件。")
        return

    print(f"🚀 找到 {len(xls_files)} 个 .xls 文件，开始使用 xlrd 引擎转换...")
    
    success_count = 0
    fail_count = 0

    for xls_path in xls_files:
        xlsx_path = xls_path.with_suffix('.xlsx')
        
        if xlsx_path.exists():
            print(f"⚠️ 跳过：{xlsx_path.name} 已存在。")
            continue
            
        try:
            print(f"正在转换: {xls_path.name} -> {xlsx_path.name}")
            
            # 💡 关键修改点：将 engine 改为 'xlrd'
            with pd.ExcelFile(xls_path, engine='xlrd') as excel_file:
                sheet_data = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            
            # 写入新的 xlsx
            with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
                for sheet_name, df in sheet_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            success_count += 1
            
        except Exception as e:
            print(f"❌ 转换失败: {xls_path.name}，原因: {str(e)}")
            fail_count += 1

    print(f"\n任务结束 -> 成功: {success_count} 个 | 失败: {fail_count} 个")

if __name__ == "__main__":
    batch_convert_xls_to_xlsx_xlrd(".")
