import pandas as pd
import os
import time
import subprocess
import json
import zipfile

class SplitAndPrint:
    def __init__(self, interval_seconds):
        # 读取配置文件
        with open('./config/bartender.json', 'r', encoding="utf-8") as f:
            config = json.load(f)
            # ==== 参数 ====
            self.bartender_path = config['bartender_path']
            self.btw_file = config['btw_file']  
            self.printer_name = config['printer_name']
            self.group_column = config['group_column']
            self.print_info = config['print_info']
            self.tag_name = config['tag_name']
        self.interval_seconds = interval_seconds

    def split(self, excel_file, sheet_name, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        # 删除输出目录下的所有文件
        for file_name in os.listdir(output_dir):    
            file_path = os.path.join(output_dir, file_name)
            if os.path.isfile(file_path) and (file_name.endswith('.xls') or file_name.endswith('.zip')):
                os.remove(file_path)
        # 读取指定 sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='xlrd')

        # 获取列名列表（即表头）
        column_names = df.columns.tolist()
        print (column_names)
        # 遍历每一行，按考点拆分
        current_group = []
        last_site = None
        # 读取第一个考点的标签信息 
        tag_info = df[self.tag_name].iloc[0]
        # 获取标签元数据信息
        meta_tag_info = tag_info.split(' ')[0]
        # 遍历 Excel 表格的每一行
        for _, row in df.iterrows():
            current_site = row[self.group_column]
            curr_tag = row[self.tag_name]
            # 遇到新的考点时，保存上一个考点数据
            if pd.notna(curr_tag) and meta_tag_info in curr_tag and curr_tag != tag_info:
                output_file = os.path.join(output_dir, f"{last_site}.xls")
                # 写入新的 Excel 文件
                with pd.ExcelWriter(output_file, engine='xlwt') as writer:
                    pd.DataFrame(current_group).to_excel(writer, index=False, sheet_name=sheet_name)
                print(f"已保存考点：{last_site} 的文件：{output_file}")
                tag_info = curr_tag
                current_group = []

            current_group.append(row)
            # 更新考点信息
            if not pd.isna(current_site):
                last_site = current_site
            elif '考点' in row[self.print_info]:
                last_site = row[self.print_info].split(':')[1].strip()
            else:
                last_site = None

        # 保存最后一个考点的数据
        if current_group:
            output_file = os.path.join(output_dir, f"{last_site}.xls")
                # 写入新的 Excel 文件
            with pd.ExcelWriter(output_file, engine='xlwt') as writer:
                pd.DataFrame(current_group).to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"已保存考点：{last_site} 的文件：{output_file}")
    
    def print_labels_with_xls_files(self):
        # 遍历输出目录中的所有 Excel 文件
        for file_name in os.listdir(self.output_dir):
            if file_name.endswith('.xls'):
                file_path = os.path.join(self.output_dir, file_name)
                # 调用 Bartender 打印标签
                command = f'{self.bartender_path} /F={self.btw_file} /PRN={self.printer_name} /P /DD'
                subprocess.run(command, shell=True, capture_output=True, text=True)
                print(f"已打印标签文件：{file_name} 到打印机：{self.printer_name}")
                time.sleep(self.interval_seconds)

    def zip_output_files(self, output_dir, extended_name='.xls'):
        with zipfile.ZipFile(f"{output_dir}/result.zip", 'w') as zipf:
            for file_name in os.listdir(output_dir):
                if file_name.endswith(extended_name):
                    file_path = os.path.join(output_dir, file_name)
                    zipf.write(file_path, arcname=file_name)
        print(f"已将输出文件打包为：" f"{output_dir}/result.zip")

