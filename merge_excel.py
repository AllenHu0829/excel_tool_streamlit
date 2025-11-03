import pandas as pd
import os
from pathlib import Path

def merge_excel_files(data_dir, output_file):
    """
    合并 data 文件夹下的所有 Excel 文件
    
    参数:
        data_dir: 包含 Excel 文件的目录路径
        output_file: 输出合并后的 Excel 文件路径
    """
    try:
        # 获取所有 Excel 文件
        excel_files = []
        for file in os.listdir(data_dir):
            if file.endswith('.xlsx') or file.endswith('.xls'):
                excel_files.append(os.path.join(data_dir, file))
        
        if not excel_files:
            print("data 文件夹下没有找到 Excel 文件")
            return
        
        print(f"找到 {len(excel_files)} 个 Excel 文件")
        
        # 存储所有数据框
        dataframes = []
        
        # 读取每个 Excel 文件
        for idx, file_path in enumerate(excel_files, 1):
            try:
                # 读取 Excel 文件，使用第一行作为列名
                df = pd.read_excel(file_path, header=0)
                
                # 添加源文件名列，用于追踪数据来源
                if '源文件' not in df.columns:
                    df.insert(0, '源文件', os.path.basename(file_path))
                
                dataframes.append(df)
                print(f"已读取 [{idx}/{len(excel_files)}]: {os.path.basename(file_path)} - {df.shape[0]} 行, {df.shape[1]} 列")
                
            except Exception as e:
                print(f"读取文件失败 {os.path.basename(file_path)}: {str(e)}")
                continue
        
        if not dataframes:
            print("没有成功读取任何文件")
            return
        
        # 合并所有数据框
        # 使用 concat 时会自动对齐列名，相同的列会合并，不同的列会保留
        print("\n正在合并数据...")
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
        
        # 统计信息
        print(f"\n合并完成!")
        print(f"总行数: {len(merged_df)}")
        print(f"总列数: {len(merged_df.columns)}")
        print(f"列名: {list(merged_df.columns)}")
        
        # 保存合并后的文件
        print(f"\n正在保存到: {output_file}")
        merged_df.to_excel(output_file, index=False, engine='openpyxl')
        print("保存完成!")
        
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(current_dir, "data")
    output_file = os.path.join(current_dir, "合并后的Excel.xlsx")
    
    merge_excel_files(data_dir, output_file)


