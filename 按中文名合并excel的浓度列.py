import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

def select_folder():
    """
    弹出文件夹选择对话框，让用户选择包含 Excel 文件的文件夹。
    返回选择的文件夹路径。
    """
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    folder_path = askdirectory(title="请选择包含 Excel 文件的文件夹")
    return folder_path

def merge_excel_files_in_folder(folder_path):
    """
    合并指定文件夹中的多个 Excel 文件，根据 "用户定义的谱库化合物" 进行去重处理，
    并保存排序后的结果为新的 Excel 文件。
    """
    # 获取指定文件夹内所有的 .xlsx 文件路径
    excel_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if
                   f.endswith('.xlsx') and not f.startswith('~$')]

    print("找到的 Excel 文件：", excel_files)  # 打印找到的文件列表

    if not excel_files:
        print("指定的文件夹中没有找到任何有效的 Excel 文件。")
        return

    # 定义包含基本化合物信息的列
    compound_info_columns = ["CAS 编号", "化合物名称", "用户定义的谱库化合物", "组分 RI", "谱库 RI", "谱库化合物描述"]

    dfs = []

    # 遍历每个 Excel 文件，加载数据并重命名浓度列
    for file_path in excel_files:
        print(f"正在处理文件：{file_path}")
        df = pd.read_excel(file_path)

        print("文件内容预览：")
        print(df.head())  # 打印文件的前几行，方便调试

        # 如果 "估计的浓度." 列存在，则重命名
        if "估计的浓度." in df.columns:
            df = df.rename(columns={"估计的浓度.": f"{os.path.splitext(os.path.basename(file_path))[0]}_浓度"})
        else:
            print(f"警告：文件 {file_path} 中缺少 '估计的浓度.' 列，跳过重命名。")

        dfs.append(df)

    # 确保 DataFrame 列表非空
    if not dfs:
        print("未加载到有效的数据文件。")
        return

    # 初始化合并的浓度数据
    first_file_name = os.path.splitext(os.path.basename(excel_files[0]))[0]
    first_concentration_column = f"{first_file_name}_浓度"

    if first_concentration_column not in dfs[0].columns:
        print(f"列 {first_concentration_column} 在第一个文件中不存在。请检查文件内容。")
        return

    concentration_merged = dfs[0][["用户定义的谱库化合物", first_concentration_column]]

    # 逐个合并其余文件的浓度数据
    for df, file_path in zip(dfs[1:], excel_files[1:]):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        concentration_column = f"{file_name}_浓度"

        # 检查当前文件的浓度列是否存在，避免 KeyError
        if concentration_column in df.columns:
            concentration_merged = pd.merge(concentration_merged,
                                            df[["用户定义的谱库化合物", concentration_column]],
                                            on="用户定义的谱库化合物", how="outer")

    concentration_merged = concentration_merged.fillna("--")

    # 合并化合物信息，并根据 "用户定义的谱库化合物" 去重处理
    compound_info_priority = pd.concat([df[compound_info_columns] for df in dfs], axis=0)
    compound_info_priority = (
        compound_info_priority.sort_values(by=["用户定义的谱库化合物"], na_position="last")
        .drop_duplicates(subset=["用户定义的谱库化合物"], keep="first")  # 根据 用户定义的谱库化合物 去重
    )

    # 最终合并化合物信息与浓度数据
    final_merged_df = pd.merge(compound_info_priority, concentration_merged, on="用户定义的谱库化合物", how="left")

    # 根据“组分 RI”列进行排序
    final_sorted_df = final_merged_df.sort_values(by="组分 RI").reset_index(drop=True)

    # 设置输出文件路径和文件名
    output_file = os.path.join(folder_path, "化合物合并处理数据_按RI排序_用户定义谱库化合物匹配.xlsx")
    final_sorted_df.to_excel(output_file, index=False)

    print(f"合并后的数据已成功保存到 {output_file}")

# 主程序运行入口
if __name__ == "__main__":
    folder_path = select_folder()  # 调用文件夹选择对话框
    if folder_path:
        merge_excel_files_in_folder(folder_path)
    else:
        print("未选择任何文件夹。")
