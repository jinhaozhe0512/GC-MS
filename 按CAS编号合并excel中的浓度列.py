import pandas as pd
import os


def merge_excel_files_in_folder(folder_path):
    # 获取指定文件夹内所有的 .xlsx 文件路径，排除以 ~$ 开头的临时文件
    excel_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if
                   f.endswith('.xlsx') and not f.startswith('~$')]

    if not excel_files:
        print("指定的文件夹中没有找到任何有效的 Excel 文件。")
        return

    # 定义包含基本化合物信息的列
    compound_info_columns = ["CAS 编号", "化合物名称", "用户定义的谱库化合物", "组分 RI", "谱库 RI", "谱库化合物描述"]

    # 用于存放数据的列表
    dfs = []

    # 遍历每个 Excel 文件，加载数据并重命名浓度列
    for file_path in excel_files:
        # 获取文件名（不带扩展名）作为列的前缀
        file_name = os.path.splitext(os.path.basename(file_path))[0]

        df = pd.read_excel(file_path)

        # 如果 "估计的浓度." 列存在，则重命名；否则跳过该列
        if "估计的浓度." in df.columns:
            df = df.rename(columns={"估计的浓度.": f"{file_name}_浓度"})

        # 过滤掉 CAS 编号 为 "38818-55-2" 的记录
        df = df[df["CAS 编号"] != "38818-55-2"]

        # 将 DataFrame 添加到列表
        dfs.append(df)

    # 检查第一个文件的浓度列名是否存在
    first_file_name = os.path.splitext(os.path.basename(excel_files[0]))[0]
    first_concentration_column = f"{first_file_name}_浓度"

    if first_concentration_column not in dfs[0].columns:
        print(f"列 {first_concentration_column} 在第一个文件中不存在。请检查文件内容。")
        return

    # 初始化合并的浓度数据
    concentration_merged = dfs[0][["CAS 编号", first_concentration_column]]

    # 逐个合并其余文件的浓度数据
    for df, file_path in zip(dfs[1:], excel_files[1:]):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        concentration_column = f"{file_name}_浓度"

        # 检查当前文件的浓度列是否存在，避免 KeyError
        if concentration_column in df.columns:
            concentration_merged = pd.merge(concentration_merged,
                                            df[["CAS 编号", concentration_column]],
                                            on="CAS 编号", how="outer")

    concentration_merged = concentration_merged.fillna("--")

    # 最终合并化合物信息与浓度数据，按 CAS 编号 和 用户定义的谱库化合物 去重
    compound_info_priority = pd.concat([df[compound_info_columns] for df in dfs],
                                       axis=0).drop_duplicates(subset=["CAS 编号", "用户定义的谱库化合物"]).reset_index(
        drop=True)

    final_merged_df = pd.merge(compound_info_priority, concentration_merged, on="CAS 编号", how="left")

    # 根据“组分 RI”列进行排序
    final_sorted_df = final_merged_df.sort_values(by="组分 RI").reset_index(drop=True)

    # 设置输出文件路径和文件名
    output_file = os.path.join(folder_path, "化合物合并处理数据_按RI排序_剔除巨豆三烯酮.xlsx")
    final_sorted_df.to_excel(output_file, index=False)

    print(f"合并后的数据已成功保存到 {output_file}")


# 使用示例：指定文件夹路径
folder_path = input("请输入包含 Excel 文件的文件夹路径：")
merge_excel_files_in_folder(folder_path)