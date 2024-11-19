import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Button
import os


def select_file():
    """使用文件对话框选择Excel文件"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showerror("错误", "未选择文件，程序终止。")
        exit()
    return file_path


def load_data(file_path):
    """加载Excel文件数据"""
    data = pd.read_excel(file_path)
    return data


def select_columns_gui(columns, title="选择列", select_mode="multiple"):
    """弹出窗口让用户点选需要的列"""
    def on_submit():
        selected_indices = listbox.curselection()
        selected_columns = [columns[i] for i in selected_indices]
        if not selected_columns:
            messagebox.showerror("错误", "未选择任何列，程序终止。")
            root.quit()
        else:
            root.selected_columns = selected_columns
            root.quit()

    # 创建GUI窗口
    root = tk.Tk()
    root.title(title)
    root.geometry("400x300")

    # 列表框 + 滚动条
    scrollbar = Scrollbar(root)
    scrollbar.pack(side="right", fill="y")

    listbox = Listbox(root, selectmode=select_mode, yscrollcommand=scrollbar.set, width=50, height=15)
    for col in columns:
        listbox.insert("end", col)
    listbox.pack(padx=10, pady=10)

    scrollbar.config(command=listbox.yview)

    # 确认按钮
    button = Button(root, text="确认选择", command=on_submit)
    button.pack(pady=10)

    root.mainloop()

    # 返回用户选择的列
    return getattr(root, "selected_columns", [])


def transform_data(data, sample_columns, compound_column):
    """
    将原始数据转换为目标格式：
    样品列填充样本名称，化合物作为列名，浓度作为值。
    """
    # 将样本列转置为行
    melted_data = data[sample_columns].copy()
    melted_data = melted_data.T  # 转置，样本列名变为行
    melted_data.columns = data[compound_column].values  # 使用化合物列值作为新列名
    melted_data.insert(0, "样品", sample_columns)  # 添加样品列

    return melted_data


def save_transformed_file(data, original_file_path):
    """保存转换后的文件到原始文件路径"""
    # 获取原始文件名和目录
    base_dir = os.path.dirname(original_file_path)
    base_name = os.path.splitext(os.path.basename(original_file_path))[0]

    # 构建保存路径
    save_path = os.path.join(base_dir, f"{base_name}_转换后.xlsx")
    data.to_excel(save_path, index=False)
    print(f"转换后的数据已保存到: {save_path}")
    return save_path


def main():
    # 文件选择与加载
    print("请选择用于转换的Excel文件...")
    file_path = select_file()
    data = load_data(file_path)

    # 选择样本列
    print("\n请选择样本列（多选，代表不同样本浓度列）：")
    sample_columns = select_columns_gui(data.columns.tolist(), title="选择样本列", select_mode="multiple")

    # 选择化合物列
    print("\n请选择化合物列（单选，代表化合物信息）：")
    compound_column = select_columns_gui(data.columns.tolist(), title="选择化合物列", select_mode="single")[0]

    # 数据转换
    transformed_data = transform_data(data, sample_columns, compound_column)

    # 保存转换后的文件
    save_transformed_file(transformed_data, file_path)


if __name__ == "__main__":
    main()
