import pandas as pd
import numpy as np
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar, Button, simpledialog
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os


# 设置中文字体（解决中文显示问题）
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


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
    base_dir = os.path.dirname(original_file_path)
    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    save_path = os.path.join(base_dir, f"{base_name}_转换后.xlsx")
    data.to_excel(save_path, index=False)
    print(f"转换后的数据已保存到: {save_path}")
    return save_path


def group_samples(sample_names):
    """
    可视化界面让用户为样本分组
    :param sample_names: 样本名称列表
    :return: 样本分组字典
    """
    groups = {}

    def on_submit():
        for i, sample_name in enumerate(sample_names):
            group = entry_widgets[i].get()
            groups[sample_name] = group
        root.quit()

    # 创建GUI窗口
    root = tk.Tk()
    root.title("样本分组")
    root.geometry("400x600")

    # 滚动条
    scrollbar = Scrollbar(root)
    scrollbar.pack(side="right", fill="y")

    # 样本列表和分组输入框
    canvas = tk.Canvas(root, yscrollcommand=scrollbar.set)
    frame = tk.Frame(canvas)
    scrollbar.config(command=canvas.yview)

    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    entry_widgets = []
    for i, sample_name in enumerate(sample_names):
        tk.Label(frame, text=sample_name, width=30, anchor="w").grid(row=i, column=0, padx=5, pady=5)
        entry = tk.Entry(frame, width=15)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entry_widgets.append(entry)

    # 确认按钮
    Button(root, text="确认分组", command=on_submit).pack(pady=10)

    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    root.mainloop()

    return groups


def pca_analysis(data, n_components, groups=None):
    """主成分分析 (PCA)"""
    # 标准化数据
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(data.iloc[:, 1:])

    # PCA分析
    pca = PCA(n_components=n_components)
    principal_components = pca.fit_transform(X_scaled)

    # 将主成分添加到DataFrame中
    pca_result = pd.DataFrame(principal_components, index=data["样品"],
                              columns=[f"主成分 {i+1}" for i in range(n_components)])

    # 方差解释比例
    explained_variance = pca.explained_variance_ratio_
    print("\n主成分的方差解释比例：")
    for i, variance in enumerate(explained_variance, start=1):
        print(f"主成分 {i}: {variance:.2%}")

    # 可视化
    if n_components >= 2:
        plt.figure(figsize=(10, 8))
        colors = None
        if groups:
            unique_groups = list(set(groups.values()))
            group_colors = {group: plt.cm.tab10(i) for i, group in enumerate(unique_groups)}
            colors = [group_colors[groups[sample]] for sample in data["样品"]]

        plt.scatter(principal_components[:, 0], principal_components[:, 1], alpha=0.7, edgecolor='k', c=colors)
        for i, sample in enumerate(data["样品"]):
            plt.annotate(sample, (principal_components[i, 0], principal_components[i, 1]), fontsize=8)
        plt.title('PCA 主成分分析 (2D)')
        plt.xlabel('主成分 1')
        plt.ylabel('主成分 2')
        plt.grid()
        plt.show()

    return pca_result


def main():
    # 文件选择与加载
    print("请选择用于PCA分析的Excel文件...")
    file_path = select_file()
    data = load_data(file_path)

    # 选择样本列
    print("\n请选择样本列（多选，代表不同样本浓度列）：")
    sample_columns = select_columns_gui(data.columns.tolist(), title="选择样本列", select_mode="multiple")

    # 选择化合物列
    print("\n请选择化合物列（单选，代表化合物信息）：")
    compound_column = select_columns_gui(data.columns.tolist(), title="选择化合物列", select_mode="single")[0]

    # 数据格式转换
    transformed_data = transform_data(data, sample_columns, compound_column)
    save_transformed_file(transformed_data, file_path)

    # 样本分组
    print("\n开始样本分组...")
    sample_groups = group_samples(transformed_data["样品"].tolist())

    # 用户选择主成分数量
    n_components = simpledialog.askinteger("主成分数", "请输入主成分数量 n_components（建议2或3）：", initialvalue=2)

    # PCA分析
    pca_result = pca_analysis(transformed_data, n_components, groups=sample_groups)


if __name__ == "__main__":
    main()
