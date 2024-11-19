import pandas as pd
import numpy as np
from sklearn.cross_decomposition import PLSRegression
from sklearn.preprocessing import StandardScaler, LabelEncoder
import matplotlib.pyplot as plt
import os
from tkinter import Tk, filedialog, Button, Label, Entry, Text, Toplevel, Listbox, MULTIPLE, SINGLE, END, messagebox

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


class OPLSDA_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OPLS-DA 多文件处理工具")

        # 文件选择
        self.label_file = Label(root, text="选择文件：")
        self.label_file.grid(row=0, column=0, padx=5, pady=5, sticky='w')

        self.file_button = Button(root, text="选择文件", command=self.load_file)
        self.file_button.grid(row=0, column=1, padx=5, pady=5)

        self.file_text = Text(root, height=5, width=60)
        self.file_text.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        # 参数输入
        self.label_vip = Label(root, text="VIP 值阈值：")
        self.label_vip.grid(row=2, column=0, padx=5, pady=5, sticky='w')

        self.entry_vip = Entry(root)
        self.entry_vip.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_vip.insert(0, "1.0")  # 默认值

        self.label_components = Label(root, text="主成分数量 (n_components)：")
        self.label_components.grid(row=3, column=0, padx=5, pady=5, sticky='w')

        self.entry_components = Entry(root)
        self.entry_components.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        self.entry_components.insert(0, "2")  # 默认值

        # 新增功能按钮
        self.select_samples_button = Button(root, text="选择样品列", command=self.select_samples)
        self.select_samples_button.grid(row=4, column=0, padx=5, pady=5)

        self.select_compound_button = Button(root, text="选择化合物种类列", command=self.select_compound)
        self.select_compound_button.grid(row=4, column=1, padx=5, pady=5)

        self.add_groups_button = Button(root, text="手动添加分组", command=self.add_groups)
        self.add_groups_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        # 执行按钮
        self.run_button = Button(root, text="开始分析", command=self.run_analysis)
        self.run_button.grid(row=6, column=0, padx=5, pady=10)

        self.quit_button = Button(root, text="退出", command=root.quit)
        self.quit_button.grid(row=6, column=1, padx=5, pady=10)

        # 数据选择变量
        self.selected_samples = None
        self.selected_compound = None
        self.sample_groups = {}

    def load_file(self):
        file_paths = filedialog.askopenfilenames(title="选择Excel文件", filetypes=(("Excel文件", "*.xlsx"), ("所有文件", "*.*")))
        self.file_text.delete(1.0, END)
        self.file_text.insert(END, "\n".join(file_paths))

    def select_samples(self):
        file_path = self.file_text.get(1.0, END).strip().split("\n")[0]
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        data = pd.read_excel(file_path)
        columns = data.columns.tolist()

        select_window = Toplevel(self.root)
        select_window.title("选择样品列")

        Label(select_window, text="多选样品列：").pack(pady=5)
        sample_listbox = Listbox(select_window, selectmode=MULTIPLE, height=15, width=50)
        sample_listbox.pack(padx=10, pady=10)

        for col in columns:
            sample_listbox.insert(END, col)

        def confirm_selection():
            selected_indices = sample_listbox.curselection()
            self.selected_samples = [columns[i] for i in selected_indices]
            if self.selected_samples:
                messagebox.showinfo("成功", f"已选择样品列：{', '.join(self.selected_samples)}")
            select_window.destroy()

        Button(select_window, text="确定", command=confirm_selection).pack(pady=10)

    def select_compound(self):
        file_path = self.file_text.get(1.0, END).strip().split("\n")[0]
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        data = pd.read_excel(file_path)
        columns = data.columns.tolist()

        select_window = Toplevel(self.root)
        select_window.title("选择化合物种类列")

        Label(select_window, text="单选化合物种类列：").pack(pady=5)
        compound_listbox = Listbox(select_window, selectmode=SINGLE, height=15, width=50)
        compound_listbox.pack(padx=10, pady=10)

        for col in columns:
            compound_listbox.insert(END, col)

        def confirm_selection():
            selected_index = compound_listbox.curselection()
            if selected_index:
                self.selected_compound = columns[selected_index[0]]
                messagebox.showinfo("成功", f"已选择化合物种类列：{self.selected_compound}")
            select_window.destroy()

        Button(select_window, text="确定", command=confirm_selection).pack(pady=10)
    def add_groups(self):
        if not self.selected_samples:
            messagebox.showwarning("警告", "请先选择样品列！")
            return

        group_window = Toplevel(self.root)
        group_window.title("样本分组")

        Label(group_window, text="请为每个样品分配分组：").pack(pady=5)

        # 创建一个字典，用于存储样品对应的输入框
        entry_fields = {}

        # 动态生成输入框
        for sample in self.selected_samples:
            frame = Label(group_window)
            frame.pack(padx=10, pady=2, anchor='w')

            sample_label = Label(frame, text=f"{sample}：", width=20, anchor='w')
            sample_label.pack(side="left")

            entry = Entry(frame, width=10)
            entry.pack(side="left")
            entry_fields[sample] = entry

        def confirm_groups():
            # 从输入框中获取每个样品的分组信息
            self.sample_groups = {}
            for sample, entry in entry_fields.items():
                group = entry.get().strip()
                if not group:
                    messagebox.showwarning("警告", f"样品 {sample} 的分组不能为空！")
                    return
                self.sample_groups[sample] = group

            messagebox.showinfo("成功", f"分组已设置：{self.sample_groups}")
            group_window.destroy()

        Button(group_window, text="确定", command=confirm_groups).pack(pady=10)
    def run_analysis(self):
        # 验证是否完成设置
        if not self.selected_samples or not self.selected_compound or not self.sample_groups:
            messagebox.showwarning("警告", "请完成所有设置！")
            return

        files = self.file_text.get(1.0, END).strip().split("\n")
        if not files:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        try:
            vip_threshold = float(self.entry_vip.get())
            n_components = int(self.entry_components.get())
        except ValueError:
            messagebox.showerror("错误", "VIP 阈值或主成分数量输入有误！")
            return

        for file_path in files:
            try:
                self.process_file(file_path, vip_threshold, n_components)
            except Exception as e:
                messagebox.showerror("错误", f"处理文件 {file_path} 时出错：{e}")
                continue

        messagebox.showinfo("完成", "所有文件处理完成！")

    def process_file(self, file_path, vip_threshold, n_components):
        # 加载数据
        data = pd.read_excel(file_path)
        print(f"正在处理文件：{file_path}")

        # 转换数据格式
        reshaped_data = self.reshape_data(data, self.selected_compound, self.selected_samples, file_path)
        reshaped_data["分组"] = reshaped_data["样品"].map(self.sample_groups)

        # 执行 OPLS-DA 分析
        important_compounds_df = self.opls_da_analysis(reshaped_data, vip_threshold, n_components)

        # 保存结果
        self.save_results(important_compounds_df, file_path)

    def reshape_data(self, data, cas_column, sample_columns, file_path):
        melted_data = data.melt(id_vars=[cas_column], value_vars=sample_columns, var_name="样品", value_name="浓度")
        reshaped_data = melted_data.pivot(index="样品", columns=cas_column, values="浓度").reset_index()

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        save_path = os.path.join(os.path.dirname(file_path), f"{base_name}_opls-da转换后格式.xlsx")
        reshaped_data.to_excel(save_path, index=False)
        print(f"转换后的数据已保存至: {save_path}")

        return reshaped_data

    def opls_da_analysis(self, reshaped_data, vip_threshold, n_components):
        X = reshaped_data.drop(columns=["样品", "分组"]).values
        y = LabelEncoder().fit_transform(reshaped_data["分组"])

        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)

        pls = PLSRegression(n_components=n_components)
        pls.fit(X_scaled, y)

        T = pls.x_scores_
        P = pls.x_loadings_
        num_features = X_scaled.shape[1]
        weights_squared = np.square(pls.x_weights_)
        explained_variance = np.var(T, axis=0)
        vip_scores = np.sqrt(num_features * np.sum(weights_squared * explained_variance / explained_variance.sum(), axis=1))

        compound_names = reshaped_data.columns[1:-1]
        important_compounds = compound_names[vip_scores > vip_threshold]
        important_vips = vip_scores[vip_scores > vip_threshold]
        important_compounds_df = pd.DataFrame({'化合物名称': important_compounds, 'VIP 值': important_vips})

        # 绘制得分图
        plt.figure(figsize=(10, 6))
        scatter = plt.scatter(T[:, 0], T[:, 1], c=y, cmap='viridis', edgecolor='k', s=100)
        plt.title('OPLS-DA Score Plot')
        plt.xlabel('Component 1')
        plt.ylabel('Component 2')
        plt.colorbar(label='Group')
        plt.grid(True)

        # 为每个样品添加标注
        for i, sample_name in enumerate(reshaped_data["样品"]):
            plt.annotate(sample_name, (T[i, 0], T[i, 1]), fontsize=8, ha='right')

        plt.show()

        return important_compounds_df

    def save_results(self, important_compounds_df, file_path):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        save_path = os.path.join(os.path.dirname(file_path), f"{base_name}_opls-da分析.xlsx")
        important_compounds_df.to_excel(save_path, index=False)
        print(f"VIP 分析结果已保存到: {save_path}")


# 主程序入口
if __name__ == "__main__":
    root = Tk()
    app = OPLSDA_GUI(root)
    root.mainloop()
