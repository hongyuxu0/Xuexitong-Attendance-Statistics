import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from pathlib import Path
import datetime

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("学习通签到统计工具")
        self.root.geometry("900x700")

        # 全局变量
        self.file_paths = []
        self.summary_data = []

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 顶部说明
        ttk.Label(self.root, text="学习通签到统计工具", font=("微软雅黑", 16)).pack(pady=10)
        ttk.Label(self.root, text="支持CSV/Excel文件导入，自动统计签到次数").pack(pady=5)

        # 按钮区域
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10, fill=tk.X, padx=20)

        ttk.Button(button_frame, text="导入单个文件", command=self.import_single_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="批量导入文件", command=self.import_multiple_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导入文件夹", command=self.import_folder).pack(side=tk.LEFT, padx=5)
        # 新增：重置按钮
        ttk.Button(button_frame, text="重置所有", command=self.reset_all).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="处理选中文件", command=self.process_files).pack(side=tk.RIGHT, padx=5)
        # 修复：padx参数移到pack方法中（ttk.Button初始化不支持padx）
        ttk.Button(button_frame, text="生成汇总表", command=self.generate_summary_button).pack(side=tk.RIGHT, padx=10)

        # 文件列表显示
        file_frame = ttk.Frame(self.root)
        file_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=20)

        ttk.Label(file_frame, text="已导入文件列表：").pack(anchor=tk.W)
        self.file_listbox = tk.Listbox(file_frame, height=8)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 滚动条
        file_scroll = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=file_scroll.set)

        # 进度条
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=5)

        # 状态标签
        self.status_label = ttk.Label(self.root, text="等待导入文件...")
        self.status_label.pack(pady=5)

        # 日志区域
        log_frame = ttk.Frame(self.root)
        log_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=20)

        ttk.Label(log_frame, text="操作日志：").pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 清空日志按钮
        ttk.Button(log_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, pady=5)

    def log_message(self, message, level="INFO"):
        """添加日志信息"""
        # 生成时间戳
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # 格式化日志内容
        log_entry = f"[{timestamp}] [{level}] {message}\n"

        # 启用文本框，添加内容，然后禁用
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_entry)
        # 自动滚动到最后
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_message("日志已清空", "INFO")

    def reset_all(self):
        """重置所有状态（重新统计）"""
        # 确认重置操作
        if messagebox.askyesno("确认重置", "是否确定重置所有数据？这将清空已导入的文件和统计数据！"):
            # 清空全局变量
            self.file_paths = []
            self.summary_data = []
            # 清空文件列表
            self.file_listbox.delete(0, tk.END)
            # 重置进度条
            self.progress['value'] = 0
            # 重置状态标签
            self.status_label.config(text="等待导入文件...")
            # 日志记录
            self.log_message("已重置所有数据：清空文件列表和汇总数据", "WARNING")
            messagebox.showinfo("重置完成", "所有数据已重置，可重新导入文件进行统计")

    def import_single_file(self):
        """导入单个文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        if file_path and file_path not in self.file_paths and not self.is_temp_file(file_path):
            self.file_paths.append(file_path)
            self.update_file_list()
            self.log_message(f"成功导入单个文件：{os.path.basename(file_path)}")
        elif self.is_temp_file(file_path):
            self.log_message(f"跳过临时文件：{os.path.basename(file_path)}", "WARNING")

    def import_multiple_files(self):
        """批量导入多个文件"""
        files = filedialog.askopenfilenames(
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        imported_count = 0
        skipped_count = 0
        for file in files:
            if file not in self.file_paths and not self.is_temp_file(file):
                self.file_paths.append(file)
                imported_count += 1
            elif self.is_temp_file(file):
                skipped_count += 1
        self.update_file_list()
        self.log_message(f"批量导入完成：成功导入{imported_count}个文件，跳过{skipped_count}个临时文件")

    def import_folder(self):
        """导入整个文件夹"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            imported_count = 0
            skipped_count = 0
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.endswith(('.xlsx', '.csv')):
                        file_path = os.path.join(root, file)
                        if file_path not in self.file_paths and not self.is_temp_file(file_path):
                            self.file_paths.append(file_path)
                            imported_count += 1
                        elif self.is_temp_file(file_path):
                            skipped_count += 1
            self.update_file_list()
            self.log_message(f"文件夹导入完成：从{folder_path}导入{imported_count}个文件，跳过{skipped_count}个临时文件")

    def is_temp_file(self, file_path):
        """判断是否是临时文件（包含~或隐藏文件）"""
        filename = os.path.basename(file_path)
        return filename.startswith('~$') or filename.startswith('.')

    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        for file in self.file_paths:
            self.file_listbox.insert(tk.END, os.path.basename(file))
        self.status_label.config(text=f"已导入 {len(self.file_paths)} 个文件")

    def process_files(self):
        """处理所有导入的文件（仅生成单个文件统计结果，不汇总）"""
        if not self.file_paths:
            messagebox.showwarning("警告", "请先导入文件！")
            self.log_message("处理文件失败：未导入任何文件", "ERROR")
            return

        self.progress['value'] = 0
        self.progress['maximum'] = len(self.file_paths)
        self.summary_data = []  # 重置汇总数据

        self.log_message(f"开始处理{len(self.file_paths)}个文件...")

        success_count = 0
        fail_count = 0

        for i, file_path in enumerate(self.file_paths):
            try:
                self.log_message(f"正在处理文件：{os.path.basename(file_path)}")
                self.process_single_file(file_path)
                self.progress['value'] = i + 1
                self.root.update_idletasks()
                success_count += 1
                self.log_message(f"文件处理成功：{os.path.basename(file_path)}")
            except PermissionError:
                fail_count += 1
                error_msg = f"权限错误：无法访问文件{os.path.basename(file_path)}，请确保文件未被打开且有读写权限"
                messagebox.showerror("权限错误", error_msg)
                self.log_message(error_msg, "ERROR")
            except KeyError as e:
                fail_count += 1
                error_msg = f"列名错误：文件{os.path.basename(file_path)}缺少关键列{str(e)}"
                messagebox.showerror("列名错误", error_msg)
                self.log_message(error_msg, "ERROR")
            except Exception as e:
                fail_count += 1
                error_msg = f"处理文件{os.path.basename(file_path)}时出错：{str(e)}"
                messagebox.showerror("错误", error_msg)
                self.log_message(error_msg, "ERROR")

        self.progress['value'] = 0
        self.status_label.config(text=f"文件处理完成：成功{success_count}个，失败{fail_count}个")
        self.log_message(f"文件处理结束：成功{success_count}个，失败{fail_count}个")

        if success_count > 0:
            messagebox.showinfo("完成",
                                f"文件处理完成！成功{success_count}个，失败{fail_count}个\n可点击'生成汇总表'按钮创建汇总文件")

    def find_header_row(self, df):
        """动态查找包含'签到状态'的表头行"""
        for i, row in df.iterrows():
            if '签到状态' in str(row.values):
                return i
        return None

    def process_single_file(self, file_path):
        """处理单个签到文件"""
        # 读取整个文件，用于动态检测表头
        if file_path.endswith('.xlsx'):
            full_df = pd.read_excel(file_path, header=None)
        elif file_path.endswith('.csv'):
            full_df = pd.read_csv(file_path, header=None)

        # 动态查找表头行
        header_row = self.find_header_row(full_df)
        if header_row is None:
            raise ValueError("未找到包含'签到状态'的表头行，请确认文件格式")

        # 用找到的表头行重新读取文件
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=header_row)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=header_row)

        # 清理列名中的空格和特殊字符
        df.columns = [col.strip() for col in df.columns]

        # 验证关键列是否存在
        required_columns = ['姓名', '学号/工号', '学校', '院系', '专业', '行政班级', '签到状态']
        for col in required_columns:
            if col not in df.columns:
                raise KeyError(col)

        # 生成统计列
        df['签到统计'] = df['签到状态'].apply(
            lambda x: 1 if str(x).strip() == '已签' else 0 if str(x).strip() == '未参与' else None
        )

        # 保存处理后的文件
        try:
            output_path = f"{os.path.splitext(file_path)[0]}_统计结果.xlsx"
            df.to_excel(output_path, index=False)
            self.log_message(f"统计文件已保存：{os.path.basename(output_path)}")
        except PermissionError:
            # 如果原路径无法写入，让用户选择保存位置
            self.log_message(f"原路径无写入权限，弹出保存对话框：{os.path.basename(file_path)}", "WARNING")
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"{os.path.splitext(os.path.basename(file_path))[0]}_统计结果.xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )
            if output_path:
                df.to_excel(output_path, index=False)
                self.log_message(f"统计文件已保存到指定位置：{os.path.basename(output_path)}")
            else:
                raise PermissionError("用户取消了文件保存操作")

        # 收集汇总数据（包含所有需要的字段）
        file_name = os.path.basename(file_path)
        for _, row in df.iterrows():
            self.summary_data.append({
                '姓名': row['姓名'],
                '学号/工号': row['学号/工号'],
                '学校': row['学校'],
                '院系': row['院系'],
                '专业': row['专业'],
                '行政班级': row['行政班级'],
                '文件名': file_name,
                '签到统计': row['签到统计']
            })

    def generate_summary_button(self):
        """独立的汇总表生成按钮处理函数"""
        if not self.summary_data:
            messagebox.showwarning("警告", "暂无汇总数据！请先处理文件")
            self.log_message("生成汇总表失败：暂无汇总数据", "WARNING")
            return

        self.log_message("开始生成汇总表...")

        try:
            summary_df = pd.DataFrame(self.summary_data)
            # 按姓名和学号分组，保留学校、院系、专业、行政班级信息并统计总签到次数
            final_summary = summary_df.groupby(
                ['姓名', '学号/工号', '学校', '院系', '专业', '行政班级']
            )['签到统计'].sum().reset_index()
            final_summary.rename(columns={'签到统计': '总签到次数'}, inplace=True)

            # 保存汇总表
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="学习通签到汇总表.xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )

            if save_path:
                final_summary.to_excel(save_path, index=False)
                self.log_message(f"汇总表生成成功：{os.path.basename(save_path)}")
                self.log_message(f"汇总数据：共统计{len(final_summary)}名学生的签到情况")
                messagebox.showinfo("成功", f"汇总表已保存！\n文件路径：{save_path}\n共统计{len(final_summary)}名学生")
            else:
                self.log_message("用户取消了汇总表保存操作", "WARNING")

        except Exception as e:
            error_msg = f"生成汇总表时出错：{str(e)}"
            messagebox.showerror("错误", error_msg)
            self.log_message(error_msg, "ERROR")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
