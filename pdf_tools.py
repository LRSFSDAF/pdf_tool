'''
Description: Merge and split PDFs
Author: Damocles_lin
Date: 2025-07-30 13:30:34
LastEditTime: 2025-07-30 14:39:10
LastEditors: Damocles_lin
'''
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io
import threading

class PDFToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF工具箱")
        self.root.geometry("1100x850")  # 修改窗口大小
        self.root.configure(bg="#f0f0f0")
        
        # 设置应用图标
        try:
            self.root.iconbitmap("pdf_icon.ico")
        except:
            pass
        
        # 定义分割文件名变量
        self.split_filename_var = tk.StringVar()
        self.split_filename_var.set("分割文件")
        
        # 当前加载的PDF
        self.current_pdf = None
        self.pages = []
        self.page_images = []
        self.selected_pages = set()
        
        # 设置全局样式
        self.setup_styles()
        
        # 创建标签页
        self.notebook = ttk.Notebook(root)
        
        # 创建标签页
        self.merge_tab = ttk.Frame(self.notebook, padding=10)
        self.split_tab = ttk.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.merge_tab, text="合并PDF")
        self.notebook.add(self.split_tab, text="分割PDF")
        
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)
        
        # 初始化标签页
        self.setup_merge_tab()
        self.setup_split_tab()
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief="sunken", padding=(10, 5))
        self.status_bar.pack(side="bottom", fill="x")
        self.update_status("就绪 - 选择操作开始")

    def setup_styles(self):
        """设置全局UI样式"""
        style = ttk.Style()
        
        # 设置按钮基本样式
        style.configure("TButton", 
                        padding=6, 
                        relief="flat",
                        font=("Arial", 9))
        
        # 设置主要操作按钮样式（合并/分割按钮），改为浅蓝底深蓝字，鼠标悬停更深色
        style.configure("Primary.TButton", 
                        background="#e3f0fc",  # 浅蓝底
                        foreground="#205080",  # 深蓝字
                        font=("Arial", 10, "bold"),
                        borderwidth=1)
        style.map("Primary.TButton", 
                 background=[("active", "#b3d8fa"), ("pressed", "#a0c8e8")],
                 foreground=[("active", "#103060")])
        
        # 设置普通按钮样式
        style.configure("Normal.TButton", 
                        background="#e0e0e0", 
                        foreground="black")
        style.map("Normal.TButton", 
                 background=[("active", "#d0d0d0"), ("pressed", "#c0c0c0")])
        
        # 设置标签页样式
        style.configure("TNotebook.Tab", 
                        padding=[10, 5], 
                        background="#e0e0e0")
        style.map("TNotebook.Tab", 
                 background=[("selected", "#4a86e8")])
        
        # 分割区域样式：浅灰底，深灰边框
        style.configure("Split.TLabelframe", background="#f8f8fa", bordercolor="#b0b0b0", borderwidth=2)
        style.configure("Split.TLabelframe.Label", background="#f8f8fa")

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def setup_merge_tab(self):
        # 合并标签页样式
        merge_frame = ttk.LabelFrame(self.merge_tab, text="PDF合并", padding=10)
        merge_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # 文件选择部分
        ttk.Label(merge_frame, text="选择要合并的PDF文件:").pack(anchor="w", pady=5)

        # 文件列表框
        list_frame = ttk.Frame(merge_frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.merge_listbox = tk.Listbox(
            list_frame,
            height=8,
            selectmode="extended",
            yscrollcommand=scrollbar.set,
            bg="white",
            relief="solid",
            borderwidth=1
        )
        self.merge_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.merge_listbox.yview)

        # 按钮框架
        btn_frame = ttk.Frame(merge_frame)
        btn_frame.pack(fill="x", pady=10)

        # 使用 Normal.TButton 样式
        ttk.Button(btn_frame, text="添加文件", command=self.add_merge_files, width=12, style="Normal.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected, width=12, style="Normal.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="上移", command=lambda: self.move_item(-1), width=8, style="Normal.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="下移", command=lambda: self.move_item(1), width=8, style="Normal.TButton").pack(side="left", padx=5)

        # 合并文件名输入（移动到添加文件按钮下方，合并PDF按钮上方）
        filename_frame = ttk.Frame(merge_frame)
        filename_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(filename_frame, text="合并文件名:").pack(side="left", padx=5)
        self.merge_filename_var = tk.StringVar()
        self.merge_filename_var.set("合并文件")
        filename_entry = ttk.Entry(filename_frame, textvariable=self.merge_filename_var, width=30)
        filename_entry.pack(side="left", padx=5, fill="x", expand=True)

        # 合并PDF按钮，样式与分割PDF一致
        ttk.Button(merge_frame, text="合并PDF", command=self.merge_pdfs, style="Primary.TButton").pack(fill="x", pady=10)

    def setup_split_tab(self):
        # 分割标签页样式
        split_frame = ttk.LabelFrame(self.split_tab, text="PDF分割", padding=10)
        split_frame.pack(fill="both", expand=True, padx=5, pady=5)
        split_frame.configure(style="Split.TLabelframe")

        # 文件选择部分
        file_frame = ttk.Frame(split_frame)
        file_frame.pack(fill="x", pady=5)
        
        ttk.Label(file_frame, text="选择要分割的PDF文件:").pack(anchor="w")
        
        self.split_file_var = tk.StringVar()
        entry = ttk.Entry(file_frame, textvariable=self.split_file_var, state="readonly")
        entry.pack(fill="x", pady=5)
        
        # 浏览文件按钮使用 Normal.TButton 样式
        ttk.Button(file_frame, text="浏览文件", command=self.select_split_file, width=12, style="Normal.TButton").pack(anchor="e", pady=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(split_frame, text="页面预览", padding=10)
        preview_frame.pack(fill="both", expand=True, pady=10)
        
        # 创建Canvas和Scrollbar
        self.canvas = tk.Canvas(preview_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # 分割选项
        options_frame = ttk.Frame(split_frame)
        options_frame.pack(fill="x", pady=10)
        
        # 文件名输入
        ttk.Label(options_frame, text="分割文件名:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        filename_entry = ttk.Entry(options_frame, textvariable=self.split_filename_var, width=30)
        filename_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        
        # 页面范围
        ttk.Label(options_frame, text="页面范围 (例如: 1-3,5,7-9):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.page_range_var = tk.StringVar()
        page_range_entry = ttk.Entry(options_frame, textvariable=self.page_range_var)
        page_range_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        
        # 提示信息
        ttk.Label(options_frame, text="提示: 在预览图上点击选择/取消页面", foreground="#666", font=("Arial", 9)).grid(
            row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        # 分割按钮使用 Primary.TButton 样式，固定在底部
        ttk.Button(split_frame, text="分割PDF", command=self.split_pdf, style="Primary.TButton").pack(side="bottom", fill="x", pady=10)
        
        # 配置网格列权重
        options_frame.columnconfigure(1, weight=1)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if files:
            for f in files:
                self.merge_listbox.insert("end", f)
            self.update_status(f"添加了 {len(files)} 个PDF文件")

    def remove_selected(self):
        selected = self.merge_listbox.curselection()
        if not selected:
            return
        for i in selected[::-1]:
            self.merge_listbox.delete(i)
        self.update_status(f"移除了 {len(selected)} 个文件")

    def move_item(self, direction):
        selected = self.merge_listbox.curselection()
        if not selected:
            return
        index = selected[0]
        if (direction < 0 and index == 0) or (direction > 0 and index == self.merge_listbox.size()-1):
            return
        
        text = self.merge_listbox.get(index)
        self.merge_listbox.delete(index)
        new_index = index + direction
        self.merge_listbox.insert(new_index, text)
        self.merge_listbox.select_set(new_index)

    def select_split_file(self):
        file = filedialog.askopenfilename(
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file:
            self.split_file_var.set(file)
            self.update_status(f"已选择分割文件: {os.path.basename(file)}")
            self.load_pdf_preview(file)

    def load_pdf_preview(self, file_path):
        self.update_status("加载PDF预览...")
        self.selected_pages = set()
        
        # 清除现有预览
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # 显示加载提示
        loading_label = ttk.Label(self.scrollable_frame, text="加载PDF预览中，请稍候...")
        loading_label.pack(pady=50)
        self.root.update()
        
        try:
            # 打开PDF
            doc = fitz.open(file_path)
            self.current_pdf = doc
            self.pages = []
            self.page_images = []
            
            # 清除加载提示
            loading_label.destroy()
            
            # 创建缩略图
            cols = 4  # 每行显示4个缩略图
            row_frame = None
            
            for page_num in range(len(doc)):
                if page_num % cols == 0:
                    row_frame = ttk.Frame(self.scrollable_frame)
                    row_frame.pack(fill="x", padx=5, pady=5)
                
                # 获取页面
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))  # 缩略图缩放
                img_data = pix.tobytes("ppm")
                
                # 转换为PIL图像
                img = Image.open(io.BytesIO(img_data))
                photo = ImageTk.PhotoImage(img)
                
                # 保存引用
                self.page_images.append(photo)
                
                # 创建带边框的框架
                frame = ttk.Frame(row_frame, relief="solid", borderwidth=1)
                frame.grid(row=0, column=page_num % cols, padx=5, pady=5)
                
                # 创建标签显示缩略图
                label = ttk.Label(frame, image=photo)
                label.image = photo  # 保持引用
                label.pack(padx=5, pady=5)
                
                # 页码标签
                page_label = ttk.Label(frame, text=f"第 {page_num+1} 页", padding=3)
                page_label.pack()
                
                # 绑定点击事件
                label.bind("<Button-1>", lambda e, p=page_num: self.toggle_page_selection(p))
                page_label.bind("<Button-1>", lambda e, p=page_num: self.toggle_page_selection(p))
                frame.bind("<Button-1>", lambda e, p=page_num: self.toggle_page_selection(p))
                
                self.pages.append(frame)
            
            self.update_status(f"已加载预览: {len(doc)} 页")
            doc.close()
            
        except Exception as e:
            loading_label.destroy()
            self.update_status(f"加载预览失败: {str(e)}")
            messagebox.showerror("错误", f"无法加载PDF预览:\n{str(e)}")

    def toggle_page_selection(self, page_num):
        if page_num in self.selected_pages:
            self.selected_pages.remove(page_num)
            self.pages[page_num].configure(relief="solid")  # 恢复普通边框
        else:
            self.selected_pages.add(page_num)
            self.pages[page_num].configure(relief="solid", borderwidth=3)  # 加粗边框表示选中
        
        # 更新页面范围显示
        if self.selected_pages:
            sorted_pages = sorted(self.selected_pages)
            ranges = []
            start = end = sorted_pages[0]
            
            for page in sorted_pages[1:]:
                if page == end + 1:
                    end = page
                else:
                    if start == end:
                        ranges.append(str(start+1))
                    else:
                        ranges.append(f"{start+1}-{end+1}")
                    start = end = page
            
            if start == end:
                ranges.append(str(start+1))
            else:
                ranges.append(f"{start+1}-{end+1}")
            
            self.page_range_var.set(",".join(ranges))
        else:
            self.page_range_var.set("")

    def merge_pdfs(self):
        files = self.merge_listbox.get(0, "end")
        if not files:
            messagebox.showwarning("警告", "请添加要合并的PDF文件")
            return

        # 新增：获取合并文件名
        filename = self.merge_filename_var.get().strip()
        if not filename:
            messagebox.showwarning("警告", "请输入合并文件名")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=filename,  # 新增：默认文件名
            filetypes=[("PDF文件", "*.pdf")]
        )
        if not output_path:
            return

        try:
            merger = fitz.open()
            for pdf in files:
                doc = fitz.open(pdf)
                merger.insert_pdf(doc)
                doc.close()
            merger.save(output_path)
            merger.close()
            self.update_status(f"合并完成: {os.path.basename(output_path)}")
            messagebox.showinfo("成功", f"PDF文件合并成功!\n保存位置: {output_path}")
        except Exception as e:
            self.update_status(f"合并失败: {str(e)}")
            messagebox.showerror("错误", f"合并过程中出错:\n{str(e)}")

    def split_pdf(self):
        input_pdf = self.split_file_var.get()
        if not input_pdf:
            messagebox.showwarning("警告", "请选择要分割的PDF文件")
            return
        
        # 获取自定义文件名
        filename = self.split_filename_var.get().strip()
        if not filename:
            messagebox.showwarning("警告", "请输入分割文件名")
            return
        
        # 获取页面范围
        page_range = self.page_range_var.get().strip()
        if not page_range:
            messagebox.showwarning("警告", "请选择要分割的页面")
            return
        
        try:
            # 解析页面范围
            pages = self.parse_page_range(page_range)
            if not pages:
                raise ValueError("无效的页面范围")
            
            # 打开PDF
            doc = fitz.open(input_pdf)
            total_pages = len(doc)
            
            # 验证页面是否有效
            if max(pages) > total_pages or min(pages) < 1:
                raise ValueError(f"页面范围超出有效范围 (1-{total_pages})")
            
            # 选择输出文件夹
            output_dir = filedialog.askdirectory(title="选择保存分割文件的文件夹")
            if not output_dir:
                return
            
            # 创建新文档
            new_doc = fitz.open()
            
            # 添加选中的页面
            for page_num in pages:
                new_doc.insert_pdf(doc, from_page=page_num-1, to_page=page_num-1)
            
            # 保存分割后的文件
            output_path = os.path.join(output_dir, f"{filename}.pdf")
            new_doc.save(output_path)
            new_doc.close()
            doc.close()
            
            self.update_status(f"分割完成: {filename}.pdf")
            messagebox.showinfo("成功", f"PDF文件分割成功!\n保存位置: {output_path}")
        except Exception as e:
            self.update_status(f"分割失败: {str(e)}")
            messagebox.showerror("错误", f"分割过程中出错:\n{str(e)}")

    def parse_page_range(self, range_str):
        """解析页面范围字符串 (如 '1-3,5,7-9')"""
        pages = []
        parts = range_str.split(',')
        
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                try:
                    start = int(start.strip())
                    end = int(end.strip())
                    pages.extend(range(start, end+1))
                except ValueError:
                    return []
            else:
                try:
                    pages.append(int(part.strip()))
                except ValueError:
                    return []
        
        # 去重并排序
        return sorted(set(pages))

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolsApp(root)
    root.mainloop()