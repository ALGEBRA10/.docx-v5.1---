#!/usr/bin/env python3
"""
📄 .docx文件名分类处理器 v5.0 - 完整无控制台版
保存为 .pyw 双击运行，CMD窗口自动隐藏
"""
import ctypes
try:
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
except:
    pass

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


class DocXClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 .docx文件名分类处理器 v5.0（Ctrl+Z/Y/X/C/V全支持）")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)

        self.rows = []
        self.file_names = []
        self.original_file_names = []

        self.setup_ui()
        self.bind_shortcuts()
        self.status_var.set("✅ 启动成功！快捷键：Ctrl+Z撤销 Ctrl+Y恢复 Enter分析 Ctrl+E导出")

    def setup_ui(self):
        self.status_var = tk.StringVar()
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

        control_frame = ttk.LabelFrame(self.root, text="配置", padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(control_frame, text="📁 目录路径：").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.path_var = tk.StringVar(value=os.getcwd())
        path_entry = ttk.Entry(control_frame, textvariable=self.path_var, width=50)
        path_entry.grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)

        ttk.Button(control_frame, text="📂 浏览", command=self.browse_directory).grid(row=0, column=2, padx=5)

        ttk.Label(control_frame, text="✂️ 分隔符：").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.delimiter_var = tk.StringVar(value="-")
        ttk.Entry(control_frame, textvariable=self.delimiter_var, width=10).grid(row=1, column=1, sticky=tk.W, padx=5)

        ttk.Button(control_frame, text="▶️ 开始分析", command=self.start_analysis).grid(row=1, column=2, padx=5, pady=5)

        control_frame.columnconfigure(1, weight=1)

        replace_frame = ttk.LabelFrame(self.root, text="🔄 文件名替换（只影响显示和导出）", padding=10)
        replace_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(replace_frame, text="替换：").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.replace_from_var = tk.StringVar()
        ttk.Entry(replace_frame, textvariable=self.replace_from_var, width=15).grid(row=0, column=1, padx=5)

        ttk.Label(replace_frame, text="→").grid(row=0, column=2, padx=5)

        self.replace_to_var = tk.StringVar()
        ttk.Entry(replace_frame, textvariable=self.replace_to_var, width=15).grid(row=0, column=3, padx=5)

        ttk.Button(replace_frame, text="✨ 执行替换", command=self.apply_replace).grid(row=0, column=4, padx=10)

        ttk.Button(replace_frame, text="↺ 重置文件名", command=self.reset_file_names).grid(row=0, column=5, padx=5)

        output_frame = ttk.LabelFrame(self.root, text="分类结果（逐行一一对应）", padding=10)
        output_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        file_frame = ttk.Frame(output_frame)
        file_frame.grid(row=0, column=0, rowspan=2, sticky=tk.NSEW, padx=5, pady=5)
        file_frame.rowconfigure(1, weight=1)
        file_frame.columnconfigure(0, weight=1)

        ttk.Label(file_frame, text="📂 文件列表（第 N 行 = 第 N 个文件，可自由编辑）").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.file_list_widget = scrolledtext.ScrolledText(file_frame, height=20, width=38, font=("Consolas", 10), wrap=tk.NONE, undo=True)
        self.file_list_widget.grid(row=1, column=0, sticky=tk.NSEW)

        ttk.Button(file_frame, text="💾 导出文件列表", command=lambda: self.export_single_field("文件列表")).grid(row=2, column=0, pady=(5, 0), sticky=tk.E)

        self.text_widgets = {}
        field_layout = [("字段1", 0, 1), ("字段2", 0, 2), ("字段3", 1, 1), ("字段4", 1, 2)]

        for label, row, col in field_layout:
            frame = ttk.Frame(output_frame)
            frame.grid(row=row, column=col, sticky=tk.NSEW, padx=5, pady=5)
            frame.rowconfigure(1, weight=1)
            frame.columnconfigure(0, weight=1)

            ttk.Label(frame, text=f"📌 {label}（第 N 行对应第 N 个文件）").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

            text_area = scrolledtext.ScrolledText(frame, height=20, width=25, font=("Consolas", 10), wrap=tk.NONE)
            text_area.grid(row=1, column=0, sticky=tk.NSEW)
            text_area.config(state=tk.DISABLED)
            self.text_widgets[label] = text_area

            ttk.Button(frame, text=f"💾 导出 {label}", command=lambda l=label: self.export_single_field(l)).grid(row=2, column=0, pady=(5, 0), sticky=tk.E)

        output_frame.columnconfigure(0, weight=2)
        output_frame.columnconfigure(1, weight=1)
        output_frame.columnconfigure(2, weight=1)
        output_frame.rowconfigure(0, weight=1)
        output_frame.rowconfigure(1, weight=1)

        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(bottom_frame, text="🗑️ 清空全部", command=self.clear_all).pack(side=tk.LEFT)
        ttk.Button(bottom_frame, text="📦 全部导出为 result.txt", command=self.export_all_fields).pack(side=tk.RIGHT)

    def bind_shortcuts(self):
        w = self.file_list_widget
        w.bind('<Control-a>', self.select_all_text)
        w.bind('<Return>', lambda e: (self.start_analysis(), "break"))
        w.bind('<Delete>', self.delete_current_line)
        w.bind('<Control-z>', self.undo_text)
        w.bind('<Control-y>', self.redo_text)
        self.root.bind('<F5>', lambda e: self.start_analysis())
        self.root.bind('<Control-e>', lambda e: self.export_all_fields())
        self.root.bind('<Control-r>', lambda e: self.apply_replace())

    def select_all_text(self, event=None):
        self.file_list_widget.focus_set()
        self.file_list_widget.tag_add(tk.SEL, "1.0", tk.END)
        return "break"

    def undo_text(self, event=None):
        try:
            self.file_list_widget.edit_undo()
            self.status_var.set("↶ 已撤销")
        except:
            pass
        return "break"

    def redo_text(self, event=None):
        try:
            self.file_list_widget.edit_redo()
            self.status_var.set("↷ 已恢复")
        except:
            pass
        return "break"

    def delete_current_line(self, event=None):
        try:
            line_num = int(self.file_list_widget.index(tk.INSERT).split('.')[0])
            self.file_list_widget.delete(f"{line_num}.0", f"{line_num}.end")
            self.status_var.set(f"🗑️ 删除第{line_num}行")
        except:
            pass
        return "break"

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.path_var.get())
        if directory:
            self.path_var.set(directory)

    def find_docx_files(self, directory):
        if not os.path.exists(directory) or not os.path.isdir(directory):
            return None, "路径错误"
        files = []
        try:
            for entry in os.listdir(directory):
                full_path = os.path.join(directory, entry)
                if os.path.isfile(full_path) and entry.lower().endswith('.docx') and not entry.startswith('~$'):
                    mtime = os.path.getmtime(full_path)
                    files.append((entry, mtime))
            files.sort(key=lambda x: x[1])
            names = [os.path.splitext(entry)[0] for entry, _ in files]
            return names, f"找到{len(names)}个文件"
        except Exception as e:
            return None, str(e)

    def split_names_to_rows(self, names, delimiter):
        rows = []
        for name in names:
            parts = [p.strip() for p in name.split(delimiter) if p.strip()]
            row = parts[:4] + [""] * (4 - len(parts))
            rows.append(row)
        return rows

    def get_names_from_textbox(self):
        content = self.file_list_widget.get("1.0", tk.END)
        lines = [line.strip() for line in content.splitlines() if line.strip()]
        names = []
        for line in lines:
            base = os.path.basename(line)
            if base.lower().endswith(".docx"):
                base = os.path.splitext(base)[0]
            names.append(base)
        return names

    def start_analysis(self):
        try:
            delimiter = self.delimiter_var.get().strip()
            if not delimiter:
                messagebox.showwarning("警告", "分隔符不能为空")
                return

            names = self.get_names_from_textbox()
            if not names:
                directory = self.path_var.get()
                if not os.path.exists(directory):
                    messagebox.showerror("错误", "目录不存在")
                    return
                names, msg = self.find_docx_files(directory)
                if names is None:
                    messagebox.showerror("错误", msg)
                    return
                if not names:
                    messagebox.showinfo("提示", "未找到docx文件")
                    return

            self.original_file_names = names.copy()
            self.file_names = names.copy()
            self.rows = self.split_names_to_rows(names, delimiter)

            self.file_list_widget.delete("1.0", tk.END)
            for name in self.file_names:
                self.file_list_widget.insert(tk.END, name + "\n")

            self.display_results()
            self.status_var.set(f"✅ 完成！{len(names)}个文件")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.status_var.set("❌ 分析失败")

    def apply_replace(self):
        if not self.file_names:
            messagebox.showwarning("提示", "请先分析")
            return
        from_text = self.replace_from_var.get().strip()
        if not from_text:
            return
        count = 0
        for i, name in enumerate(self.file_names):
            new_name = name.replace(from_text, self.replace_to_var.get())
            if new_name != name:
                count += 1
            self.file_names[i] = new_name
        self.file_list_widget.delete("1.0", tk.END)
        for name in self.file_names:
            self.file_list_widget.insert(tk.END, name + "\n")
        self.display_results()
        self.status_var.set(f"✨ 替换{count}个")

    def reset_file_names(self):
        if self.original_file_names:
            self.file_names = self.original_file_names.copy()
            self.file_list_widget.delete("1.0", tk.END)
            for name in self.file_names:
                self.file_list_widget.insert(tk.END, name + "\n")
            self.display_results()
            self.status_var.set("↺ 已重置")

    def display_results(self):
        for label, widget in self.text_widgets.items():
            widget.config(state=tk.NORMAL)
            widget.delete("1.0", tk.END)
            idx = int(label[-1]) - 1
            for row in self.rows:
                value = row[idx] if idx < len(row) else ""
                widget.insert(tk.END, value + "\n")
            widget.config(state=tk.DISABLED)

    def export_single_field(self, field_key):
        try:
            if field_key == "文件列表":
                content = self.file_list_widget.get("1.0", tk.END).strip()
            else:
                idx = int(field_key[-1]) - 1
                content = "\n".join([row[idx] if idx < len(row) else "" for row in self.rows])
            
            filename = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{field_key}.txt")
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.status_var.set(f"💾 已保存{field_key}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def export_all_fields(self):
        try:
            lines = []
            for i, row in enumerate(self.rows):
                name = self.file_names[i] if i < len(self.file_names) else ""
                lines.extend([
                    f"# 文件 {i+1}: {name}",
                    f"字段1: {row[0]}",
                    f"字段2: {row[1] if len(row)>1 else ''}",
                    f"字段3: {row[2] if len(row)>2 else ''}",
                    f"字段4: {row[3] if len(row)>3 else ''}",
                    ""
                ])
            filename = filedialog.asksaveasfilename(defaultextension=".txt", initialfile="docx_分类结果.txt")
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("\n".join(lines))
                self.status_var.set("💾 全部导出完成")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def clear_all(self):
        self.file_list_widget.delete("1.0", tk.END)
        for widget in self.text_widgets.values():
            widget.config(state=tk.NORMAL)
            widget.delete("1.0", tk.END)
            widget.config(state=tk.DISABLED)
        self.rows.clear()
        self.file_names.clear()
        self.original_file_names.clear()
        self.status_var.set("已清空全部")


def main():
    try:
        root = tk.Tk()
        app = DocXClassifierApp(root)
        root.mainloop()
    except Exception as e:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动失败", str(e))
        except:
            pass


if __name__ == "__main__":
    main()
