import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess # Added for opening files
import sys # Added for platform detection
# from pypdf import PdfWriter, PdfReader # No longer directly used here
from merge_pdfs import merge_files # Import the new function

class PdfMergerApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF 合并工具")
        master.geometry("600x450")

        self.pdf_files = []

        # --- UI Elements ---
        # Frame for file list and controls
        list_frame = ttk.Frame(master, padding="10")
        list_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Listbox to display selected files
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, width=60, height=10)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Frame for list control buttons
        button_frame_list = ttk.Frame(master, padding="5")
        button_frame_list.pack(fill=tk.X)

        self.add_button = ttk.Button(button_frame_list, text="添加文件", command=self.add_files)
        self.add_button.pack(side=tk.LEFT, padx=5)

        self.remove_button = ttk.Button(button_frame_list, text="移除选中", command=self.remove_file)
        self.remove_button.pack(side=tk.LEFT, padx=5)

        self.up_button = ttk.Button(button_frame_list, text="上移", command=self.move_up)
        self.up_button.pack(side=tk.LEFT, padx=5)

        self.down_button = ttk.Button(button_frame_list, text="下移", command=self.move_down)
        self.down_button.pack(side=tk.LEFT, padx=5)

        self.preview_button = ttk.Button(button_frame_list, text="预览选中", command=self.preview_selected_file) # New Preview Button
        self.preview_button.pack(side=tk.LEFT, padx=5)

        self.preview_all_button = ttk.Button(button_frame_list, text="预览全部", command=self.preview_all_files) # New Preview All Button
        self.preview_all_button.pack(side=tk.LEFT, padx=5)

        # Frame for output file selection
        output_frame = ttk.Frame(master, padding="10")
        output_frame.pack(fill=tk.X)

        output_label = ttk.Label(output_frame, text="输出文件:")
        output_label.pack(side=tk.LEFT, padx=5)

        self.output_entry = ttk.Entry(output_frame, width=50)
        self.output_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.output_entry.insert(0, "merged_output.pdf") # Default output name

        self.browse_button = ttk.Button(output_frame, text="浏览...", command=self.select_output_file)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        # Merge button
        self.merge_button = ttk.Button(master, text="开始合并", command=self.merge_selected_files, padding="5")
        self.merge_button.pack(pady=10)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择要合并的文件",
            filetypes=[
                ("支持的文件", "*.pdf *.docx *.ppt *.pptx *.jpg *.jpeg *.png"),
                ("PDF files", "*.pdf"),
                ("Word documents", "*.docx"),
                ("PowerPoint files", "*.ppt *.pptx"),
                ("Image files", "*.jpg *.jpeg *.png")
            ]
        )
        if files:
            for file_path in files:
                if file_path not in self.pdf_files:
                    self.pdf_files.append(file_path)
                    self.listbox.insert(tk.END, os.path.basename(file_path))

    def remove_file(self):
        selected_index = self.listbox.curselection()
        if selected_index:
            index = selected_index[0]
            del self.pdf_files[index]
            self.listbox.delete(index)

    def move_up(self):
        selected_index = self.listbox.curselection()
        if selected_index:
            index = selected_index[0]
            if index > 0:
                # Swap in list
                self.pdf_files[index], self.pdf_files[index - 1] = self.pdf_files[index - 1], self.pdf_files[index]
                # Swap in listbox
                text = self.listbox.get(index)
                self.listbox.delete(index)
                self.listbox.insert(index - 1, text)
                self.listbox.selection_set(index - 1)
                self.listbox.activate(index - 1)

    def move_down(self):
        selected_index = self.listbox.curselection()
        if selected_index:
            index = selected_index[0]
            if index < self.listbox.size() - 1:
                # Swap in list
                self.pdf_files[index], self.pdf_files[index + 1] = self.pdf_files[index + 1], self.pdf_files[index]
                # Swap in listbox
                text = self.listbox.get(index)
                self.listbox.delete(index)
                self.listbox.insert(index + 1, text)
                self.listbox.selection_set(index + 1)
                self.listbox.activate(index + 1)

    def preview_selected_file(self):
        """Opens the selected file using the system's default application."""
        selected_index = self.listbox.curselection()
        if not selected_index:
            messagebox.showwarning("未选择文件", "请先在列表中选择一个文件进行预览。")
            return

        index = selected_index[0]
        file_path = self.pdf_files[index]

        try:
            # Use os.startfile on Windows (more direct)
            # Use subprocess.Popen(['open', file_path]) on macOS
            # Use subprocess.Popen(['xdg-open', file_path]) on Linux
            if os.name == 'nt': # Windows
                os.startfile(file_path)
            elif sys.platform == 'darwin': # macOS
                subprocess.Popen(['open', file_path])
            else: # Linux and other Unix-like
                subprocess.Popen(['xdg-open', file_path])
        except FileNotFoundError:
            messagebox.showerror("错误", f"文件未找到: {file_path}")
        except Exception as e:
            messagebox.showerror("预览错误", f"无法打开文件 '{os.path.basename(file_path)}'.\n错误: {e}")

    def preview_all_files(self):
        """Opens all files in the list using the system's default application."""
        if not self.pdf_files:
            messagebox.showwarning("无文件", "列表中没有文件可预览。")
            return

        errors = []
        for file_path in self.pdf_files:
            try:
                if os.name == 'nt': # Windows
                    os.startfile(file_path)
                elif sys.platform == 'darwin': # macOS
                    subprocess.Popen(['open', file_path])
                else: # Linux and other Unix-like
                    subprocess.Popen(['xdg-open', file_path])
            except FileNotFoundError:
                errors.append(f"文件未找到: {os.path.basename(file_path)}")
            except Exception as e:
                errors.append(f"无法打开 '{os.path.basename(file_path)}': {e}")

        if errors:
            messagebox.showerror("预览错误", "预览部分文件时出错:\n" + "\n".join(errors))

    def select_output_file(self):
        output_filename = filedialog.asksaveasfilename(
            title="保存合并后的 PDF 文件",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if output_filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_filename)

    def merge_selected_files(self):
        if not self.pdf_files:
            messagebox.showwarning("无文件", "请先添加要合并的文件。")
            return

        output_filename = self.output_entry.get()
        if not output_filename:
            messagebox.showwarning("无输出文件", "请指定输出文件名和路径。")
            return

        try:
            # Call the new merge_files function with the list of selected files
            success, message = merge_files(self.pdf_files, output_filename)

            if success:
                messagebox.showinfo("成功", message)
                # Optionally clear the list after successful merge
                # self.clear_list() # You might need to implement clear_list if you uncomment this
            else:
                messagebox.showerror("合并失败", f"合并过程中发生错误:\n{message}")

        except Exception as e:
            # Catch any unexpected errors during the merge_files call itself
            messagebox.showerror("错误", f"合并过程中发生意外错误: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PdfMergerApp(root)
    root.mainloop()