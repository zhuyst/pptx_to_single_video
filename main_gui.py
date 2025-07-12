# -*- coding: UTF-8 -*-

import os
import time
import threading
import subprocess
import win32com.client
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import uuid
import tempfile


class PPTToVideoConverter:
    def __init__(self):
        self.prs = None
        self.powerpoint = None
        self.is_converting = False
        
    def export_single_slide_to_video(self, slide_index, output_wmv, progress_callback=None):
        """导出单个幻灯片为视频"""
        try:
            powerpoint = self.prs.Application
            single_prs = powerpoint.Presentations.Add()
            self.prs.Slides(slide_index).Copy()
            single_prs.Slides.Paste()
            
            # 使用临时目录和英文文件名避免中文路径问题
            temp_dir = tempfile.gettempdir()
            temp_filename = f"temp_slide_{uuid.uuid4().hex}.pptx"
            temp_pptx = os.path.join(temp_dir, temp_filename)
            temp_pptx = os.path.normpath(temp_pptx)
            
            # 确保输出路径格式正确
            output_wmv = os.path.normpath(os.path.abspath(output_wmv))
            
            single_prs.SaveAs(temp_pptx)
            single_prs.CreateVideo(output_wmv, False, 5, 720, 30, 100)
            
            while single_prs.CreateVideoStatus != 3:
                if not self.is_converting:  # 检查是否被取消
                    break
                if progress_callback:
                    status = single_prs.CreateVideoStatus
                    progress_callback(f"正在处理第{slide_index}页 (状态: {status})")
                time.sleep(1)
                
            single_prs.Close()
            if os.path.exists(temp_pptx):
                try:
                    os.remove(temp_pptx)
                except:
                    pass
            return True
        except Exception as e:
            print(f"导出第{slide_index}页时出错: {e}")
            return False
    
    def convert_ppt_to_videos(self, pptx_path, progress_callback=None, completion_callback=None):
        """转换PPT为视频"""
        try:
            self.is_converting = True
            
            # 规范化输入路径
            pptx_path = os.path.normpath(os.path.abspath(pptx_path))
            
            # 检查文件是否存在
            if not os.path.exists(pptx_path):
                raise FileNotFoundError(f"文件不存在: {pptx_path}")
            
            # 根据pptx文件名创建输出目录
            pptx_name = os.path.splitext(os.path.basename(pptx_path))[0]
            output_dir = pptx_name
            
            # 确保输出目录路径正确
            output_dir = os.path.normpath(os.path.abspath(output_dir))
            os.makedirs(output_dir, exist_ok=True)
            
            if progress_callback:
                progress_callback("正在打开PowerPoint...")
            
            # 启动PowerPoint应用程序
            try:
                self.powerpoint = win32com.client.Dispatch('PowerPoint.Application')
                self.powerpoint.Visible = 1
            except Exception as e:
                raise Exception(f"无法启动PowerPoint应用程序: {e}")
            
            # 打开PPT文件
            try:
                self.prs = self.powerpoint.Presentations.Open(pptx_path, WithWindow=False)
            except Exception as e:
                raise Exception(f"无法打开PPT文件: {e}\n文件路径: {pptx_path}")
            
            slide_count = self.prs.Slides.Count
            
            if progress_callback:
                progress_callback(f"开始转换，共{slide_count}页幻灯片")
            
            success_count = 0
            for i in range(1, slide_count + 1):
                if not self.is_converting:  # 检查是否被取消
                    break
                    
                # 使用英文文件名避免中文路径问题
                wmv_filename = f"{pptx_name}_{i}.wmv"
                wmv_path = os.path.join(output_dir, wmv_filename)
                wmv_path = os.path.normpath(os.path.abspath(wmv_path))
                
                if progress_callback:
                    progress_callback(f"正在导出第{i}页为视频...")
                
                if self.export_single_slide_to_video(i, wmv_path, progress_callback):
                    success_count += 1
                    if progress_callback:
                        progress_callback(f"第{i}页导出完成 ({success_count}/{slide_count})")
                else:
                    if progress_callback:
                        progress_callback(f"第{i}页导出失败")
            
            self.cleanup()
            
            if completion_callback:
                completion_callback(output_dir, success_count, slide_count)
                
        except Exception as e:
            error_msg = f"转换过程中出错: {str(e)}"
            print(error_msg)
            if progress_callback:
                progress_callback(error_msg)
            self.cleanup()
            if completion_callback:
                completion_callback(None, 0, 0)
    
    def cleanup(self):
        """清理资源"""
        try:
            if self.prs:
                self.prs.Close()
        except:
            pass
        
        try:
            if self.powerpoint:
                self.powerpoint.Quit()
        except:
            pass
        
        self.prs = None
        self.powerpoint = None
        self.is_converting = False
    
    def stop_conversion(self):
        """停止转换"""
        self.is_converting = False


class PPTToVideoGUI:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("PPT转视频工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.converter = PPTToVideoConverter()
        self.selected_file = tk.StringVar()
        self.conversion_thread = None
        
        self.setup_ui()
        self.setup_drag_drop()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="PPT转视频工具", font=("微软雅黑", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="选择PPT文件", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 文件路径显示
        self.file_label = ttk.Label(file_frame, textvariable=self.selected_file, 
                                   relief="sunken", width=60, background="white")
        self.file_label.grid(row=0, column=0, padx=(0, 10), sticky=(tk.W, tk.E))
        
        # 浏览按钮
        browse_btn = ttk.Button(file_frame, text="浏览", command=self.browse_file)
        browse_btn.grid(row=0, column=1)
        
        # 拖拽提示
        drop_label = ttk.Label(file_frame, text="或将.pptx文件拖拽到上方文件路径框", 
                              font=("微软雅黑", 9), foreground="gray")
        drop_label.grid(row=1, column=0, columnspan=2, pady=(5, 0))
        
        # 转换控制区域
        control_frame = ttk.LabelFrame(main_frame, text="转换控制", padding="10")
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 开始导出按钮
        self.start_btn = ttk.Button(control_frame, text="开始导出", command=self.start_conversion)
        self.start_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 停止按钮
        self.stop_btn = ttk.Button(control_frame, text="停止转换", command=self.stop_conversion, state="disabled")
        self.stop_btn.grid(row=0, column=1)
        
        # 进度显示区域
        progress_frame = ttk.LabelFrame(main_frame, text="转换进度", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 进度条
        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 进度文本
        self.progress_text = tk.StringVar(value="请选择PPT文件后开始转换")
        progress_label = ttk.Label(progress_frame, textvariable=self.progress_text)
        progress_label.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 状态区域
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.status_text = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_text, foreground="blue")
        status_label.grid(row=0, column=0)
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
        file_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def setup_drag_drop(self):
        """设置拖拽功能"""
        self.file_label.drop_target_register(DND_FILES)
        self.file_label.dnd_bind('<<Drop>>', self.on_file_drop)
    
    def on_file_drop(self, event):
        """处理文件拖拽"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.pptx'):
                # 规范化路径
                file_path = os.path.normpath(file_path)
                self.selected_file.set(file_path)
                self.status_text.set("文件已选择")
            else:
                messagebox.showerror("错误", "请选择.pptx文件")
    
    def browse_file(self):
        """浏览文件"""
        file_path = filedialog.askopenfilename(
            title="选择PPT文件",
            filetypes=[("PowerPoint文件", "*.pptx"), ("所有文件", "*.*")]
        )
        if file_path:
            # 规范化路径
            file_path = os.path.normpath(file_path)
            self.selected_file.set(file_path)
            self.status_text.set("文件已选择")
    
    def start_conversion(self):
        """开始转换"""
        if not self.selected_file.get():
            messagebox.showerror("错误", "请先选择PPT文件")
            return
        
        file_path = os.path.normpath(self.selected_file.get())
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 检查路径和文件名
        try:
            # 测试路径是否可以正常访问
            with open(file_path, 'rb') as f:
                pass
        except Exception as e:
            messagebox.showerror("错误", f"无法访问文件，可能是路径或文件名包含特殊字符: {e}")
            return
        
        # 更新UI状态
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress_var.set(0)
        self.status_text.set("正在转换...")
        
        # 在新线程中执行转换
        self.conversion_thread = threading.Thread(
            target=self.converter.convert_ppt_to_videos,
            args=(file_path, self.update_progress, self.conversion_complete)
        )
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def stop_conversion(self):
        """停止转换"""
        self.converter.stop_conversion()
        self.reset_ui()
        self.status_text.set("转换已停止")
    
    def update_progress(self, message):
        """更新进度"""
        self.root.after(0, lambda: self.progress_text.set(message))
        
        # 尝试从消息中提取进度信息
        if "(" in message and "/" in message and ")" in message:
            try:
                # 提取 (x/y) 格式的进度
                progress_part = message[message.find("("):message.find(")")+1]
                if "/" in progress_part:
                    current, total = progress_part.strip("()").split("/")
                    progress_percent = int((int(current) / int(total)) * 100)
                    self.root.after(0, lambda: self.progress_var.set(progress_percent))
            except:
                pass
    
    def conversion_complete(self, output_dir, success_count, total_count):
        """转换完成"""
        self.root.after(0, lambda: self.reset_ui())
        
        if output_dir and success_count > 0:
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.progress_text.set(f"转换完成！成功导出 {success_count}/{total_count} 个视频"))
            self.root.after(0, lambda: self.status_text.set("转换完成"))
            
            # 询问是否打开输出文件夹
            result = messagebox.askyesno("转换完成", 
                                       f"转换完成！成功导出 {success_count}/{total_count} 个视频\n\n是否打开输出文件夹？")
            if result:
                self.open_output_folder(output_dir)
        else:
            self.root.after(0, lambda: self.progress_text.set("转换失败或被取消"))
            self.root.after(0, lambda: self.status_text.set("转换失败"))
    
    def reset_ui(self):
        """重置UI状态"""
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
    
    def open_output_folder(self, output_dir):
        """打开输出文件夹"""
        try:
            output_dir = os.path.normpath(os.path.abspath(output_dir))
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}")
    
    def on_closing(self):
        """窗口关闭事件"""
        if self.converter.is_converting:
            result = messagebox.askyesno("确认", "转换正在进行中，确定要退出吗？")
            if not result:
                return
            self.converter.stop_conversion()
        
        self.converter.cleanup()
        self.root.destroy()
    
    def run(self):
        """运行应用程序"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()


if __name__ == "__main__":
    app = PPTToVideoGUI()
    app.run() 