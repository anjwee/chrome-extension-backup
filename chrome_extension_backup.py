import os
import shutil
import json
import platform
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import time
import subprocess
import sys
import ctypes
import winreg

class ChromeExtensionBackupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Chrome扩展设置备份与恢复工具")
        self.root.geometry("630x330")  # 设置窗口大小为640x370
        
        # 初始化备份工具
        self.backup_tool = ChromeExtensionBackup()
        
        # 创建UI元素
        self.create_widgets()
        
        # 加载可用备份
        self.refresh_backups_list()
        
        # 获取可用配置
        self.profiles = self.backup_tool.get_available_profiles()
        self.profile_combobox['values'] = self.profiles
        if self.profiles:
            self.profile_combobox.current(0)

    def create_widgets(self):
        # 创建选项卡
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 备份选项卡
        self.backup_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.backup_frame, text="备份")
        
        # 恢复选项卡
        self.restore_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.restore_frame, text="恢复")
        
        # 备份界面
        ttk.Label(self.backup_frame, text="选择Chrome配置:").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        self.profile_combobox = ttk.Combobox(self.backup_frame, state="readonly")
        self.profile_combobox.grid(row=0, column=1, padx=5, pady=3, sticky=tk.EW)
        
        ttk.Label(self.backup_frame, text="备份名称:").grid(row=1, column=0, padx=5, pady=3, sticky=tk.W)
        self.backup_name_entry = ttk.Entry(self.backup_frame)
        self.backup_name_entry.grid(row=1, column=1, padx=5, pady=3, sticky=tk.EW)
        
        self.backup_button = ttk.Button(self.backup_frame, text="创建备份", command=self.create_backup)
        self.backup_button.grid(row=2, column=0, columnspan=2, pady=5)
        
        # 恢复界面
        ttk.Label(self.restore_frame, text="选择备份:").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        
        # 创建Treeview来显示备份详情
        self.backups_tree = ttk.Treeview(self.restore_frame, columns=('profile', 'timezone', 'date'), selectmode='browse', height=8)
        self.backups_tree.heading('#0', text='备份名称')
        self.backups_tree.heading('profile', text='Chrome配置')
        self.backups_tree.heading('timezone', text='时区')
        self.backups_tree.heading('date', text='备份时间')
        
        # 调整列宽以适应新的窗口大小
        self.backups_tree.column('#0', width=120)
        self.backups_tree.column('profile', width=80)
        self.backups_tree.column('timezone', width=200)
        self.backups_tree.column('date', width=120)
        
        self.backups_tree.grid(row=1, column=0, columnspan=3, padx=5, pady=3, sticky=tk.NSEW)
        
        scrollbar = ttk.Scrollbar(self.restore_frame, orient="vertical", command=self.backups_tree.yview)
        scrollbar.grid(row=1, column=3, sticky=tk.NS)
        self.backups_tree.configure(yscrollcommand=scrollbar.set)
        
        # 操作按钮框架
        button_frame = ttk.Frame(self.restore_frame)
        button_frame.grid(row=2, column=0, columnspan=4, pady=3)
        
        self.restore_button = ttk.Button(button_frame, text="恢复选中备份", command=self.restore_backup)
        self.restore_button.pack(side=tk.LEFT, padx=3)
        
        self.delete_button = ttk.Button(button_frame, text="删除选中备份", command=self.delete_backup)
        self.delete_button.pack(side=tk.LEFT, padx=3)

        # 添加手动时区设置
        timezone_frame = ttk.Frame(self.restore_frame)
        timezone_frame.grid(row=3, column=0, columnspan=4, pady=3)
        
        ttk.Label(timezone_frame, text="手动设置时区:").pack(side=tk.LEFT, padx=3)
        
        # 获取可用时区
        available_timezones = self.backup_tool.get_available_timezones()
        # 减小Combobox的宽度
        self.timezone_combobox = ttk.Combobox(timezone_frame, values=available_timezones, width=40)
        self.timezone_combobox.pack(side=tk.LEFT, padx=3)
        
        ttk.Button(timezone_frame, text="应用时区", command=self.apply_timezone).pack(side=tk.LEFT, padx=3)
        
        # 配置网格权重
        self.backup_frame.columnconfigure(1, weight=1)
        self.restore_frame.columnconfigure(0, weight=1)
        self.restore_frame.rowconfigure(1, weight=1)

    def apply_timezone(self):
        """应用选择的时区"""
        timezone_display = self.timezone_combobox.get()
        if timezone_display:
            if self.backup_tool.set_timezone(timezone_display):
                messagebox.showinfo("成功", f"时区已设置为: {timezone_display}")
            else:
                messagebox.showerror("错误", "设置时区失败")

    def refresh_backups_list(self):
        """刷新备份列表"""
        for item in self.backups_tree.get_children():
            self.backups_tree.delete(item)
        
        backups = self.backup_tool.list_backups_with_details()
        
        for backup_name, details in backups.items():
            timezone = details.get('timezone', '未知时区')
            self.backups_tree.insert('', 'end', text=backup_name, 
                                   values=(details['profile'], 
                                          timezone, 
                                          details['timestamp']))

    def create_backup(self):
        """创建备份"""
        profile = self.profile_combobox.get()
        backup_name = self.backup_name_entry.get()
        
        if not profile:
            messagebox.showerror("错误", "请选择Chrome配置")
            return
        
        if not backup_name:
            messagebox.showerror("错误", "请输入备份名称")
            return
        
        try:
            if not messagebox.askyesno("警告", 
                                    "请确保Chrome浏览器已完全关闭(包括后台进程)\n是否继续?"):
                return
                
            self.backup_tool.backup_extensions(profile, backup_name)
            messagebox.showinfo("成功", f"备份 '{backup_name}' 创建成功")
            self.backup_name_entry.delete(0, tk.END)
            self.refresh_backups_list()
        except Exception as e:
            error_msg = f"备份失败: {str(e)}\n\n可能原因:\n"
            error_msg += "1. Chrome浏览器未完全关闭\n"
            error_msg += "2. 没有足够的权限\n"
            error_msg += "3. 备份名称已存在"
            messagebox.showerror("错误", error_msg)

    def restore_backup(self):
        """恢复备份"""
        selected_item = self.backups_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要恢复的备份")
            return
        
        backup_name = self.backups_tree.item(selected_item, 'text')
        
        if not messagebox.askyesno("确认", 
                                 f"确定要恢复备份 '{backup_name}' 吗?\n这将覆盖当前的Chrome扩展设置和时区配置。\n请确保Chrome浏览器已关闭"):
            return
        
        try:
            self.backup_tool.restore_extensions(backup_name)
            messagebox.showinfo("成功", "恢复完成!\n请重新启动Chrome浏览器")
        except Exception as e:
            messagebox.showerror("错误", f"恢复失败: {str(e)}")

    def delete_backup(self):
        """删除备份"""
        selected_item = self.backups_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要删除的备份")
            return
        
        backup_name = self.backups_tree.item(selected_item, 'text')
        
        if not messagebox.askyesno("确认", f"确定要删除备份 '{backup_name}' 吗?\n此操作不可恢复!"):
            return
        
        try:
            self.backup_tool.delete_backup(backup_name)
            # 移除成功提示，直接刷新列表
            self.refresh_backups_list()
        except Exception as e:
            messagebox.showerror("错误", f"删除失败: {str(e)}")

class ChromeExtensionBackup:
    def __init__(self):
        # 确定Chrome用户数据目录位置
        self.chrome_data_dir = os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data")
        
        # 备份存储目录
        self.backup_dir = os.path.expanduser("~/chrome_extension_backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        
        # 备份元数据文件
        self.metadata_file = os.path.join(self.backup_dir, "backups_metadata.json")
        if not os.path.exists(self.metadata_file):
            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump({"backups": {}}, f, ensure_ascii=False, indent=2)

    def get_available_profiles(self):
        """获取所有用户配置"""
        profiles = []
        try:
            for item in os.listdir(self.chrome_data_dir):
                if item.startswith("Profile") or item == "Default":
                    profiles.append(item)
        except FileNotFoundError:
            pass
        return profiles

    def get_current_timezone(self):
        """获取当前时区设置"""
        try:
            # 使用tzutil命令获取当前时区ID
            result = subprocess.run(['tzutil', '/g'], capture_output=True, text=True)
            if result.returncode == 0:
                tz_id = result.stdout.strip()
                # 获取时区显示名称
                result = subprocess.run(['tzutil', '/l'], capture_output=True, text=True)
                if result.returncode == 0:
                    tz_list = result.stdout.strip().split('\n')
                    # 查找对应的显示名称
                    for tz in tz_list:
                        if tz.strip() == tz_id:
                            return tz.strip()
            return "未知时区"
        except Exception as e:
            print(f"获取时区失败: {str(e)}")
            return "未知时区"

    def get_available_timezones(self):
        """获取可用的时区列表"""
        try:
            # 使用tzutil命令获取时区列表
            result = subprocess.run(['tzutil', '/l'], capture_output=True, text=True)
            if result.returncode == 0:
                # 过滤和处理时区列表
                timezones = [tz.strip() for tz in result.stdout.strip().split('\n') if tz.strip()]
                return sorted(timezones)
            return []
        except Exception as e:
            print(f"获取时区列表失败: {str(e)}")
            return []

    def set_timezone(self, timezone_name):
        """设置系统时区"""
        try:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                raise PermissionError("需要管理员权限才能更改时区")

            print(f"正在尝试将时区设置为: {timezone_name}")
            
            # 使用tzutil命令设置时区
            result = subprocess.run(['tzutil', '/s', timezone_name], 
                                 capture_output=True, 
                                 text=True)
            
            if result.returncode != 0:
                error_msg = result.stderr.strip() or "未知错误"
                raise RuntimeError(f"设置时区失败: {error_msg}")
            
            return True

        except PermissionError as e:
            print(f"权限错误: {str(e)}")
            messagebox.showerror("错误", "需要管理员权限才能更改时区。\n请以管理员身份运行程序。")
            return False
        except RuntimeError as e:
            print(f"运行时错误: {str(e)}")
            messagebox.showerror("错误", str(e))
            return False
        except Exception as e:
            print(f"未知错误: {str(e)}")
            messagebox.showerror("错误", f"设置时区时发生未知错误: {str(e)}")
            return False

    def backup_extensions(self, profile, backup_name):
        """备份指定配置的扩展设置"""
        profile_path = os.path.join(self.chrome_data_dir, profile)
        if not os.path.exists(profile_path):
            raise FileNotFoundError(f"Profile {profile} not found")
        
        backup_path = os.path.join(self.backup_dir, backup_name)
        if os.path.exists(backup_path):
            raise FileExistsError(f"备份名称 '{backup_name}' 已存在")
        
        os.makedirs(backup_path, exist_ok=True)
        
        files_to_backup = [
            os.path.join(profile_path, "Preferences"),
            os.path.join(profile_path, "Local Extension Settings")
        ]
        
        for file_path in files_to_backup:
            if os.path.exists(file_path):
                try:
                    if os.path.isdir(file_path):
                        for root, dirs, files in os.walk(file_path):
                            dest_dir = os.path.join(backup_path, 
                                                  os.path.basename(file_path), 
                                                  os.path.relpath(root, file_path))
                            os.makedirs(dest_dir, exist_ok=True)
                            for file in files:
                                src_file = os.path.join(root, file)
                                dest_file = os.path.join(dest_dir, file)
                                if not file.endswith('LOCK'):
                                    try:
                                        shutil.copy2(src_file, dest_file)
                                    except PermissionError:
                                        print(f"跳过无法访问的文件: {src_file}")
                    else:
                        shutil.copy2(file_path, backup_path)
                except PermissionError as e:
                    print(f"跳过无法访问的路径: {file_path} - {str(e)}")
        
        current_timezone = self.get_current_timezone()
        
        with open(self.metadata_file, 'r+', encoding='utf-8') as f:
            metadata = json.load(f)
            if "backups" not in metadata:
                metadata["backups"] = {}
            metadata["backups"][backup_name] = {
                "profile": profile,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "timezone": current_timezone,
                "backup_path": backup_path
            }
            f.seek(0)
            json.dump(metadata, f, ensure_ascii=False, indent=2)
            f.truncate()
        
        return True

    def restore_extensions(self, backup_name):
        """恢复指定备份(包括时区设置)"""
        with open(self.metadata_file, 'r', encoding='utf-8') as f:
            metadata = json.load(f)
        
        if backup_name not in metadata.get("backups", {}):
            raise ValueError(f"Backup {backup_name} not found")
        
        backup_info = metadata["backups"][backup_name]
        profile = backup_info["profile"]
        backup_path = backup_info["backup_path"]
        timezone = backup_info.get("timezone", "")
        
        profile_path = os.path.join(self.chrome_data_dir, profile)
        if not os.path.exists(profile_path):
            raise FileNotFoundError(f"Profile {profile} not found")
        
        for item in os.listdir(backup_path):
            src = os.path.join(backup_path, item)
            dst = os.path.join(profile_path, item)
            
            if os.path.exists(dst):
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                else:
                    os.remove(dst)
            
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
        
        if timezone and timezone != "未知时区":
            try:
                print(f"正在尝试将时区设置为: {timezone}")
                if self.set_timezone(timezone):
                    messagebox.showinfo("成功", f"时区已成功设置为: {timezone}")
                else:
                    messagebox.showwarning("警告",
                        "时区设置失败。\n" +
                        "可能的原因：\n" +
                        "1. 程序没有管理员权限\n" +
                        "2. 时区名称无效\n" +
                        "3. 系统不支持自动设置时区\n" +
                        "\n请尝试以管理员身份运行程序，或手动设置时区。")
            except Exception as e:
                messagebox.showerror("错误", f"设置时区时出错: {str(e)}\n请手动设置时区。")
        
        return True

    def list_backups_with_details(self):
        """列出所有备份及其详细信息"""
        try:
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
            return metadata.get("backups", {})
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def delete_backup(self, backup_name):
        """删除指定备份"""
        with open(self.metadata_file, 'r+', encoding='utf-8') as f:
            try:
                metadata = json.load(f)
            except json.JSONDecodeError:
                metadata = {"backups": {}}
            
            if backup_name not in metadata.get("backups", {}):
                raise ValueError(f"Backup {backup_name} not found")
            
            backup_path = metadata["backups"][backup_name].get("backup_path", "")
            
            if os.path.exists(backup_path):
                shutil.rmtree(backup_path)
            
            if "backups" in metadata and backup_name in metadata["backups"]:
                del metadata["backups"][backup_name]
            
            f.seek(0)
            json.dump(metadata, f, ensure_ascii=False, indent=2)
            f.truncate()
        
        return True

def is_admin():
    """检查是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == "__main__":
    # 检查是否具有管理员权限
    if not is_admin():
        # 如果不是管理员权限，显示警告
        if messagebox.askyesno("警告", 
            "程序未以管理员身份运行。\n" +
            "这可能导致无法自动恢复时区设置。\n" +
            "是否以管理员身份重新启动程序？"):
            # 请求管理员权限重新运行
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
    
    # 创建并运行GUI
    root = tk.Tk()
    app = ChromeExtensionBackupGUI(root)
    root.mainloop()
