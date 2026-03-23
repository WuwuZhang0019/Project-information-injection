import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import win32com.client

DATA_FILE = "projects_data.json"
LSP_FILE = "current_project.lsp"

class CadInfoInjector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CAD 字段变量管理器")
        self.geometry("700x500")
        
        self.projects = {}
        self.current_project = None
        
        self.load_data()
        self.setup_ui()
        
    def load_data(self):
        default_project = {
            "默认项目组": {
                "1#": {
                    "xmmc": "津武（挂）2023-017号地块居住项目（三期）",
                    "jsdw": "天津万新澜城置业有限公司",
                    "zxmc": "1#楼",
                    "xmbh": "STG2023422-\t1#",
                    "ctrq": "2025-01",
                    "BBH": "B"
                }
            }
        }
        
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if not data:
                    self.projects = default_project.copy()
                else:
                    # 兼容老数据结构 (一层字典) 转化为两层结构
                    needs_migration = False
                    for key, val in data.items():
                        if isinstance(val, dict) and any(not isinstance(v, dict) for v in val.values()):
                            needs_migration = True
                            break
                    if needs_migration:
                        self.projects = {"导入的项目": data}
                    else:
                        self.projects = data
            except Exception as e:
                messagebox.showerror("错误", f"读取数据文件失败: {e}")
                self.projects = default_project.copy()
        else:
            self.projects = default_project.copy()

    def save_data(self):
        try:
            with open(DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.projects, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("错误", f"保存数据失败: {e}")

    def setup_ui(self):
        # Left Panel (Projects List)
        left_frame = tk.Frame(self, width=250, bg="#f0f0f0")
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        
        tk.Label(left_frame, text="项目列表").pack(pady=5)
        
        # 使用 Treeview 替代 Listbox，并添加滚动条
        tree_frame = tk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.proj_tree = ttk.Treeview(tree_frame, selectmode="browse", show="tree")
        self.proj_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 项目树滚动条
        proj_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.proj_tree.yview)
        proj_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.proj_tree.configure(yscrollcommand=proj_scrollbar.set)
        
        self.proj_tree.bind('<<TreeviewSelect>>', self.on_project_select)
        
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # 第一排按钮：文件夹操作
        btn_frame1 = tk.Frame(btn_frame)
        btn_frame1.pack(fill=tk.X, pady=2)
        tk.Button(btn_frame1, text="建组", command=self.add_group).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(btn_frame1, text="重命名组", command=self.rename_group).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        
        # 第二排按钮：项目操作
        btn_frame2 = tk.Frame(btn_frame)
        btn_frame2.pack(fill=tk.X, pady=2)
        tk.Button(btn_frame2, text="加项目", command=self.add_project).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(btn_frame2, text="复制", command=self.copy_project).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(btn_frame2, text="删除", command=self.delete_item).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        
        # Right Panel (Variables Table)
        right_frame = tk.Frame(self)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.proj_title = tk.Label(right_frame, text="未选择项目", font=("Arial", 14, "bold"))
        self.proj_title.pack(pady=5)
        
        # Treeview for Variables
        columns = ("Variable", "Value")
        self.tree = ttk.Treeview(right_frame, columns=columns, show="headings")
        self.tree.heading("Variable", text="Lisp 变量名 (如 xmbh)")
        self.tree.heading("Value", text="变量值")
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<Double-1>', self.on_double_click)
        
        var_btn_frame = tk.Frame(right_frame)
        var_btn_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(var_btn_frame, text="添加变量", command=self.add_variable).pack(side=tk.LEFT, padx=5)
        tk.Button(var_btn_frame, text="删除变量", command=self.delete_variable).pack(side=tk.LEFT, padx=5)
        tk.Button(var_btn_frame, text="保存当前项目", command=self.save_current_project).pack(side=tk.LEFT, padx=5)
        
        action_frame = tk.Frame(right_frame)
        action_frame.pack(fill=tk.X, pady=10)
        tk.Button(action_frame, text="一键注入到 CAD (生成LSP并执行)", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", command=self.inject_to_cad).pack(fill=tk.X, ipady=10)
        
        self.refresh_project_list()
        
    def refresh_project_list(self):
        # 清空树
        for item in self.proj_tree.get_children():
            self.proj_tree.delete(item)
            
        for group_name, projects_in_group in self.projects.items():
            # 插入组节点
            group_id = self.proj_tree.insert("", tk.END, iid=f"group_{group_name}", text=group_name, open=True)
            # 插入项目节点
            for proj_name in projects_in_group.keys():
                self.proj_tree.insert(group_id, tk.END, iid=f"proj_{group_name}_{proj_name}", text=proj_name)
            
    def get_selected_item_info(self):
        selection = self.proj_tree.selection()
        if not selection:
            return None, None, None
        item_id = selection[0]
        if item_id.startswith("group_"):
            group_name = item_id[len("group_"):]
            return "group", group_name, None
        elif item_id.startswith("proj_"):
            # 格式是 proj_{group_name}_{proj_name}
            parts = item_id.split("_", 2)
            if len(parts) == 3:
                return "project", parts[1], parts[2]
        return None, None, None

    def on_project_select(self, event):
        item_type, group_name, proj_name = self.get_selected_item_info()
        
        if item_type == "project":
            self.current_project = (group_name, proj_name)
            self.proj_title.config(text=f"项目: {group_name} / {proj_name}")
            self.refresh_variables()
        else:
            self.current_project = None
            if item_type == "group":
                self.proj_title.config(text=f"项目组: {group_name}")
            else:
                self.proj_title.config(text="未选择项目")
            self.refresh_variables()
            
    def refresh_variables(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
            
        if self.current_project:
            group_name, proj_name = self.current_project
            if group_name in self.projects and proj_name in self.projects[group_name]:
                vars_dict = self.projects[group_name][proj_name]
                for var, val in vars_dict.items():
                    self.tree.insert("", tk.END, values=(var, val))
                    
    def add_group(self):
        name = self.prompt_dialog("添加项目组", "输入项目组名称:")
        if name:
            if name in self.projects:
                messagebox.showwarning("警告", "项目组已存在")
                return
            self.projects[name] = {}
            self.save_data()
            self.refresh_project_list()
            
    def rename_group(self):
        item_type, group_name, _ = self.get_selected_item_info()
        if item_type != "group":
            messagebox.showwarning("提示", "请先选择一个项目组（文件夹）")
            return
            
        new_name = self.prompt_dialog("重命名项目组", f"将项目组 '{group_name}' 重命名为:", initialvalue=group_name)
        if new_name and new_name != group_name:
            if new_name in self.projects:
                messagebox.showwarning("警告", "项目组名已存在！")
                return
            # 更新字典键
            self.projects[new_name] = self.projects.pop(group_name)
            self.save_data()
            self.refresh_project_list()
                
    def add_project(self):
        item_type, group_name, _ = self.get_selected_item_info()
        if not group_name:
            messagebox.showwarning("提示", "请先选择一个项目组或该组下的项目，以确定在哪个组内添加项目。")
            return
            
        name = self.prompt_dialog("添加项目", f"在组 '{group_name}' 中输入新项目名称:")
        if name:
            if name in self.projects[group_name]:
                messagebox.showwarning("警告", "该组下项目已存在")
                return
            self.projects[group_name][name] = {}
            self.save_data()
            self.refresh_project_list()
            
    def copy_project(self):
        if not self.current_project:
            messagebox.showwarning("提示", "请先选择一个要复制的具体项目")
            return
            
        group_name, proj_name = self.current_project
        new_name = self.prompt_dialog("复制项目", f"输入新项目名称 (复制自 '{proj_name}'):")
        if new_name:
            if new_name in self.projects[group_name]:
                messagebox.showwarning("警告", "该组下项目已存在")
                return
            # 复制变量
            self.projects[group_name][new_name] = dict(self.projects[group_name][proj_name])
            self.save_data()
            self.refresh_project_list()
            
            # 自动选中新复制的项目
            new_item_id = f"proj_{group_name}_{new_name}"
            self.proj_tree.selection_set(new_item_id)
            self.proj_tree.focus(new_item_id)
            self.proj_tree.event_generate("<<TreeviewSelect>>")

    def delete_item(self):
        item_type, group_name, proj_name = self.get_selected_item_info()
        if not item_type:
            return
            
        if item_type == "group":
            if messagebox.askyesno("确认", f"确定删除整个项目组 '{group_name}' 及其包含的所有项目吗?"):
                del self.projects[group_name]
                self.current_project = None
                self.proj_title.config(text="未选择项目")
                self.save_data()
                self.refresh_project_list()
                self.refresh_variables()
        elif item_type == "project":
            if messagebox.askyesno("确认", f"确定删除项目 '{proj_name}' 吗?"):
                del self.projects[group_name][proj_name]
                self.current_project = None
                self.proj_title.config(text="未选择项目")
                self.save_data()
                self.refresh_project_list()
                self.refresh_variables()
            
    def add_variable(self):
        if not self.current_project:
            messagebox.showwarning("提示", "请先选择一个具体项目")
            return
            
        var_name = self.prompt_dialog("添加变量", "输入 Lisp 变量名 (不需要加括号):")
        if var_name:
            var_value = self.prompt_dialog("添加变量", f"输入 '{var_name}' 的值:")
            if var_value is not None:
                group_name, proj_name = self.current_project
                self.projects[group_name][proj_name][var_name] = var_value
                self.save_data()
                self.refresh_variables()
                
    def delete_variable(self):
        if not self.current_project:
            return
        selected = self.tree.selection()
        if not selected:
            return
        
        item = self.tree.item(selected[0])
        var_name = item['values'][0]
        
        if messagebox.askyesno("确认", f"确定删除变量 '{var_name}' 吗?"):
            group_name, proj_name = self.current_project
            if var_name in self.projects[group_name][proj_name]:
                del self.projects[group_name][proj_name][var_name]
                self.save_data()
                self.refresh_variables()
                
    def on_double_click(self, event):
        if not self.current_project:
            return
            
        # 获取点击的区域
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        # 获取点击的列和行
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        
        if not item_id:
            return
            
        item = self.tree.item(item_id)
        old_var_name = item['values'][0]
        old_val = item['values'][1]
        
        group_name, proj_name = self.current_project
        
        if column == '#1':  # 点击了第一列 (变量名)
            new_var_name = self.prompt_dialog("修改变量名", f"修改变量名 (原: {old_var_name}):", initialvalue=old_var_name)
            if new_var_name and new_var_name != old_var_name:
                if new_var_name in self.projects[group_name][proj_name]:
                    messagebox.showwarning("警告", "该变量名已存在于当前项目中！")
                    return
                # 更新字典的键：取出旧值并赋给新键
                val = self.projects[group_name][proj_name].pop(old_var_name)
                self.projects[group_name][proj_name][new_var_name] = val
                self.save_data()
                self.refresh_variables()
                
        elif column == '#2':  # 点击了第二列 (变量值)
            new_val = self.prompt_dialog("修改变量值", f"修改 '{old_var_name}' 的值:", initialvalue=old_val)
            if new_val is not None:
                self.projects[group_name][proj_name][old_var_name] = new_val
                self.save_data()
                self.refresh_variables()
            
    def save_current_project(self):
        self.save_data()
        messagebox.showinfo("成功", "保存成功")

    def inject_to_cad(self):
        if not self.current_project:
            messagebox.showwarning("提示", "请先选择一个具体项目")
            return
            
        group_name, proj_name = self.current_project
        vars_dict = self.projects[group_name][proj_name]
        if not vars_dict:
            messagebox.showwarning("提示", "该项目没有变量")
            return
            
        # 1. Generate LSP file content
        lsp_content = ";; 自动生成的CAD注入脚本\n"
        for var, val in vars_dict.items():
            # 简单处理，如果值是数字可以不加引号，这里默认全当字符串处理，CAD中字段通常也是字符串
            lsp_content += f'(setq {var} "{val}")\n'
        
        # 添加自动刷新命令
        lsp_content += '(command "REGEN")\n'
        lsp_content += f'(princ "\\n[{group_name} / {proj_name}] 变量已更新并刷新图纸。")\n'
        lsp_content += '(princ)\n'
        
        # 保存到当前目录
        lsp_path = os.path.abspath(LSP_FILE)
        try:
            with open(lsp_path, 'w', encoding='utf-8') as f:
                f.write(lsp_content)
        except Exception as e:
            messagebox.showerror("错误", f"生成LSP文件失败: {e}")
            return
            
        # 2. 尝试通过 COM 接口发送给 AutoCAD
        try:
            # 尝试连接当前运行的 AutoCAD
            acad = win32com.client.Dispatch("AutoCAD.Application")
            doc = acad.ActiveDocument
            # 转换路径，将 \ 替换为 / 以防 Lisp 解析错误
            lsp_path_cad = lsp_path.replace("\\", "/")
            # 发送 load 命令并触发 regen
            command_str = f'(load "{lsp_path_cad}") '
            doc.SendCommand(command_str)
            messagebox.showinfo("成功", f"成功注入到 AutoCAD！\n项目: {group_name} / {proj_name}\n已执行 REGEN。")
        except Exception as e:
            messagebox.showwarning("COM 注入失败", f"无法直接控制 AutoCAD (可能未打开或权限不足)。\n\nLSP 脚本已生成在:\n{lsp_path}\n您可以将其直接拖入 CAD 窗口中。\n\n错误信息: {e}")

    def prompt_dialog(self, title, prompt, initialvalue=""):
        # 简单的输入对话框替代方案
        import tkinter.simpledialog as sd
        return sd.askstring(title, prompt, initialvalue=initialvalue, parent=self)

if __name__ == "__main__":
    app = CadInfoInjector()
    app.mainloop()
