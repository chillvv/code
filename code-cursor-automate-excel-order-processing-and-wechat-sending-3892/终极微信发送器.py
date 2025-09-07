#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
终极微信发送器 - 使用最简单直接的方法
完全避开输入框查找问题，使用模拟用户操作
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sys
import re
import time
import threading
from datetime import datetime
import tempfile
import platform

# 可选的拖放支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# 基础自动化支持
try:
    import pyperclip
    import pyautogui
    HAS_AUTO = True
    # 设置pyautogui参数
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1
except ImportError:
    HAS_AUTO = False

# Win32支持
try:
    import win32gui
    import win32con
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


class UltimateWeChatSender:
    """终极微信发送器 - 使用最直接的方法"""
    
    def __init__(self):
        """初始化程序"""
        self.log("🚀 启动终极微信发送器...")
        
        # 初始化变量
        self.data = []
        self.columns = []
        self.lunch_orders = ""
        self.dinner_orders = ""
        self.is_sending = False
        self.stop_sending = False
        self.wechat_hwnd = None
        
        # 创建主窗口
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
            
        self.setup_ui()
        self.log("✅ 程序初始化完成")
    
    def log(self, message):
        """输出日志信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")
    
    def setup_ui(self):
        """设置用户界面"""
        self.root.title("终极微信发送器 - 直接操作方案")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="📁 Excel文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="请选择Excel文件或拖拽文件到此区域" if HAS_DND else "请选择Excel文件")
        self.file_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(file_frame, text="选择文件", command=self.select_file).grid(row=1, column=0, padx=(0, 5))
        ttk.Button(file_frame, text="创建测试文件", command=self.create_test_file).grid(row=1, column=1, padx=(5, 0))
        
        # 拖放支持
        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_file_drop)
        
        # 参数设置区域
        param_frame = ttk.LabelFrame(main_frame, text="⚙️ 发送参数", padding="10")
        param_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 起始编号设置
        ttk.Label(param_frame, text="午餐起始编号:").grid(row=0, column=0, sticky=tk.W)
        self.lunch_start = tk.StringVar(value="1")
        ttk.Entry(param_frame, textvariable=self.lunch_start, width=10).grid(row=0, column=1, padx=(5, 20))
        
        ttk.Label(param_frame, text="晚餐起始编号:").grid(row=0, column=2, sticky=tk.W)
        self.dinner_start = tk.StringVar(value="1")
        ttk.Entry(param_frame, textvariable=self.dinner_start, width=10).grid(row=0, column=3, padx=(5, 0))
        
        # 群设置
        ttk.Label(param_frame, text="午餐群:").grid(row=1, column=0, sticky=tk.W)
        self.lunch_group = tk.StringVar(value="简知午餐群")
        ttk.Entry(param_frame, textvariable=self.lunch_group, width=15).grid(row=1, column=1, padx=(5, 20))
        
        ttk.Label(param_frame, text="晚餐群:").grid(row=1, column=2, sticky=tk.W)
        self.dinner_group = tk.StringVar(value="简知晚餐群")
        ttk.Entry(param_frame, textvariable=self.dinner_group, width=15).grid(row=1, column=3, padx=(5, 0))
        
        # 发送选择
        send_select_frame = ttk.Frame(param_frame)
        send_select_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(send_select_frame, text="发送选择:").grid(row=0, column=0, sticky=tk.W)
        self.send_lunch = tk.BooleanVar(value=True)
        self.send_dinner = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(send_select_frame, text="发送午餐订单", variable=self.send_lunch).grid(row=0, column=1, padx=(10, 20), sticky=tk.W)
        ttk.Checkbutton(send_select_frame, text="发送晚餐订单", variable=self.send_dinner).grid(row=0, column=2, padx=(0, 20), sticky=tk.W)
        
        # 测试模式
        self.test_mode = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="测试模式（发送到'末'群）", variable=self.test_mode).grid(row=3, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="📋 订单预览", padding="10")
        preview_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=15, width=80)
        self.preview_text.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 按钮区域
        ttk.Button(preview_frame, text="处理订单", command=self.process_orders).grid(row=1, column=0, padx=(0, 5))
        ttk.Button(preview_frame, text="直接发送到微信", command=self.send_to_wechat).grid(row=1, column=1, padx=5)
        ttk.Button(preview_frame, text="停止发送", command=self.stop_sending_orders).grid(row=1, column=2, padx=(5, 0))
        ttk.Button(preview_frame, text="测试微信", command=self.test_wechat_window).grid(row=1, column=3, padx=(5, 0))
        
        # 第二行按钮
        ttk.Button(preview_frame, text="测试群聊搜索", command=self.test_group_search).grid(row=2, column=0, padx=(0, 5), pady=(5, 0))
        ttk.Button(preview_frame, text="测试发送", command=self.test_send_message).grid(row=2, column=1, padx=5, pady=(5, 0))
        ttk.Button(preview_frame, text="测试输入框定位", command=self.test_input_location).grid(row=2, column=2, padx=(5, 0), pady=(5, 0))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # 绑定快捷键
        self.root.bind('<Control-s>', lambda e: self.stop_sending_orders())
    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.load_excel_file(file_path)
    
    def on_file_drop(self, event):
        """处理拖放文件"""
        files = self.root.tk.splitlist(event.data)
        if files:
            self.load_excel_file(files[0])
    
    def load_excel_file(self, file_path):
        """加载Excel文件"""
        try:
            self.log(f"📂 正在加载文件: {os.path.basename(file_path)}")
            self.status_var.set(f"正在加载: {os.path.basename(file_path)}")
            
            # 使用强化的Excel读取方法
            df = self._load_dataframe(file_path)
            
            self.data = df.values.tolist()
            self.columns = df.columns.tolist()
            
            self.log(f"✅ 成功加载 {len(self.data)} 行数据，{len(self.columns)} 列")
            self.file_label.config(text=f"已加载: {os.path.basename(file_path)} ({len(self.data)}行)")
            self.status_var.set(f"已加载: {len(self.data)}行数据")
            
            # 显示列信息
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, "📋 文件列信息:\n")
            for i, col in enumerate(self.columns):
                self.preview_text.insert(tk.END, f"{i+1}. {col}\n")
            self.preview_text.insert(tk.END, f"\n总共 {len(self.data)} 行数据\n")
            
        except Exception as e:
            error_msg = f"❌ 加载文件失败: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            self.status_var.set("加载失败")
    
    def _load_dataframe(self, file_path):
        """强化的Excel加载方法"""
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
            try:
                return pd.read_excel(file_path, engine='openpyxl')
            except Exception:
                try:
                    return pd.read_excel(file_path)
                except Exception:
                    if HAS_WIN32 and platform.system().lower() == "windows":
                        fixed_file = self._repair_excel_via_com(file_path)
                        if fixed_file:
                            return pd.read_excel(fixed_file, engine='openpyxl')
                    raise
        
        elif ext == ".xls":
            try:
                return pd.read_excel(file_path, engine='xlrd')
            except Exception:
                try:
                    if HAS_WIN32 and platform.system().lower() == "windows":
                        fixed_file = self._repair_excel_via_com(file_path)
                        if fixed_file:
                            return pd.read_excel(fixed_file, engine='openpyxl')
                except Exception:
                    pass
                raise
        
        elif ext == ".csv":
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
            for encoding in encodings:
                try:
                    return pd.read_csv(file_path, encoding=encoding)
                except:
                    continue
            raise Exception("无法解析CSV文件编码")
        
        else:
            return pd.read_excel(file_path)
    
    def _repair_excel_via_com(self, file_path):
        """使用Excel COM修复文件"""
        if not HAS_WIN32:
            return None
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            temp_path = os.path.join(tempfile.gettempdir(), f"repaired_{int(time.time()*1000)}.xlsx")
            wb.SaveAs(temp_path, 51)
            wb.Close(False)
            excel.Quit()
            
            self.log("✅ Excel文件已通过COM修复")
            return temp_path
            
        except Exception as e:
            self.log(f"⚠️ COM修复失败: {str(e)}")
            try:
                excel.Quit()
            except:
                pass
            return None
    
    def create_test_file(self):
        """创建测试Excel文件"""
        try:
            test_data = {
                '商品信息': [
                    '明日午餐 x1', '明日晚餐 x1', '明日午餐 x1', '明日晚餐 x1',
                    '明日午餐 x1', '明日晚餐 x1', '明日午餐 x1'
                ],
                '支付状态': [
                    '已支付', '已支付', '未支付', '已支付',
                    '已支付', '已退款', '已支付'
                ],
                '订单状态': [
                    '已完成', '制作中', '待支付', '已完成',
                    '商品中', '已取消', '已完成'
                ],
                '收货地址': [
                    '张三-13800138000-光谷A座101室',
                    '李四-13900139000-南湖B栋202室',
                    '王五-13700137000-卓刀泉C区303号',
                    '赵六-13600136000-关山D园404室',
                    '孙七-13500135000-鲁巷E座505室',
                    '周八-13400134000-华科F栋606室',
                    '吴九-13300133000-珞喻路G号707室'
                ],
                '用户备注': [
                    '', '12点前送达', '', '不要辣',
                    '多加米饭', '', '少盐少油'
                ]
            }
            
            df = pd.DataFrame(test_data)
            test_file = "测试订单数据.xlsx"
            df.to_excel(test_file, index=False)
            
            self.log(f"✅ 创建测试文件: {test_file}")
            messagebox.showinfo("成功", f"测试文件已创建: {test_file}")
            self.load_excel_file(test_file)
            
        except Exception as e:
            error_msg = f"创建测试文件失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("错误", error_msg)
    
    def process_orders(self):
        """处理订单数据"""
        if not self.data:
            messagebox.showwarning("警告", "请先加载Excel文件")
            return
        
        try:
            self.log("🔄 开始处理订单数据...")
            self.status_var.set("正在处理订单...")
            
            # 自动识别列
            column_mapping = self._detect_columns()
            if not column_mapping:
                messagebox.showerror("错误", "无法识别必要的列")
                return
            
            # 处理数据
            lunch_orders, dinner_orders = self._process_order_data(column_mapping)
            
            # 保存订单列表用于一条一条发送
            self.lunch_order_list = lunch_orders
            self.dinner_order_list = dinner_orders
            
            # 生成输出文本用于预览
            self.lunch_orders = self._generate_output(lunch_orders, int(self.lunch_start.get()), "午餐", "明日午餐 x1")
            self.dinner_orders = self._generate_output(dinner_orders, int(self.dinner_start.get()), "晚餐", "明日晚餐 x1")
            
            # 显示预览 - 根据选择显示对应订单
            self.preview_text.delete(1.0, tk.END)
            
            # 显示发送状态提示
            send_status = []
            if self.send_lunch.get():
                send_status.append("午餐订单")
            if self.send_dinner.get():
                send_status.append("晚餐订单")
            
            if send_status:
                status_text = " + ".join(send_status)
                if self.test_mode.get():
                    self.preview_text.insert(tk.END, f"【将发送 {status_text} 到测试群'末'】\n\n")
                else:
                    self.preview_text.insert(tk.END, f"【将发送 {status_text}】\n\n")
            else:
                self.preview_text.insert(tk.END, "【未选择发送任何订单】\n\n")
            
            # 显示午餐订单（如果选中）
            if self.send_lunch.get() and self.lunch_orders.strip():
                target_group = "末" if self.test_mode.get() else self.lunch_group.get()
                self.preview_text.insert(tk.END, f"🍽️ 午餐订单 → {target_group}\n")
                self.preview_text.insert(tk.END, "=" * 50 + "\n")
                self.preview_text.insert(tk.END, self.lunch_orders)
                self.preview_text.insert(tk.END, "\n\n")
            
            # 显示晚餐订单（如果选中）
            if self.send_dinner.get() and self.dinner_orders.strip():
                target_group = "末" if self.test_mode.get() else self.dinner_group.get()
                self.preview_text.insert(tk.END, f"🍽️ 晚餐订单 → {target_group}\n")
                self.preview_text.insert(tk.END, "=" * 50 + "\n")
                self.preview_text.insert(tk.END, self.dinner_orders)
            
            # 检查是否有订单显示
            if ((not self.send_lunch.get() or not self.lunch_orders.strip()) and 
                (not self.send_dinner.get() or not self.dinner_orders.strip())):
                if not send_status:
                    self.preview_text.insert(tk.END, "请选择要发送的订单类型")
                else:
                    self.preview_text.insert(tk.END, "没有对应的订单数据")
            
            total_orders = len(lunch_orders) + len(dinner_orders)
            self.log(f"✅ 处理完成: 午餐{len(lunch_orders)}条, 晚餐{len(dinner_orders)}条")
            self.status_var.set(f"处理完成: 午餐{len(lunch_orders)}条, 晚餐{len(dinner_orders)}条")
            
        except Exception as e:
            error_msg = f"处理订单失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("错误", error_msg)
            self.status_var.set("处理失败")
    
    def _detect_columns(self):
        """自动检测列名"""
        mapping = {}
        keywords = {
            'product_info': ['商品信息', '商品', '产品'],
            'payment_status': ['支付状态', '付款状态'],
            'order_status': ['订单状态'],
            'address': ['收货地址', '地址', '收货人'],
            'user_note': ['用户备注', '备注', '说明']
        }
        
        for key, words in keywords.items():
            for col in self.columns:
                if any(word in str(col) for word in words):
                    mapping[key] = col
                    break
        
        required = ['product_info', 'payment_status', 'address']
        for req in required:
            if req not in mapping:
                return None
        
        return mapping
    
    def _process_order_data(self, mapping):
        """处理订单数据"""
        df = pd.DataFrame(self.data, columns=self.columns)
        df = df.fillna("")
        df['__row__'] = range(len(df))
        
        # 筛选已支付订单
        payment_col = mapping['payment_status']
        paid_orders = df[df[payment_col].astype(str).str.strip() == '已支付']
        
        # 排除无效订单
        def is_valid_order(row):
            payment_status = str(row[payment_col]).strip()
            order_status = str(row.get(mapping.get('order_status', ''), '')).strip()
            
            if payment_status in ['未支付', '已退款']:
                return False
            if order_status in ['已取消', '用户申请退款']:
                return False
            return True
        
        valid_orders = paid_orders[paid_orders.apply(is_valid_order, axis=1)]
        
        # 按商品信息分类
        product_col = mapping['product_info']
        lunch_orders = valid_orders[valid_orders[product_col].astype(str).str.contains('明日午餐', na=False)]
        dinner_orders = valid_orders[valid_orders[product_col].astype(str).str.contains('明日晚餐', na=False)]
        
        # 按行号倒序排列
        lunch_orders = lunch_orders.sort_values('__row__', ascending=False)
        dinner_orders = dinner_orders.sort_values('__row__', ascending=False)
        
        # 转换为列表格式
        def to_order_list(orders_df):
            order_list = []
            for _, row in orders_df.iterrows():
                order_list.append({
                    'address': self._format_address(str(row[mapping['address']])),
                    'user_note': str(row.get(mapping.get('user_note', ''), '')).strip()
                })
            return order_list
        
        return to_order_list(lunch_orders), to_order_list(dinner_orders)
    
    def _format_address(self, address):
        """格式化地址"""
        address = str(address).strip()
        if not address:
            return "地址信息缺失"
        
        if re.match(r'^[^-]+-[^-]+-', address):
            return address
        
        phone_pattern = r'1[3-9]\d{9}'
        phone_match = re.search(phone_pattern, address)
        
        if phone_match:
            phone = phone_match.group()
            parts = address.split(phone)
            if len(parts) >= 2:
                name = parts[0].strip(' -')
                addr = phone.join(parts[1:]).strip(' -')
                return f"{name}-{phone}-{addr}"
        
        return address
    
    def _generate_output(self, orders, start_num, title, product_label):
        """生成输出文本"""
        if not orders:
            return ""  # 如果没有订单就返回空字符串
        
        lines = []
        
        for i, order in enumerate(orders):
            lines.append(str(start_num + i))
            lines.append(order['address'])
            if order['user_note']:
                lines.append(f"（用户备注：{order['user_note']}）")
            if i < len(orders) - 1:  # 不是最后一个订单时添加空行
                lines.append("")
        
        return "\n".join(lines)
    
    def test_wechat_window(self):
        """测试微信窗口"""
        if not HAS_WIN32:
            messagebox.showerror("错误", "需要安装 pywin32 包")
            return
        
        try:
            self.log("🔍 测试微信窗口...")
            self.status_var.set("正在测试微信窗口...")
            
            hwnd = self._find_wechat_window()
            if hwnd:
                self.wechat_hwnd = hwnd
                window_title = win32gui.GetWindowText(hwnd)
                self.log(f"✅ 找到微信窗口: {window_title}")
                messagebox.showinfo("测试成功", f"微信窗口已找到!\n窗口标题: {window_title}")
                self.status_var.set("微信窗口正常")
            else:
                self.log("❌ 未找到微信窗口")
                messagebox.showerror("测试失败", "未找到微信窗口，请确保微信已启动")
                self.status_var.set("未找到微信窗口")
                
        except Exception as e:
            error_msg = f"测试微信窗口失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("测试失败", error_msg)
            self.status_var.set("测试失败")
    
    def _find_wechat_window(self):
        """查找微信窗口 - 改进版"""
        if not HAS_WIN32:
            return None
        
        def enum_windows_callback(hwnd, windows):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                window_text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                # 更准确的微信窗口识别
                wechat_indicators = [
                    ("WeChatMainWndForPC" in class_name, "主窗口类名"),
                    ("微信" in window_text and len(window_text) < 10, "窗口标题"),
                    ("WeChat" in window_text and "PC" not in window_text, "英文标题"),
                    (class_name.startswith("Qt") and "微信" in window_text, "Qt框架窗口"),
                    (class_name == "Chrome_WidgetWin_1" and "微信" in window_text, "Chrome内核窗口")
                ]
                
                # 计算匹配度
                match_score = 0
                match_reasons = []
                for condition, reason in wechat_indicators:
                    if condition:
                        match_score += 1
                        match_reasons.append(reason)
                
                if match_score > 0:
                    # 排除一些明显不是主窗口的
                    if any(keyword in window_text.lower() for keyword in ['update', 'installer', 'setup']):
                        return True
                    
                    windows.append((hwnd, window_text, class_name, match_score, match_reasons))
                    
            except Exception:
                pass  # 忽略获取窗口信息时的异常
            
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if not windows:
            self.log("❌ 未找到任何微信窗口")
            return None
        
        # 按匹配度排序，选择最佳匹配
        windows.sort(key=lambda x: x[3], reverse=True)
        
        for hwnd, title, class_name, score, reasons in windows:
            try:
                # 验证窗口是否真的可用
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    self.log(f"✅ 找到微信窗口: {title}")
                    self.log(f"   类名: {class_name}")
                    self.log(f"   匹配度: {score} ({', '.join(reasons)})")
                    return hwnd
            except Exception:
                continue
        
        self.log("❌ 找到微信窗口但都不可用")
        return None
    
    def test_group_search(self):
        """测试群聊搜索功能"""
        if not HAS_AUTO:
            messagebox.showerror("错误", "需要安装 pyautogui 和 pyperclip 包")
            return
        
        # 获取测试群名
        test_group = "末" if self.test_mode.get() else self.lunch_group.get()
        
        result = messagebox.askyesno("测试群聊搜索", 
            f"将测试搜索群聊: {test_group}\n\n"
            "请确保：\n"
            "1. 微信已打开并登录\n"
            "2. 微信窗口可见\n"
            "3. 测试期间不要操作电脑\n\n"
            "开始测试？")
        
        if not result:
            return
        
        try:
            self.log(f"🧪 开始测试群聊搜索: {test_group}")
            self.status_var.set("正在测试群聊搜索...")
            
            # 激活微信窗口
            if not self._activate_wechat():
                messagebox.showerror("测试失败", "无法激活微信窗口")
                return
            
            # 测试搜索功能
            success = self._switch_to_group(test_group)
            
            if success:
                self.log("✅ 群聊搜索测试成功!")
                messagebox.showinfo("测试成功", f"成功找到并进入群聊: {test_group}")
                self.status_var.set("群聊搜索测试成功")
            else:
                self.log("❌ 群聊搜索测试失败")
                messagebox.showerror("测试失败", f"无法找到或进入群聊: {test_group}")
                self.status_var.set("群聊搜索测试失败")
                
        except Exception as e:
            error_msg = f"测试群聊搜索失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("测试失败", error_msg)
            self.status_var.set("测试失败")
    
    def test_send_message(self):
        """测试发送消息功能"""
        if not HAS_AUTO:
            messagebox.showerror("错误", "需要安装 pyautogui 和 pyperclip 包")
            return
        
        test_message = "这是一条测试消息，用于验证微信自动发送功能。\n如果看到此消息，说明发送功能正常！"
        test_group = "末"
        
        result = messagebox.askyesno("测试发送消息", 
            f"将向群聊'{test_group}'发送测试消息:\n\n"
            f"{test_message}\n\n"
            "请确保：\n"
            "1. 微信已打开并登录\n"
            "2. 微信窗口可见\n"
            "3. 存在名为'末'的群聊\n"
            "4. 测试期间不要操作电脑\n\n"
            "开始测试？")
        
        if not result:
            return
        
        try:
            self.log(f"🧪 开始测试发送消息到: {test_group}")
            self.status_var.set("正在测试发送消息...")
            
            # 激活微信窗口
            if not self._activate_wechat():
                messagebox.showerror("测试失败", "无法激活微信窗口")
                return
            
            # 切换到测试群
            if not self._switch_to_group(test_group):
                messagebox.showerror("测试失败", f"无法进入群聊: {test_group}")
                return
            
            # 发送测试消息
            success = self._send_single_order(test_message)
            
            if success:
                self.log("✅ 消息发送测试成功!")
                messagebox.showinfo("测试成功", "测试消息已发送，请检查微信群聊")
                self.status_var.set("消息发送测试成功")
            else:
                self.log("❌ 消息发送测试失败")
                messagebox.showerror("测试失败", "消息发送失败")
                self.status_var.set("消息发送测试失败")
                
        except Exception as e:
            error_msg = f"测试发送消息失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("测试失败", error_msg)
            self.status_var.set("测试失败")
    
    def test_input_location(self):
        """测试输入框定位功能"""
        if not HAS_AUTO:
            messagebox.showerror("错误", "需要安装 pyautogui 和 pyperclip 包")
            return
        
        result = messagebox.askyesno("测试输入框定位", 
            "将测试输入框定位功能\n\n"
            "请确保：\n"
            "1. 微信已打开并登录\n"
            "2. 微信窗口可见\n"
            "3. 已进入任意群聊或个人聊天\n"
            "4. 测试期间不要操作电脑\n\n"
            "测试会在输入框位置显示红色标记\n"
            "开始测试？")
        
        if not result:
            return
        
        try:
            self.log(f"🧪 开始测试输入框定位")
            self.status_var.set("正在测试输入框定位...")
            
            # 激活微信窗口
            if not self._activate_wechat():
                messagebox.showerror("测试失败", "无法激活微信窗口")
                return
            
            # 测试各种定位方法
            self.log("📍 测试方法1: 控件识别")
            pos1 = self._find_input_by_control()
            if pos1:
                self.log(f"✅ 控件识别成功: {pos1}")
                self._mark_position(pos1, "控件识别", "red")
            else:
                self.log("❌ 控件识别失败")
            
            time.sleep(1)
            
            self.log("📍 测试方法2: 窗口计算")
            pos2 = self._find_input_by_window_calc()
            if pos2:
                self.log(f"✅ 窗口计算成功: {pos2}")
                self._mark_position(pos2, "窗口计算", "blue")
            else:
                self.log("❌ 窗口计算失败")
            
            time.sleep(1)
            
            self.log("📍 测试方法3: 智能点击")
            success = self._smart_click_input_area()
            if success:
                self.log("✅ 智能点击成功")
            else:
                self.log("❌ 智能点击失败")
            
            # 汇总结果
            results = []
            if pos1:
                results.append(f"控件识别: {pos1}")
            if pos2:
                results.append(f"窗口计算: {pos2}")
            if success:
                results.append("智能点击: 成功")
            
            if results:
                result_text = "\n".join(results)
                messagebox.showinfo("测试成功", f"输入框定位测试完成!\n\n{result_text}\n\n请查看微信窗口上的标记点")
                self.status_var.set("输入框定位测试成功")
            else:
                messagebox.showerror("测试失败", "所有定位方法都失败了")
                self.status_var.set("输入框定位测试失败")
                
        except Exception as e:
            error_msg = f"测试输入框定位失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("测试失败", error_msg)
            self.status_var.set("测试失败")
    
    def _mark_position(self, position, method_name, color):
        """在指定位置显示标记"""
        try:
            import tkinter as tk
            x, y = position
            
            # 创建一个小的标记窗口
            marker = tk.Toplevel()
            marker.title(f"定位标记 - {method_name}")
            marker.geometry(f"20x20+{x-10}+{y-10}")
            marker.configure(bg=color)
            marker.attributes("-topmost", True)
            marker.overrideredirect(True)
            
            # 3秒后自动关闭
            marker.after(3000, marker.destroy)
            
            self.log(f"🔴 在位置 {position} 显示{color}色标记 ({method_name})")
            
        except Exception as e:
            self.log(f"⚠️ 显示标记失败: {e}")
    
    def send_to_wechat(self):
        """直接发送到微信"""
        if not hasattr(self, 'lunch_orders') or not hasattr(self, 'dinner_orders'):
            messagebox.showwarning("警告", "请先处理订单数据")
            return
        
        if not HAS_AUTO:
            messagebox.showerror("错误", "需要安装 pyautogui 和 pyperclip 包")
            return
        
        # 检查发送选择
        if not self.send_lunch.get() and not self.send_dinner.get():
            messagebox.showwarning("警告", "请至少选择一种订单类型进行发送")
            return
        
        # 生成确认信息
        send_items = []
        if self.send_lunch.get() and hasattr(self, 'lunch_order_list') and self.lunch_order_list:
            target = "末" if self.test_mode.get() else self.lunch_group.get()
            send_items.append(f"午餐订单({len(self.lunch_order_list)}条) → {target}")
        
        if self.send_dinner.get() and hasattr(self, 'dinner_order_list') and self.dinner_order_list:
            target = "末" if self.test_mode.get() else self.dinner_group.get()
            send_items.append(f"晚餐订单({len(self.dinner_order_list)}条) → {target}")
        
        if not send_items:
            messagebox.showwarning("警告", "没有可发送的订单数据")
            return
        
        send_info = "\n".join(send_items)
        
        result = messagebox.askyesno("确认发送", 
            f"即将发送以下订单：\n\n{send_info}\n\n"
            "发送方式：完全模拟用户操作\n"
            "请确保：\n"
            "1. 微信已打开并登录\n"
            "2. 微信窗口可见\n"
            "3. 发送期间不要操作电脑\n"
            "4. 可以按Ctrl+S停止发送\n\n"
            "是否开始发送？")
        if not result:
            return
        
        # 在新线程中执行发送
        self.is_sending = True
        self.stop_sending = False
        thread = threading.Thread(target=self._send_orders_thread, daemon=True)
        thread.start()
    
    def _send_orders_thread(self):
        """发送订单的线程函数"""
        try:
            self.log("🚀 开始直接发送到微信...")
            self.status_var.set("正在直接发送到微信...")
            
            # 准备发送项目 - 根据用户选择发送
            items = []
            
            # 处理午餐订单（如果选中）
            if self.send_lunch.get() and hasattr(self, 'lunch_order_list') and self.lunch_order_list:
                target_group = "末" if self.test_mode.get() else self.lunch_group.get()
                self.log(f"📋 准备午餐订单: {len(self.lunch_order_list)}条 → {target_group}")
                for i, order in enumerate(self.lunch_order_list):
                    order_text = str(int(self.lunch_start.get()) + i) + "\n" + order['address']
                    if order['user_note']:
                        order_text += f"\n（用户备注：{order['user_note']}）"
                    items.append((target_group, order_text, "午餐"))
            
            # 处理晚餐订单（如果选中）
            if self.send_dinner.get() and hasattr(self, 'dinner_order_list') and self.dinner_order_list:
                target_group = "末" if self.test_mode.get() else self.dinner_group.get()
                self.log(f"📋 准备晚餐订单: {len(self.dinner_order_list)}条 → {target_group}")
                for i, order in enumerate(self.dinner_order_list):
                    order_text = str(int(self.dinner_start.get()) + i) + "\n" + order['address']
                    if order['user_note']:
                        order_text += f"\n（用户备注：{order['user_note']}）"
                    items.append((target_group, order_text, "晚餐"))
            
            if not items:
                self.status_var.set("没有订单需要发送")
                return
            
            # 确保微信窗口激活
            if not self._activate_wechat():
                self.log("❌ 无法激活微信窗口")
                self.status_var.set("无法激活微信窗口")
                return
            
            # 一条一条发送
            current_group = None
            lunch_count = 0
            dinner_count = 0
            
            for i, (group, content, meal_type) in enumerate(items):
                if self.stop_sending:
                    break
                
                # 统计发送数量
                if meal_type == "午餐":
                    lunch_count += 1
                    order_num = lunch_count
                else:
                    dinner_count += 1
                    order_num = dinner_count
                
                # 如果切换群，需要重新搜索
                if current_group != group:
                    self.log(f"📤 切换到群: {group}")
                    self.status_var.set(f"正在切换到: {group}")
                    success = self._switch_to_group(group)
                    if not success:
                        self.log(f"❌ 切换群失败: {group}")
                        continue
                    current_group = group
                    time.sleep(0.5)
                
                self.log(f"📤 发送{meal_type}第 {order_num} 条 (总进度: {i+1}/{len(items)})")
                self.status_var.set(f"正在发送{meal_type}: {order_num} ({i+1}/{len(items)})")
                
                success = self._send_single_order(content)
                
                if success:
                    self.log(f"✅ {meal_type}第{order_num}条发送成功")
                else:
                    self.log(f"❌ {meal_type}第{order_num}条发送失败")
                
                # 间隔1-1.5秒
                if i < len(items) - 1 and not self.stop_sending:
                    time.sleep(1.2)
            
            if not self.stop_sending:
                self.log("✅ 所有订单发送完成!")
                self.status_var.set("发送完成")
            else:
                self.log("⏹️ 发送已停止")
                self.status_var.set("发送已停止")
                
        except Exception as e:
            error_msg = f"发送失败: {str(e)}"
            self.log(f"❌ {error_msg}")
            self.status_var.set("发送失败")
        finally:
            self.is_sending = False
    
    def _activate_wechat(self):
        """激活微信窗口 - 改进版"""
        try:
            # 重新查找微信窗口，确保窗口仍然有效
            self.wechat_hwnd = self._find_wechat_window()
            
            if not self.wechat_hwnd:
                self.log("❌ 未找到微信窗口")
                return False
            
            # 检查窗口是否仍然有效
            try:
                if not win32gui.IsWindow(self.wechat_hwnd):
                    self.log("⚠️ 微信窗口句柄无效，重新查找")
                    self.wechat_hwnd = self._find_wechat_window()
                    if not self.wechat_hwnd:
                        return False
            except:
                self.wechat_hwnd = self._find_wechat_window()
                if not self.wechat_hwnd:
                    return False
            
            # 多步骤激活窗口
            try:
                # 1. 先恢复窗口（如果被最小化）
                win32gui.ShowWindow(self.wechat_hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
                
                # 2. 将窗口置顶
                win32gui.SetWindowPos(self.wechat_hwnd, win32con.HWND_TOP, 0, 0, 0, 0, 
                                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
                time.sleep(0.3)
                
                # 3. 设置为前台窗口
                win32gui.SetForegroundWindow(self.wechat_hwnd)
                time.sleep(0.5)
                
                # 4. 验证窗口是否真的在前台
                current_window = win32gui.GetForegroundWindow()
                if current_window != self.wechat_hwnd:
                    self.log("⚠️ 微信窗口可能未完全激活，但继续尝试")
                else:
                    self.log("✅ 微信窗口已成功激活")
                
            except Exception as e:
                self.log(f"⚠️ 窗口激活过程中出现异常: {e}")
                # 尝试备用方法
                try:
                    win32gui.SetForegroundWindow(self.wechat_hwnd)
                    time.sleep(0.5)
                except:
                    pass
            
            return True
            
        except Exception as e:
            self.log(f"❌ 激活微信窗口失败: {str(e)}")
            return False
    
    def _switch_to_group(self, group):
        """切换到指定群 - 改进版"""
        try:
            self.log(f"🔍 搜索群聊: {group}")
            
            # 多次尝试打开搜索框
            for attempt in range(3):
                try:
                    pyautogui.hotkey('ctrl', 'f')
                    time.sleep(0.8)
                    break
                except Exception as e:
                    if attempt == 2:
                        raise e
                    time.sleep(0.5)
            
            # 确保搜索框激活并清空
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.3)
            pyautogui.press('delete')
            time.sleep(0.2)
            
            # 输入群名 - 分步骤确保准确
            pyperclip.copy(group)
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.8)
            
            # 验证输入内容
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            
            # 检查剪贴板内容是否正确
            try:
                clipboard_content = pyperclip.paste()
                if clipboard_content != group:
                    self.log(f"⚠️ 剪贴板内容不匹配，重新输入")
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.2)
                    pyperclip.copy(group)
                    time.sleep(0.2)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.5)
            except:
                pass
            
            # 按回车进入群聊
            pyautogui.press('enter')
            time.sleep(2.0)  # 增加等待时间确保进入
            
            self.log(f"✅ 成功切换到群: {group}")
            return True
            
        except Exception as e:
            self.log(f"❌ 切换到群 {group} 失败: {str(e)}")
            return False
    
    def _send_single_order(self, content):
        """发送单条订单 - 智能输入框识别版"""
        try:
            # 复制内容到剪贴板并验证
            pyperclip.copy(content)
            time.sleep(0.3)
            
            # 验证剪贴板内容
            try:
                clipboard_check = pyperclip.paste()
                if clipboard_check != content:
                    self.log("⚠️ 剪贴板验证失败，重新复制")
                    pyperclip.copy(content)
                    time.sleep(0.3)
            except:
                pass
            
            # 尝试找到输入框位置
            input_position = self._find_input_box_position()
            
            if input_position:
                x, y = input_position
                self.log(f"🎯 找到输入框位置: ({x}, {y})")
                pyautogui.click(x, y)
                time.sleep(0.4)
            else:
                self.log("⚠️ 未找到输入框，使用智能点击策略")
                # 使用智能点击策略
                success = self._smart_click_input_area()
                if not success:
                    self.log("⚠️ 智能点击也失败，使用默认位置")
                    screen_width, screen_height = pyautogui.size()
                    pyautogui.click(screen_width // 2, int(screen_height * 0.85))
                    time.sleep(0.4)
            
            # 清空输入框 - 多次尝试确保清空
            for attempt in range(2):
                try:
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.2)
                    pyautogui.press('delete')
                    time.sleep(0.2)
                    break
                except:
                    time.sleep(0.3)
            
            # 粘贴内容
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            
            # 验证粘贴是否成功
            try:
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                pasted_content = pyperclip.paste()
                if content not in pasted_content:
                    self.log("⚠️ 粘贴验证失败，但继续发送")
            except:
                pass
            
            # 发送消息
            pyautogui.press('enter')
            time.sleep(0.3)
            
            return True
            
        except Exception as e:
            self.log(f"❌ 发送单条订单失败: {str(e)}")
            return False
    
    def _find_input_box_position(self):
        """查找微信输入框的实际位置"""
        try:
            # 方法1: 尝试使用控件识别（如果可用）
            position = self._find_input_by_control()
            if position:
                return position
            
            # 方法2: 基于窗口位置的智能估算
            position = self._find_input_by_window_calc()
            if position:
                return position
            
            return None
            
        except Exception as e:
            self.log(f"⚠️ 查找输入框位置失败: {e}")
            return None
    
    def _find_input_by_control(self):
        """通过控件识别查找输入框"""
        try:
            # 尝试导入uiautomation
            try:
                import uiautomation as auto
            except:
                return None
            
            if not self.wechat_hwnd:
                return None
            
            # 通过句柄创建窗口控件
            main_window = auto.WindowControl(handle=self.wechat_hwnd)
            if not main_window.Exists():
                return None
            
            # 查找编辑控件（输入框）
            edit_controls = main_window.EditControls()
            if not edit_controls:
                return None
            
            # 通常最后一个编辑控件是消息输入框
            for edit_ctrl in reversed(edit_controls):
                try:
                    rect = edit_ctrl.BoundingRectangle
                    if rect.width() > 100 and rect.height() > 20:  # 输入框应该有一定大小
                        center_x = rect.left + rect.width() // 2
                        center_y = rect.top + rect.height() // 2
                        self.log(f"🎯 通过控件找到输入框: ({center_x}, {center_y})")
                        return (center_x, center_y)
                except Exception:
                    continue
            
            return None
            
        except Exception as e:
            self.log(f"⚠️ 控件识别输入框失败: {e}")
            return None
    
    def _find_input_by_window_calc(self):
        """通过窗口计算查找输入框位置"""
        try:
            if not HAS_WIN32 or not self.wechat_hwnd:
                return None
            
            # 获取微信窗口的位置和大小
            rect = win32gui.GetWindowRect(self.wechat_hwnd)
            left, top, right, bottom = rect
            window_width = right - left
            window_height = bottom - top
            
            self.log(f"🔍 微信窗口位置: ({left}, {top}) 大小: {window_width}x{window_height}")
            
            # 根据窗口大小动态调整输入框位置
            if window_height < 400:  # 小窗口
                input_offset = 30
            elif window_height < 600:  # 中等窗口
                input_offset = 50
            else:  # 大窗口
                input_offset = 70
            
            # 输入框位置计算
            input_x = left + window_width // 2
            input_y = bottom - input_offset
            
            # 验证位置是否合理
            screen_width, screen_height = pyautogui.size()
            if (0 <= input_x <= screen_width and 
                0 <= input_y <= screen_height and 
                input_y > top + 100):  # 确保不在窗口标题栏
                
                self.log(f"🎯 计算得出输入框位置: ({input_x}, {input_y})")
                return (input_x, input_y)
            
            return None
            
        except Exception as e:
            self.log(f"⚠️ 窗口计算输入框位置失败: {e}")
            return None
    
    def _smart_click_input_area(self):
        """智能点击输入区域"""
        try:
            if not HAS_WIN32 or not self.wechat_hwnd:
                return False
            
            # 获取微信窗口信息
            rect = win32gui.GetWindowRect(self.wechat_hwnd)
            left, top, right, bottom = rect
            
            # 在窗口底部区域尝试多个点击位置
            click_positions = [
                (left + (right - left) // 2, bottom - 60),  # 窗口中下部
                (left + (right - left) // 2, bottom - 80),  # 稍微往上一点
                (left + (right - left) // 2, bottom - 40),  # 更靠近底部
                (left + (right - left) * 3 // 4, bottom - 60),  # 右侧区域
                (left + (right - left) // 4, bottom - 60),   # 左侧区域
            ]
            
            for i, (x, y) in enumerate(click_positions):
                try:
                    # 确保点击位置在屏幕范围内
                    screen_width, screen_height = pyautogui.size()
                    if 0 <= x <= screen_width and 0 <= y <= screen_height:
                        pyautogui.click(x, y)
                        time.sleep(0.3)
                        
                        # 测试是否点击成功（尝试输入测试字符）
                        test_char = "t"
                        pyautogui.typewrite(test_char)
                        time.sleep(0.2)
                        
                        # 如果能删除测试字符，说明点击成功
                        pyautogui.press('backspace')
                        time.sleep(0.2)
                        
                        self.log(f"✅ 智能点击成功: 位置 {i+1} ({x}, {y})")
                        return True
                        
                except Exception:
                    continue
            
            return False
            
        except Exception as e:
            self.log(f"⚠️ 智能点击失败: {e}")
            return False
    
    
    def stop_sending_orders(self):
        """停止发送订单"""
        if self.is_sending:
            self.stop_sending = True
            self.log("⏹️ 正在停止发送...")
            self.status_var.set("正在停止...")
        else:
            self.log("ℹ️ 当前没有发送任务")
    
    def run(self):
        """运行程序"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.log("👋 程序被用户中断")


def main():
    """主函数"""
    try:
        print("🚀 启动终极微信发送器...")
        print("使用最直接的模拟用户操作方案")
        
        # 检查依赖
        missing_deps = []
        
        if not HAS_AUTO:
            missing_deps.append("pyautogui pyperclip")
        
        if not HAS_WIN32:
            missing_deps.append("pywin32")
        
        try:
            import pandas
        except ImportError:
            missing_deps.append("pandas")
        
        if missing_deps:
            print("⚠️ 缺少以下依赖包:")
            for dep in missing_deps:
                print(f"  - {dep}")
            print("\n建议运行: pip install pandas pyautogui pyperclip pywin32 openpyxl")
        
        # 启动程序
        app = UltimateWeChatSender()
        app.run()
        
    except Exception as e:
        print(f"❌ 程序启动失败: {str(e)}")
        input("按回车键退出...")


if __name__ == "__main__":
    main()
