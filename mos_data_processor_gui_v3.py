import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# 设置Matplotlib中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei', 'Heiti TC']
plt.rcParams['axes.unicode_minus'] = False

# 现代化配色方案
COLORS = {
    'primary': '#2563EB',
    'primary_hover': '#1D4ED8',
    'secondary': '#10B981',
    'danger': '#EF4444',
    'warning': '#F59E0B',
    'bg_dark': '#1E293B',
    'bg_light': '#F8FAFC',
    'text_primary': '#1E293B',
    'text_secondary': '#64748B',
    'border': '#E2E8F0',
}

class MOSDataProcessorV3:
    def __init__(self, root):
        self.root = root
        self.root.title("半导体测试数据处理工具 - MOS特性分析")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        self.root.configure(bg=COLORS['bg_light'])
        self.root.minsize(900, 600)
        
        # 初始化变量
        self.file_path = None
        self.data = None
        self.processed_data = None
        self.fig = None
        self.canvas = None
        self.W = 100
        self.L = 10
        self.Cox = 3.28
        self.vth = None
        self.ss = None
        self.mobility = None
        self.idsat = None
        
        # 创建界面
        self.create_widgets()
        
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', self.on_window_resize)
    
    def create_widgets(self):
        """创建主界面 - 采用专业响应式布局"""
        # 创建主容器 - 使用网格布局实现精确控制
        self.main_container = tk.Frame(self.root, bg=COLORS['bg_light'])
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # 配置网格权重 - 关键：让图表区域获得更多空间
        self.main_container.grid_rowconfigure(2, weight=3)  # 图表区域获得更多权重
        self.main_container.grid_rowconfigure(1, weight=1)  # 参数区域
        self.main_container.grid_columnconfigure(0, weight=1)
        
        # ========== 顶部标题栏 (固定高度) ==========
        self.create_header()
        
        # ========== 工具栏 (固定高度) ==========
        self.create_toolbar()
        
        # ========== 中间内容区域 - 使用PanedWindow实现灵活分割 ==========
        self.create_content_area()
        
        # ========== 底部状态栏 (固定高度) ==========
        self.create_statusbar()
    
    def create_header(self):
        """创建标题栏"""
        header = tk.Frame(self.main_container, bg=COLORS['bg_dark'], height=45)
        header.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        header.pack_propagate(False)
        
        tk.Label(header, text="◈ 半导体测试数据处理工具", 
                font=('Microsoft YaHei', 13, 'bold'),
                bg=COLORS['bg_dark'], fg='white').pack(side=tk.LEFT, padx=15, pady=8)
        
        tk.Label(header, text="MOS特性分析", 
                font=('Microsoft YaHei', 10),
                bg=COLORS['bg_dark'], fg='#94A3B8').pack(side=tk.LEFT, padx=10, pady=8)
    
    def create_toolbar(self):
        """创建工具栏"""
        toolbar = tk.Frame(self.main_container, bg='white', 
                          highlightbackground=COLORS['border'], highlightthickness=1)
        toolbar.grid(row=1, column=0, sticky='ew', pady=(0, 8))
        
        # 文件信息区域
        file_frame = tk.Frame(toolbar, bg='white', padx=12, pady=8)
        file_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(file_frame, text="当前文件:", font=('Microsoft YaHei', 9),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
        
        self.file_label = tk.Label(file_frame, text="未选择文件", 
                                  font=('Microsoft YaHei', 9, 'bold'),
                                  bg='white', fg=COLORS['text_primary'])
        self.file_label.pack(side=tk.LEFT, padx=8)
        
        # 按钮组
        btn_frame = tk.Frame(toolbar, bg='white', padx=8, pady=5)
        btn_frame.pack(side=tk.RIGHT)
        
        buttons = [
            ("选择文件", self.select_file, COLORS['primary']),
            ("处理数据", self.process_data, COLORS['secondary']),
            ("批量处理", self.select_folder, COLORS['warning']),
            ("导出参数", self.export_parameters, COLORS['primary']),
            ("保存结果", self.save_results, COLORS['secondary']),
        ]
        
        for text, cmd, color in buttons:
            btn = tk.Button(btn_frame, text=text, command=cmd,
                          bg=color, fg='white', font=('Microsoft YaHei', 9, 'bold'),
                          relief='flat', padx=10, pady=4, cursor='hand2',
                          activebackground=self._darken_color(color))
            btn.pack(side=tk.LEFT, padx=2)
    
    def create_content_area(self):
        """创建内容区域 - 使用PanedWindow实现参数和图表的灵活布局"""
        # 创建水平分割的PanedWindow
        self.content_paned = tk.PanedWindow(self.main_container, orient=tk.HORIZONTAL,
                                           bg=COLORS['bg_light'], sashwidth=6,
                                           sashrelief=tk.RAISED, sashpad=2)
        self.content_paned.grid(row=2, column=0, sticky='nsew', pady=(0, 8))
        
        # ========== 左侧：参数面板 (可调整宽度) ==========
        self.param_panel = tk.Frame(self.content_paned, bg=COLORS['bg_light'], width=320)
        self.content_paned.add(self.param_panel, minsize=280, stretch="never")
        
        # 器件参数卡片
        self.create_device_params_card()
        
        # 电学性能参数卡片
        self.create_electrical_params_card()
        
        # ========== 右侧：图表面板 (占据剩余空间) ==========
        self.plot_panel = tk.Frame(self.content_paned, bg=COLORS['bg_light'])
        self.content_paned.add(self.plot_panel, minsize=400, stretch="always")
        
        # 转移特性曲线
        self.create_plot_area()
    
    def create_device_params_card(self):
        """创建设备参数卡片"""
        card = tk.Frame(self.param_panel, bg='white', 
                       highlightbackground=COLORS['border'], highlightthickness=1)
        card.pack(fill=tk.X, pady=(0, 8), padx=2)
        
        # 标题
        header = tk.Frame(card, bg='white', padx=10, pady=8)
        header.pack(fill=tk.X)
        
        tk.Frame(header, bg=COLORS['secondary'], width=3, height=16).pack(side=tk.LEFT, padx=6)
        tk.Label(header, text="器件参数", font=('Microsoft YaHei', 11, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT, padx=6)
        
        # 参数内容 - 使用紧凑布局
        content = tk.Frame(card, bg='white', padx=10, pady=8)
        content.pack(fill=tk.X)
        
        # 使用网格布局使参数更紧凑
        params = [
            ("沟道宽度 W", "μm", "w_entry", self.W),
            ("沟道长度 L", "μm", "l_entry", self.L),
            ("栅氧电容 Cox", "F/cm²", "cox_entry", self.Cox),
        ]
        
        for i, (label, unit, attr, default) in enumerate(params):
            row = tk.Frame(content, bg='white')
            row.pack(fill=tk.X, pady=3)
            
            tk.Label(row, text=f"{label}:", font=('Microsoft YaHei', 9),
                    bg='white', fg=COLORS['text_primary'], width=12, anchor='w').pack(side=tk.LEFT)
            
            entry = tk.Entry(row, font=('Microsoft YaHei', 9), width=10,
                           relief='solid', bd=1, highlightthickness=1,
                           highlightcolor=COLORS['primary'])
            entry.insert(0, str(default))
            entry.pack(side=tk.LEFT, padx=4)
            setattr(self, attr, entry)
            
            tk.Label(row, text=unit, font=('Microsoft YaHei', 8),
                    bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
        
        # 计算按钮
        calc_btn = tk.Button(content, text="计算电学性能", command=self.calculate_parameters,
                            bg=COLORS['primary'], fg='white', font=('Microsoft YaHei', 9, 'bold'),
                            relief='flat', padx=12, pady=5, cursor='hand2',
                            activebackground=COLORS['primary_hover'])
        calc_btn.pack(pady=10)
    
    def create_electrical_params_card(self):
        """创建电学性能参数卡片 - 使用2列网格布局节省空间"""
        card = tk.Frame(self.param_panel, bg='white',
                       highlightbackground=COLORS['border'], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=2)
        
        # 标题
        header = tk.Frame(card, bg='white', padx=10, pady=8)
        header.pack(fill=tk.X)
        
        tk.Frame(header, bg=COLORS['warning'], width=3, height=16).pack(side=tk.LEFT, padx=6)
        tk.Label(header, text="电学性能参数", font=('Microsoft YaHei', 11, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT, padx=6)
        
        # 创建带滚动条的参数区域
        container = tk.Frame(card, bg='white')
        container.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        
        # Canvas用于滚动
        canvas = tk.Canvas(container, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        self.params_content = tk.Frame(canvas, bg='white')
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.create_window((0, 0), window=self.params_content, anchor='nw', width=290)
        
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        self.params_content.bind('<Configure>', configure_canvas)
        
        # 创建结果变量
        self.vth_var = tk.StringVar(value="--")
        self.ss_var = tk.StringVar(value="--")
        self.mobility_var = tk.StringVar(value="--")
        self.idsat_var = tk.StringVar(value="--")
        self.gm_var = tk.StringVar(value="--")
        
        # 使用2列网格布局显示参数
        results = [
            ("阈值电压 Vth", self.vth_var, "V", COLORS['primary']),
            ("亚阈值斜率 SS", self.ss_var, "V/dec", COLORS['secondary']),
            ("迁移率 μFE", self.mobility_var, "cm²/V·s", COLORS['warning']),
            ("饱和电流 Idsat", self.idsat_var, "A", COLORS['danger']),
            ("跨导 gm", self.gm_var, "S", COLORS['primary']),
        ]
        
        for i, (label, var, unit, color) in enumerate(results):
            self._create_compact_param_card(self.params_content, label, var, unit, color, i)
    
    def _create_compact_param_card(self, parent, label, value_var, unit, color, index):
        """创建紧凑的参数卡片 - 用于网格布局"""
        card = tk.Frame(parent, bg='white', highlightbackground=COLORS['border'],
                       highlightthickness=1)
        card.pack(fill=tk.X, pady=3, padx=3)
        
        # 顶部色条
        tk.Frame(card, bg=color, height=2).pack(fill=tk.X)
        
        content = tk.Frame(card, bg='white', padx=8, pady=6)
        content.pack(fill=tk.X)
        
        tk.Label(content, text=label, font=('Microsoft YaHei', 8),
                bg='white', fg=COLORS['text_secondary']).pack(anchor='w')
        
        value_frame = tk.Frame(content, bg='white')
        value_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(value_frame, textvariable=value_var, font=('Microsoft YaHei', 13, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT)
        
        tk.Label(value_frame, text=f" {unit}", font=('Microsoft YaHei', 8),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT, pady=3)
    
    def create_plot_area(self):
        """创建图表区域 - 关键：确保图表能自适应大小"""
        # 图表卡片
        card = tk.Frame(self.plot_panel, bg='white',
                       highlightbackground=COLORS['border'], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=2)
        
        # 标题
        header = tk.Frame(card, bg='white', padx=10, pady=8)
        header.pack(fill=tk.X)
        
        tk.Frame(header, bg=COLORS['danger'], width=3, height=16).pack(side=tk.LEFT, padx=6)
        tk.Label(header, text="转移特性曲线", font=('Microsoft YaHei', 11, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT, padx=6)
        
        # 图表容器 - 使用pack填充
        self.plot_container = tk.Frame(card, bg='white', padx=8, pady=8)
        self.plot_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建图形 - 使用相对尺寸
        self.fig = plt.Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        
        # 创建Canvas
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_container)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True)
        
        # 初始化图表样式
        self._setup_chart_style()
        
        # 绑定Canvas大小变化事件
        self.canvas_widget.bind('<Configure>', self.on_canvas_resize)
    
    def create_statusbar(self):
        """创建状态栏"""
        status = tk.Frame(self.main_container, bg='white', height=26,
                         highlightbackground=COLORS['border'], highlightthickness=1)
        status.grid(row=3, column=0, sticky='ew')
        status.pack_propagate(False)
        
        self.status_indicator = tk.Canvas(status, width=10, height=10, 
                                         bg='white', highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=10, pady=6)
        self._draw_status_circle('ready')
        
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(status, textvariable=self.status_var,
                font=('Microsoft YaHei', 9),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
    
    def on_window_resize(self, event):
        """窗口大小变化时的处理"""
        # 延迟更新以避免频繁重绘
        if hasattr(self, '_resize_job'):
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(150, self.update_layout)
    
    def on_canvas_resize(self, event):
        """Canvas大小变化时的处理"""
        if self.fig and self.canvas:
            # 获取新的尺寸（转换为英寸）
            width = event.width / self.fig.dpi
            height = event.height / self.fig.dpi
            
            # 更新图形尺寸
            self.fig.set_size_inches(width, height, forward=True)
            self.fig.tight_layout(pad=1.5)
            self.canvas.draw()
    
    def update_layout(self):
        """更新布局 - 确保所有元素正确显示"""
        if self.fig and self.canvas:
            # 重新调整图表布局
            self.fig.tight_layout(pad=1.5)
            self.canvas.draw()
    
    def _darken_color(self, color, factor=0.85):
        """使颜色变暗"""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(int(c * factor) for c in rgb)
        return f'#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}'
    
    def _draw_status_circle(self, status):
        """绘制状态指示器"""
        self.status_indicator.delete('all')
        colors = {
            'ready': COLORS['secondary'],
            'busy': COLORS['warning'],
            'error': COLORS['danger'],
        }
        color = colors.get(status, COLORS['secondary'])
        self.status_indicator.create_oval(2, 2, 8, 8, fill=color, outline='')
    
    def _set_status(self, text, status='ready'):
        """设置状态"""
        self.status_var.set(text)
        self._draw_status_circle(status)
        self.root.update_idletasks()
    
    def _setup_chart_style(self):
        """设置图表样式"""
        self.ax.set_facecolor('white')
        self.ax.tick_params(axis='both', direction='in', length=5, width=1.2)
        
        for spine in self.ax.spines.values():
            spine.set_linewidth(1.2)
            spine.set_color(COLORS['border'])
        
        self.ax.set_yscale('log')
        self.ax.set_xlabel('V$_{GS}$ (V)', fontsize=10, fontweight='bold')
        self.ax.set_ylabel('I$_{DS}$ (A)', fontsize=10, fontweight='bold')
        self.ax.grid(True, alpha=0.3, linestyle='--', color=COLORS['border'])
        
        # 添加初始提示文字
        self.ax.text(0.5, 0.5, '请选择文件并点击"处理数据"\n以显示转移特性曲线',
                    transform=self.ax.transAxes, ha='center', va='center',
                    fontsize=11, color=COLORS['text_secondary'])
        
        self.fig.tight_layout(pad=1.5)
        self.canvas.draw()
    
    def select_file(self):
        """选择文件"""
        try:
            self.file_path = filedialog.askopenfilename(
                title="选择CSV文件",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
            if self.file_path:
                self.file_label.config(text=os.path.basename(self.file_path))
                self._set_status(f"已选择: {os.path.basename(self.file_path)}")
                self.read_data()
        except Exception as e:
            messagebox.showerror("错误", f"选择文件时出错: {e}")
            self._set_status('错误', 'error')
    
    def read_data(self):
        """读取数据"""
        if not self.file_path:
            return False
        
        try:
            self._set_status('正在读取数据...', 'busy')
            self.data = pd.read_csv(self.file_path, skiprows=1)
            
            vg = self.data.iloc[:, 0]
            id_raw = self.data.iloc[:, 1]
            
            self._set_status(f'已加载 {len(vg)} 个数据点', 'ready')
            return True
            
        except Exception as e:
            self._set_status('读取失败', 'error')
            messagebox.showerror("错误", f"读取文件时出错: {e}")
            return False
    
    def process_data(self):
        """处理数据"""
        if not self.file_path:
            messagebox.showerror("错误", "请先选择文件")
            return
        
        if self.data is None:
            if not self.read_data():
                return
        
        try:
            self._set_status('正在处理数据...', 'busy')
            
            vg = self.data.iloc[:, 0]
            id_raw = self.data.iloc[:, 1]
            abs_id = np.abs(id_raw)
            log_id = np.log10(abs_id + 1e-15)
            
            self.processed_data = pd.DataFrame({
                'Vg': vg,
                'log_abs_Id': log_id
            })
            
            self.plot_curve(vg, id_raw)
            
            self._set_status('数据处理完成', 'ready')
            messagebox.showinfo("成功", "数据处理完成！")
            
        except Exception as e:
            self._set_status('处理失败', 'error')
            messagebox.showerror("错误", f"处理数据时出错: {e}")
    
    def plot_curve(self, vg, id_raw):
        """绘制转移特性曲线"""
        try:
            # 清除坐标轴
            self.ax.clear()
            
            # 重新设置图表样式
            self.ax.set_facecolor('white')
            self.ax.tick_params(axis='both', direction='in', length=5, width=1.2)
            
            for spine in self.ax.spines.values():
                spine.set_linewidth(1.2)
                spine.set_color(COLORS['border'])
            
            self.ax.grid(True, linestyle='--', alpha=0.3, color=COLORS['border'])
            self.ax.set_xlabel('V$_{GS}$ (V)', fontsize=10, fontweight='bold')
            self.ax.set_ylabel('I$_{DS}$ (A)', fontsize=10, fontweight='bold')
            self.ax.set_yscale('log')
            
            # 绘制曲线
            self.ax.plot(vg, np.abs(id_raw), color=COLORS['primary'], 
                        linewidth=2, marker='o', markersize=2, alpha=0.8)
            self.ax.set_title('MOS转移特性曲线', fontsize=12, fontweight='bold', pad=10)
            
            # 调整布局
            self.fig.tight_layout(pad=1.5)
            
            # 强制刷新画布
            self.canvas.draw()
            self.canvas.flush_events()
            
        except Exception as e:
            messagebox.showerror("错误", f"绘图时出错: {e}")
    
    def calculate_parameters(self):
        """计算电学性能参数"""
        if self.data is None:
            messagebox.showerror("错误", "请先选择文件并处理数据")
            return
        
        try:
            self.W = float(self.w_entry.get())
            self.L = float(self.l_entry.get())
            self.Cox = float(self.cox_entry.get())
            
            vg = self.data.iloc[:, 0]
            id_raw = self.data.iloc[:, 1]
            
            params = self._calculate_electrical_parameters(vg, id_raw, self.W, self.L, self.Cox)
            
            self.vth = params['vth']
            self.ss = params['ss']
            self.mobility = params['mobility']
            self.idsat = params['idsat']
            gm = params['gm']
            
            # 更新显示
            self.vth_var.set(f"{self.vth:.4f}")
            self.ss_var.set(f"{self.ss:.4f}")
            self.mobility_var.set(f"{self.mobility:.2f}")
            self.idsat_var.set(f"{self.idsat:.4e}")
            self.gm_var.set(f"{gm:.4e}")
            
            self._set_status('参数计算完成', 'ready')
            messagebox.showinfo("成功", "电学性能参数计算完成！")
            
        except Exception as e:
            self._set_status('计算失败', 'error')
            messagebox.showerror("错误", f"计算参数时出错: {e}")
    
    def _calculate_electrical_parameters(self, vg, id_raw, W, L, Cox):
        """计算电学性能参数"""
        abs_id = np.abs(id_raw)
        
        # 阈值电压 Vth (使用最大跨导法)
        sqrt_id = np.sqrt(abs_id)
        diff_sqrt_id = np.diff(sqrt_id)
        diff_vg = np.diff(vg)
        gm_sqrt = diff_sqrt_id / diff_vg
        
        # 找到最大跨导点
        max_gm_idx = np.argmax(gm_sqrt)
        vg_max_gm = vg.iloc[max_gm_idx]
        sqrt_id_max_gm = sqrt_id.iloc[max_gm_idx]
        
        # 线性外推
        slope = gm_sqrt[max_gm_idx]
        intercept = sqrt_id_max_gm - slope * vg_max_gm
        vth = -intercept / slope
        
        # 亚阈值斜率 SS
        log_id = np.log10(abs_id + 1e-15)
        subthreshold_region = vg < vth
        if np.sum(subthreshold_region) > 5:
            sub_vg = vg[subthreshold_region]
            sub_log_id = log_id[subthreshold_region]
            ss = np.abs(1 / np.polyfit(sub_vg, sub_log_id, 1)[0])
        else:
            ss = 0
        
        # 迁移率 μFE
        mu_fe = (slope ** 2 * 2 * L) / (W * Cox * 1e-4) if Cox > 0 else 0
        
        # 饱和电流 Idsat
        idsat = np.max(abs_id)
        
        # 跨导 gm
        gm = np.max(np.abs(np.diff(abs_id) / np.diff(vg)))
        
        return {
            'vth': vth,
            'ss': ss,
            'mobility': mu_fe,
            'idsat': idsat,
            'gm': gm
        }
    
    def export_parameters(self):
        """导出参数"""
        if self.vth is None:
            messagebox.showerror("错误", "请先计算电学性能参数")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                title="保存参数"
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("MOS器件电学性能参数\\n")
                    f.write("=" * 40 + "\\n")
                    f.write(f"沟道宽度 W: {self.W} μm\\n")
                    f.write(f"沟道长度 L: {self.L} μm\\n")
                    f.write(f"栅氧电容 Cox: {self.Cox} F/cm²\\n")
                    f.write("-" * 40 + "\\n")
                    f.write(f"阈值电压 Vth: {self.vth:.4f} V\\n")
                    f.write(f"亚阈值斜率 SS: {self.ss:.4f} V/dec\\n")
                    f.write(f"迁移率 μFE: {self.mobility:.2f} cm²/V·s\\n")
                    f.write(f"饱和电流 Idsat: {self.idsat:.4e} A\\n")
                    f.write(f"跨导 gm: {self.gm_var.get()} S\\n")
                
                self._set_status('参数已导出', 'ready')
                messagebox.showinfo("成功", f"参数已保存到:\\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出参数时出错: {e}")
    
    def save_results(self):
        """保存结果"""
        if self.processed_data is None:
            messagebox.showerror("错误", "请先处理数据")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")],
                title="保存处理结果"
            )
            if file_path:
                self.processed_data.to_csv(file_path, index=False, encoding='utf-8-sig')
                self._set_status('结果已保存', 'ready')
                messagebox.showinfo("成功", f"结果已保存到:\\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存结果时出错: {e}")
    
    def select_folder(self):
        """批量处理文件夹"""
        messagebox.showinfo("提示", "批量处理功能开发中...")


def main():
    root = tk.Tk()
    app = MOSDataProcessorV3(root)
    root.mainloop()


if __name__ == "__main__":
    main()
