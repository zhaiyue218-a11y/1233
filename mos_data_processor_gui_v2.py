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

class MOSDataProcessorV2:
    def __init__(self, root):
        self.root = root
        self.root.title("半导体测试数据处理工具 - MOS特性分析")
        self.root.geometry("1400x950")
        self.root.resizable(True, True)
        self.root.configure(bg=COLORS['bg_light'])
        
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
    
    def create_widgets(self):
        # 创建主框架 - 使用PanedWindow实现可调整的分割
        main_frame = tk.Frame(self.root, bg=COLORS['bg_light'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========== 顶部标题栏 ==========
        header_frame = tk.Frame(main_frame, bg=COLORS['bg_dark'], height=50)
        header_frame.pack(fill=tk.X, pady=5)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="◈ 半导体测试数据处理工具", 
                font=('Microsoft YaHei', 14, 'bold'),
                bg=COLORS['bg_dark'], fg='white').pack(side=tk.LEFT, padx=15, pady=10)
        
        tk.Label(header_frame, text="MOS特性分析", 
                font=('Microsoft YaHei', 11),
                bg=COLORS['bg_dark'], fg='#94A3B8').pack(side=tk.LEFT, padx=15, pady=10)
        
        # ========== 工具栏 ==========
        toolbar_frame = tk.Frame(main_frame, bg='white', highlightbackground=COLORS['border'],
                                highlightthickness=1)
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # 文件信息
        file_info_frame = tk.Frame(toolbar_frame, bg='white', padx=15, pady=10)
        file_info_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(file_info_frame, text="当前文件:", font=('Microsoft YaHei', 10),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
        
        self.file_label = tk.Label(file_info_frame, text="未选择文件", 
                                  font=('Microsoft YaHei', 10, 'bold'),
                                  bg='white', fg=COLORS['text_primary'])
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # 按钮组
        btn_frame = tk.Frame(toolbar_frame, bg='white', padx=10, pady=6)
        btn_frame.pack(side=tk.RIGHT)
        
        # 创建样式化的按钮
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
                          relief='flat', padx=12, pady=5, cursor='hand2',
                          activebackground=self._darken_color(color))
            btn.pack(side=tk.LEFT, padx=3)
        
        # ========== 中间内容区域 - 仅参数区域 ==========
        right_frame = tk.Frame(main_frame, bg=COLORS['bg_light'])
        right_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 器件参数卡片
        param_card = tk.Frame(right_frame, bg='white', highlightbackground=COLORS['border'],
                             highlightthickness=1)
        param_card.pack(fill=tk.X, pady=5)
        
        param_header = tk.Frame(param_card, bg='white', padx=12, pady=10)
        param_header.pack(fill=tk.X)
        
        tk.Frame(param_header, bg=COLORS['secondary'], width=4, height=18).pack(side=tk.LEFT, padx=8)
        tk.Label(param_header, text="器件参数", font=('Microsoft YaHei', 12, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT)
        
        param_content = tk.Frame(param_card, bg='white', padx=12, pady=12)
        param_content.pack(fill=tk.X)
        
        # 参数输入
        params = [
            ("沟道宽度 W", "μm", "w_entry", self.W),
            ("沟道长度 L", "μm", "l_entry", self.L),
            ("栅氧电容 Cox", "F/cm²", "cox_entry", self.Cox),
        ]
        
        for i, (label, unit, attr, default) in enumerate(params):
            row = tk.Frame(param_content, bg='white')
            row.pack(fill=tk.X, pady=4)
            
            tk.Label(row, text=f"{label}:", font=('Microsoft YaHei', 10),
                    bg='white', fg=COLORS['text_primary'], width=11, anchor='w').pack(side=tk.LEFT)
            
            entry = tk.Entry(row, font=('Microsoft YaHei', 10), width=10,
                           relief='solid', bd=1, highlightthickness=1,
                           highlightcolor=COLORS['primary'])
            entry.insert(0, str(default))
            entry.pack(side=tk.LEFT, padx=4)
            setattr(self, attr, entry)
            
            tk.Label(row, text=unit, font=('Microsoft YaHei', 9),
                    bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
        
        # 计算按钮
        calc_btn = tk.Button(param_content, text="计算电学性能", command=self.calculate_parameters,
                            bg=COLORS['primary'], fg='white', font=('Microsoft YaHei', 10, 'bold'),
                            relief='flat', padx=15, pady=6, cursor='hand2',
                            activebackground=COLORS['primary_hover'])
        calc_btn.pack(pady=12)
        
        # 计算结果卡片 - 使用固定高度确保显示
        result_card = tk.Frame(right_frame, bg='white', highlightbackground=COLORS['border'],
                              highlightthickness=1, height=280)
        result_card.pack(fill=tk.X, pady=5)
        result_card.pack_propagate(False)
        
        result_header = tk.Frame(result_card, bg='white', padx=12, pady=10)
        result_header.pack(fill=tk.X)
        
        tk.Frame(result_header, bg=COLORS['warning'], width=4, height=18).pack(side=tk.LEFT, padx=8)
        tk.Label(result_header, text="电学性能参数", font=('Microsoft YaHei', 12, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT)
        
        # 创建带滚动条的结果区域
        result_container = tk.Frame(result_card, bg='white')
        result_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        result_canvas = tk.Canvas(result_container, bg='white', highlightthickness=0)
        result_scrollbar = ttk.Scrollbar(result_container, orient=tk.VERTICAL, command=result_canvas.yview)
        result_content = tk.Frame(result_canvas, bg='white')
        
        result_canvas.configure(yscrollcommand=result_scrollbar.set)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        result_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        result_canvas.create_window((0, 0), window=result_content, anchor='nw', width=340)
        
        def configure_result_canvas(event):
            result_canvas.configure(scrollregion=result_canvas.bbox('all'))
        result_content.bind('<Configure>', configure_result_canvas)
        
        # 创建结果变量
        self.vth_var = tk.StringVar(value="--")
        self.ss_var = tk.StringVar(value="--")
        self.mobility_var = tk.StringVar(value="--")
        self.idsat_var = tk.StringVar(value="--")
        self.gm_var = tk.StringVar(value="--")
        
        # 参数显示
        results = [
            ("阈值电压 Vth", self.vth_var, "V", COLORS['primary']),
            ("亚阈值斜率 SS", self.ss_var, "V/dec", COLORS['secondary']),
            ("迁移率 μFE", self.mobility_var, "cm²/V·s", COLORS['warning']),
            ("饱和电流 Idsat", self.idsat_var, "A", COLORS['danger']),
            ("跨导 gm", self.gm_var, "(A^0.5)/V", COLORS['primary']),
        ]
        
        for label, var, unit, color in results:
            self._create_param_card(result_content, label, var, unit, color)
        
        # ========== 底部图表区域 ==========
        plot_card = tk.Frame(main_frame, bg='white', highlightbackground=COLORS['border'],
                            highlightthickness=1, height=320)
        plot_card.pack(fill=tk.X, pady=5)
        plot_card.pack_propagate(False)
        
        plot_header = tk.Frame(plot_card, bg='white', padx=12, pady=10)
        plot_header.pack(fill=tk.X)
        
        tk.Frame(plot_header, bg=COLORS['danger'], width=4, height=18).pack(side=tk.LEFT, padx=8)
        tk.Label(plot_header, text="转移特性曲线", font=('Microsoft YaHei', 12, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT)
        
        # 图表区域
        plot_frame = tk.Frame(plot_card, bg='white', padx=12, pady=12)
        plot_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建图形 - 使用响应式尺寸
        self.fig = plt.Figure(figsize=(8, 3.5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 初始化图表样式
        self._setup_chart_style()
        
        # ========== 状态栏 ==========
        status_frame = tk.Frame(main_frame, bg='white', height=28)
        status_frame.pack(fill=tk.X, pady=5)
        status_frame.pack_propagate(False)
        
        # 状态指示器
        self.status_indicator = tk.Canvas(status_frame, width=10, height=10, 
                                         bg='white', highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=12)
        self._draw_status_circle('ready')
        
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(status_frame, textvariable=self.status_var,
                font=('Microsoft YaHei', 9),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT)
    
    def _darken_color(self, color, factor=0.85):
        """使颜色变暗"""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(int(c * factor) for c in rgb)
        return f'#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}'
    
    def _create_param_card(self, parent, label, value_var, unit, color):
        """创建参数显示卡片"""
        card = tk.Frame(parent, bg='white', highlightbackground=COLORS['border'],
                       highlightthickness=1)
        card.pack(fill=tk.X, pady=3, padx=5)
        
        # 顶部色条
        tk.Frame(card, bg=color, height=3).pack(fill=tk.X)
        
        content = tk.Frame(card, bg='white', padx=12, pady=10)
        content.pack(fill=tk.X)
        
        tk.Label(content, text=label, font=('Microsoft YaHei', 9),
                bg='white', fg=COLORS['text_secondary']).pack(anchor='w')
        
        value_frame = tk.Frame(content, bg='white')
        value_frame.pack(fill=tk.X, pady=3)
        
        tk.Label(value_frame, textvariable=value_var, font=('Microsoft YaHei', 16, 'bold'),
                bg='white', fg=COLORS['text_primary']).pack(side=tk.LEFT)
        
        tk.Label(value_frame, text=f" {unit}", font=('Microsoft YaHei', 9),
                bg='white', fg=COLORS['text_secondary']).pack(side=tk.LEFT, pady=5)
    
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
        self.ax.tick_params(axis='both', direction='in', length=6, width=1.5)
        
        for spine in self.ax.spines.values():
            spine.set_linewidth(1.5)
            spine.set_color(COLORS['border'])
        
        self.ax.set_yscale('log')
        self.ax.set_xlabel('V$_{GS}$ (V)', fontsize=11, fontweight='bold')
        self.ax.set_ylabel('I$_{DS}$ (A)', fontsize=11, fontweight='bold')
        self.ax.grid(True, alpha=0.3, linestyle='--', color=COLORS['border'])
        self.canvas.draw()
    
    def select_file(self):
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
        try:
            # 清除坐标轴
            self.ax.clear()
            
            # 重新设置图表样式
            self.ax.set_facecolor('white')
            self.ax.tick_params(axis='both', direction='in', length=6, width=1.5)
            
            for spine in self.ax.spines.values():
                spine.set_linewidth(1.5)
                spine.set_color('#E2E8F0')
            
            self.ax.grid(True, linestyle='--', alpha=0.3, color='#94A3B8')
            self.ax.set_xlabel('Vgs (V)', fontsize=11, fontweight='bold')
            self.ax.set_ylabel('Ids (A)', fontsize=11, fontweight='bold')
            
            # 绘制曲线
            self.ax.plot(vg, np.abs(id_raw), color=COLORS['primary'], linewidth=2.5, marker='o', markersize=3)
            self.ax.set_title('MOS转移特性曲线', fontsize=13, fontweight='bold', pad=15)
            
            # 使用更安全的布局方式
            self.fig.tight_layout(pad=2.0)
            
            # 强制刷新画布
            self.canvas.draw()
            self.canvas.flush_events()
            
            # 更新状态
            self.root.update_idletasks()
        except Exception as e:
            messagebox.showerror("错误", f"绘图时出错: {e}")
    
    def calculate_parameters(self):
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
            
            self.vth_var.set(f"{self.vth:.4f}")
            self.ss_var.set(f"{self.ss:.4f}")
            self.mobility_var.set(f"{self.mobility:.2f}")
            self.idsat_var.set(f"{self.idsat:.4e}")
            self.gm_var.set(f"{gm:.4e}")
            
            self._set_status('电学参数计算完成', 'ready')
            messagebox.showinfo("成功", "电学性能参数计算完成！")
            
        except Exception as e:
            self._set_status('计算失败', 'error')
            messagebox.showerror("错误", f"计算电学性能时出错: {e}")
    
    def _calculate_electrical_parameters(self, vg, id_raw, W, L, Cox):
        """计算MOS器件电学参数"""
        abs_id = np.abs(id_raw)
        
        # 阈值电压
        target_id = 1e-8
        diff = np.abs(abs_id - target_id)
        closest_idx = diff.idxmin()
        vth = vg.iloc[closest_idx]
        
        # 亚阈值斜率
        target_id1, target_id2 = 1e-8, 1e-10
        diff1 = np.abs(abs_id - target_id1)
        diff2 = np.abs(abs_id - target_id2)
        vg1 = vg.iloc[diff1.idxmin()]
        vg2 = vg.iloc[diff2.idxmin()]
        ss = (vg1 - vg2) / 2
        
        # 迁移率
        sqrt_id = np.sqrt(abs_id)
        vth_range = 5.0
        mask = (vg >= vth - vth_range) & (vg <= vth + vth_range)
        linear_vg = vg[mask] if mask.sum() >= 3 else vg
        linear_sqrt_id = sqrt_id[mask] if mask.sum() >= 3 else sqrt_id
        
        coefficients = np.polyfit(linear_vg, linear_sqrt_id, 1)
        slope = coefficients[0]
        
        W_cm, L_cm = W * 1e-4, L * 1e-4
        mobility = (2 * L_cm) / (W_cm * Cox) * (slope ** 2) / 1e-8
        
        # 饱和电流
        max_vg = vg.max()
        idsat = (W / (2 * L)) * Cox * mobility * ((max_vg - vth) ** 2)
        
        return {
            'vth': vth,
            'ss': ss,
            'mobility': mobility,
            'idsat': idsat,
            'gm': slope
        }
    
    def select_folder(self):
        try:
            folder_path = filedialog.askdirectory(title="选择文件夹")
            if not folder_path:
                return
            
            self._set_status('正在批量处理...', 'busy')
            
            csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
            
            if not csv_files:
                messagebox.showinfo("提示", "该文件夹下没有CSV文件")
                self._set_status('就绪', 'ready')
                return
            
            output_folder = os.path.join(folder_path, 'processed_results')
            os.makedirs(output_folder, exist_ok=True)
            
            all_params = []
            processed_count = 0
            
            for csv_file in csv_files:
                try:
                    file_path = os.path.join(folder_path, csv_file)
                    data = pd.read_csv(file_path, skiprows=1)
                    
                    vg = data.iloc[:, 0]
                    id_raw = data.iloc[:, 1]
                    
                    params = self._calculate_electrical_parameters(vg, id_raw, self.W, self.L, self.Cox)
                    
                    base_name = os.path.splitext(csv_file)[0]
                    
                    # 保存参数
                    params_csv = os.path.join(output_folder, f'{base_name}_parameters.csv')
                    params_df = pd.DataFrame({
                        'Parameter': ['Vth (V)', 'SS (V/dec)', 'Mobility (cm²/V·s)',
                                     'Idsat (A)', 'gm ((A^0.5)/V)'],
                        'Value': [params['vth'], params['ss'], params['mobility'],
                                 params['idsat'], params['gm']]
                    })
                    params_df.to_csv(params_csv, index=False)
                    
                    # 保存图表
                    fig = plt.Figure(figsize=(8, 6), dpi=150)
                    ax = fig.add_subplot(111)
                    ax.set_facecolor('white')
                    ax.tick_params(axis='both', direction='in', length=6, width=1.5)
                    ax.set_yscale('log')
                    ax.plot(vg, np.abs(id_raw), color=COLORS['primary'], linewidth=2.5)
                    ax.set_xlabel('V$_{GS}$(V)', fontsize=12)
                    ax.set_ylabel('I$_{DS}$(A)', fontsize=12)
                    
                    output_jpg = os.path.join(output_folder, f'{base_name}_transfer_curve.jpg')
                    fig.savefig(output_jpg, dpi=300, bbox_inches='tight', format='jpg')
                    plt.close(fig)
                    
                    all_params.append({
                        'File': csv_file,
                        'Vth (V)': params['vth'],
                        'SS (V/dec)': params['ss'],
                        'Mobility (cm²/V·s)': params['mobility'],
                        'Idsat (A)': params['idsat'],
                        'gm ((A^0.5)/V)': params['gm']
                    })
                    
                    processed_count += 1
                    
                except Exception as e:
                    print(f"处理文件 {csv_file} 时出错: {e}")
            
            if all_params:
                summary_df = pd.DataFrame(all_params)
                summary_csv = os.path.join(output_folder, 'all_parameters_summary.csv')
                summary_df.to_csv(summary_csv, index=False)
            
            self._set_status(f'批量处理完成 ({processed_count} 个文件)', 'ready')
            messagebox.showinfo("成功", f"批量处理完成！\n共处理 {processed_count} 个文件\n"
                              f"结果已保存至: {output_folder}")
            
        except Exception as e:
            self._set_status('批量处理失败', 'error')
            messagebox.showerror("错误", f"批量处理时出错: {e}")
    
    def export_parameters(self):
        if self.vth is None:
            messagebox.showerror("错误", "请先计算电学性能参数")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                title="导出参数",
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
            
            if file_path:
                params_data = pd.DataFrame({
                    'Parameter': ['Vth (V)', 'SS (V/dec)', 'Mobility (cm²/V·s)',
                                 'Idsat (A)', 'gm ((A^0.5)/V)'],
                    'Value': [self.vth, self.ss, self.mobility, self.idsat, self.gm_var.get()]
                })
                params_data.to_csv(file_path, index=False)
                messagebox.showinfo("成功", f"参数已导出至: {file_path}")
                self._set_status('参数导出完成', 'ready')
                
        except Exception as e:
            messagebox.showerror("错误", f"导出参数时出错: {e}")
    
    def save_results(self):
        if self.processed_data is None:
            messagebox.showerror("错误", "请先处理数据")
            return
        
        try:
            base_name = os.path.splitext(self.file_path)[0]
            output_csv = base_name + '_processed_log.csv'
            output_jpg = base_name + '_transfer_curve_log.jpg'
            
            self.processed_data.to_csv(output_csv, index=False)
            self.fig.savefig(output_jpg, dpi=300, bbox_inches='tight', format='jpg')
            
            messagebox.showinfo("成功", f"结果已保存:\n{output_csv}\n{output_jpg}")
            self._set_status('结果保存完成', 'ready')
            
        except Exception as e:
            messagebox.showerror("错误", f"保存结果时出错: {e}")

def main():
    root = tk.Tk()
    app = MOSDataProcessorV2(root)
    root.mainloop()

if __name__ == "__main__":
    main()
