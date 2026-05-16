import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread
from ..modules.csv_to_report import CsvToReportConverter
from ..modules.report_to_85 import ReportTo85Converter
from ..modules.header_fill import HeaderFiller
from ..modules.auto_plot import AutoPlotter
from ..modules.plot_summary import PlotSummary
from ..modules.folder_integration import FolderIntegrator
from ..modules.storage_calculation import StorageCalculator
from ..modules.data_check import DataChecker
from ..utils import get_all_csv_files, backup_file, log_message
from ..config import DEFAULT_DATA_DIR, DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_DIR

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("断面数据处理系统")
        self.root.geometry("900x700")
        
        self.data_dir = DEFAULT_DATA_DIR
        self.output_dir = DEFAULT_OUTPUT_DIR
        self.template_dir = DEFAULT_TEMPLATE_DIR
        self.root_dir = os.path.dirname(DEFAULT_DATA_DIR)
        
        self.setup_ui()
        
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.create_settings_tab()
        self.create_conversion_tab()
        self.create_processing_tab()
        self.create_check_tab()
        self.create_summary_tab()
        
    def create_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="路径设置")
        
        ttk.Label(tab, text="请选择包含'断面'文件夹的根目录，程序会自动查找模板并创建成果文件夹", 
                  font=('Arial', 10, 'bold')).pack(pady=5)
        
        settings_frame = ttk.LabelFrame(tab, text="目录设置")
        settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        row = 0
        
        ttk.Label(settings_frame, text="根目录:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.root_dir_entry = ttk.Entry(settings_frame, width=60)
        self.root_dir_entry.insert(0, self.root_dir)
        self.root_dir_entry.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(settings_frame, text="浏览", command=self.browse_root_dir).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        ttk.Label(settings_frame, text="断面数据目录:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.data_dir_entry = ttk.Entry(settings_frame, width=60)
        self.data_dir_entry.insert(0, self.data_dir)
        self.data_dir_entry.config(state='readonly')
        self.data_dir_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1
        
        ttk.Label(settings_frame, text="模板文件目录:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.template_dir_entry = ttk.Entry(settings_frame, width=60)
        self.template_dir_entry.insert(0, self.template_dir)
        self.template_dir_entry.config(state='readonly')
        self.template_dir_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1
        
        ttk.Label(settings_frame, text="成果输出目录:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.output_dir_entry = ttk.Entry(settings_frame, width=60)
        self.output_dir_entry.insert(0, self.output_dir)
        self.output_dir_entry.config(state='readonly')
        self.output_dir_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1
        
        self.apply_button = ttk.Button(tab, text="应用设置", command=self.apply_settings)
        self.apply_button.pack(pady=10)
        
        self.settings_status = ttk.Label(tab, text="", foreground="green")
        self.settings_status.pack(pady=5)
        
        self.verify_button = ttk.Button(tab, text="验证路径", command=self.verify_paths)
        self.verify_button.pack(pady=5)
        
        self.verify_result = tk.Text(tab, height=8, width=80, state=tk.DISABLED)
        self.verify_result.pack(padx=10, pady=5)
    
    def browse_root_dir(self):
        path = filedialog.askdirectory(title="选择根目录（包含断面和模板文件夹）", initialdir=self.root_dir)
        if path:
            self.root_dir_entry.delete(0, tk.END)
            self.root_dir_entry.insert(0, path)
            
            self.data_dir = os.path.join(path, '断面')
            self.template_dir = os.path.join(path, '模板')
            self.output_dir = os.path.join(path, '成果')
            
            self.data_dir_entry.config(state='normal')
            self.template_dir_entry.config(state='normal')
            self.output_dir_entry.config(state='normal')
            
            self.data_dir_entry.delete(0, tk.END)
            self.data_dir_entry.insert(0, self.data_dir)
            self.template_dir_entry.delete(0, tk.END)
            self.template_dir_entry.insert(0, self.template_dir)
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, self.output_dir)
            
            self.data_dir_entry.config(state='readonly')
            self.template_dir_entry.config(state='readonly')
            self.output_dir_entry.config(state='readonly')
    
    def apply_settings(self):
        self.root_dir = self.root_dir_entry.get()
        self.data_dir = os.path.join(self.root_dir, '断面')
        self.template_dir = os.path.join(self.root_dir, '模板')
        self.output_dir = os.path.join(self.root_dir, '成果')
        
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.template_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
        
        self.csv_path_label.config(text="源目录: " + self.data_dir)
        
        self.settings_status.config(text="设置已应用！", foreground="green")
        
        settings = f"""# 断面数据处理系统配置文件
ROOT_DIR={self.root_dir}
DATA_DIR={self.data_dir}
OUTPUT_DIR={self.output_dir}
TEMPLATE_DIR={self.template_dir}
"""
        config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'settings.ini')
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(settings)
    
    def save_settings(self):
        self.apply_settings()
    
    def verify_paths(self):
        self.verify_result.config(state=tk.NORMAL)
        self.verify_result.delete(1.0, tk.END)
        
        results = []
        
        data_path = self.data_dir_entry.get()
        if os.path.exists(data_path):
            csv_count = len([f for f in os.listdir(data_path) if f.endswith('.csv')])
            results.append(f"✓ 数据源目录: {data_path} (包含 {csv_count} 个CSV文件)")
        else:
            results.append(f"✗ 数据源目录不存在: {data_path}")
        
        output_path = self.output_dir_entry.get()
        if os.path.exists(output_path):
            xlsx_count = len([f for f in os.listdir(output_path) if f.endswith('.xlsx')])
            results.append(f"✓ 输出目录: {output_path} (包含 {xlsx_count} 个Excel文件)")
        else:
            results.append(f"⚠ 输出目录不存在，将自动创建: {output_path}")
        
        template_path = self.template_dir_entry.get()
        if os.path.exists(template_path):
            templates = ['横断面成果表模板.xlsx', '纵断面模板.xlsx', '成图模板.xlsx', '对应表.xlsx', '对应表.xlsm']
            found = [t for t in templates if os.path.exists(os.path.join(template_path, t))]
            missing = [t for t in templates if not os.path.exists(os.path.join(template_path, t))]
            results.append(f"✓ 模板目录: {template_path}")
            if found:
                results.append(f"  已找到模板: {', '.join(found)}")
            if missing:
                results.append(f"  缺失模板: {', '.join(missing)}")
        else:
            results.append(f"✗ 模板目录不存在: {template_path}")
        
        corr_path = self.correspondence_entry.get()
        if corr_path and os.path.exists(corr_path):
            results.append(f"✓ 对应表文件: {corr_path}")
        elif corr_path:
            results.append(f"✗ 对应表文件不存在: {corr_path}")
        else:
            results.append(f"⚠ 未指定对应表文件")
        
        self.verify_result.insert(tk.END, "\n".join(results))
        self.verify_result.config(state=tk.DISABLED)
    
    def create_conversion_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="CSV转成果表")
        
        frame = ttk.LabelFrame(tab, text="选择CSV文件")
        frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.csv_path_label = ttk.Label(frame, text="源目录: " + self.data_dir)
        self.csv_path_label.pack(side=tk.LEFT, padx=5)
        
        self.browse_button = ttk.Button(frame, text="浏览", command=self.browse_csv)
        self.browse_button.pack(side=tk.RIGHT, padx=5)
        
        self.csv_files_list = tk.Listbox(tab, width=100, height=15)
        self.csv_files_list.pack(padx=10, pady=5)
        
        self.load_csv_button = ttk.Button(tab, text="加载CSV文件", command=self.load_csv_files)
        self.load_csv_button.pack(pady=5)
        
        self.convert_button = ttk.Button(tab, text="开始转换", command=self.start_conversion)
        self.convert_button.pack(pady=5)
        
        self.progress_text = tk.Text(tab, height=10, width=100, state=tk.DISABLED)
        self.progress_text.pack(padx=10, pady=5)
        
    def create_processing_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="成果处理")
        
        self.process_buttons_frame = ttk.LabelFrame(tab, text="处理操作")
        self.process_buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.convert_85_button = ttk.Button(self.process_buttons_frame, text="转85高程", command=self.convert_to_85)
        self.convert_85_button.grid(row=0, column=0, padx=5, pady=5)
        
        self.fill_header_button = ttk.Button(self.process_buttons_frame, text="填写表头", command=self.fill_header)
        self.fill_header_button.grid(row=0, column=1, padx=5, pady=5)
        
        self.auto_plot_button = ttk.Button(self.process_buttons_frame, text="自动成图", command=self.auto_plot)
        self.auto_plot_button.grid(row=1, column=0, padx=5, pady=5)
        
        self.export_txt_button = ttk.Button(self.process_buttons_frame, text="导出TXT", command=self.export_txt)
        self.export_txt_button.grid(row=1, column=1, padx=5, pady=5)
        
        self.merge_button = ttk.Button(self.process_buttons_frame, text="整合合并", command=self.integrate_merge)
        self.merge_button.grid(row=2, column=0, padx=5, pady=5)
        
        self.storage_button = ttk.Button(self.process_buttons_frame, text="库容计算", command=self.calculate_storage)
        self.storage_button.grid(row=2, column=1, padx=5, pady=5)
        
        self.process_progress = tk.Text(tab, height=15, width=100, state=tk.DISABLED)
        self.process_progress.pack(padx=10, pady=5)
        
    def create_check_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="数据检查")
        
        self.check_buttons_frame = ttk.LabelFrame(tab, text="检查操作")
        self.check_buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.check_section_button = ttk.Button(self.check_buttons_frame, text="检查断面", command=self.check_section)
        self.check_section_button.grid(row=0, column=0, padx=5, pady=5)
        
        self.find_empty_button = ttk.Button(self.check_buttons_frame, text="查找空白断面", command=self.find_empty_sections)
        self.find_empty_button.grid(row=0, column=1, padx=5, pady=5)
        
        self.check_roughness_button = ttk.Button(self.check_buttons_frame, text="检查糙率", command=self.check_roughness)
        self.check_roughness_button.grid(row=0, column=2, padx=5, pady=5)
        
        self.generate_report_button = ttk.Button(self.check_buttons_frame, text="生成报告", command=self.generate_check_report)
        self.generate_report_button.grid(row=1, column=0, padx=5, pady=5)
        
        self.check_results = tk.Text(tab, height=20, width=100, state=tk.DISABLED)
        self.check_results.pack(padx=10, pady=5)
        
    def create_summary_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="统计摘要")
        
        self.summary_frame = ttk.LabelFrame(tab, text="处理统计")
        self.summary_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.stats_labels = {
            'total': ttk.Label(self.summary_frame, text="总文件数: 0"),
            'success': ttk.Label(self.summary_frame, text="成功: 0"),
            'failed': ttk.Label(self.summary_frame, text="失败: 0")
        }
        
        row = 0
        for key, label in self.stats_labels.items():
            label.grid(row=row, column=0, sticky=tk.W, padx=20, pady=5)
            row += 1
        
        self.backup_button = ttk.Button(tab, text="备份数据", command=self.backup_data)
        self.backup_button.pack(pady=10)
        
        self.open_output_button = ttk.Button(tab, text="打开成果目录", command=self.open_output_dir)
        self.open_output_button.pack(pady=5)
        
    def browse_csv(self):
        folder = filedialog.askdirectory(initialdir=self.data_dir)
        if folder:
            self.data_dir = folder
            self.csv_path_label.config(text="源目录: " + folder)
            self.data_dir_entry.delete(0, tk.END)
            self.data_dir_entry.insert(0, folder)
        
    def load_csv_files(self):
        self.csv_files_list.delete(0, tk.END)
        csv_files = get_all_csv_files(self.data_dir)
        for f in csv_files:
            self.csv_files_list.insert(tk.END, os.path.basename(f))
        
    def update_progress(self, text_widget, message):
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, message + "\n")
        text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)
        self.root.update_idletasks()
        
    def start_conversion(self):
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.config(state=tk.DISABLED)
        
        def run_conversion():
            converter = CsvToReportConverter(output_dir=self.output_dir, template_dir=self.template_dir)
            csv_files = get_all_csv_files(self.data_dir)
            
            def progress_callback(msg):
                self.update_progress(self.progress_text, msg)
            
            results = converter.process_all(csv_files, progress_callback)
            
            self.update_progress(self.progress_text, f"\n转换完成！")
            self.update_progress(self.progress_text, f"成功: {len(results['success'])}")
            self.update_progress(self.progress_text, f"失败: {len(results['failed'])}")
            
            self.update_stats(results)
            
        Thread(target=run_conversion).start()
        
    def convert_to_85(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_conversion():
            converter = ReportTo85Converter(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx') and not f.endswith('_成图.xlsx')]
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = converter.process_all(report_files, progress_callback)
            
            self.update_progress(self.process_progress, f"\n转换完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_conversion).start()
        
    def fill_header(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_fill():
            filler = HeaderFiller(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = filler.process_all(report_files, progress_callback)
            
            self.update_progress(self.process_progress, f"\n填写完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_fill).start()
        
    def auto_plot(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_plot():
            plotter = AutoPlotter(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = plotter.process_all(report_files, progress_callback)
            
            self.update_progress(self.process_progress, f"\n成图完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_plot).start()
        
    def export_txt(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_export():
            exporter = PlotSummary(output_dir=self.output_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = exporter.process_all(report_files, progress_callback)
            
            self.update_progress(self.process_progress, f"\n导出完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_export).start()
        
    def integrate_merge(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_integrate():
            integrator = FolderIntegrator(output_dir=self.output_dir, template_dir=self.template_dir)
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = integrator.process_all(self.output_dir, progress_callback)
            
            self.update_progress(self.process_progress, f"\n整合完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_integrate).start()
        
    def calculate_storage(self):
        self.process_progress.config(state=tk.NORMAL)
        self.process_progress.delete(1.0, tk.END)
        self.process_progress.config(state=tk.DISABLED)
        
        def run_calculate():
            calculator = StorageCalculator(output_dir=self.output_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.startswith('K') and f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.update_progress(self.process_progress, msg)
            
            results = calculator.process_all(report_files, progress_callback)
            
            self.update_progress(self.process_progress, f"\n计算完成！")
            self.update_progress(self.process_progress, f"成功: {len(results['success'])}")
            self.update_progress(self.process_progress, f"失败: {len(results['failed'])}")
            
        Thread(target=run_calculate).start()
        
    def check_section(self):
        self.check_results.config(state=tk.NORMAL)
        self.check_results.delete(1.0, tk.END)
        self.check_results.config(state=tk.DISABLED)
        
        def run_check():
            checker = DataChecker()
            csv_files = get_all_csv_files(self.data_dir)
            
            def progress_callback(msg):
                self.update_progress(self.check_results, msg)
            
            for csv_file in csv_files:
                checker.check_section(csv_file, progress_callback)
            
            report_path = checker.generate_report()
            if report_path:
                self.update_progress(self.check_results, f"\n报告已生成: {report_path}")
            
        Thread(target=run_check).start()
        
    def find_empty_sections(self):
        self.check_results.config(state=tk.NORMAL)
        self.check_results.delete(1.0, tk.END)
        self.check_results.config(state=tk.DISABLED)
        
        def run_find():
            checker = DataChecker()
            csv_files = get_all_csv_files(self.data_dir)
            
            def progress_callback(msg):
                self.update_progress(self.check_results, msg)
            
            empty_files = checker.find_empty_sections(csv_files, progress_callback)
            
            if empty_files:
                self.update_progress(self.check_results, f"\n发现空白断面:")
                for f in empty_files:
                    self.update_progress(self.check_results, f"  - {f}")
            else:
                self.update_progress(self.check_results, f"\n未发现空白断面")
            
        Thread(target=run_find).start()
        
    def check_roughness(self):
        self.check_results.config(state=tk.NORMAL)
        self.check_results.delete(1.0, tk.END)
        self.check_results.config(state=tk.DISABLED)
        
        def run_check():
            checker = DataChecker()
            csv_files = get_all_csv_files(DATA_DIR)
            
            def progress_callback(msg):
                self.update_progress(self.check_results, msg)
            
            for csv_file in csv_files:
                checker.check_roughness(csv_file, progress_callback)
            
            report_path = checker.generate_report()
            if report_path:
                self.update_progress(self.check_results, f"\n报告已生成: {report_path}")
            
        Thread(target=run_check).start()
        
    def generate_check_report(self):
        checker = DataChecker()
        csv_files = get_all_csv_files(DATA_DIR)
        
        for csv_file in csv_files:
            checker.check_section(csv_file)
        
        report_path = checker.generate_report()
        if report_path:
            messagebox.showinfo("报告生成", f"报告已生成: {report_path}")
            os.startfile(report_path)
        else:
            messagebox.showerror("错误", "生成报告失败")
            
    def update_stats(self, results):
        self.stats_labels['total'].config(text=f"总文件数: {results['total']}")
        self.stats_labels['success'].config(text=f"成功: {len(results['success'])}")
        self.stats_labels['failed'].config(text=f"失败: {len(results['failed'])}")
        
    def backup_data(self):
        files_to_backup = []
        
        for root, dirs, files in os.walk(DATA_DIR):
            for file in files:
                files_to_backup.append(os.path.join(root, file))
        
        for f in files_to_backup[:10]:
            backup_file(f)
        
        messagebox.showinfo("备份完成", f"已备份 {min(len(files_to_backup), 10)} 个文件")
        
    def open_output_dir(self):
        os.startfile(OUTPUT_DIR)

def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()