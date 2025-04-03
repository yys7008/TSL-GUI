import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                           QRadioButton, QButtonGroup, QFrame, QTabWidget,
                           QGroupBox, QTextEdit, QScrollArea, QGridLayout,
                           QSpacerItem, QSizePolicy)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette
import pyvisa as visa

class TSL570():
    def __init__(self, model="TSL-570"):
        self.rm = visa.ResourceManager()
        self.device = None
        self.connected = False
        self.model = model  # 添加型号属性，默认为TSL-570
        self.device_info = {
            "model": model,
            "wavelength_range": "",
            "max_power": ""
        }

    def get_model(self):
        """获取当前设备型号"""
        return self.model

    def set_model(self, model):
        """设置设备型号"""
        if model in ["TSL-570", "TSL-550"]:
            self.model = model
            return f"设备型号已设置为: {model}"
        else:
            return f"不支持的设备型号: {model}"

    def search_gpib_addresses(self):
        """搜索可用的 GPIB 设备地址"""
        try:
            return [addr for addr in self.rm.list_resources() if 'GPIB' in addr]
        except:
            return []

    def connect_device(self, address):
        """连接到指定地址的设备"""
        try:
            self.device = self.rm.open_resource(address)
            self.connected = True
            # 连接后立即读取设备信息
            self.get_device_info()
            return f"成功连接到设备: {address}"
        except Exception as e:
            self.connected = False
            return f"连接设备失败: {str(e)}"

    def disconnect(self):
        """断开设备连接"""
        if self.device:
            try:
                self.device.close()
                self.connected = False
                return "设备已断开连接"
            except Exception as e:
                return f"断开连接失败: {str(e)}"
        return "设备未连接"

    def is_connected(self):
        """返回设备连接状态"""
        return self.connected

    def device_shut_down(self):
        """关闭设备"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write("*RST")
            return "设备已关闭"
        except Exception as e:
            return f"关闭设备失败: {str(e)}"

    def device_restart(self):
        """重启设备"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write("*RST")
            return "设备已重启"
        except Exception as e:
            return f"重启设备失败: {str(e)}"

    def set_wavelength(self, wavelength):
        """设置波长"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength {wavelength}")
            return f"波长已设置为 {wavelength}"
        except Exception as e:
            return f"设置波长失败: {str(e)}"

    def set_wave_unit(self, unit):
        """设置波长单位"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":UNIT:WAVelength {unit}")
            return f"波长单位已设置为 {unit}"
        except Exception as e:
            return f"设置波长单位失败: {str(e)}"

    def set_power_status(self, status):
        """设置激光器输出状态"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":POWer:STATe {status}")
            return "激光器输出已" + ("开启" if status == '1' else "关闭")
        except Exception as e:
            return f"设置输出状态失败: {str(e)}"

    def set_power_level(self, power):
        """设置输出功率"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":POWer:LEVel {power}")
            return f"输出功率已设置为 {power}"
        except Exception as e:
            return f"设置功率失败: {str(e)}"

    def set_sweep_mode(self, mode):
        """设置扫描模式
        mode: 
            0: 步进式扫描 单向
            1: 连续式扫描 单向
            2: 步进式扫描 往复
            3: 连续式扫描 往复
        """
        if not self.connected:
            return "设备未连接"
        try:
            # 扫描模式映射表
            MODE = {
                'STEP_ONE_WAY': '0',      # 步进式单向
                'CONTINUOUS_ONE_WAY': '1', # 连续式单向
                'STEP_TWO_WAY': '2',      # 步进式往复
                'CONTINUOUS_TWO_WAY': '3'  # 连续式往复
            }
            
            # 根据GUI中的组合值转换为设备需要的模式值
            mode_value = MODE.get(mode, '0')
            
            self.device.write(f":WAVelength:SWEep:MODe {mode_value}")
            return f"扫描模式已设置为 {mode}"
        except Exception as e:
            return f"设置扫描模式失败: {str(e)}"

    def set_sweep_start(self, start):
        """设置扫描起始波长"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:STARt {start}")
            return f"扫描起始波长已设置为 {start}nm"
        except Exception as e:
            return f"设置起始波长失败: {str(e)}"

    def set_sweep_stop(self, stop):
        """设置扫描结束波长"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:STOP {stop}")
            return f"扫描结束波长已设置为 {stop}nm"
        except Exception as e:
            return f"设置结束波长失败: {str(e)}"

    def set_sweep_step(self, step):
        """设置扫描步长(nm)"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:STEP {step}")
            return f"扫描步长已设置为 {step}nm"
        except Exception as e:
            return f"设置步长失败: {str(e)}"

    def set_sweep_speed(self, speed):
        """设置扫描速度(nm/s)"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:SPEed {speed}")
            return f"扫描速度已设置为 {speed}nm/s"
        except Exception as e:
            return f"设置扫描速度失败: {str(e)}"
            
    def set_dwell_time(self, dwell):
        """设置驻留时间(秒)
        Range: 0 to 999.9 sec
        Step: 0.1 sec
        """
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:DWELl {dwell}")
            return f"驻留时间已设置为 {dwell}秒"
        except Exception as e:
            return f"设置驻留时间失败: {str(e)}"
    
    def set_sweep_cycles(self, cycles):
        """设置扫描循环次数
        Range: 0 to 999
        Step: 1
        """
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(f":WAVelength:SWEep:CYCLes {cycles}")
            return f"扫描循环次数已设置为 {cycles}次"
        except Exception as e:
            return f"设置循环次数失败: {str(e)}"

    def read_sweep_count(self):
        """读取当前扫描次数"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(":WAVelength:SWEep:COUNt?")
            count = self.device.read()
            return f"当前扫描次数: {count}"
        except Exception as e:
            return f"读取扫描次数失败: {str(e)}"

    def start_sweep(self):
        """开始扫描"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(":WAVelength:SWEep:STATe 1")
            return "扫描已开始"
        except Exception as e:
            return f"开始扫描失败: {str(e)}"

    def stop_sweep(self):
        """停止扫描"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(":WAVelength:SWEep:STATe 0")
            return "扫描已停止"
        except Exception as e:
            return f"停止扫描失败: {str(e)}"
            
    def sweep_repeat(self):
        """重复扫描"""
        if not self.connected:
            return "设备未连接"
        try:
            self.device.write(":WAVelength:SWEep:REPeat")
            return "扫描已重复启动"
        except Exception as e:
            return f"重复扫描失败: {str(e)}"

    def get_device_info(self):
        """从设备读取设备信息"""
        if not self.connected:
            return "设备未连接"
        
        try:
            # 读取设备型号
            self.device.write("*IDN?")
            idn = self.device.read()
            self.device_info["model"] = idn.split(",")[1] if "," in idn else self.model
            
            # 根据型号确定波长范围和功率限制
            # 实际应用中应从设备读取这些信息
            if "570" in self.device_info["model"]:
                # 读取波长范围
                self.device.write(":WAVelength:RANGe?")
                wavelength_range = self.device.read()
                self.device_info["wavelength_range"] = wavelength_range
                
                # 读取最大功率
                self.device.write(":POWer:RANGe?")
                max_power = self.device.read()
                self.device_info["max_power"] = max_power
            elif "550" in self.device_info["model"]:
                # 读取波长范围
                self.device.write(":WAVelength:RANGe?")
                wavelength_range = self.device.read()
                self.device_info["wavelength_range"] = wavelength_range
                
                # 读取最大功率
                self.device.write(":POWer:RANGe?")
                max_power = self.device.read()
                self.device_info["max_power"] = max_power
            
            return "设备信息已更新"
        except Exception as e:
            return f"读取设备信息失败: {str(e)}"

class ColorScheme:
    """颜色方案类"""
    PRIMARY = "#2980b9"    # 深蓝色
    SUCCESS = "#27ae60"    # 深绿色
    WARNING = "#f39c12"    # 深黄色
    DANGER = "#c0392b"     # 深红色
    INFO = "#7f8c8d"       # 深灰色
    DARK = "#2c3e50"       # 深色
    LIGHT = "#f8f9fa"      # 浅灰色
    TEXT = "#000000"       # 纯黑色文本
    BORDER = "#bdc3c7"     # 边框颜色
    
    @staticmethod
    def adjust_color(color, amount):
        """调整颜色深浅"""
        color = color.lstrip('#')
        r = int(color[:2], 16) + amount
        g = int(color[2:4], 16) + amount
        b = int(color[4:], 16) + amount
        
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        
        return f'#{r:02x}{g:02x}{b:02x}'

class StyleSheet:
    """样式表类"""
    @staticmethod
    def get_button_style(color, text_color="white"):
        return f"""
            QPushButton {{
                background-color: {color};
                color: {text_color};
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 12px;
            }}
            QPushButton:hover {{
                background-color: {ColorScheme.adjust_color(color, -20)};
            }}
            QPushButton:pressed {{
                background-color: {ColorScheme.adjust_color(color, -40)};
            }}
        """

    @staticmethod
    def get_group_box_style():
        return f"""
            QGroupBox {{
                border: 1px solid {ColorScheme.BORDER};
                border-radius: 4px;
                margin-top: 1em;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }}
        """

    @staticmethod
    def get_line_edit_style():
        return f"""
            QLineEdit {{
                padding: 5px;
                border: 1px solid {ColorScheme.BORDER};
                border-radius: 4px;
                background-color: white;
            }}
            QLineEdit:focus {{
                border: 1px solid {ColorScheme.PRIMARY};
            }}
        """

class TSL570GUI(QMainWindow):
    def __init__(self, model="TSL-570"):
        super().__init__()
        self.tsl = TSL570(model)
        self.setup_ui()
        
    def setup_ui(self):
        """设置主窗口UI"""
        self.setWindowTitle(f"{self.tsl.get_model()} 激光器控制")
        self.setGeometry(100, 100, 1000, 700)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建状态栏
        self.create_status_bar()
        
        # 创建选项卡部件
        tab_widget = QTabWidget()
        
        # 添加三个主要选项卡
        tab_widget.addTab(self.create_system_tab(), "系统控制")
        tab_widget.addTab(self.create_optical_tab(), "光学参数")
        tab_widget.addTab(self.create_sweep_tab(), "扫频设置")
        
        main_layout.addWidget(tab_widget)
        
        # 创建日志显示区域
        self.create_log_area()
        
        # 设置样式
        self.apply_styles()

    def create_status_bar(self):
        """创建状态栏"""
        status_bar = self.statusBar()
        
        # 创建状态指示器
        self.status_indicator = QLabel()
        self.status_indicator.setFixedSize(15, 15)
        self.status_indicator.setStyleSheet(f"background-color: {ColorScheme.DANGER}; border-radius: 7px;")
        
        # 创建状态文本
        self.status_text = QLabel("未连接")
        self.status_text.setStyleSheet(f"color: {ColorScheme.DANGER};")
        
        # 添加到状态栏
        status_bar.addWidget(self.status_indicator)
        status_bar.addWidget(self.status_text)

    def create_system_tab(self):
        """创建系统控制选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 设备连接控制组
        connection_group = QGroupBox("设备连接控制")
        connection_layout = QHBoxLayout()
        
        connect_btn = QPushButton("连接设备")
        connect_btn.clicked.connect(self.connect_device)
        connect_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.SUCCESS))
        
        disconnect_btn = QPushButton("断开连接")
        disconnect_btn.clicked.connect(self.disconnect_device)
        disconnect_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.DANGER))
        
        connection_layout.addWidget(connect_btn)
        connection_layout.addWidget(disconnect_btn)
        connection_group.setLayout(connection_layout)
        
        # 设备型号选择组
        model_group = QGroupBox("设备型号")
        model_layout = QHBoxLayout()
        
        self.model_570 = QRadioButton("TSL-570")
        self.model_550 = QRadioButton("TSL-550")
        self.model_570.setChecked(True)
        
        model_layout.addWidget(self.model_570)
        model_layout.addWidget(self.model_550)
        model_group.setLayout(model_layout)
        
        # 设备控制组
        control_group = QGroupBox("设备控制")
        control_layout = QHBoxLayout()
        
        restart_btn = QPushButton("重启设备")
        restart_btn.clicked.connect(self.restart_device)
        restart_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.WARNING))
        
        shutdown_btn = QPushButton("关闭设备")
        shutdown_btn.clicked.connect(self.shutdown_device)
        shutdown_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.DANGER))
        
        refresh_btn = QPushButton("刷新设备信息")
        refresh_btn.clicked.connect(self.refresh_device_info)
        refresh_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.INFO))
        
        control_layout.addWidget(restart_btn)
        control_layout.addWidget(shutdown_btn)
        control_layout.addWidget(refresh_btn)
        control_group.setLayout(control_layout)
        
        # 添加到布局
        layout.addWidget(connection_group)
        layout.addWidget(model_group)
        layout.addWidget(control_group)
        layout.addStretch()
        
        return tab

    def create_optical_tab(self):
        """创建光学参数选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 波长设置组
        wavelength_group = QGroupBox("波长设置")
        wavelength_layout = QGridLayout()
        
        wavelength_layout.addWidget(QLabel("波长值(nm):"), 0, 0)
        self.wavelength_input = QLineEdit()
        self.wavelength_input.setStyleSheet(StyleSheet.get_line_edit_style())
        wavelength_layout.addWidget(self.wavelength_input, 0, 1)
        
        set_wavelength_btn = QPushButton("设置")
        set_wavelength_btn.clicked.connect(self.set_wavelength)
        set_wavelength_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.PRIMARY))
        wavelength_layout.addWidget(set_wavelength_btn, 0, 2)
        
        wavelength_group.setLayout(wavelength_layout)
        
        # 功率控制组
        power_group = QGroupBox("功率控制")
        power_layout = QGridLayout()
        
        power_layout.addWidget(QLabel("功率值(dBm):"), 0, 0)
        self.power_input = QLineEdit()
        self.power_input.setStyleSheet(StyleSheet.get_line_edit_style())
        power_layout.addWidget(self.power_input, 0, 1)
        
        set_power_btn = QPushButton("设置")
        set_power_btn.clicked.connect(self.set_power_level)
        set_power_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.PRIMARY))
        power_layout.addWidget(set_power_btn, 0, 2)
        
        # 激光器开关控制
        power_layout.addWidget(QLabel("激光器输出:"), 1, 0)
        power_ctrl_layout = QHBoxLayout()
        
        power_on_btn = QPushButton("开启")
        power_on_btn.clicked.connect(lambda: self.set_power(True))
        power_on_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.SUCCESS))
        
        power_off_btn = QPushButton("关闭")
        power_off_btn.clicked.connect(lambda: self.set_power(False))
        power_off_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.DANGER))
        
        power_ctrl_layout.addWidget(power_on_btn)
        power_ctrl_layout.addWidget(power_off_btn)
        power_layout.addLayout(power_ctrl_layout, 1, 1, 1, 2)
        
        power_group.setLayout(power_layout)
        
        # 参数监控组
        monitor_group = QGroupBox("参数监控")
        monitor_layout = QGridLayout()
        
        # 创建状态标签
        self.status_labels = {}
        status_items = [
            ('current_wavelength', '当前波长:'),
            ('current_power', '当前功率:'),
            ('output_status', '输出状态:')
        ]
        
        for i, (key, label) in enumerate(status_items):
            monitor_layout.addWidget(QLabel(label), i, 0)
            self.status_labels[key] = QLabel("--")
            monitor_layout.addWidget(self.status_labels[key], i, 1)
        
        refresh_status_btn = QPushButton("刷新状态")
        refresh_status_btn.clicked.connect(self.refresh_optical_status)
        refresh_status_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.PRIMARY))
        monitor_layout.addWidget(refresh_status_btn, len(status_items), 0, 1, 2)
        
        monitor_group.setLayout(monitor_layout)
        
        # 添加到布局
        layout.addWidget(wavelength_group)
        layout.addWidget(power_group)
        layout.addWidget(monitor_group)
        layout.addStretch()
        
        return tab

    def create_sweep_tab(self):
        """创建扫频设置选项卡"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 扫描参数设置组
        params_group = QGroupBox("扫描参数设置")
        params_layout = QGridLayout()
        
        # 创建参数输入框
        self.sweep_inputs = {}
        sweep_params = [
            ('start', '起始波长(nm):'),
            ('stop', '结束波长(nm):'),
            ('step', '步长(nm):'),
            ('speed', '扫描速度(nm/s):'),
            ('dwell', '驻留时间(s):'),
            ('cycles', '循环次数:')
        ]
        
        for i, (key, label) in enumerate(sweep_params):
            row = i // 3
            col = (i % 3) * 2
            
            params_layout.addWidget(QLabel(label), row, col)
            self.sweep_inputs[key] = QLineEdit()
            self.sweep_inputs[key].setStyleSheet(StyleSheet.get_line_edit_style())
            params_layout.addWidget(self.sweep_inputs[key], row, col + 1)
        
        params_group.setLayout(params_layout)
        
        # 扫描模式选择组
        mode_group = QGroupBox("扫描模式选择")
        mode_layout = QVBoxLayout()
        
        # 扫描类型
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("扫描类型:"))
        self.sweep_type_step = QRadioButton("步进式扫描")
        self.sweep_type_continuous = QRadioButton("连续式扫描")
        self.sweep_type_step.setChecked(True)
        type_layout.addWidget(self.sweep_type_step)
        type_layout.addWidget(self.sweep_type_continuous)
        mode_layout.addLayout(type_layout)
        
        # 扫描方向
        direction_layout = QHBoxLayout()
        direction_layout.addWidget(QLabel("扫描方向:"))
        self.sweep_direction_one = QRadioButton("单向扫描")
        self.sweep_direction_two = QRadioButton("往复扫描")
        self.sweep_direction_one.setChecked(True)
        direction_layout.addWidget(self.sweep_direction_one)
        direction_layout.addWidget(self.sweep_direction_two)
        mode_layout.addLayout(direction_layout)
        
        mode_group.setLayout(mode_layout)
        
        # 扫描状态组
        status_group = QGroupBox("扫描状态")
        status_layout = QHBoxLayout()
        
        self.sweep_status_label = QLabel("未开始")
        self.sweep_count_label = QLabel("扫描次数: 0")
        status_layout.addWidget(self.sweep_status_label)
        status_layout.addWidget(self.sweep_count_label)
        
        status_group.setLayout(status_layout)
        
        # 控制按钮组
        control_group = QGroupBox("扫描控制")
        control_layout = QHBoxLayout()
        
        setup_btn = QPushButton("设置参数")
        setup_btn.clicked.connect(self.setup_sweep)
        setup_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.PRIMARY))
        
        start_btn = QPushButton("开始扫描")
        start_btn.clicked.connect(self.start_sweep)
        start_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.SUCCESS))
        
        stop_btn = QPushButton("停止扫描")
        stop_btn.clicked.connect(self.stop_sweep)
        stop_btn.setStyleSheet(StyleSheet.get_button_style(ColorScheme.DANGER))
        
        control_layout.addWidget(setup_btn)
        control_layout.addWidget(start_btn)
        control_layout.addWidget(stop_btn)
        
        control_group.setLayout(control_layout)
        
        # 添加到布局
        layout.addWidget(params_group)
        layout.addWidget(mode_group)
        layout.addWidget(status_group)
        layout.addWidget(control_group)
        layout.addStretch()
        
        return tab

    def create_log_area(self):
        """创建日志显示区域"""
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        self.centralWidget().layout().addWidget(log_group)

    def apply_styles(self):
        """应用全局样式"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
            }
            QGroupBox {
                font-weight: bold;
            }
            QLabel {
                font-size: 12px;
            }
            QRadioButton {
                font-size: 12px;
            }
        """)

    # 设备控制方法
    def connect_device(self):
        devices = self.tsl.search_gpib_addresses()
        if not devices:
            self.update_status("未连接", False)
            self.show_log("未找到 GPIB 设备")
            return
            
        result = self.tsl.connect_device(devices[0])
        if self.tsl.is_connected():
            self.update_status("已连接", True)
            self.update_device_info()
        else:
            self.update_status("未连接", False)
        self.show_log(result)

    def disconnect_device(self):
        result = self.tsl.disconnect()
        self.update_status("未连接", False)
        self.show_log(result)

    def shutdown_device(self):
        result = self.tsl.device_shut_down()
        self.show_log(result)

    def restart_device(self):
        result = self.tsl.device_restart()
        self.show_log(result)

    def set_wavelength(self):
        wavelength = self.wavelength_input.text()
        if not wavelength:
            self.show_log("请输入波长值")
            return
        result = self.tsl.set_wavelength(wavelength)
        self.show_log(result)

    def set_power(self, state):
        result = self.tsl.set_power_status('1' if state else '0')
        self.show_log(result)

    def set_power_level(self):
        power = self.power_input.text()
        if not power:
            self.show_log("请输入功率值")
            return
        result = self.tsl.set_power_level(power)
        self.show_log(result)

    def setup_sweep(self):
        if not self.tsl.is_connected():
            self.show_log("设备未连接，无法设置扫描参数")
            return
            
        # 获取扫描模式
        sweep_type = "STEP" if self.sweep_type_step.isChecked() else "CONTINUOUS"
        sweep_direction = "ONE_WAY" if self.sweep_direction_one.isChecked() else "TWO_WAY"
        mode = f"{sweep_type}_{sweep_direction}"
        
        # 设置扫描模式
        result = self.tsl.set_sweep_mode(mode)
        self.show_log(result)
        
        # 设置其他参数
        for key, input_widget in self.sweep_inputs.items():
            value = input_widget.text()
            if value:
                if key == 'start':
                    self.show_log(self.tsl.set_sweep_start(value))
                elif key == 'stop':
                    self.show_log(self.tsl.set_sweep_stop(value))
                elif key == 'step':
                    self.show_log(self.tsl.set_sweep_step(value))
                elif key == 'speed':
                    self.show_log(self.tsl.set_sweep_speed(value))
                elif key == 'dwell':
                    self.show_log(self.tsl.set_dwell_time(value))
                elif key == 'cycles':
                    self.show_log(self.tsl.set_sweep_cycles(value))
        
        self.sweep_status_label.setText("已设置参数，等待开始")
        self.show_log("扫描参数已全部设置")

    def start_sweep(self):
        if not self.tsl.is_connected():
            self.show_log("设备未连接，无法开始扫描")
            return
            
        # 先尝试重复扫描
        result = self.tsl.sweep_repeat()
        if "失败" in result:
            # 如果重复扫描失败，则尝试开始新的扫描
            result = self.tsl.start_sweep()
            
        if "失败" not in result:
            self.sweep_status_label.setText("扫描中...")
            self.sweep_count_label.setText("扫描次数: 0")
            self._start_count_update()
            
        self.show_log(result)

    def stop_sweep(self):
        if not self.tsl.is_connected():
            self.show_log("设备未连接，无法停止扫描")
            return
            
        result = self.tsl.stop_sweep()
        if "失败" not in result:
            self.sweep_status_label.setText("已停止")
            
        self.show_log(result)

    def _start_count_update(self):
        """启动自动更新扫描次数"""
        if hasattr(self, 'count_timer'):
            self.count_timer.stop()
            
        self.count_timer = QTimer()
        self.count_timer.timeout.connect(self.update_sweep_count)
        self.count_timer.start(1000)  # 每秒更新一次

    def update_sweep_count(self):
        """更新扫描次数"""
        if self.tsl.is_connected() and "扫描中" in self.sweep_status_label.text():
            try:
                result = self.tsl.read_sweep_count()
                if "失败" not in result and ":" in result:
                    count = result.split(":")[-1].strip()
                    self.sweep_count_label.setText(f"扫描次数: {count}")
            except Exception:
                pass

    def update_status(self, text, connected):
        """更新状态栏显示"""
        self.status_text.setText(text)
        color = ColorScheme.SUCCESS if connected else ColorScheme.DANGER
        self.status_indicator.setStyleSheet(f"background-color: {color}; border-radius: 7px;")
        self.status_text.setStyleSheet(f"color: {color};")

    def show_log(self, message):
        """显示日志信息"""
        self.log_text.append(str(message))
        # 设置不同消息类型的颜色
        if "错误" in str(message) or "失败" in str(message):
            self.log_text.setTextColor(QColor(ColorScheme.DANGER))
        elif "成功" in str(message) or "已连接" in str(message):
            self.log_text.setTextColor(QColor(ColorScheme.SUCCESS))
        else:
            self.log_text.setTextColor(QColor(ColorScheme.TEXT))
        
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def update_device_info(self):
        """更新设备信息显示"""
        if self.tsl.is_connected():
            self.tsl.get_device_info()
            device_info = self.tsl.device_info
            self.show_log(f"已读取设备信息:\n型号: {device_info['model']}\n"
                         f"波长范围: {device_info['wavelength_range']}\n"
                         f"最大功率: {device_info['max_power']}")

    def refresh_device_info(self):
        """刷新设备信息"""
        if self.tsl.is_connected():
            result = self.tsl.get_device_info()
            self.update_device_info()
            self.show_log(result)
        else:
            self.show_log("设备未连接，无法获取信息")

    def refresh_optical_status(self):
        """刷新光学参数状态"""
        if not self.tsl.is_connected():
            self.show_log("设备未连接，无法获取状态")
            return
            
        try:
            # 获取当前波长
            self.tsl.device.write(":WAVelength?")
            wavelength = self.tsl.device.read()
            self.status_labels['current_wavelength'].setText(f"{wavelength} nm")
            
            # 获取当前功率
            self.tsl.device.write(":POWer?")
            power = self.tsl.device.read()
            self.status_labels['current_power'].setText(f"{power} dBm")
            
            # 获取输出状态
            self.tsl.device.write(":POWer:STATe?")
            output_state = self.tsl.device.read()
            if output_state == "1" or output_state.strip() == "1":
                self.status_labels['output_status'].setText("开启")
            else:
                self.status_labels['output_status'].setText("关闭")
                
            self.show_log("参数状态已刷新")
        except Exception as e:
            self.show_log(f"获取参数状态失败: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = TSL570GUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
