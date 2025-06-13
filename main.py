import os, re
import pandas as pd
from datetime import datetime, timedelta
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.anchorlayout import AnchorLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.popup import Popup
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.widget import Widget
from kivy.metrics import dp
from kivy.core.text import LabelBase
from kivy.core.window import Window
from kivy.properties import ObjectProperty, StringProperty, ListProperty, NumericProperty
from kivy.clock import Clock
from kivy.config import Config
import os
from kivy.resources import resource_add_path, resource_find
resource_add_path(os.path.abspath('.'))
from kivy.core.text import LabelBase
LabelBase.register('Roboto', 'msyh.ttc')
# 设置窗口大小便于开发
Config.set('graphics', 'width', '400')
Config.set('graphics', 'height', '700')

class AttendanceApp(App):
    def __init__(self, **kwargs):
        super(AttendanceApp, self).__init__(**kwargs)
        # 初始化所有必要属性
        self.leave_records = {
            "按周请假": pd.DataFrame(columns=["人名", "类型", "每周次数", "周数", "剩余次数"]),  # 添加剩余次数列
            "固定时段": pd.DataFrame(columns=["人名", "类型", "日期", "时段"]),
            "长期请假": pd.DataFrame(columns=["人名", "类型"])
        }
        self.week_start_day = "星期五"  # 默认值
        self.current_attendance_file = ""
        self.file_list = []
        self.students = []
        self.time_slots = []
        self.leave_types = []
        self.current_year = datetime.now().year
        self.current_week = datetime.now().isocalendar()[1]

    def build(self):
        self.load_class_data()
        return MainScreen()

    def load_class_data(self):
        """加载班级基础数据 - 完整修复版"""
        try:
            self.class_info = pd.ExcelFile("班级信息.xlsx")
            
            # 读取学生名单
            self.students = pd.read_excel(self.class_info, sheet_name="名单").iloc[:, 0].tolist()
            
            # 读取时段
            self.time_slots = pd.read_excel(self.class_info, sheet_name="时段").iloc[:, 0].tolist()
            
            # 读取请假类型
            self.leave_types = pd.read_excel(self.class_info, sheet_name="类型").iloc[:, 0].tolist()
            if "出勤" not in self.leave_types:
                self.leave_types.insert(0, "出勤")
            if "缺勤" not in self.leave_types:
                self.leave_types.append("缺勤")
            
            # 读取文件名列表
            self.file_list = pd.read_excel(self.class_info, sheet_name="文件").iloc[:, 0].tolist()
            
            # 读取杂项设置
            misc_df = pd.read_excel(self.class_info, sheet_name="杂项")
            self.current_attendance_file = misc_df[misc_df.iloc[:, 0] == "当前文件"].iloc[0, 1]
            self.week_start_day = misc_df[misc_df.iloc[:, 0] == "周起始日"].iloc[0, 1]
            
            # 加载请假记录
            self.load_leave_records()
            
        except Exception as e:
            print(f"加载班级数据出错: {e}")
            # 提供默认值防止程序崩溃
            self.students = ["学生1", "学生2", "学生3"]
            self.time_slots = ["早殿", "上午", "下午", "晚殿", "夜间"]
            self.leave_types = ["出勤", "病假", "事假", "公假", "缺勤"]
            self.file_list = []
            self.current_attendance_file = f"{datetime.now().year}年第{datetime.now().isocalendar()[1]}周班级考勤记录表.xlsx"
            self.week_start_day = "星期五"
            self.leave_records = {
                "按周请假": pd.DataFrame(columns=["人名", "类型", "每周次数", "周数"]),
                "固定时段": pd.DataFrame(columns=["人名", "类型", "日期", "时段"]),
                "长期请假": pd.DataFrame(columns=["人名", "类型"])
            }
    
    def save_misc_settings(self):
        """保存杂项设置"""
        try:
            misc_data = [
                ["当前文件", self.current_attendance_file],
                ["周起始日", self.week_start_day]
            ]
            misc_df = pd.DataFrame(misc_data, columns=["名称", "值"])
            
            with pd.ExcelWriter("班级信息.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                misc_df.to_excel(writer, sheet_name="杂项", index=False)
        except Exception as e:
            print(f"保存设置出错: {e}")
    
    def save_file_list(self):
        """保存文件列表"""
        try:
            file_df = pd.DataFrame(self.file_list, columns=["文件名"])
            
            with pd.ExcelWriter("班级信息.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                file_df.to_excel(writer, sheet_name="文件", index=False)
        except Exception as e:
            print(f"保存文件列表出错: {e}")

    def load_leave_records(self):
        """加载请假记录 - 修复版"""
        try:
            if os.path.exists("请假信息记录表.xlsx"):
                leave_file = pd.ExcelFile("请假信息记录表.xlsx")
                
                # 确保基础结构存在
                if not hasattr(self, 'leave_records'):
                    self.leave_records = {
                        "按周请假": pd.DataFrame(columns=["人名", "类型", "每周次数", "周数", "剩余次数"]),
                        "固定时段": pd.DataFrame(columns=["人名", "类型", "日期", "时段"]),
                        "长期请假": pd.DataFrame(columns=["人名", "类型"])
                    }
                
                # 加载各sheet
                for sheet in ["按周请假", "固定时段", "长期请假"]:
                    if sheet in leave_file.sheet_names:
                        df = pd.read_excel(leave_file, sheet_name=sheet)
                        # 如果是按周请假且没有剩余次数列，则添加该列并初始化
                        if sheet == "按周请假" and "剩余次数" not in df.columns.values:
                            df["剩余次数"] = df["每周次数"]
                        self.leave_records[sheet] = df
        except Exception as e:
            print(f"加载请假记录出错: {e}")
            # 确保有默认值
            self.leave_records = {
                "按周请假": pd.DataFrame(columns=["人名", "类型", "每周次数", "周数", "剩余次数"]),
                "固定时段": pd.DataFrame(columns=["人名", "类型", "日期", "时段"]),
                "长期请假": pd.DataFrame(columns=["人名", "类型"])
            }

    def save_leave_records(self):
        """保存请假记录"""
        with pd.ExcelWriter("请假信息记录表.xlsx") as writer:
            for sheet, df in self.leave_records.items():
                df.to_excel(writer, sheet_name=sheet, index=False)

    def load_attendance_data(self, day_time):
        """加载指定时段的考勤数据"""
        if not os.path.exists(self.current_attendance_file):
            return None
            
        try:
            attendance_df = pd.read_excel(self.current_attendance_file)
            if day_time in attendance_df.columns:
                return dict(zip(attendance_df["姓名"], attendance_df[day_time]))
            return None
        except:
            return None

    def save_attendance_data(self, day_time, attendance_data, filename=None):
        """保存考勤数据"""
        if filename is None:
            filename = self.current_attendance_file
        
        # 创建基础DataFrame
        df = pd.DataFrame({"姓名": self.students})
        
        # 如果文件已存在，读取现有数据
        if os.path.exists(filename):
            existing_df = pd.read_excel(filename)
            # 合并现有数据
            df = existing_df.copy()
            # 如果该时段已存在，则更新
            if day_time in df.columns:
                df[day_time] = df["姓名"].map(attendance_data).fillna(df[day_time])
            else:
                df[day_time] = df["姓名"].map(attendance_data)
        else:
            # 新文件，直接添加列
            df[day_time] = df["姓名"].map(attendance_data)
        
        # 保存文件
        df.to_excel(filename, index=False)


class MainScreen(BoxLayout):
    """主界面"""
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = [20, 20]
        self.spacing = 15
        self.selected_date = None
        self.selected_time_slot = None 
        # 标题
        self.add_widget(Label(text="班级考勤管理系统", font_size=24, size_hint_y=0.15))
        
        # 按钮区域
        btn_layout = BoxLayout(orientation='vertical', spacing=10, size_hint_y=0.85)
        
        buttons = [
            ("开始点名", self.start_attendance),
            ("请假记录", self.leave_record),
            ("销假记录", self.cancel_leave_record),
            ("考勤管理", self.attendance_management)
        ]
        
        for text, callback in buttons:
            btn = Button(text=text, size_hint_y=None, height=dp(60))
            btn.bind(on_release=callback)
            btn_layout.add_widget(btn)
        
        self.add_widget(btn_layout)

    def attendance_management(self, instance):
        """进入考勤管理界面"""
        self.clear_widgets()
        self.add_widget(AttendanceManagementScreen())
    
    def generate_report(self, instance):
        """生成报表"""
        # 直接调用考勤管理页面的生成报表方法
        content = AttendanceManagementScreen()
        content.generate_report()


    def start_attendance(self, instance):
        """开始点名 - 新版：先选择日期再选择时段"""
        app = App.get_running_app()
        
        # 创建日期选择布局
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text="请选择日期:", size_hint_y=0.1, font_size='14sp'))
        
        # 从班级信息读取起始日期
        try:
            start_day = app.week_start_day
        except:
            start_day = "星期五"  # 默认值
        
        # 生成日期选项（从起始日开始）
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        start_index = weekdays.index(start_day) if start_day in weekdays else 4
        display_days = weekdays[start_index:] + weekdays[:start_index]
        
        # 日期按钮网格 (2排: 每排4个)
        date_grid = GridLayout(cols=4, rows=2, spacing=10, size_hint_y=0.5)
        self.date_buttons = []
        
        for i, day in enumerate(display_days[:7]):  # 最多显示7天
            btn = Button(
                text=day,
                size_hint=(None, None),
                size=(dp(80), dp(40)),
                font_size='14sp'  # 统一字号为14sp
            )
            btn.bind(on_release=lambda btn, d=day: self.select_date(d))
            date_grid.add_widget(btn)
            self.date_buttons.append(btn)
        
        # 如果不足8个，添加空Widget填充
        for _ in range(8 - len(display_days[:7])):
            date_grid.add_widget(Widget(size_hint=(None, None), size=(dp(80), dp(40))))
        
        content.add_widget(date_grid)
        
        # 添加时段选择部分
        time_slot_screen = BoxLayout(orientation='vertical', spacing=10)
        time_slot_screen.add_widget(Label(text="请选择时段:", size_hint_y=0.1, font_size='14sp'))
        
        # 时段按钮网格 (2排: 每排4个)
        time_grid = GridLayout(cols=4, rows=2, spacing=10, size_hint_y=0.4)
        time_slots = app.time_slots[:8]  # 最多显示8个时段
        
        for slot in time_slots:
            btn = Button(
                text=slot,
                size_hint=(None, None),
                size=(dp(80), dp(40)),
                font_size='14sp'  # 统一字号为14sp
            )
            btn.bind(on_release=lambda btn, s=slot: self.select_time_slot(s))
            time_grid.add_widget(btn)
        
        # 如果不足8个，添加空Widget填充
        for _ in range(8 - len(time_slots)):
            time_grid.add_widget(Widget(size_hint=(None, None), size=(dp(80), dp(40))))
        
        time_slot_screen.add_widget(time_grid)
        
        # 添加返回按钮
        btn_back = Button(text='返回', size_hint_y=0.1, font_size='14sp')
        btn_back.bind(on_release=lambda x: self.clear_widgets() or self.add_widget(MainScreen()))
        
        # 将各部分添加到主布局
        content.add_widget(time_slot_screen)
        content.add_widget(btn_back)
        
        # 初始化选择状态
        self.selected_date = None
        self.selected_time_slot = None
        
        # 显示新界面
        self.clear_widgets()
        self.add_widget(content)

    def select_date(self, date):
        """选择日期"""
        self.selected_date = date
        for btn in self.date_buttons:
            btn.background_color = (0.5, 0.5, 1, 1) if btn.text == date else (1, 1, 1, 1)
    
    def select_time_slot(self, time_slot):
        """选择时段"""
        if not self.selected_date:
            self.show_message("请先选择日期")
            return
        
        self.selected_time_slot = time_slot
        self.go_to_attendance(time_slot)  # 修复：传递time_slot参数
    
    def go_to_attendance(self, time_slot):
        """进入点名界面"""
        if not self.selected_date:
            self.show_message("请先选择日期")
            return
            
        app = App.get_running_app()
        full_time_slot = f"{self.selected_date}{time_slot}"
        self.clear_widgets()
        self.add_widget(AttendanceScreen(time_slot=time_slot, day_time=full_time_slot))
    def leave_record(self, instance):
        """请假记录"""
        app = App.get_running_app()
        app.root.clear_widgets()
        app.root.add_widget(LeaveStudentSelectScreen())

    def cancel_leave_record(self, instance):
        """销假记录"""
        app = App.get_running_app()
        app.root.clear_widgets()
        app.root.add_widget(LeaveStudentSelectScreen(is_cancel=True))

    def misc_info(self, instance):
        """考勤管理"""
        self.show_message("功能开发中...")

    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

class AttendanceScreen(BoxLayout):
    """点名界面"""
    def __init__(self, time_slot, day_time, **kwargs):
        super(AttendanceScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.time_slot = time_slot
        self.day_time = day_time  # 完整的日期时段标识，如"星期五上午"
        self.app = App.get_running_app()
        self.attendance_data = {}
        
        # 加载已有考勤数据
        self.existing_data = self.load_existing_attendance()
        
        # 构建界面
        self.build_attendance_ui()
    def add_bottom_buttons(self):
        """添加底部按钮"""
        btn_layout = BoxLayout(
            size_hint_y=0.1,
            spacing=10,
            padding=[10, 5]
        )
        
        # 自动标记按钮
        btn_auto = Button(text='自动标记')
        btn_auto.bind(on_release=self.auto_mark)
        
        # 保存按钮
        btn_save = Button(text='保存')
        btn_save.bind(on_release=self.save_attendance)
        
        # 返回按钮
        btn_back = Button(text='返回')
        btn_back.bind(on_release=self.go_back)
        
        btn_layout.add_widget(btn_auto)
        btn_layout.add_widget(btn_save)
        btn_layout.add_widget(btn_back)
        
        self.add_widget(btn_layout)

    def get_current_attendance(self, student, time_slot):
        """获取本周该学生此时段的考勤记录"""
        # 获取今天是周几
        weekday = datetime.now().weekday()  # 0是周一，6是周日
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        today = weekdays[weekday]
        
        day_time = f"{today}{self.time_slot}"
        
        # 加载已有考勤数据
        self.existing_data = self.load_existing_attendance()
        
        # 创建界面元素
        self.build_attendance_ui()
    def load_existing_attendance(self):
        """加载已有考勤数据"""
        try:
            if os.path.exists(self.app.current_attendance_file):
                df = pd.read_excel(self.app.current_attendance_file)
                if self.day_time in df.columns:
                    # 返回 {姓名: 考勤类型} 的字典
                    return dict(zip(df["姓名"], df[self.day_time]))
        except Exception as e:
            print(f"加载考勤数据出错: {e}")
        return {}
    def build_attendance_ui(self):
        """构建考勤界面"""
        self.clear_widgets()
        
        # 标题显示完整日期时段
        self.add_widget(Label(text=f"点名 - {self.day_time}", font_size=20, size_hint_y=0.1))
        
        # 创建滚动区域
        scroll = ScrollView(size_hint=(1, 0.8))
        grid = GridLayout(cols=1, spacing=dp(15), size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        
        # 为每个学生创建考勤选项
        for student in self.app.students:
            # 获取该学生当前时段已有的考勤记录
            current_status = self.existing_data.get(student, "出勤")
            
            # 检查按周请假
            weekly_leave = self.check_weekly_leave(student)
            
            # 创建学生行
            student_row = self.create_student_row(student, current_status, weekly_leave)
            grid.add_widget(student_row)
        
        scroll.add_widget(grid)
        self.add_widget(scroll)
        
        # 添加底部按钮
        self.add_bottom_buttons()
    def create_student_row(self, student, current_status, weekly_leave):
        """创建单个学生的考勤行"""
        student_row = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=dp(50) + dp(40)*((len(self.app.leave_types)+4)//5),
            spacing=dp(5),
            padding=[dp(5), dp(5)]
        )
        
        # 学生名字标签
        name_label = Label(
            text=student,
            size_hint=(1, None),
            height=dp(40),
            halign='left',
            valign='middle'
        )
        student_row.add_widget(name_label)
        
        # 请假类型按钮组
        btn_rows = []
        items_per_row = 5
        total_rows = (len(self.app.leave_types) + items_per_row - 1) // items_per_row
        
        for row in range(total_rows):
            btn_row = GridLayout(cols=items_per_row, spacing=dp(5), size_hint_y=None, height=dp(40))
            btn_rows.append(btn_row)
        
        # 分配按钮到各行
        for i, ltype in enumerate(self.app.leave_types):
            row_idx = i // items_per_row
            # 添加请假次数到按钮文本
            btn_text = ltype + (f"({weekly_leave.get(ltype, '')})" if ltype in weekly_leave else "")
            btn = ToggleButton(
                text=btn_text,
                group=student,
                state='down' if ltype == current_status else 'normal',
                size_hint=(None, None),
                size=(dp(70), dp(40))
            )
            btn.bind(on_release=lambda btn, s=student, t=ltype: self.update_attendance(s, t))
            btn_rows[row_idx].add_widget(btn)
        
        # 添加所有按钮行
        for btn_row in btn_rows:
            student_row.add_widget(btn_row)
        
        return student_row
    
    def check_long_leave(self, student):
        """检查长期请假"""
        long_leaves = self.app.leave_records["长期请假"]
        if not long_leaves.empty:
            records = long_leaves[long_leaves["人名"] == student]
            if not records.empty:
                return records.iloc[0]["类型"]
        return None

    def check_fixed_leave(self, student, time_slot):
        """检查固定时段请假"""
        # 获取今天是周几
        weekday = datetime.now().weekday()  # 0是周一，6是周日
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        today = weekdays[weekday]
        
        fixed_leaves = self.app.leave_records["固定时段"]
        if not fixed_leaves.empty:
            records = fixed_leaves[(fixed_leaves["人名"] == student) & 
                                  (fixed_leaves["日期"] == today) & 
                                  (fixed_leaves["时段"] == time_slot)]
            if not records.empty:
                return records.iloc[0]["类型"]
        return None

    def check_weekly_leave(self, student):
        """检查按周请假"""
        weekly_leaves = self.app.leave_records["按周请假"]
        if not weekly_leaves.empty:
            records = weekly_leaves[weekly_leaves["人名"] == student]
            if not records.empty:
                return dict(zip(records["类型"], records["剩余次数"]))  # 修改为返回剩余次数
        return {}

    def update_attendance(self, student, leave_type):
        """更新考勤数据"""
        self.attendance_data[student] = leave_type
        
        # 如果是按周请假类型，减少剩余次数
        weekly_leaves = self.app.leave_records["按周请假"]
        if not weekly_leaves.empty:
            records = weekly_leaves[(weekly_leaves["人名"] == student) & (weekly_leaves["类型"] == leave_type)]
            if not records.empty:
                # 减少剩余次数而不是每周次数
                self.app.leave_records["按周请假"].loc[records.index, "剩余次数"] -= 1
                # 如果剩余次数用完，删除记录
                self.app.leave_records["按周请假"] = self.app.leave_records["按周请假"][
                    ~((self.app.leave_records["按周请假"]["剩余次数"] <= 0) & 
                      (self.app.leave_records["按周请假"]["人名"] == student) & 
                      (self.app.leave_records["按周请假"]["类型"] == leave_type))
                ]
                self.app.save_leave_records()

    def auto_mark(self, instance):
        """自动标记未选择的学生（仅标记既无用户修改又无原有记录的学生）"""
        for student in self.app.students:
            # 只有当该学生没有手动标记且表格中也没有记录时才标记为缺勤
            if student not in self.attendance_data and student not in self.existing_data:
                self.attendance_data[student] = "缺勤"
        
        # 调用保存方法
        self.save_attendance(instance)

    def save_attendance(self, instance):
        """保存考勤数据，用户手动修改的优先于原有数据"""
        final_data = {}
        
        # 1. 先添加表格中原有的考勤数据（作为默认值）
        for student in self.app.students:
            if student in self.existing_data:
                final_data[student] = self.existing_data[student]
            else:
                final_data[student] = "缺勤"  # 默认值
        
        # 2. 用用户手动修改的数据覆盖原有数据（用户修改的优先级更高）
        final_data.update(self.attendance_data)
        
        # 3. 保存最终合并后的数据
        self.app.save_attendance_data(self.day_time, final_data, self.app.current_attendance_file)
        
        # 显示成功消息
        self.show_message("考勤记录保存成功")
        
        # 返回主界面
        self.go_back(None)
    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())
    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())

    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

class AttendanceManagementScreen(BoxLayout):
    def __init__(self, **kwargs):
        super(AttendanceManagementScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = [10, 10]
        self.spacing = 10
        self.app = App.get_running_app()
        self.selected_file = None  # 记录当前选中的文件
        
        # 起始日期设置 - 缩小尺寸
        start_day_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        start_day_layout.add_widget(Label(
            text="起始日期:", 
            size_hint=(0.3, 1),
            halign='left'
        ))
        
        self.week_start_spinner = Spinner(
            text=self.app.week_start_day,
            values=["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"],
            size_hint=(0.7, 1),
            font_size='12sp'
        )
        self.week_start_spinner.bind(text=self.on_week_start_changed)
        start_day_layout.add_widget(self.week_start_spinner)
        self.add_widget(start_day_layout)
        
        # 操作按钮行 - 两个按钮并排
        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=10)
        
        # 减扣一周按钮 - 缩小尺寸
        btn_deduct = Button(
            text="减扣一周", 
            size_hint=(0.5, 1),
            font_size='12sp'
        )
        btn_deduct.bind(on_release=self.deduct_week)
        btn_row.add_widget(btn_deduct)
        
        # 新的一周按钮 - 缩小尺寸
        btn_new_week = Button(
            text="新的一周", 
            size_hint=(0.5, 1),
            font_size='12sp'
        )
        btn_new_week.bind(on_release=self.create_new_week)
        btn_row.add_widget(btn_new_week)
        
        self.add_widget(btn_row)
        
        # 考勤记录列表
        self.add_widget(Label(
            text="考勤记录列表:", 
            size_hint_y=None, 
            height=30,
            font_size='12sp'
        ))
        
        scroll = ScrollView()
        self.file_grid = GridLayout(
            cols=1, 
            spacing=5, 
            size_hint_y=None,
            padding=[5, 5]
        )
        self.file_grid.bind(minimum_height=self.file_grid.setter('height'))
        
        self.refresh_file_list()
        scroll.add_widget(self.file_grid)
        self.add_widget(scroll)
        
        # 底部按钮 - 改为生成报表
        bottom_btn_layout = BoxLayout(
            size_hint_y=None,
            height=40,
            spacing=10,
            padding=[10, 0]
        )
        # 生成报表按钮
        btn_report = Button(
            text="生成报表",
            size_hint=(0.7, 1),  # 占据70%宽度
            font_size='12sp'
        )
        btn_report.bind(on_release=self.generate_report)
        # 新增返回按钮 - 尺寸自适应文字
        btn_back = Button(
            text="返回",
            size_hint=(0.3, 1),  # 占据30%宽度
            font_size='12sp'
        )
        btn_back.bind(on_release=self.go_back)
        # 将按钮添加到布局
        bottom_btn_layout.add_widget(btn_report)
        bottom_btn_layout.add_widget(btn_back)
        # 将底部按钮布局添加到主布局
        self.add_widget(bottom_btn_layout)
    
    def refresh_file_list(self):
        """刷新文件列表显示 - 带删除功能版"""
        self.file_grid.clear_widgets()
        
        for filename in self.app.file_list:
            item = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=40,
                spacing=5
            )
            
            # 文件选择按钮
            btn_file = Button(
                text=filename,
                size_hint=(0.7, 1),
                font_size='12sp',
                background_color=(0.9, 0.9, 0.9, 1) if filename == self.app.current_attendance_file else (1, 1, 1, 1)
            )
            btn_file.bind(on_release=lambda btn: self.select_file(btn.text))
            item.add_widget(btn_file)
            
            # 删除按钮
            btn_del = Button(
                text="×",
                size_hint=(0.3, 1),
                font_size='14sp',
                background_color=(0.9, 0.4, 0.4, 1)
            )
            btn_del.bind(on_release=lambda btn, f=filename: self.confirm_delete_file(f))
            item.add_widget(btn_del)
            
            self.file_grid.add_widget(item)

    def confirm_delete_file(self, filename):
        """确认删除文件"""
        if filename == self.app.current_attendance_file:
            self.show_message("不能删除当前使用的文件")
            return
            
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=f"确定要永久删除 {filename} 吗？", color=(1, 0, 0, 1)))
        
        btn_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        btn_yes = Button(text="确定删除", background_color=(0.9, 0.4, 0.4, 1))
        btn_no = Button(text="取消")
        
        popup = Popup(
            title="确认删除", 
            content=content, 
            size_hint=(0.7, 0.3),
            separator_color=(0.9, 0.4, 0.4, 1)
        )
        
        def do_delete(instance):
            try:
                # 从文件系统中删除
                if os.path.exists(filename):
                    os.remove(filename)
                
                # 从列表中删除
                if filename in self.app.file_list:
                    self.app.file_list.remove(filename)
                    self.app.save_file_list()
                
                # 如果删除的是当前文件，重置当前文件
                if filename == self.app.current_attendance_file:
                    self.app.current_attendance_file = ""
                    self.app.save_misc_settings()
                
                self.refresh_file_list()
                popup.dismiss()
                self.show_message(f"已删除 {filename}")
            except Exception as e:
                popup.dismiss()
                self.show_message(f"删除失败: {e}")
        
        btn_yes.bind(on_release=do_delete)
        btn_no.bind(on_release=popup.dismiss)
        
        btn_layout.add_widget(btn_yes)
        btn_layout.add_widget(btn_no)
        content.add_widget(btn_layout)
        
        popup.open()

    def select_file(self, filename):
        """选择考勤记录文件"""
        self.selected_file = filename
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=f"确定要使用 {filename} 作为当前考勤表吗？"))
        
        btn_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        btn_yes = Button(text="确定", size_hint=(0.5, 1))
        btn_no = Button(text="取消", size_hint=(0.5, 1))
        
        popup = Popup(
            title="确认", 
            content=content, 
            size_hint=(0.7, 0.3)
        )
        def confirm(instance):
            self.app.current_attendance_file = filename
            self.app.save_misc_settings()
            self.refresh_file_list()
            popup.dismiss()
            self.show_message(f"已切换到 {filename}")
        
        btn_yes.bind(on_release=confirm)
        btn_no.bind(on_release=popup.dismiss)
        
        btn_layout.add_widget(btn_yes)
        btn_layout.add_widget(btn_no)
        content.add_widget(btn_layout)
        
        popup.open()
    
    def on_week_start_changed(self, spinner, text):
        """起始日期变更处理"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=f"确定将起始日期改为 {text} 吗？"))
        
        btn_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        btn_yes = Button(text="确定", size_hint=(0.5, 1))
        btn_no = Button(text="取消", size_hint=(0.5, 1))
        
        popup = Popup(
            title="确认", 
            content=content, 
            size_hint=(0.7, 0.3)
        )
        def confirm(instance):
            self.app.week_start_day = text
            self.app.save_misc_settings()
            popup.dismiss()
            self.show_message(f"起始日期已设为 {text}")
        
        btn_yes.bind(on_release=confirm)
        btn_no.bind(on_release=popup.dismiss)
        
        btn_layout.add_widget(btn_yes)
        btn_layout.add_widget(btn_no)
        content.add_widget(btn_layout)
        
        popup.open()
    
    def create_new_week(self, instance):
        """创建新的一周考勤表 - 按起始日排序并应用请假信息"""
        # 计算下周的年份和周数
        today = datetime.now()
        next_week = today + timedelta(weeks=1)
        year = next_week.year
        week_num = next_week.isocalendar()[1]
        
        new_filename = f"{year}年第{week_num}周班级考勤记录表.xlsx"
        
        if new_filename in self.app.file_list:
            self.show_message(f"{new_filename} 已存在")
            return
        
        try:
            # 获取起始日索引 (0=星期一, 6=星期日)
            start_day_index = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"].index(self.app.week_start_day)
            
            # 按起始日重新排序星期
            week_days = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
            reordered_days = week_days[start_day_index:] + week_days[:start_day_index]
            
            # 创建基础DataFrame
            df = pd.DataFrame({"姓名": self.app.students})
            
            # 添加所有时段列（按新顺序）
            for day in reordered_days:
                for slot in self.app.time_slots:
                    col_name = f"{day}{slot}"
                    df[col_name] = ""  # 初始化为空
            
            # 应用长期请假
            long_leaves = self.app.leave_records["长期请假"]
            if not long_leaves.empty:
                for _, record in long_leaves.iterrows():
                    student = record["人名"]
                    leave_type = record["类型"]
                    if student in df["姓名"].values:
                        # 将该学生所有时段标记为请假类型
                        for col in df.columns:
                            if col != "姓名":
                                df.loc[df["姓名"] == student, col] = leave_type
            
            # 应用固定时段请假
            fixed_leaves = self.app.leave_records["固定时段"]
            if not fixed_leaves.empty:
                for _, record in fixed_leaves.iterrows():
                    student = record["人名"]
                    leave_type = record["类型"]
                    day = record["日期"]
                    slot = record["时段"]
                    col_name = f"{day}{slot}"
                    
                    if student in df["姓名"].values and col_name in df.columns:
                        df.loc[df["姓名"] == student, col_name] = leave_type
            
            # 更新按周请假的剩余次数
            weekly_leaves = self.app.leave_records["按周请假"]
            if not weekly_leaves.empty:
                # 将每周次数复制到剩余次数
                self.app.leave_records["按周请假"]["剩余次数"] = self.app.leave_records["按周请假"]["每周次数"]
                self.app.save_leave_records()
            
            # 保存新文件
            df.to_excel(new_filename, index=False)
            
            # 更新文件列表和当前文件
            self.app.file_list.append(new_filename)
            self.app.current_attendance_file = new_filename
            self.app.save_file_list()
            self.app.save_misc_settings()
            
            self.refresh_file_list()
            self.show_message(f"已创建 {new_filename}\n并应用请假信息")
        except Exception as e:
            self.show_message(f"创建失败: {e}")
    
    def deduct_week(self, instance):
        """减扣一周操作 - 修复版"""
        if not hasattr(self.app, 'leave_records'):
            self.show_message("请假记录未初始化")
            return
        
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text="本次操作将对按周请假的所有记录减扣一周，是否继续？"))
        
        btn_layout = BoxLayout(size_hint_y=None, height=50, spacing=10)
        btn_yes = Button(text="确定")
        btn_no = Button(text="取消")
        
        popup = Popup(title="确认", content=content, size_hint=(0.8, 0.4))
        
        def do_deduct(btn):
            try:
                if "按周请假" not in self.app.leave_records:
                    self.show_message("没有按周请假记录")
                    return
                    
                weekly_leaves = self.app.leave_records["按周请假"]
                if not weekly_leaves.empty:
                    # 修改1：先复制一份数据避免修改时的问题
                    weekly_leaves = weekly_leaves.copy()
                    
                    # 修改2：按不同情况处理
                    for index, record in weekly_leaves.iterrows():
                        if record["周数"] > 1:
                            # 周数大于1，只减扣周数
                            self.app.leave_records["按周请假"].at[index, "周数"] -= 1
                            self.app.leave_records["按周请假"].at[index, "剩余次数"] = self.app.leave_records["按周请假"].at[index, "每周次数"]
                        else:
                            # 周数等于1，删除记录
                            self.app.leave_records["按周请假"] = self.app.leave_records["按周请假"].drop(index)
                    
                    self.app.save_leave_records()
                    self.show_message("已成功减扣一周")
                else:
                    self.show_message("没有按周请假记录")
            except Exception as e:
                self.show_message(f"操作失败: {e}")
            popup.dismiss()
        
        btn_yes.bind(on_release=do_deduct)
        btn_no.bind(on_release=popup.dismiss)
        
        btn_layout.add_widget(btn_yes)
        btn_layout.add_widget(btn_no)
        content.add_widget(btn_layout)
        
        popup.open()
    def create_new_week(self, instance):
        """创建新的一周考勤表 - 按起始日排序并应用请假信息"""
        # 计算下周的年份和周数
        today = datetime.now()
        next_week = today + timedelta(weeks=1)
        year = next_week.year
        week_num = next_week.isocalendar()[1]
        
        new_filename = f"{year}年第{week_num}周班级考勤记录表.xlsx"
        
        if new_filename in self.app.file_list:
            self.show_message(f"{new_filename} 已存在")
            return
        
        try:
            # 获取起始日索引 (0=星期一, 6=星期日)
            start_day_index = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"].index(self.app.week_start_day)
            
            # 按起始日重新排序星期
            week_days = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
            reordered_days = week_days[start_day_index:] + week_days[:start_day_index]
            
            # 创建基础DataFrame
            df = pd.DataFrame({"姓名": self.app.students})
            
            # 添加所有时段列（按新顺序）
            for day in reordered_days:
                for slot in self.app.time_slots:
                    col_name = f"{day}{slot}"
                    df[col_name] = ""  # 初始化为空
            
            # 应用长期请假
            long_leaves = self.app.leave_records["长期请假"]
            if not long_leaves.empty:
                for _, record in long_leaves.iterrows():
                    student = record["人名"]
                    leave_type = record["类型"]
                    if student in df["姓名"].values:
                        # 将该学生所有时段标记为请假类型
                        for col in df.columns:
                            if col != "姓名":
                                df.loc[df["姓名"] == student, col] = leave_type
            
            # 应用固定时段请假
            fixed_leaves = self.app.leave_records["固定时段"]
            if not fixed_leaves.empty:
                for _, record in fixed_leaves.iterrows():
                    student = record["人名"]
                    leave_type = record["类型"]
                    day = record["日期"]
                    slot = record["时段"]
                    col_name = f"{day}{slot}"
                    
                    if student in df["姓名"].values and col_name in df.columns:
                        df.loc[df["姓名"] == student, col_name] = leave_type
            
            # 保存新文件
            df.to_excel(new_filename, index=False)
            
            # 更新文件列表和当前文件
            self.app.file_list.append(new_filename)
            self.app.current_attendance_file = new_filename
            self.app.save_file_list()
            self.app.save_misc_settings()
            
            self.refresh_file_list()
            self.show_message(f"已创建 {new_filename}\n并应用请假信息")
        except Exception as e:
            self.show_message(f"创建失败: {e}")

    def delete_file(self, filename):
        """删除考勤记录文件"""
        if filename == self.app.current_attendance_file:
            self.show_message("不能删除当前使用的文件")
            return
            
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=f"确定要删除 {filename} 吗？"))
        
        btn_layout = BoxLayout(size_hint_y=None, height=50, spacing=10)
        btn_yes = Button(text="确定")
        btn_no = Button(text="取消")
        
        popup = Popup(title="确认删除", content=content, size_hint=(0.8, 0.4))
        
        def do_delete(btn):
            try:
                # 从文件系统中删除
                if os.path.exists(filename):
                    os.remove(filename)
                
                # 从列表中删除
                if filename in self.app.file_list:
                    self.app.file_list.remove(filename)
                    self.app.save_file_list()
                
                self.refresh_file_list()
                self.show_message(f"已删除 {filename}")
            except Exception as e:
                self.show_message(f"删除失败: {e}")
            popup.dismiss()
        
        btn_yes.bind(on_release=do_delete)
        btn_no.bind(on_release=popup.dismiss)
        
        btn_layout.add_widget(btn_yes)
        btn_layout.add_widget(btn_no)
        content.add_widget(btn_layout)
        
        popup.open()
    
    def apply_changes(self, instance):
        """应用所有更改"""
        # 更新周起始日
        if self.app.week_start_day != self.week_start_spinner.text:
            self.app.week_start_day = self.week_start_spinner.text
            self.app.save_misc_settings()
        
        # 更新当前文件
        for check in self.file_checks:
            if check.state == 'down' and check.text != self.app.current_attendance_file:
                self.app.current_attendance_file = check.text
                self.app.save_misc_settings()
                break
        
        self.show_message("更改已保存")
    
    def generate_report(self, instance=None):  # 添加instance参数并设置默认值
        """生成统计报表"""
        try:
            if not hasattr(self.app, 'current_attendance_file') or not self.app.current_attendance_file:
                self.show_message("请先选择考勤记录文件")
                return
                
            # 读取当前考勤数据
            df = pd.read_excel(self.app.current_attendance_file)
            
            # 准备统计结果
            report_data = []
            
            for student in self.app.students:
                student_data = {"姓名": student}
                # 初始化所有类型计数为0
                for leave_type in self.app.leave_types:
                    student_data[leave_type] = 0
                
                # 统计各类型出现次数
                if student in df["姓名"].values:
                    student_row = df[df["姓名"] == student].iloc[0]
                    for col in df.columns:
                        if col != "姓名":
                            value = student_row[col]
                            if pd.notna(value) and value in student_data:
                                student_data[value] += 1
                
                report_data.append(student_data)
            
            # 创建报表DataFrame
            report_df = pd.DataFrame(report_data)
            
            # 从文件名中提取年份和周数
            filename = os.path.basename(self.app.current_attendance_file)
            match = re.search(r"(\d{4})年第(\d+)周", filename)
            if match:
                year, week_num = match.groups()
            else:
                year = datetime.now().year
                week_num = datetime.now().isocalendar()[1]
            
            # 保存报表
            report_filename = f"{year}年第{week_num}周班级考勤统计表.xlsx"
            report_df.to_excel(report_filename, index=False)
            
            self.show_message(f"报表已生成: {report_filename}")
        except Exception as e:
            print(f"生成报表错误: {e}")
            self.show_message(f"生成报表失败: {str(e)}")
    
    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())
    
    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text="确定", size_hint_y=0.3)
        popup = Popup(title="提示", content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

class LeaveStudentSelectScreen(BoxLayout):
    """学生选择界面（专用于请假或销假）"""
    def __init__(self, is_cancel=False, **kwargs):
        super(LeaveStudentSelectScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.app = App.get_running_app()
        self.is_cancel = is_cancel  # 直接通过参数确定是请假还是销假
        
        # 标题直接显示当前模式
        title = "销假记录 - 选择人员" if is_cancel else "请假记录 - 选择人员"
        self.add_widget(Label(text=title, font_size=20, size_hint_y=0.1))
        
        # 创建滚动区域
        scroll = ScrollView(size_hint=(1, 0.8))
        grid = GridLayout(
            cols=4,
            spacing=dp(10),
            size_hint_y=None,
            padding=[dp(10), dp(10)]
        )
        grid.bind(minimum_height=grid.setter('height'))
        
        # 为每个学生创建选择控件
        self.student_checks = {}
        for student in self.app.students:
            item = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=dp(50),
                spacing=dp(5)
            )
            
            # 销假模式使用ToggleButton（单选），请假模式使用CheckBox（多选）
            if is_cancel:
                check = ToggleButton(
                    group='students',  # 单选分组
                    size_hint=(None, None),
                    size=(dp(30), dp(30)),
                    allow_no_selection=False
                )
            else:
                check = CheckBox(
                    size_hint=(None, None),
                    size=(dp(30), dp(30))
                )
            label = Label(
                text=student,
                size_hint=(1, None),
                height=dp(30),
                halign='left',
                text_size=(dp(100), None)
            )
            
            item.add_widget(check)
            item.add_widget(label)
            grid.add_widget(item)
            
            self.student_checks[student] = check
        
        scroll.add_widget(grid)
        self.add_widget(scroll)
        
        # 底部按钮
        btn_layout = BoxLayout(
            size_hint_y=0.1,
            spacing=dp(10),
            padding=[dp(20), dp(5)]
        )
        btn_confirm = Button(
            text='确定销假' if is_cancel else '确定请假',
            size_hint=(None, None),
            size=(dp(100), dp(40))
        )
        btn_back = Button(
            text='返回',
            size_hint=(None, None),
            size=(dp(80), dp(40))
        )
        
        btn_confirm.bind(on_release=self.confirm_selection)
        btn_back.bind(on_release=self.go_back)
        
        btn_layout.add_widget(Widget())  # 左边填充
        btn_layout.add_widget(btn_back)
        btn_layout.add_widget(btn_confirm)
        btn_layout.add_widget(Widget())  # 右边填充
        
        self.add_widget(btn_layout)

    def confirm_selection(self, instance):
        """确认选择"""
        if self.is_cancel:
            # 销假模式 - 单选，使用state属性检查ToggleButton状态
            selected_students = [s for s, c in self.student_checks.items() if c.state == 'down']
            if not selected_students:
                self.show_message("请选择一名学生")
                return
            # 进入销假详情界面
            self.app.root.clear_widgets()
            self.app.root.add_widget(CancelLeaveScreen(selected_students[0]))
        else:
            # 请假模式 - 多选，使用active属性检查CheckBox状态
            selected_students = [s for s, c in self.student_checks.items() if c.active]
            if not selected_students:
                self.show_message("请至少选择一名学生")
                return
            # 进入请假类型选择界面
            self.app.root.clear_widgets()
            self.app.root.add_widget(LeaveTypeScreen(selected_students))

    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())

    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

class CancelLeaveScreen(BoxLayout):
    """销假界面 - 支持按类型筛选和删除记录"""
    def __init__(self, student, **kwargs):
        super(CancelLeaveScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.app = App.get_running_app()
        self.student = student
        self.selected_records = []  # 存储选中的记录
        
        # 标题
        self.add_widget(Label(
            text=f"为 {student} 销假", 
            font_size=20, 
            size_hint_y=0.1
        ))
        
        # 请假类型选择
        type_layout = BoxLayout(
            orientation='horizontal',
            size_hint_y=0.1,
            spacing=10
        )
        type_layout.add_widget(Label(
            text="请假类型:", 
            size_hint=(0.3, 1)
        ))
        
        self.type_spinner = Spinner(
            text='选择请假类型',
            values=['按周请假', '固定时段', '长期请假'],
            size_hint=(0.7, 1)
        )
        self.type_spinner.bind(text=self.on_type_selected)
        type_layout.add_widget(self.type_spinner)
        self.add_widget(type_layout)
        
        # 记录列表区域
        self.records_scroll = ScrollView(size_hint=(1, 0.7))
        self.records_layout = GridLayout(
            cols=1,
            spacing=10,
            size_hint_y=None,
            padding=[10, 10]
        )
        self.records_layout.bind(minimum_height=self.records_layout.setter('height'))
        self.records_scroll.add_widget(self.records_layout)
        self.add_widget(self.records_scroll)
        
        # 底部按钮
        btn_layout = BoxLayout(
            size_hint_y=0.1,
            spacing=10,
            padding=[10, 5]
        )
        btn_delete = Button(
            text='删除选中记录',
            size_hint=(0.5, 1)
        )
        btn_back = Button(
            text='返回',
            size_hint=(0.5, 1)
        )
        
        btn_delete.bind(on_release=self.delete_selected)
        btn_back.bind(on_release=self.go_back)
        
        btn_layout.add_widget(btn_delete)
        btn_layout.add_widget(btn_back)
        self.add_widget(btn_layout)

    def on_type_selected(self, spinner, text):
        """当选择请假类型时加载对应记录"""
        self.records_layout.clear_widgets()
        self.selected_records = []  # 存储选中的原始索引

        if text not in self.app.leave_records:
            return

        records = self.app.leave_records[text]
        student_records = records[records["人名"] == self.student]

        if student_records.empty:
            self.records_layout.add_widget(Label(
                text="没有找到相关请假记录",
                size_hint_y=None,
                height=40
            ))
            return

        # 重置索引前保留原始索引用于删除
        student_records = student_records.reset_index()  # 添加 'index' 列

        for idx, record in student_records.iterrows():
            record_item = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=40,
                spacing=10
            )

            try:
                original_index = int(record['index'])  # 确保是整数
            except (ValueError, TypeError) as e:
                print(f"[ERROR] 无法解析索引: {e}")
                continue

            print(f"原始索引: {original_index}, 类型: {type(original_index)}")  # 调试

            # 添加复选框
            chk = CheckBox(size_hint=(None, 1), width=40)

            # 封装避免闭包陷阱
            def make_handler(index):
                def handler(checkbox, value):
                    self.toggle_record(checkbox, index)
                return handler

            chk.bind(active=make_handler(original_index))

            # 显示信息
            if text == "按周请假":
                info = f"类型: {record['类型']}  每周次数: {record['每周次数']}  剩余周数: {record['周数']}"
            elif text == "固定时段":
                info = f"类型: {record['类型']}  日期: {record['日期']}  时段: {record['时段']}"
            else:  # 长期请假
                info = f"类型: {record['类型']}"

            record_item.add_widget(chk)
            record_item.add_widget(Label(
                text=info,
                halign='left',
                text_size=(Window.width - 100, None)
            ))

            self.records_layout.add_widget(record_item)

    def toggle_record(self, checkbox, record_index):
        """切换记录选择状态"""
        if checkbox.active:
            if record_index not in self.selected_records:
                self.selected_records.append(record_index)
        else:
            if record_index in self.selected_records:
                self.selected_records.remove(record_index)

    def delete_selected(self, instance):
        """删除选中的记录"""
        if not self.selected_records:
            self.show_message("请至少选择一条记录")
            return

        current_type = self.type_spinner.text
        if current_type not in self.app.leave_records:
            return

        df = self.app.leave_records[current_type]

        try:
            # 删除指定索引的记录
            df = df.drop(self.selected_records).reset_index(drop=True)
            self.app.leave_records[current_type] = df
            self.app.save_leave_records()
            self.show_message(f"已删除{len(self.selected_records)}条记录")

            # 清空选中记录并刷新界面
            self.selected_records = []
            self.on_type_selected(self.type_spinner, current_type)

        except KeyError as e:
            self.show_message(f"删除失败: 找不到记录索引 {e}")
    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())

    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()
class LeaveTypeScreen(BoxLayout):
    """请假类型选择界面"""
    def __init__(self, selected_students, **kwargs):
        super(LeaveTypeScreen, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.app = App.get_running_app()
        self.selected_students = selected_students
        self.leave_type = None
        self.leave_category = None
        
        # 标题
        self.add_widget(Label(
            text="请假信息", 
            font_size=20, 
            size_hint_y=0.1
        ))
        
        # 创建滚动区域
        scroll = ScrollView(size_hint=(1, 0.7))
        content = BoxLayout(
            orientation='vertical', 
            spacing=15, 
            size_hint_y=None,
            padding=[10, 10]
        )
        content.bind(minimum_height=content.setter('height'))
        
        # 请假类型选择 - 改为每行4个
        content.add_widget(Label(
            text="请选择请假类型:", 
            size_hint_y=None, 
            height=dp(30)
        ))
        
        self.type_buttons = {}
        type_group = GridLayout(
            cols=4,  # 改为4列
            spacing=dp(5),
            size_hint_y=None,
            height=dp(40)*((len(self.app.leave_types)+3)//4)  # 计算合适的高度
        )
        
        for ltype in self.app.leave_types:
            btn = ToggleButton(
                text=ltype,
                group='leave_type',
                size_hint=(None, None),
                size=(dp(80), dp(40)),  # 缩小按钮尺寸
                font_size='14sp'  # 减小字体
            )
            btn.bind(on_press=lambda btn, t=ltype: self.set_leave_type(t))
            type_group.add_widget(btn)
            self.type_buttons[ltype] = btn
        
        content.add_widget(type_group)
        
        # 请假类别选择
        content.add_widget(Label(
            text="请选择请假类别:", 
            size_hint_y=None, 
            height=dp(30)
        ))
        
        categories = ["按周请假", "固定时段", "长期请假"]
        self.category_buttons = {}
        category_group = GridLayout(
            cols=3, 
            size_hint_y=None, 
            height=dp(40)*len(categories)  # 缩小高度
        )
        for category in categories:
            btn = ToggleButton(
                text=category,
                group='leave_category',
                size_hint_y=None,
                height=dp(40),  # 缩小高度
                font_size='14sp'  # 减小字体
            )
            btn.bind(on_press=lambda btn, c=category: self.set_leave_category(c))
            category_group.add_widget(btn)
            self.category_buttons[category] = btn
        
        content.add_widget(category_group)
        
        # 按周请假选项
        self.weekly_options = BoxLayout(
            orientation='vertical', 
            size_hint_y=None, 
            height=dp(90),
            spacing=dp(5)
        )

        # 每周次数 - 一行布局
        times_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(40),
            spacing=dp(10)
        )
        times_row.add_widget(Label(
            text="每周次数:", 
            size_hint=(0.4, 1),
            halign='right'
        ))
        self.times_per_week = TextInput(
            text='1',
            input_filter='int',  # 只允许整数输入
            size_hint=(0.6, 1),
            multiline=False
        )
        times_row.add_widget(self.times_per_week)
        self.weekly_options.add_widget(times_row)

        # 请假周数 - 一行布局
        weeks_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(40),
            spacing=dp(10)
        )
        weeks_row.add_widget(Label(
            text="请假周数:", 
            size_hint=(0.4, 1),
            halign='right'
        ))
        self.weeks = TextInput(
            text='1',
            input_filter='int',  # 只允许整数输入
            size_hint=(0.6, 1),
            multiline=False
        )
        weeks_row.add_widget(self.weeks)
        self.weekly_options.add_widget(weeks_row)

        content.add_widget(self.weekly_options)
        self.weekly_options.opacity = 0
        
        # 固定时段选项
        self.fixed_options = BoxLayout(
            orientation='vertical', 
            size_hint_y=None, 
            height=dp(250)  # 调整高度
        )
        self.fixed_options.add_widget(Label(
            text="选择固定时段:", 
            size_hint_y=None, 
            height=dp(25)
        ))

        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        self.time_slot_checks = {}

        for day in weekdays:
            day_label = Label(
                text=day, 
                size_hint_y=None, 
                height=dp(25)
            )
            self.fixed_options.add_widget(day_label)
            
            day_layout = GridLayout(
                cols=len(self.app.time_slots)*2,  # 每列包含复选框和标签
                size_hint_y=None, 
                height=dp(35),
                spacing=dp(5)
            )
            
            for slot in self.app.time_slots:
                # 添加复选框
                check = CheckBox(
                    size_hint=(None, None),
                    size=(dp(30), dp(30))
                )
                day_layout.add_widget(check)
                self.time_slot_checks[(day, slot)] = check
                
                # 添加时段名称标签
                slot_label = Label(
                    text=slot,
                    size_hint=(None, None),
                    size=(dp(70), dp(30)),
                    halign='left'
                )
                day_layout.add_widget(slot_label)
            
            self.fixed_options.add_widget(day_layout)

        content.add_widget(self.fixed_options)
        self.fixed_options.opacity = 0
        
        scroll.add_widget(content)
        self.add_widget(scroll)
        
        # 底部按钮 - 缩小尺寸
        btn_layout = BoxLayout(
            size_hint_y=0.1,
            spacing=dp(10),
            padding=[dp(20), dp(5)]
        )
        btn_save = Button(
            text='保存',
            size_hint=(None, None),
            size=(dp(80), dp(40)),  # 缩小按钮
            font_size='14sp'
        )
        btn_back = Button(
            text='返回',
            size_hint=(None, None),
            size=(dp(80), dp(40)),  # 缩小按钮
            font_size='14sp'
        )
        
        btn_save.bind(on_release=self.save_leave)
        btn_back.bind(on_release=self.go_back)
        
        btn_layout.add_widget(Widget())  # 左边填充
        btn_layout.add_widget(btn_back)
        btn_layout.add_widget(btn_save)
        btn_layout.add_widget(Widget())  # 右边填充
        
        self.add_widget(btn_layout)
        
        Clock.schedule_once(self.update_ui)

    def set_leave_type(self, ltype):
        """设置请假类型"""
        self.leave_type = ltype

    def set_leave_category(self, category):
        """设置请假类别"""
        self.leave_category = category
        self.update_ui()

    def update_ui(self, dt=None):
        """更新UI显示"""
        # 隐藏所有选项
        self.weekly_options.opacity = 0
        self.weekly_options.disabled = True
        self.fixed_options.opacity = 0
        self.fixed_options.disabled = True
        
        # 根据选择的类别显示相应选项
        if self.leave_category == "按周请假":
            self.weekly_options.opacity = 1
            self.weekly_options.disabled = False
        elif self.leave_category == "固定时段":
            self.fixed_options.opacity = 1
            self.fixed_options.disabled = False

    def save_leave(self, instance):
        """保存请假信息"""
        if not self.leave_type:
            self.show_message("请选择请假类型")
            return
            
        if not self.leave_category:
            self.show_message("请选择请假类别")
            return
            
        if self.leave_category == "按周请假":
            try:
                times = int(self.times_per_week.text)
                weeks = int(self.weeks.text)
                
                if times <= 0 or weeks <= 0:
                    raise ValueError
                    
                # 为每个学生创建记录
                for student in self.selected_students:
                    new_record = pd.DataFrame({
                        "人名": [student],
                        "类型": [self.leave_type],
                        "每周次数": [times],
                        "周数": [weeks],
                        "剩余次数": [times]  # 初始化剩余次数等于每周次数
                    })
                    self.app.leave_records["按周请假"] = pd.concat([self.app.leave_records["按周请假"], new_record], ignore_index=True)
                
                self.app.save_leave_records()
                self.show_message("请假信息保存成功")
                self.go_back(None)
                
            except ValueError:
                self.show_message("请输入有效的数字")
                
        elif self.leave_category == "固定时段":
            # 收集选中的时段
            selected_slots = []
            for (day, slot), check in self.time_slot_checks.items():
                if check.active:
                    selected_slots.append((day, slot))
            
            if not selected_slots:
                self.show_message("请至少选择一个时段")
                return
                
            # 为每个学生和每个时段创建记录
            for student in self.selected_students:
                for day, slot in selected_slots:
                    new_record = pd.DataFrame({
                        "人名": [student],
                        "类型": [self.leave_type],
                        "日期": [day],
                        "时段": [slot]
                    })
                    self.app.leave_records["固定时段"] = pd.concat([self.app.leave_records["固定时段"], new_record], ignore_index=True)
            
            self.app.save_leave_records()
            self.show_message("请假信息保存成功")
            self.go_back(None)
            
        elif self.leave_category == "长期请假":
            # 为每个学生创建记录
            for student in self.selected_students:
                new_record = pd.DataFrame({
                    "人名": [student],
                    "类型": [self.leave_type]
                })
                self.app.leave_records["长期请假"] = pd.concat([self.app.leave_records["长期请假"], new_record], ignore_index=True)
            
            self.app.save_leave_records()
            self.show_message("请假信息保存成功")
            self.go_back(None)
            


    def go_back(self, instance):
        """返回主界面"""
        self.app.root.clear_widgets()
        self.app.root.add_widget(MainScreen())

    def show_message(self, message):
        """显示消息弹窗"""
        content = BoxLayout(orientation='vertical', spacing=10)
        content.add_widget(Label(text=message))
        btn = Button(text='确定', size_hint_y=0.3)
        popup = Popup(title='提示', content=content, size_hint=(0.7, 0.4))
        btn.bind(on_release=popup.dismiss)
        content.add_widget(btn)
        popup.open()

if __name__ == '__main__':
    AttendanceApp().run()