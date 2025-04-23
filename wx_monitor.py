import sys
import os
import time
import json
import threading
import datetime
import requests
import re  # 在文件顶部添加re模块引入
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QTextEdit, QLineEdit, QListWidget, 
                             QListWidgetItem, QCheckBox, QGroupBox, QSpinBox, QTabWidget,
                             QFileDialog, QMessageBox, QSplitter, QInputDialog)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread

from chat_monitor import WeChatMonitor
from chat_summarizer import DeepSeekSummarizer

class WeChatMonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("微信监控工具 by晚风(推特x.com/pl_wanfeng)")
        self.setGeometry(100, 100, 1000, 700)
        
        self.monitor = None
        self.monitor_thread = None
        self.is_monitoring = False
        self.chat_records = {}
        self.config_file = "monitor_config.json"
        
        # 初始化AI提示模板
        self.ai_prompt = "你是一个Web3撸毛的人，你非常擅长撸毛，你加入了一个群聊，你看过了所有人的聊天后，对他们聊的内容进行了重点分析，分析了哪些是项目相关的，哪些是要空投相关的，哪些是做任务的，并把看到的项目地址，需要做什么任务都分析出来，根据聊天内容的前后顺序，进行关联分析，要进行聊天的上下文关联，确保上下文关联的准确性，然后进行总结"
        
        self.init_ui()
        self.load_config()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        
        # 创建选项卡
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # 监控选项卡
        monitor_tab = QWidget()
        monitor_layout = QVBoxLayout(monitor_tab)
        
        # 上部分：配置面板
        config_group = QGroupBox("监控配置")
        config_layout = QVBoxLayout()
        
        # 微信群聊选择
        chat_layout = QHBoxLayout()
        chat_layout.addWidget(QLabel("监控的群聊:"))
        self.chat_list = QListWidget()
        self.chat_list.setMaximumHeight(120)
        chat_layout.addWidget(self.chat_list)
        
        # 群聊操作按钮垂直布局
        chat_buttons = QVBoxLayout()
        
        # 刷新群聊列表按钮
        refresh_btn = QPushButton("刷新群聊列表")
        refresh_btn.clicked.connect(self.refresh_chat_list)
        chat_buttons.addWidget(refresh_btn)
        
        # 添加手动输入群聊按钮
        manual_add_btn = QPushButton("手动添加群聊")
        manual_add_btn.clicked.connect(self.manual_add_chat)
        chat_buttons.addWidget(manual_add_btn)
        
        # 清空选择按钮
        clear_btn = QPushButton("清空选择")
        clear_btn.clicked.connect(self.clear_chat_selection)
        chat_buttons.addWidget(clear_btn)
        
        chat_layout.addLayout(chat_buttons)
        
        config_layout.addLayout(chat_layout)
        
        # 监控时间设置
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("监控时间(分钟):"))
        self.time_spin = QSpinBox()
        self.time_spin.setRange(1, 1440)  # 1分钟到24小时
        self.time_spin.setValue(60)  # 默认1小时
        time_layout.addWidget(self.time_spin)
        
        # 添加检测间隔设置
        time_layout.addWidget(QLabel("检测间隔(秒):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(5, 300)  # 5秒到5分钟
        self.interval_spin.setValue(10)  # 默认10秒
        time_layout.addWidget(self.interval_spin)
        
        time_layout.addStretch()
        
        # DeepSeek API配置
        time_layout.addWidget(QLabel("DeepSeek API密钥:"))
        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.Password)
        time_layout.addWidget(self.api_key_input)
        
        config_layout.addLayout(time_layout)
        
        # Webhook配置
        webhook_layout = QHBoxLayout()
        webhook_layout.addWidget(QLabel("飞书Webhook URL:"))
        self.webhook_input = QLineEdit()
        webhook_layout.addWidget(self.webhook_input)
        
        self.webhook_enabled = QCheckBox("启用Webhook")
        webhook_layout.addWidget(self.webhook_enabled)
        
        config_layout.addLayout(webhook_layout)
        
        # 添加AI提示设置按钮
        ai_prompt_layout = QHBoxLayout()
        self.set_ai_prompt_btn = QPushButton("设置AI提示模板")
        self.set_ai_prompt_btn.clicked.connect(self.set_ai_prompt)
        ai_prompt_layout.addWidget(self.set_ai_prompt_btn)
        
        reset_prompt_btn = QPushButton("重置为默认提示")
        reset_prompt_btn.clicked.connect(self.reset_ai_prompt)
        ai_prompt_layout.addWidget(reset_prompt_btn)
        
        config_layout.addLayout(ai_prompt_layout)
        
        # 按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始监控")
        self.start_btn.clicked.connect(self.toggle_monitoring)
        btn_layout.addWidget(self.start_btn)
        
        self.save_config_btn = QPushButton("保存配置")
        self.save_config_btn.clicked.connect(self.save_config)
        btn_layout.addWidget(self.save_config_btn)
        
        config_layout.addLayout(btn_layout)
        config_group.setLayout(config_layout)
        monitor_layout.addWidget(config_group)
        
        # 下部分：状态和实时消息
        status_splitter = QSplitter(Qt.Vertical)
        
        # 状态面板
        status_group = QGroupBox("监控状态")
        status_layout = QVBoxLayout()
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        status_layout.addWidget(self.status_text)
        status_group.setLayout(status_layout)
        status_splitter.addWidget(status_group)
        
        # 实时消息面板
        message_group = QGroupBox("实时消息")
        message_layout = QVBoxLayout()
        self.message_text = QTextEdit()
        self.message_text.setReadOnly(True)
        message_layout.addWidget(self.message_text)
        message_group.setLayout(message_layout)
        status_splitter.addWidget(message_group)
        
        monitor_layout.addWidget(status_splitter, 1)  # 设置拉伸因子
        
        # 总结选项卡
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)
        
        # 总结列表
        self.summary_list = QListWidget()
        self.summary_list.currentRowChanged.connect(self.show_summary)
        summary_layout.addWidget(QLabel("历史总结:"))
        summary_layout.addWidget(self.summary_list)
        
        # 总结内容
        summary_layout.addWidget(QLabel("总结内容:"))
        self.summary_content = QTextEdit()
        self.summary_content.setReadOnly(True)
        summary_layout.addWidget(self.summary_content)
        
        # 手动总结按钮
        summary_btn_layout = QHBoxLayout()
        self.manual_summary_btn = QPushButton("手动总结当前记录")
        self.manual_summary_btn.clicked.connect(self.manual_summarize)
        summary_btn_layout.addWidget(self.manual_summary_btn)
        
        self.export_btn = QPushButton("导出总结")
        self.export_btn.clicked.connect(self.export_summary)
        summary_btn_layout.addWidget(self.export_btn)
        
        summary_layout.addLayout(summary_btn_layout)
        
        # 添加选项卡
        tab_widget.addTab(monitor_tab, "监控")
        tab_widget.addTab(summary_tab, "总结")
        
        # 初始状态更新
        self.update_status("应用已启动，请配置监控参数")

    def refresh_chat_list(self):
        """刷新微信群聊列表"""
        self.update_status("正在获取微信群聊列表...")
        self.chat_list.clear()  # 先清空列表
        
        # 禁用按钮，防止重复点击
        refresh_btn = self.sender()
        if refresh_btn:
            old_text = refresh_btn.text()
            refresh_btn.setText("正在刷新...")
            refresh_btn.setEnabled(False)
            QApplication.processEvents()  # 强制更新UI
        
        try:
            # 显示加载状态
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            temp_monitor = WeChatMonitor()
            chats = temp_monitor.get_chat_list()
            
            if not chats:
                self.update_status("未找到任何群聊，请确保微信已登录且有可用的聊天会话")
                QMessageBox.warning(self, "警告", "未找到任何群聊，请检查微信是否正常登录")
                return
            
            for chat in chats:
                item = QListWidgetItem(chat)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.chat_list.addItem(item)
            
            self.update_status(f"成功获取到 {len(chats)} 个群聊")
        except Exception as e:
            error_msg = str(e)
            self.update_status(f"获取群聊失败: {error_msg}")
            QMessageBox.critical(self, "错误", f"获取群聊列表失败：\n{error_msg}\n\n请确保：\n1. 微信已正常登录\n2. 微信窗口未被最小化\n3. 微信版本与wxauto兼容")
        finally:
            # 恢复按钮状态
            if refresh_btn:
                refresh_btn.setText(old_text)
                refresh_btn.setEnabled(True)
            
            # 恢复鼠标指针
            QApplication.restoreOverrideCursor()
    
    def toggle_monitoring(self):
        """开始或停止监控"""
        if not self.is_monitoring:
            # 获取选中的群聊
            selected_chats = []
            for i in range(self.chat_list.count()):
                item = self.chat_list.item(i)
                if item.checkState() == Qt.Checked:
                    selected_chats.append(item.text())
            
            if not selected_chats:
                QMessageBox.warning(self, "警告", "请至少选择一个群聊进行监控")
                return
            
            # 获取监控时间（分钟）
            monitor_time = self.time_spin.value() * 60  # 转换为秒
            
            # 获取DeepSeek API密钥
            api_key = self.api_key_input.text()
            if not api_key:
                QMessageBox.warning(self, "警告", "请输入DeepSeek API密钥")
                return
            
            # 获取Webhook URL（可选）
            webhook_url = self.webhook_input.text() if self.webhook_enabled.isChecked() else None
            
            # 获取检测间隔设置
            check_interval = self.interval_spin.value()
            
            # 提示用户确认
            msg = f"将开始监控以下群聊，持续 {self.time_spin.value()} 分钟，检测间隔: {check_interval} 秒：\n\n"
            for chat in selected_chats:
                msg += f"• {chat}\n"
            
            msg += "\n请确保：\n• 微信保持在前台或可见状态\n• 期间不要手动切换微信聊天窗口"
            
            reply = QMessageBox.question(self, "确认开始监控", msg, 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply == QMessageBox.Yes:
                # 开始监控
                try:
                    self.start_monitoring(selected_chats, monitor_time, api_key, webhook_url)
                    self.start_btn.setText("停止监控")
                except Exception as e:
                    error_msg = str(e)
                    self.update_status(f"启动监控失败: {error_msg}")
                    QMessageBox.critical(self, "错误", f"启动监控失败：\n{error_msg}")
        else:
            # 停止监控
            reply = QMessageBox.question(self, "确认停止", "确定要停止正在进行的监控吗？\n已收集的消息将会保留。", 
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.stop_monitoring()
                self.start_btn.setText("开始监控")
                self.update_status("监控已手动停止")
    
    def start_monitoring(self, chats, duration, api_key, webhook_url=None):
        """启动监控线程"""
        try:
            self.monitor = WeChatMonitor()
            
            # 创建总结器
            self.summarizer = DeepSeekSummarizer(api_key)
            
            # 状态初始化
            self.chat_records = {chat: [] for chat in chats}
            self.is_monitoring = True
            
            # 获取检测间隔设置
            check_interval = self.interval_spin.value()
            
            # 创建并启动监控线程
            self.monitor_thread = MonitorThread(self.monitor, chats, duration, check_interval)
            self.monitor_thread.message_signal.connect(self.handle_new_message)
            self.monitor_thread.complete_signal.connect(lambda: self.handle_monitor_complete(webhook_url))
            self.monitor_thread.status_signal.connect(self.update_status)
            self.monitor_thread.start()
            
            # 更新状态
            chats_str = ", ".join(chats)
            self.update_status(f"开始监控群聊: {chats_str}，持续时间: {duration//60} 分钟")
            
            # 添加预计结束时间
            end_time = datetime.datetime.now() + datetime.timedelta(seconds=duration)
            self.update_status(f"预计结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            # 添加检测间隔信息
            self.update_status(f"检测间隔: {check_interval} 秒")
            
            # 添加操作建议
            self.update_status("建议: 请保持微信窗口可见，不要最小化或遮挡微信窗口")
        except Exception as e:
            self.is_monitoring = False
            raise Exception(f"启动监控失败: {str(e)}")
    
    def stop_monitoring(self):
        """停止监控线程"""
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.stop()
            self.monitor_thread.wait()
        
        self.is_monitoring = False
        self.update_status("监控已停止")
    
    def handle_new_message(self, chat_name, sender, content, timestamp):
        """处理新消息"""
        try:
            # 添加到记录
            if chat_name in self.chat_records:
                self.chat_records[chat_name].append({
                    "sender": sender,
                    "content": content,
                    "timestamp": timestamp
                })
            
            # 更新实时消息显示
            time_str = datetime.datetime.fromtimestamp(timestamp).strftime("%H:%M:%S")
            
            # 对内容进行HTML转义，防止特殊字符破坏格式
            import html
            safe_content = html.escape(content)
            
            # 创建更醒目的消息格式
            message_html = f"""
            <div style="margin: 5px 0; padding: 5px; border-left: 3px solid #4a7ebb;">
                <b style="color: #2c3e50;">[{chat_name}]</b> 
                <span style="color: #3498db; font-weight: bold;">{sender}</span> 
                <span style="color: #7f8c8d; font-size: 0.9em;">({time_str})</span><br/>
                <span style="margin-left: 10px;">{safe_content}</span>
            </div>
            """
            
            # 更新UI
            self.message_text.append(message_html)
            
            # 确保滚动到最新消息
            self.message_text.ensureCursorVisible()
            
            # 如果现在正在监控，更新状态栏
            if self.is_monitoring:
                self.update_status(f"收到新消息 [{chat_name}] {sender}: {content[:30]}{'...' if len(content) > 30 else ''}")
        except Exception as e:
            self.update_status(f"处理新消息时出错: {str(e)}")
    
    def handle_monitor_complete(self, webhook_url=None):
        """监控完成后的处理"""
        self.is_monitoring = False
        self.start_btn.setText("开始监控")
        
        # 自动生成总结
        self.update_status("监控完成，正在生成总结...")
        
        empty_chats = []
        for chat_name, messages in self.chat_records.items():
            if not messages:
                empty_chats.append(chat_name)
                continue
                
            try:
                self.summarize_chat(chat_name, messages, webhook_url)
            except Exception as e:
                self.update_status(f"总结群聊 {chat_name} 失败: {str(e)}")
        
        if empty_chats:
            empty_str = ", ".join(empty_chats)
            self.update_status(f"以下群聊没有消息，跳过总结: {empty_str}")
        
        self.update_status("所有总结完成")
    
    def summarize_chat(self, chat_name, messages, webhook_url=None):
        """总结群聊记录"""
        # 转换消息格式
        messages_text = []
        for msg in messages:
            time_str = datetime.datetime.fromtimestamp(msg["timestamp"]).strftime("%H:%M:%S")
            messages_text.append(f"{time_str} {msg['sender']}: {msg['content']}")
        
        messages_str = "\n".join(messages_text)
        
        # 生成总结，传递AI提示模板
        summary = self.summarizer.summarize(messages_str, self.ai_prompt)
        
        # 生成时间戳和标题
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
        title = f"{chat_name} - {timestamp}"
        
        # 保存总结
        summary_obj = {
            "title": title,
            "chat_name": chat_name,
            "timestamp": timestamp,
            "summary": summary,
            "messages": messages
        }
        
        # 添加到总结列表
        self.add_summary_to_list(summary_obj)
        
        # 发送webhook（如果启用）
        if webhook_url:
            try:
                self.send_webhook(webhook_url, chat_name, summary, timestamp)
                self.update_status(f"已发送总结到Webhook: {chat_name}")
            except Exception as e:
                self.update_status(f"发送Webhook失败: {str(e)}")
        
        self.update_status(f"完成群聊总结: {chat_name}")
    
    def add_summary_to_list(self, summary_obj):
        """添加总结到列表"""
        # 存储总结对象
        if not hasattr(self, 'summaries'):
            self.summaries = []
        
        self.summaries.append(summary_obj)
        
        # 添加到UI列表
        self.summary_list.addItem(summary_obj["title"])
        
        # 选中新添加的项
        self.summary_list.setCurrentRow(self.summary_list.count() - 1)
    
    def _format_summary_content(self, content):
        """格式化总结内容，增强可读性"""
        if not content:
            return "<p class='summary-empty'>无内容</p>"
            
        # 将内容中的HTML特殊字符进行转义
        import html
        content = html.escape(content)
        
        # 将星号标记(**文本**)转换为HTML加粗标签
        formatted = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', content)
        
        # 处理Markdown风格的标题 (## 标题)
        formatted = re.sub(r'^##\s+(.+)$', r'<h4>\1</h4>', formatted, flags=re.MULTILINE)
        
        # 处理列表项的不同形式
        # 1. 数字序号（1. 2. 3.）
        formatted = re.sub(r'^(\d+)\.\s+(.+)$', r'<div class="summary-point"><span class="point-number">\1.</span> \2</div>', 
                           formatted, flags=re.MULTILINE)
        
        # 2. 短横线（- 项目）
        formatted = re.sub(r'^-\s+(.+)$', r'<div class="dash-item">\1</div>', 
                          formatted, flags=re.MULTILINE)
        
        # 3. 星号项（* 项目）
        formatted = re.sub(r'^\*\s+(.+)$', r'<div class="dash-item">\1</div>', 
                          formatted, flags=re.MULTILINE)
        
        # 给重要的关键词加粗
        keywords = ["主要", "重点", "建议", "结论", "计划", "任务", "链接", "地址", "收益", "项目", "空投", "机会", "警告"]
        
        # 使用分词方式替换，避免在HTML标签内替换
        parts = re.split(r'(<[^>]*>)', formatted)
        result = []
        
        for part in parts:
            # 如果不是HTML标签，处理关键词
            if not part.startswith('<') or not part.endswith('>'):
                for keyword in keywords:
                    # 确保只替换完整的词，而不是部分匹配
                    part = re.sub(r'(?<!\w)(' + keyword + r')(?!\w)', r'<b>\1</b>', part)
            result.append(part)
        
        formatted = "".join(result)
        
        # 处理段落：空行转换为段落分隔
        paragraphs = formatted.split('\n\n')
        formatted_paragraphs = []
        
        for p in paragraphs:
            if p.strip():
                if not (p.startswith('<div') or p.startswith('<h4')):
                    formatted_paragraphs.append(f"<p>{p.replace('\n', '<br>')}</p>")
                else:
                    formatted_paragraphs.append(p)
        
        return "".join(formatted_paragraphs)

    def show_summary(self, row):
        """显示选中的总结内容"""
        if row >= 0 and hasattr(self, 'summaries') and row < len(self.summaries):
            summary = self.summaries[row]
            
            # 修改HTML内容格式，使用更现代美观的样式
            html_content = f"""
            <html>
            <head>
            <style>
                body {{
                    font-family: 'Microsoft YaHei', Arial, sans-serif;
                    line-height: 1.6;
                    margin: 0;
                    padding: 15px;
                    background-color: #f9f9f9;
                    color: #333;
                }}
                .summary-container {{
                    background-color: white;
                    border-radius: 8px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    padding: 20px;
                }}
                .chat-title {{
                    color: #2c3e50;
                    border-bottom: 2px solid #3498db;
                    padding-bottom: 10px;
                    margin-bottom: 15px;
                    font-size: 18px;
                }}
                .timestamp {{
                    color: #7f8c8d;
                    font-size: 14px;
                    margin-bottom: 15px;
                }}
                .summary-title {{
                    background-color: #f0f7ff;
                    padding: 8px 15px;
                    border-left: 4px solid #3498db;
                    margin: 15px 0;
                    font-weight: bold;
                    font-size: 16px;
                }}
                .summary-content {{
                    padding: 10px;
                    line-height: 1.8;
                    white-space: pre-line;
                    text-align: justify;
                }}
                .summary-point {{
                    margin: 10px 0;
                    padding-left: 20px;
                    position: relative;
                }}
                .summary-point:before {{
                    content: "•";
                    position: absolute;
                    left: 0;
                    color: #3498db;
                    font-weight: bold;
                }}
            </style>
            </head>
            <body>
            <div class="summary-container">
                <div class="chat-title">{summary['chat_name']}</div>
                <div class="timestamp">时间: {summary['timestamp']}</div>
                <div class="summary-title">会话总结</div>
                <div class="summary-content">{self._format_summary_content(summary['summary'])}</div>
            </div>
            </body>
            </html>
            """
            
            self.summary_content.setHtml(html_content)
    
    def manual_summarize(self):
        """手动总结当前记录"""
        if not self.chat_records:
            QMessageBox.warning(self, "警告", "没有可用的聊天记录进行总结")
            return
        
        api_key = self.api_key_input.text()
        if not api_key:
            QMessageBox.warning(self, "警告", "请输入DeepSeek API密钥")
            return
        
        webhook_url = self.webhook_input.text() if self.webhook_enabled.isChecked() else None
        
        self.update_status("正在生成手动总结...")
        
        # 为每个有消息的群聊生成总结
        for chat_name, messages in self.chat_records.items():
            if not messages:
                continue
                
            try:
                self.summarize_chat(chat_name, messages, webhook_url)
            except Exception as e:
                self.update_status(f"总结群聊 {chat_name} 失败: {str(e)}")
        
        self.update_status("手动总结完成")
    
    def export_summary(self):
        """导出总结"""
        if not hasattr(self, 'summaries') or not self.summaries:
            QMessageBox.warning(self, "警告", "没有可用的总结内容可导出")
            return
        
        # 选择保存路径
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出总结", "", "文本文件 (*.txt);;HTML文件 (*.html);;JSON文件 (*.json)"
        )
        
        if not file_path:
            return
        
        try:
            # 根据文件扩展名选择导出格式
            if file_path.endswith('.txt'):
                self.export_as_text(file_path)
            elif file_path.endswith('.html'):
                self.export_as_html(file_path)
            elif file_path.endswith('.json'):
                self.export_as_json(file_path)
            else:
                # 默认为文本格式
                self.export_as_text(file_path)
            
            self.update_status(f"总结已导出到: {file_path}")
        except Exception as e:
            self.update_status(f"导出总结失败: {str(e)}")
    
    def export_as_text(self, file_path):
        """导出为文本格式"""
        with open(file_path, 'w', encoding='utf-8') as f:
            for summary in self.summaries:
                f.write(f"群聊: {summary['chat_name']}\n")
                f.write(f"时间: {summary['timestamp']}\n")
                f.write("总结:\n")
                f.write(summary['summary'])
                f.write("\n\n" + "-"*40 + "\n\n")
    
    def _generate_message_chart(self, messages):
        """生成简单的消息数量图表HTML"""
        if not messages:
            return ""
            
        # 统计每个人的发言次数（最多显示前8名）
        sender_counts = {}
        for msg in messages:
            sender = msg.get('sender', '未知用户')
            if sender in sender_counts:
                sender_counts[sender] += 1
            else:
                sender_counts[sender] = 1
        
        # 按发言次数排序并取前8名
        top_senders = sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)[:8]
        
        # 如果没有数据，返回空
        if not top_senders:
            return ""
            
        # 计算最大值用于比例缩放
        max_count = max([count for _, count in top_senders])
        
        # 生成图表HTML
        chart_html = """
        <div class="chart-container">
            <h3>活跃发言者统计</h3>
            <div class="chart">
        """
        
        # 为每个发言者创建一个条形图条目
        for sender, count in top_senders:
            # 计算百分比宽度
            percentage = (count / max_count) * 100
            chart_html += f"""
            <div class="chart-item">
                <div class="chart-label">{sender}</div>
                <div class="chart-bar-container">
                    <div class="chart-bar" style="width: {percentage}%"></div>
                    <span class="chart-value">{count}</span>
                </div>
            </div>
            """
        
        chart_html += """
            </div>
        </div>
        """
        
        return chart_html

    def _generate_word_cloud(self, messages):
        """生成简单的词云HTML"""
        if not messages:
            return ""
            
        # 收集所有消息文本
        all_text = " ".join([msg.get('content', '') for msg in messages if isinstance(msg.get('content', ''), str)])
        
        # 简单的中文分词（不使用第三方库）
        # 这里使用一个简单方法，实际应用中可以使用jieba等分词库
        words = []
        for char in all_text:
            if '\u4e00' <= char <= '\u9fff':  # 是中文字符
                words.append(char)
        
        # 过滤掉常见停用词和短字符
        stop_words = set(["的", "了", "在", "是", "我", "有", "和", "就", "不", "人", "都", "一", "一个", "上", "也", "很", "到", "说", " ", ""])
        filtered_words = [w for w in words if w not in stop_words]
        
        # 统计词频
        word_counts = {}
        for word in filtered_words:
            if word in word_counts:
                word_counts[word] += 1
            else:
                word_counts[word] = 1
        
        # 取频率最高的30个词
        top_words = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)[:30]
        
        # 如果没有足够的词，返回空
        if len(top_words) < 5:
            return ""
            
        # 计算最大和最小频率用于字体大小缩放
        max_count = max([count for _, count in top_words])
        min_count = min([count for _, count in top_words])
        
        # 生成词云HTML
        cloud_html = """
        <div class="cloud-container">
            <h3>热门词汇</h3>
            <div class="word-cloud">
        """
        
        # 生成随机颜色
        colors = ['#1890ff', '#52c41a', '#f5222d', '#fa8c16', '#722ed1', '#13c2c2', '#eb2f96']
        import random
        
        # 为每个词创建一个span
        for word, count in top_words:
            # 计算字体大小（12px-24px）
            if max_count == min_count:
                font_size = 18
            else:
                font_size = 12 + ((count - min_count) / (max_count - min_count)) * 12
                
            # 随机选择颜色
            color = random.choice(colors)
            
            cloud_html += f'<span class="cloud-word" style="font-size:{font_size}px;color:{color}">{word}</span>'
        
        cloud_html += """
            </div>
        </div>
        """
        
        return cloud_html

    def export_as_html(self, file_path):
        """导出为HTML格式"""
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>微信群聊总结</title>
            <style>
                :root {
                    --primary-color: #1890ff;
                    --primary-light: #e6f7ff;
                    --secondary-color: #52c41a;
                    --text-color: #262626;
                    --text-secondary: #595959;
                    --text-light: #8c8c8c;
                    --bg-color: #f0f2f5;
                    --card-bg: #ffffff;
                    --border-radius: 8px;
                    --shadow: rgba(0, 0, 0, 0.1) 0px 4px 12px;
                }
                
                * {
                    box-sizing: border-box;
                    margin: 0;
                    padding: 0;
                }
                
                body { 
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Microsoft YaHei', sans-serif;
                    font-size: 14px;
                    line-height: 1.6;
                    color: var(--text-color);
                    background-color: var(--bg-color);
                    padding: 20px;
                    max-width: 1000px;
                    margin: 0 auto;
                }
                
                header {
                    text-align: center;
                    margin-bottom: 40px;
                }
                
                h1 {
                    font-size: 28px;
                    color: var(--text-color);
                    margin-bottom: 10px;
                    font-weight: 600;
                }
                
                .page-subtitle {
                    color: var(--text-light);
                    font-size: 14px;
                    margin-bottom: 30px;
                }
                
                .summaries-container {
                    display: grid;
                    gap: 24px;
                }
                
                .summary { 
                    background-color: var(--card-bg);
                    border-radius: var(--border-radius);
                    box-shadow: var(--shadow);
                    overflow: hidden;
                }
                
                .summary-header {
                    padding: 16px 20px;
                    border-bottom: 1px solid rgba(0, 0, 0, 0.06);
                    background-color: #fafafa;
                }
                
                .summary h2 { 
                    color: var(--text-color);
                    font-size: 18px;
                    margin: 0;
                    font-weight: 600;
                    display: flex;
                    align-items: center;
                }
                
                .summary h2::before {
                    content: '';
                    display: inline-block;
                    width: 4px;
                    height: 18px;
                    background-color: var(--primary-color);
                    margin-right: 10px;
                    border-radius: 2px;
                }
                
                .summary-body {
                    padding: 20px;
                }
                
                .summary .timestamp {
                    color: var(--text-light);
                    font-size: 14px;
                    margin-top: 4px;
                    display: flex;
                    align-items: center;
                }
                
                .summary .timestamp::before {
                    content: '⏱️';
                    margin-right: 6px;
                    font-size: 12px;
                }
                
                .summary h3 {
                    background-color: var(--primary-light);
                    padding: 10px 16px;
                    margin: 20px 0 16px;
                    border-radius: 4px;
                    font-size: 16px;
                    font-weight: 500;
                    color: var(--primary-color);
                    position: relative;
                }
                
                .summary h4 {
                    margin: 16px 0 10px;
                    color: var(--text-color);
                    font-size: 15px;
                    font-weight: 500;
                }
                
                .summary-content {
                    line-height: 1.8;
                    text-align: justify;
                    padding: 0 5px;
                    color: var(--text-secondary);
                }
                
                .summary-content p {
                    margin-bottom: 12px;
                }
                
                .summary-point {
                    margin: 12px 0;
                    padding-left: 24px;
                    position: relative;
                }
                
                .summary-point:before {
                    content: "•";
                    position: absolute;
                    left: 8px;
                    color: var(--primary-color);
                    font-weight: bold;
                }
                
                .dash-item {
                    margin: 8px 0 8px 16px;
                    padding-left: 20px;
                    position: relative;
                }
                
                .dash-item:before {
                    content: "-";
                    position: absolute;
                    left: 0;
                    color: var(--primary-color);
                }
                
                b {
                    color: var(--primary-color);
                    font-weight: 600;
                }
                
                .footer {
                    margin-top: 40px;
                    padding-top: 20px;
                    text-align: center;
                    border-top: 1px solid rgba(0, 0, 0, 0.06);
                    color: var(--text-light);
                    font-size: 12px;
                }
                
                /* 图表样式 */
                .chart-container {
                    margin-top: 20px;
                    padding-top: 10px;
                    border-top: 1px solid rgba(0, 0, 0, 0.06);
                }
                
                .chart {
                    margin-top: 15px;
                }
                
                .chart-item {
                    display: flex;
                    margin-bottom: 12px;
                    align-items: center;
                }
                
                .chart-label {
                    width: 80px;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                    color: var(--text-secondary);
                    font-size: 13px;
                }
                
                .chart-bar-container {
                    flex: 1;
                    display: flex;
                    align-items: center;
                    height: 20px;
                }
                
                .chart-bar {
                    height: 12px;
                    background-color: var(--primary-color);
                    border-radius: 6px;
                    min-width: 4px;
                }
                
                .chart-value {
                    margin-left: 8px;
                    font-size: 12px;
                    color: var(--text-secondary);
                }
                
                /* 词云样式 */
                .cloud-container {
                    margin-top: 20px;
                    padding-top: 10px;
                    border-top: 1px solid rgba(0, 0, 0, 0.06);
                }
                
                .word-cloud {
                    margin: 15px 0;
                    text-align: center;
                    background: #fafafa;
                    padding: 20px;
                    border-radius: 4px;
                    min-height: 120px;
                    display: flex;
                    flex-wrap: wrap;
                    justify-content: center;
                    align-items: center;
                }
                
                .cloud-word {
                    display: inline-block;
                    padding: 4px 8px;
                    margin: 4px;
                    border-radius: 4px;
                    transition: transform 0.2s;
                }
                
                .cloud-word:hover {
                    transform: scale(1.2);
                }
                
                .summary-empty {
                    color: var(--text-light);
                    font-style: italic;
                }
                
                /* 信息板块 */
                .info-panels {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 20px;
                    margin-top: 20px;
                }
                
                @media screen and (max-width: 600px) {
                    .info-panels {
                        grid-template-columns: 1fr;
                    }
                }
                
                @media screen and (max-width: 768px) {
                    body {
                        padding: 16px;
                    }
                    
                    .summary-header {
                        padding: 12px 16px;
                    }
                    
                    .summary-body {
                        padding: 16px;
                    }
                    
                    h1 {
                        font-size: 24px;
                    }
                }
            </style>
        </head>
        <body>
            <header>
                <h1>微信群聊总结</h1>
                <div class="page-subtitle">由微信群聊监控工具生成于 %s</div>
            </header>
            
            <div class="summaries-container">
        """ % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        
        for summary in self.summaries:
            # 为每条总结添加更美观的HTML格式
            formatted_summary = self._format_summary_content(summary['summary'])
            
            # 查找对应的聊天记录，生成图表和词云
            messages = self.chat_records.get(summary['chat_name'], [])
            chart_html = self._generate_message_chart(messages)
            cloud_html = self._generate_word_cloud(messages)
            
            # 如果有图表和词云，将它们放在并排的布局中
            visualizations = ""
            if chart_html or cloud_html:
                visualizations = f"""
                <div class="info-panels">
                    {chart_html}
                    {cloud_html}
                </div>
                """
            else:
                # 如果只有其中一个，则使用普通布局
                visualizations = chart_html + cloud_html
            
            html_content += f"""
            <div class="summary">
                <div class="summary-header">
                    <h2>{summary['chat_name']}</h2>
                    <div class="timestamp">总结时间: {summary['timestamp']}</div>
                </div>
                <div class="summary-body">
                    <h3>会话总结</h3>
                    <div class="summary-content">{formatted_summary}</div>
                    {visualizations}
                </div>
            </div>
            """
        
        html_content += """
            </div>
            <div class="footer">
                生成于微信群聊监控工具 | 总结内容仅供参考
            </div>
        </body>
        </html>
        """
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    
    def export_as_json(self, file_path):
        """导出为JSON格式"""
        export_data = []
        for summary in self.summaries:
            # 创建简化版本，不包含原始消息
            export_item = {
                "chat_name": summary['chat_name'],
                "timestamp": summary['timestamp'],
                "summary": summary['summary']
            }
            export_data.append(export_item)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
    
    def send_webhook(self, webhook_url, chat_name, summary, timestamp):
        """发送飞书Webhook"""
        # 飞书消息卡片格式
        post_data = {
            "msg_type": "interactive",
            "card": {
                "config": {
                    "wide_screen_mode": True
                },
                "header": {
                    "title": {
                        "tag": "plain_text",
                        "content": f"微信群聊总结 - {chat_name}"
                    },
                    "template": "blue"
                },
                "elements": [
                    {
                        "tag": "div",
                        "text": {
                            "tag": "lark_md",
                            "content": f"**时间**: {timestamp}"
                        }
                    },
                    {
                        "tag": "div",
                        "text": {
                            "tag": "lark_md",
                            "content": summary  # 飞书支持markdown，星号会自动转为加粗
                        }
                    }
                ]
            }
        }
        
        # 发送请求
        response = requests.post(
            webhook_url,
            headers={"Content-Type": "application/json"},
            data=json.dumps(post_data)
        )
        
        if response.status_code != 200:
            raise Exception(f"Webhook请求失败: {response.status_code}, {response.text}")
    
    def update_status(self, message):
        """更新状态栏信息"""
        time_str = datetime.datetime.now().strftime("%H:%M:%S")
        status_message = f"[{time_str}] {message}"
        # 使用QMetaObject.invokeMethod确保在主线程中更新UI
        self.status_text.append(status_message)
        # 滚动到底部
        self.status_text.ensureCursorVisible()
        # 强制更新UI
        QApplication.processEvents()
    
    def save_config(self):
        """保存配置到文件"""
        # 收集选中的群聊
        selected_chats = []
        for i in range(self.chat_list.count()):
            item = self.chat_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_chats.append(item.text())
        
        # 收集其他配置
        config = {
            "monitor_time": self.time_spin.value(),
            "interval_time": self.interval_spin.value(),
            "api_key": self.api_key_input.text(),
            "webhook_url": self.webhook_input.text(),
            "webhook_enabled": self.webhook_enabled.isChecked(),
            "selected_chats": selected_chats,
            "ai_prompt": self.ai_prompt  # 保存AI提示模板
        }
        
        # 保存到文件
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.update_status("配置已保存")
        except Exception as e:
            self.update_status(f"保存配置失败: {str(e)}")
    
    def load_config(self):
        """从文件加载配置"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 应用配置
            self.time_spin.setValue(config.get("monitor_time", 60))
            self.interval_spin.setValue(config.get("interval_time", 10))
            self.api_key_input.setText(config.get("api_key", ""))
            self.webhook_input.setText(config.get("webhook_url", ""))
            self.webhook_enabled.setChecked(config.get("webhook_enabled", False))
            
            # 加载AI提示模板
            if "ai_prompt" in config:
                self.ai_prompt = config["ai_prompt"]
            
            # 重新加载群聊列表，然后应用选中状态
            self.refresh_chat_list()
            
            # 延迟应用选中状态，确保群聊列表已加载
            QTimer.singleShot(1000, lambda: self.apply_chat_selection(config.get("selected_chats", [])))
            
            self.update_status("配置已加载")
        except Exception as e:
            self.update_status(f"加载配置失败: {str(e)}")
    
    def apply_chat_selection(self, selected_chats):
        """应用群聊选择"""
        for i in range(self.chat_list.count()):
            item = self.chat_list.item(i)
            if item.text() in selected_chats:
                item.setCheckState(Qt.Checked)

    def manual_add_chat(self):
        """手动添加群聊名称"""
        chat_name, ok = QInputDialog.getText(self, "手动添加群聊", "请输入完整的群聊名称:")
        
        if ok and chat_name:
            # 检查是否已存在
            exists = False
            for i in range(self.chat_list.count()):
                if self.chat_list.item(i).text() == chat_name:
                    exists = True
                    # 选中已存在的项
                    self.chat_list.item(i).setCheckState(Qt.Checked)
                    break
            
            # 如果不存在，添加新项
            if not exists:
                item = QListWidgetItem(chat_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)  # 默认选中
                self.chat_list.addItem(item)
                
            self.update_status(f"已添加群聊: {chat_name}")

    def clear_chat_selection(self):
        """清空群聊选择"""
        for i in range(self.chat_list.count()):
            self.chat_list.item(i).setCheckState(Qt.Unchecked)
        
        self.update_status("已清空群聊选择")

    def set_ai_prompt(self):
        """设置AI提示模板"""
        current_prompt = self.ai_prompt
        new_prompt, ok = QInputDialog.getMultiLineText(
            self, "设置AI提示模板", 
            "请输入AI总结时使用的提示模板：\n(提示将指导AI如何总结对话内容)", 
            current_prompt
        )
        
        if ok and new_prompt:
            self.ai_prompt = new_prompt
            self.update_status("已更新AI提示模板")

    def reset_ai_prompt(self):
        """重置为默认AI提示模板"""
        default_prompt = "你是一个Web3撸毛的人，你非常擅长撸毛，你加入了一个群聊，你看过了所有人的聊天后，对他们聊的内容进行了重点分析，分析了哪些是项目相关的，哪些是要空投相关的，哪些是做任务的，并把看到的项目地址，需要做什么任务都分析出来，根据聊天内容的前后顺序，进行关联分析，要进行聊天的上下文关联，确保上下文关联的准确性，然后进行总结"
        self.ai_prompt = default_prompt
        self.update_status("已重置为默认AI提示模板")


class MonitorThread(QThread):
    message_signal = pyqtSignal(str, str, str, float)  # 群聊名称, 发送者, 内容, 时间戳
    status_signal = pyqtSignal(str)  # 状态信息
    complete_signal = pyqtSignal()  # 监控完成信号
    
    def __init__(self, monitor, chats, duration, check_interval=10):
        super().__init__()
        self.monitor = monitor
        self.chats = chats
        self.duration = duration
        self.running = True
        # 使用传入的检测间隔
        self.check_interval = check_interval  # 检测间隔，单位秒
        # 调试模式
        self.debug_mode = True
    
    def log(self, message):
        """输出调试日志"""
        if self.debug_mode:
            print(f"[监控线程] {message}")
            self.status_signal.emit(f"调试: {message}")
    
    def run(self):
        """线程主函数"""
        start_time = time.time()
        end_time = start_time + self.duration
        
        self.status_signal.emit(f"开始监控 {len(self.chats)} 个群聊，预计结束时间: {datetime.datetime.fromtimestamp(end_time).strftime('%H:%M:%S')}")
        self.log(f"监控线程启动，检测间隔: {self.check_interval}秒")
        
        last_check_time = {chat: 0 for chat in self.chats}
        chat_error_count = {chat: 0 for chat in self.chats}  # 记录每个群聊的错误次数
        max_error_count = 3  # 最大错误次数
        
        # 当前轮询的聊天索引
        current_chat_index = 0
        
        try:
            while self.running and time.time() < end_time:
                # 计算剩余时间
                remaining = int(end_time - time.time())
                if remaining % 60 == 0 and remaining > 0:  # 每分钟更新一次状态
                    minutes = remaining // 60
                    self.status_signal.emit(f"监控中，剩余时间: {minutes} 分钟")
                
                # 检查当前群聊索引
                if current_chat_index >= len(self.chats):
                    current_chat_index = 0  # 重置索引，开始新一轮检查
                    self.log("完成一轮群聊检查，开始新一轮")
                
                # 获取当前要检查的群聊
                if current_chat_index < len(self.chats):
                    chat_name = self.chats[current_chat_index]
                    
                    # 检查是否需要跳过此群聊
                    if chat_error_count[chat_name] >= max_error_count:
                        if chat_error_count[chat_name] == max_error_count:  # 只在第一次超过时通知
                            self.status_signal.emit(f"暂时跳过群聊 {chat_name}，连续错误次数过多")
                            chat_error_count[chat_name] += 1  # 增加计数但不再发送通知
                        current_chat_index += 1  # 移到下一个群聊
                        continue
                        
                    try:
                        # 切换到当前群聊
                        self.status_signal.emit(f"正在检查群聊: {chat_name}...")
                        self.monitor.switch_to_chat(chat_name)
                        
                        # 读取新消息
                        self.log(f"开始获取 {chat_name} 的新消息...")
                        messages = self.monitor.get_new_messages()  # 不再使用last_check过滤，依赖缓存机制
                        
                        if messages:
                            self.log(f"获取到 {len(messages)} 条新消息")
                            for msg in messages:
                                self.message_signal.emit(chat_name, msg["sender"], msg["content"], msg["timestamp"])
                            self.status_signal.emit(f"已读取 {chat_name} 的 {len(messages)} 条新消息")
                        else:
                            self.log(f"群聊 {chat_name} 没有新消息")
                            self.status_signal.emit(f"群聊 {chat_name} 没有新消息")
                        
                        # 更新最后检查时间
                        last_check_time[chat_name] = time.time()
                        
                        # 成功读取后重置错误计数
                        chat_error_count[chat_name] = 0
                    except Exception as e:
                        error_msg = str(e)
                        self.log(f"监控 {chat_name} 出错: {error_msg}")
                        # 增加错误计数
                        chat_error_count[chat_name] += 1
                        
                        # 根据错误次数显示不同级别的警告
                        if chat_error_count[chat_name] == 1:
                            self.status_signal.emit(f"监控群聊 {chat_name} 时出错: {error_msg}")
                        elif chat_error_count[chat_name] == 2:
                            self.status_signal.emit(f"再次尝试监控群聊 {chat_name} 失败: {error_msg}")
                        elif chat_error_count[chat_name] == max_error_count:
                            self.status_signal.emit(f"群聊 {chat_name} 多次访问失败，可能是名称不匹配或其他问题，将暂时跳过该群聊")
                    
                    # 移动到下一个群聊
                    current_chat_index += 1
                
                # 休眠指定的检测间隔时间
                self.log(f"休眠 {self.check_interval} 秒...")
                time.sleep(self.check_interval)
            
            # 监控完成
            if time.time() >= end_time:
                self.status_signal.emit("监控时间已到，正在停止监控...")
            else:
                self.status_signal.emit("监控已手动停止")
                
            self.complete_signal.emit()
                
        except Exception as e:
            error_msg = str(e)
            self.log(f"监控线程发生错误: {error_msg}")
            self.status_signal.emit(f"监控线程发生错误: {error_msg}")
            self.running = False
    
    def stop(self):
        """停止监控线程"""
        self.running = False
        self.log("正在停止监控线程...")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WeChatMonitorApp()
    window.show()
    sys.exit(app.exec_()) 