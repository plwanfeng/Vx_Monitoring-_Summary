import time
import datetime
import re
from wxauto import WeChat

class WeChatMonitor:
    def __init__(self):
        """初始化微信监控器"""
        # 检查微信是否已启动
        self.wx = WeChat()
        
        # 使用微信实例的属性来检查是否登录
        try:
            # 尝试获取会话列表，如果失败则可能未登录
            if not self.wx.GetSessionList():
                raise Exception("未检测到会话列表，请确保微信已登录")
        except Exception as e:
            raise Exception(f"微信未登录或未正确启动: {str(e)}")
        
        # 初始化缓存
        self.message_cache_by_chat = {}  # 用于缓存消息，避免重复
    
    def get_chat_list(self):
        """获取可用的群聊列表"""
        # 获取所有会话
        sessions = self.wx.GetSessionList()
        
        # 过滤出群聊（通常群聊名称至少包含几个字符）
        # 这里简单地过滤掉可能的系统会话或空名称
        chat_groups = [session for session in sessions if len(session) > 1]
        
        return chat_groups
    
    def switch_to_chat(self, chat_name):
        """切换到指定的群聊，使用更安全的方法处理特殊字符"""
        # 尝试使用不同的方法切换到群聊
        methods = [
            self._switch_direct,
            self._switch_by_search,
            self._switch_by_click_session
        ]
        
        # 记录所有错误，以便在所有方法都失败时提供详细信息
        errors = []
        
        for method_index, method in enumerate(methods):
            try:
                if method(chat_name):
                    # 成功切换后，确保滚动到最新消息
                    self._scroll_to_latest_messages()
                    return True
            except Exception as e:
                error_msg = f"方法{method_index+1}失败: {str(e)}"
                print(error_msg)
                errors.append(error_msg)
        
        # 所有方法都失败
        raise Exception(f"切换到群聊 {chat_name} 失败: 尝试了所有可用方法。错误: {'; '.join(errors)}")
    
    def _scroll_to_latest_messages(self):
        """滚动到最新消息"""
        try:
            # 尝试点击消息区域并按End键滚动到底部
            chat_box = self.wx.ChatBox
            if chat_box:
                chat_box.SetFocus()
                # 先按Home键到顶部，然后按End键到底部，确保位置正确
                chat_box.SendKeys("{END}")
                time.sleep(0.5)
                print("已滚动到最新消息")
            else:
                print("找不到聊天框，无法滚动到最新消息")
        except Exception as e:
            print(f"滚动到最新消息时出错: {str(e)}")
    
    def _switch_direct(self, chat_name):
        """方法1: 直接使用ChatWith尝试切换"""
        try:
            if self.wx.ChatWith(chat_name):
                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            raise Exception(f"直接切换失败: {str(e)}")
    
    def _switch_by_search(self, chat_name):
        """方法2: 使用搜索框搜索并点击结果"""
        try:
            # 首先切换到聊天列表
            self.wx.SwitchToChat()
            time.sleep(0.5)
            
            # 获取并清空搜索框
            search_box = self.wx.B_Search
            if not search_box:
                raise Exception("未找到搜索框")
            
            search_box.SetFocus()
            # 使用安全的方式清空搜索框
            search_box.SendKeys('{CONTROL}a{BACKSPACE}')
            time.sleep(0.2)
            
            # 安全地输入群聊名称
            # 逐字符输入，避免一次性输入导致的特殊字符问题
            for char in chat_name:
                if char in ['+', '^', '%', '(', ')', '{', '}', '[', ']', '~']:
                    # 对特殊字符进行转义
                    search_box.SendKeys('{' + char + '}')
                else:
                    search_box.SendKeys(char)
                time.sleep(0.05)  # 短暂延迟，模拟真实输入
            
            time.sleep(1)  # 等待搜索结果
            
            # 尝试点击第一个搜索结果
            session_list = self.wx.SessionItemList
            if session_list and len(session_list) > 0:
                session_list[0].Click()
                time.sleep(0.5)
                return True
            
            return False
        except Exception as e:
            raise Exception(f"搜索切换失败: {str(e)}")
    
    def _switch_by_click_session(self, chat_name):
        """方法3: 遍历会话列表并点击匹配项"""
        try:
            # 首先切换到聊天列表
            self.wx.SwitchToChat()
            time.sleep(0.5)
            
            # 获取会话列表
            session_list = self.wx.GetSessionList()
            
            # 在会话列表中找到匹配的会话
            target_index = -1
            for i, session in enumerate(session_list):
                if session == chat_name:
                    target_index = i
                    break
            
            if target_index >= 0:
                # 尝试点击会话项
                session_items = self.wx.SessionItemList
                if len(session_items) > target_index:
                    session_items[target_index].Click()
                    time.sleep(0.5)
                    return True
            
            return False
        except Exception as e:
            raise Exception(f"会话列表点击失败: {str(e)}")
    
    def get_new_messages(self, max_messages=20):
        """只获取最新消息，而不尝试提取历史消息
        
        Args:
            max_messages: 最大获取消息数量
            
        Returns:
            list: 包含新消息的列表，每条消息为一个字典，包含发送者、内容和时间戳
        """
        # 获取当前聊天窗口名称
        chat_name = self.wx.CurrentChat
        if not chat_name:
            print("无法获取当前聊天窗口名称")
            return []
        
        print(f"监控 {chat_name} 的新消息...")
        
        # 确保滚动到最新消息
        self._scroll_to_latest_messages()
        
        # 获取新消息
        try:
            messages_raw = self.wx.GetAllMessage()
            
            if not messages_raw:
                print("未获取到任何消息")
                return []
                
            # 只获取最新的几条消息
            latest_messages = messages_raw[-max_messages:] if len(messages_raw) > max_messages else messages_raw
            print(f"获取到 {len(latest_messages)} 条最新消息")
        except Exception as e:
            print(f"获取消息失败: {str(e)}")
            return []
        
        # 初始化聊天的缓存
        if chat_name not in self.message_cache_by_chat:
            self.message_cache_by_chat[chat_name] = set()
            
        # 初始化返回的消息列表
        new_messages = []
        
        # 记录现在的时间戳，用于标记所有该轮次获取的新消息
        current_time = time.time()
        
        # 处理消息
        for msg_item in latest_messages:
            try:
                # 提取发送者和内容
                sender, content = self._parse_message(msg_item)
                
                # 如果解析失败，跳过
                if not sender or not content:
                    continue
                
                # 生成消息指纹
                import hashlib
                msg_fingerprint = hashlib.md5(f"{sender}:{content}".encode()).hexdigest()
                
                # 检查是否是新消息
                if msg_fingerprint in self.message_cache_by_chat[chat_name]:
                    continue
                
                # 记录新消息
                new_messages.append({
                    "sender": sender,
                    "content": content,
                    "timestamp": current_time
                })
                
                # 更新缓存
                self.message_cache_by_chat[chat_name].add(msg_fingerprint)
                print(f"新消息: {sender} -> {content[:30]}{'...' if len(content) > 30 else ''}")
            
            except Exception as e:
                print(f"处理消息时出错: {str(e)}")
                continue
        
        # 限制缓存大小，只保留最近的消息指纹
        MAX_CACHE_SIZE = 200
        if len(self.message_cache_by_chat[chat_name]) > MAX_CACHE_SIZE:
            print(f"清理 {chat_name} 的消息缓存")
            # 转换为列表后随机保留一部分
            cache_list = list(self.message_cache_by_chat[chat_name])
            # 保留后半部分（较新的消息）
            self.message_cache_by_chat[chat_name] = set(cache_list[-int(MAX_CACHE_SIZE/2):])
        
        # 返回新消息
        return new_messages
    
    def send_message(self, chat_name, message):
        """向指定群聊发送消息"""
        # 切换到指定群聊
        self.switch_to_chat(chat_name)
        
        # 发送消息
        self.wx.SendMsg(message)
        
        return True 

    def _parse_message(self, msg_item):
        """解析消息，提取发送者和内容
        
        Args:
            msg_item: 原始消息项
            
        Returns:
            tuple: (发送者, 内容) 元组
        """
        # 打印原始消息用于调试
        # print(f"解析消息: {msg_item}")
        
        # 初始化发送者和内容
        sender = None
        content = None
        
        # 消息可能是字符串或其他类型对象，统一处理
        msg_str = str(msg_item)
        
        # 方法1: 常规的冒号分割
        if ":" in msg_str:
            parts = msg_str.split(":", 1)
            if len(parts) == 2:
                sender, content = parts
        # 方法2: 中文冒号分割
        elif "：" in msg_str:
            parts = msg_str.split("：", 1)
            if len(parts) == 2:
                sender, content = parts
        # 方法3: 处理特殊格式的消息
        else:
            # 尝试使用直接属性访问（如果是对象）
            if hasattr(msg_item, 'sender') and hasattr(msg_item, 'content'):
                sender = str(msg_item.sender)
                content = str(msg_item.content)
            # 假设是系统消息
            elif "[" in msg_str or "]" in msg_str:
                sender = "系统消息"
                content = msg_str
            else:
                print(f"无法解析消息格式: {msg_str}")
                return None, None
        
        # 清理字符串
        if sender:
            sender = sender.strip()
        if content:
            content = content.strip()
        
        # 如果发送者或内容为空，跳过
        if not sender or not content:
            print(f"发送者或内容为空: 发送者='{sender}', 内容='{content}'")
            return None, None
            
        return sender, content 