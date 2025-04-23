import json
import time
import requests

class DeepSeekSummarizer:
    def __init__(self, api_key):
        """初始化DeepSeek API总结器
        
        Args:
            api_key: DeepSeek API密钥
        """
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"  # 假设这是DeepSeek的API地址
        self.max_retries = 3  # 最大重试次数
        self.retry_delay = 2  # 重试延迟（秒）
        
    def summarize(self, messages_text, custom_prompt=None):
        """使用DeepSeek API对聊天记录进行总结
        
        Args:
            messages_text: 消息文本，每行一条消息
            custom_prompt: 自定义AI提示模板，如果为None则使用默认提示
            
        Returns:
            str: 总结的文本
        """
        # 如果消息为空，返回提示
        if not messages_text or messages_text.strip() == "":
            return "没有可用的聊天记录进行总结。"
        
        # 使用默认提示或自定义提示
        if custom_prompt:
            prompt = f"""
{custom_prompt}

聊天记录：
{messages_text}

请给出3000字以内的总结。
"""
        else:
            prompt = f"""
以下是微信群聊的聊天记录，请对这些聊天内容进行总结。
总结要点：
1. 主要讨论了哪些话题，讨论了哪些项目
2. 和项目相关的聊天内容，要进行上下文关联，确保上下文关联的准确性
3. 每个项目讨论了哪些内容，每个内容讨论了哪些方面，每个方面讨论了哪些细节
4. 有哪些值得注意的信息，有哪些项目值得关注，有哪些项目值得投资，有哪些项目值得参与

聊天记录：
{messages_text}

请给出3000字以内的总结。
"""
        
        # 使用DeepSeek API，带重试机制
        for attempt in range(self.max_retries):
            try:
                return self._call_api(prompt)
            except Exception as e:
                error_message = str(e)
                print(f"API调用失败 (尝试 {attempt+1}/{self.max_retries}): {error_message}")
                
                if attempt < self.max_retries - 1:
                    # 如果不是最后一次尝试，则等待后重试
                    print(f"等待 {self.retry_delay} 秒后重试...")
                    time.sleep(self.retry_delay)
                    # 增加重试间隔，避免频繁请求
                    self.retry_delay *= 1.5
                else:
                    # 所有重试都失败，返回错误信息
                    return f"总结生成失败: {error_message}\n\n尝试了 {self.max_retries} 次调用API但均未成功。"
    
    def _call_api(self, prompt):
        """执行实际的API调用
        
        Args:
            prompt: 发送给API的提示文本
            
        Returns:
            str: API返回的总结文本
            
        Raises:
            Exception: 如果API调用失败
        """
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        payload = {
            "model": "deepseek-chat",  # 使用适当的模型
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.3,  # 较低的温度以获得更确定性的输出
            "max_tokens": 1000
        }
        
        response = requests.post(
            self.api_url,
            headers=headers,
            data=json.dumps(payload),
            timeout=30  # 添加超时设置
        )
        
        # 检查请求是否成功
        if response.status_code == 200:
            try:
                data = response.json()
                if "choices" in data and len(data["choices"]) > 0:
                    return data["choices"][0]["message"]["content"].strip()
                else:
                    raise Exception("API返回的数据格式不正确")
            except json.JSONDecodeError:
                raise Exception("无法解析API返回的JSON数据")
        elif response.status_code == 401:
            raise Exception("API密钥无效或未授权")
        elif response.status_code == 429:
            raise Exception("API请求频率过高，请稍后再试")
        elif response.status_code >= 2000:
            raise Exception(f"DeepSeek服务器错误: {response.status_code}")
        else:
            raise Exception(f"API请求失败: HTTP {response.status_code}, {response.text}")
            
    def is_api_key_valid(self):
        """测试API密钥是否有效
        
        Returns:
            bool: API密钥是否有效
        """
        try:
            # 发送一个简单的请求来验证API密钥
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {
                        "role": "user",
                        "content": "测试"
                    }
                ],
                "max_tokens": 5
            }
            
            response = requests.post(
                self.api_url,
                headers=headers,
                data=json.dumps(payload),
                timeout=10  # 添加超时设置
            )
            
            return response.status_code == 200
            
        except Exception:
            return False 