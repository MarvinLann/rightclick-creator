#!/usr/bin/env python3
"""
MD整理工具 - 使用通义千问API修复PDF转Markdown后的混乱文档
"""

import sys
import os
import json
import argparse
import subprocess
import time
from pathlib import Path
from typing import Optional
import http.client
import urllib.parse


def show_notification(title: str, message: str):
    """显示macOS通知（非阻塞）"""
    try:
        subprocess.run([
            "osascript", "-e",
            f'display notification "{message}" with title "{title}"'
        ], check=False)
    except:
        pass


class MDCleaner:
    """Markdown文档清理器"""
    
    # 默认模型分级配置
    DEFAULT_TIERS = {
        "small": {"model": "qwen-turbo", "max_tokens": 6000, "input_price": 0.3, "output_price": 9.6},
        "medium": {"model": "qwen-plus", "max_tokens": 30000, "input_price": 0.8, "output_price": 2.0},
        "large": {"model": "qwen3.5-plus", "max_tokens": 60000, "input_price": 0.8, "output_price": 4.8}
    }
    
    def __init__(self):
        self.config = self._load_config()
        self.api_key = self.config.get('api_key', '')
        self.model = self.config.get('model', 'auto')
        self.base_url = self.config.get('base_url', 'dashscope.aliyuncs.com')
        self.model_tiers = self.config.get('model_tiers', self.DEFAULT_TIERS)
    
    def _select_model(self, token_count: int) -> tuple:
        """
        根据token数量自动选择最合适的模型
        返回: (模型名称, 最大输出tokens, 分级名称)
        """
        # 按max_tokens从小到大排序
        tiers = sorted(self.model_tiers.items(), key=lambda x: x[1]['max_tokens'])
        
        for tier_name, tier_config in tiers:
            if token_count <= tier_config['max_tokens']:
                return tier_config['model'], tier_config['max_tokens'], tier_name
        
        # 如果超过所有模型限制，使用最大的模型并分段
        largest = tiers[-1][1]
        return largest['model'], largest['max_tokens'], tiers[-1][0]
    
    def _load_config(self) -> dict:
        """加载配置文件，如果不存在则创建空文件"""
        config_paths = [
            Path.home() / '.rightclick-creator' / 'config' / 'md_cleaner.json',
            Path.home() / '.config' / 'md_cleaner' / 'config.json',
        ]
        
        for config_path in config_paths:
            if config_path.exists():
                try:
                    with open(config_path, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except Exception as e:
                    print(f"警告：无法读取配置文件 {config_path}: {e}")
        
        # 创建默认配置文件
        default_config = {"api_key": "", "model": "auto"}
        default_path = config_paths[0]
        default_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(default_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        
        return default_config
    
    def _check_api_key(self) -> bool:
        """检查 API Key 是否已配置，未配置则静默退出"""
        if not self.api_key:
            print("API Key 未配置，跳过处理")
            print("请通过 WorkBuddy / Claude 配置 API Key 后重试")
            return False
        return True
    
    def _estimate_tokens(self, text: str) -> int:
        """估算文本的token数量（中文约1.5-2 tokens/字）"""
        # 简单估算：中文字符按2 tokens，英文按0.5 tokens
        chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
        other_chars = len(text) - chinese_chars
        return int(chinese_chars * 2 + other_chars * 0.5)
    
    def _split_content(self, content: str, max_tokens: int = 6000) -> list:
        """
        智能分段处理
        按段落分割，保持上下文完整性
        """
        total_tokens = self._estimate_tokens(content)
        
        # 如果总token数在安全范围内，直接返回
        if total_tokens <= max_tokens:
            return [content]
        
        print(f"  文档较大（约{total_tokens} tokens），将分段处理...")
        
        # 按段落分割
        paragraphs = content.split('\n\n')
        segments = []
        current_segment = []
        current_tokens = 0
        
        for para in paragraphs:
            para_tokens = self._estimate_tokens(para)
            
            # 如果当前段落单独就超过限制，需要进一步分割
            if para_tokens > max_tokens:
                # 先保存当前段
                if current_segment:
                    segments.append('\n\n'.join(current_segment))
                    current_segment = []
                    current_tokens = 0
                
                # 按句子分割大段落
                sentences = para.replace('。', '。\n').replace('！', '！\n').replace('？', '？\n').split('\n')
                temp_segment = []
                temp_tokens = 0
                
                for sent in sentences:
                    sent_tokens = self._estimate_tokens(sent)
                    if temp_tokens + sent_tokens > max_tokens and temp_segment:
                        segments.append('\n'.join(temp_segment))
                        temp_segment = [sent]
                        temp_tokens = sent_tokens
                    else:
                        temp_segment.append(sent)
                        temp_tokens += sent_tokens
                
                if temp_segment:
                    segments.append('\n'.join(temp_segment))
            else:
                # 正常段落，检查是否可以加入当前段
                if current_tokens + para_tokens > max_tokens and current_segment:
                    segments.append('\n\n'.join(current_segment))
                    current_segment = [para]
                    current_tokens = para_tokens
                else:
                    current_segment.append(para)
                    current_tokens += para_tokens
        
        # 保存最后一段
        if current_segment:
            segments.append('\n\n'.join(current_segment))
        
        print(f"  已分割为 {len(segments)} 个段落")
        return segments
    
    def _call_qwen_api(self, content: str, model: str = None, max_output_tokens: int = 16384) -> Optional[str]:
        """调用通义千问API"""
        if not self._check_api_key():
            return None
        
        # 如果没有指定模型，使用配置的模型
        if model is None:
            model = self.model
        
        prompt = f"""你是一个专业的Markdown文档整理专家。请修复以下从PDF转换而来的Markdown文档中的问题：

1. 修复断行和错行问题，恢复自然段落
2. 识别并修正乱码和乱码占位符
3. 修正错别字和标点符号错误
4. 规范化Markdown格式（标题层级、列表、代码块等）
5. 保持原文内容完整，不做内容增删

请直接输出修复后的完整Markdown文档，不要添加任何解释。

---

待修复文档：

{content}"""
        
        try:
            payload = json.dumps({
                "model": model,
                "messages": [
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.3,
                "max_tokens": max_output_tokens
            })
            
            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }
            
            with http.client.HTTPSConnection(self.base_url, timeout=360) as conn:
                conn.request("POST", "/compatible-mode/v1/chat/completions", payload, headers)
                response = conn.getresponse()
                
                if response.status == 429:
                    print("API调用频率受限，等待5秒后重试...")
                    time.sleep(5)
                    conn.request("POST", "/compatible-mode/v1/chat/completions", payload, headers)
                    response = conn.getresponse()
                
                if response.status != 200:
                    error_body = response.read().decode('utf-8')
                    print(f"API调用失败: HTTP {response.status}")
                    print(f"错误详情: {error_body}")
                    return None
                
                data = json.loads(response.read().decode('utf-8'))
            
            choices = data.get('choices')
            if choices and len(choices) > 0:
                return choices[0].get('message', {}).get('content')
            else:
                print(f"API响应格式异常: {data}")
                return None
                
        except Exception as e:
            print(f"API调用出错: {e}")
            return None
    
    def clean_file(self, file_path: str) -> bool:
        """清理单个Markdown文件"""
        path = Path(file_path)
        
        if not path.exists():
            print(f"错误：文件不存在 {file_path}")
            return False
        
        if path.suffix.lower() not in ['.md', '.markdown']:
            print(f"跳过非Markdown文件: {file_path}")
            return False
        
        try:
            # 读取原文件
            with open(path, 'r', encoding='utf-8') as f:
                original_content = f.read()
            
            if not original_content.strip():
                print(f"跳过空文件: {file_path}")
                return False
            
            # 估算token数量
            total_tokens = self._estimate_tokens(original_content)
            print(f"  文档大小: 约{total_tokens} tokens")
            
            # 自动选择模型
            selected_model, max_output, tier_name = self._select_model(total_tokens)
            print(f"  自动选择模型: {selected_model} ({tier_name}级)")
            
            # 检查是否需要分段（超过最大模型限制）
            if total_tokens > self.model_tiers['large']['max_tokens']:
                print(f"  文档超过最大模型限制，将使用{self.model_tiers['large']['model']}分段处理")
                segments = self._split_content(original_content, max_tokens=self.model_tiers['large']['max_tokens'])
                total_segments = len(segments)
                
                print(f"✅ 已与通义千问连接成功！正在处理: {path.name}")
                print(f"  文档超大，将使用{self.model_tiers['large']['model']}分{total_segments}段处理")
                print(f"  ⏱️ 预计需要 {total_segments * 2}-{total_segments * 3} 分钟，在此期间请勿操作该文档")
                
                # 分段处理
                cleaned_segments = []
                for i, segment in enumerate(segments, 1):
                    print(f"  处理第 {i}/{total_segments} 段...")
                    
                    cleaned_segment = self._call_qwen_api(
                        segment, 
                        model=self.model_tiers['large']['model'],
                        max_output_tokens=self.model_tiers['large']['max_tokens']
                    )
                    
                    if cleaned_segment is None:
                        print(f"第 {i} 段处理失败")
                        print(f"❌ 第 {i}/{total_segments} 段处理失败: {path.name}")
                        return False
                    
                    cleaned_segments.append(cleaned_segment)
                    
                    if i < total_segments:
                        show_notification("MD整理", f"进度: {i}/{total_segments} 段完成，继续处理...")
                
                cleaned_content = '\n\n'.join(cleaned_segments)
            else:
                # 一次性处理
                print(f"✅ 已与通义千问连接成功！正在处理: {path.name}")
                print(f"  使用模型: {selected_model}，预计需要 1-3 分钟，在此期间请勿操作该文档")
                
                cleaned_content = self._call_qwen_api(
                    original_content,
                    model=selected_model,
                    max_output_tokens=max_output
                )
                
                if cleaned_content is None:
                    print(f"处理失败: {path.name}")
                    print(f"❌ 处理失败: {path.name}，请检查网络连接或API配置")
                    return False
            
            # 原子写入：先写临时文件，成功后再替换
            tmp_path = path.with_suffix('.md.tmp')
            try:
                with open(tmp_path, 'w', encoding='utf-8') as f:
                    f.write(cleaned_content)
                os.replace(str(tmp_path), str(path))
            except Exception:
                if tmp_path.exists():
                    tmp_path.unlink()
                raise
            
            print(f"✓ 完成: {path.name}")
            
            # 显示完成通知
            show_notification("MD整理", f"✅ {path.name} 处理完成！")
            return True
            
        except Exception as e:
            print(f"处理文件时出错 {file_path}: {e}")
            print(f"❌ 处理出错 {path.name}: {e}")
            return False
    
    def clean_files(self, file_paths: list) -> dict:
        """批量清理多个文件"""
        results = {'success': [], 'failed': []}
        
        for file_path in file_paths:
            if self.clean_file(file_path):
                results['success'].append(file_path)
            else:
                results['failed'].append(file_path)
        
        return results


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='MD整理工具 - 使用AI修复Markdown文档')
    parser.add_argument('files', nargs='+', help='要处理的Markdown文件路径')
    parser.add_argument('--config', '-c', help='配置文件路径')
    
    args = parser.parse_args()
    
    # 限制处理文件数量
    if len(args.files) > 5:
        print("警告：一次最多处理5个文件，将只处理前5个")
        args.files = args.files[:5]
    
    cleaner = MDCleaner()
    results = cleaner.clean_files(args.files)
    
    # 输出统计
    print(f"\n处理完成:")
    print(f"  成功: {len(results['success'])} 个文件")
    print(f"  失败: {len(results['failed'])} 个文件")
    
    if results['failed']:
        print(f"\n失败的文件:")
        for f in results['failed']:
            print(f"  - {f}")
    
    sys.exit(0 if len(results['failed']) == 0 else 1)


if __name__ == '__main__':
    main()
