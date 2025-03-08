import re

class ContentFilter:
    @staticmethod
    def filter_ai_symbols(text):
        """
        过滤AI回复中的特殊符号
        Args:
            text: 需要过滤的文本
        Returns:
            过滤后的文本
        """
        # 使用正则表达式移除*和-及其后续空格
        text = re.sub(r'\*{2,}', '', text)  # 删除连续双星号
        print(f'[预处理] 原始内容:\n{text}\n{"-"*40}')
        # 优化符号行过滤仅移除含*/符号的行
        result = re.sub(r'^\s*[\*/－]+\s*$', '', text, flags=re.MULTILINE)
        print(f'[处理后] 最终内容:\n{result}\n{"="*40}')
        return result