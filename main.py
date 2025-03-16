import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser
import requests
import pandas as pd
import json
import os


class DeepSeekExcelGenerator:
    def __init__(self, master):
        self.master = master
        master.title("DeepSeek Excel 生成器 v1.0")

        # 加载配置
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        self.api_settings = self.config['API']

        # 创建界面
        self.create_widgets()
        self.setup_layout()

    def create_widgets(self):
        self.input_label = ttk.Label(self.master, text="请输入需求描述:")
        self.input_text = tk.Text(self.master, height=10, width=60, wrap=tk.WORD)
        self.generate_btn = ttk.Button(self.master, text="生成Excel", command=self.generate_excel)
        self.status_bar = ttk.Label(self.master, relief=tk.SUNKEN, anchor=tk.W)

    def setup_layout(self):
        # 布局管理
        self.input_label.pack(pady=5, padx=10, anchor=tk.W)
        self.input_text.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        self.generate_btn.pack(pady=10)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def generate_excel(self):
        user_input = self.input_text.get("1.0", tk.END).strip()
        if not user_input:
            messagebox.showwarning("输入错误", "请输入需求描述")
            return

        try:
            self.set_status("正在与DeepSeek API通信...")

            # 调用API
            response = self.call_deepseek_api(user_input)
            table_data = self.parse_response(response)

            # 生成DataFrame
            df = pd.DataFrame(table_data[1:], columns=table_data[0])

            # 保存文件
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel 文件", "*.xlsx"), ("All Files", "*.*")]
            )
            if file_path:
                df.to_excel(file_path, index=False)
                self.set_status(f"文件已保存至: {file_path}")
                messagebox.showinfo("成功", "Excel文件生成成功！")

        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.set_status("就绪")

    def call_deepseek_api(self, prompt):
        headers = {
            "Authorization": f"Bearer {self.api_settings['api_key']}",
            "Content-Type": "application/json"
        }

        payload = {
            "model": self.api_settings['model'],
            "messages": [
                {
                    "role": "user",
                    "content": f"{prompt}\n请用规范的表格格式返回数据，第一行为表头，后续为数据行"
                }
            ],
            "temperature": float(self.api_settings['temperature']),
            "max_tokens": int(self.api_settings['max_tokens'])
        }

        response = requests.post(
            f"{self.api_settings['api_base']}/chat/completions",
            headers=headers,
            json=payload
        )

        if response.status_code != 200:
            raise Exception(f"API请求失败: {response.text}")

        return json.loads(response.text)

    def parse_response(self, response_data):
        try:
            content = response_data['choices'][0]['message']['content']
            lines = [line.strip() for line in content.split('\n') if line.strip()]

            # 解析表格数据
            table_data = []
            for line in lines:
                if line.startswith('|'):
                    cells = [cell.strip() for cell in line.split('|') if cell.strip()]
                    if len(cells) < 2:
                        continue
                    table_data.append(cells)

            if len(table_data) < 2:
                raise ValueError("API返回的表格数据格式不正确")

            return table_data

        except Exception as e:
            raise ValueError("解析API响应失败: " + str(e))

    def set_status(self, message):
        self.status_bar.config(text=message)
        self.master.update()


if __name__ == "__main__":
    root = tk.Tk()
    app = DeepSeekExcelGenerator(root)
    root.geometry("600x400")
    root.mainloop()