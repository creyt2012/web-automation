import openai
import pandas as pd
import os
import requests
from tqdm import tqdm
import random

def generate_text(prompt, api_key, creativity=0.5):
    endpoint = "https://api.openai.com/v1/engines/text-davinci-003/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "prompt": prompt,
        "max_tokens": 500,
        "temperature": creativity
    }

    response = requests.post(endpoint, headers=headers, json=data)
    if response.status_code != 200:
        print(f"Error: {response.status_code}, {response.text}")
        return None

    result = response.json()
    return result["choices"][0]["text"]

def main():
    # input data
    apikey_file = input("Nhập đường dẫn đến file txt chứa danh sách API keys: ")
    input_file = input("Nhập đường dẫn đến file input XLSX: ")
    output_file = input("Nhập đường dẫn đến file output: ")
    creativity = float(input("Nhập mức độ sáng tạo của API (từ 0.0 đến 0.9): "))
    
    # input keyword + pomrt
    character_before_heading = input("Nhập ký tự trước 'heading' (ví dụ: giới thiệu + khóa học java): ").strip()

    # Read list API keys từ file txt
    with open(apikey_file, 'r') as f:
        api_keys = f.read().splitlines()

    # Read data từ file input xlsx
    data = pd.read_excel(input_file)

    # open file output để write data
    output_file_exists = os.path.isfile(output_file)

    with open(output_file, "w", encoding="utf-8") as file_output:
        for heading in tqdm(data["heading"], desc="Xử lý tác vụ", dynamic_ncols=True):
            # Băm ngẫu nhiên apikey theo list
            api_key = random.choice(api_keys)

            if character_before_heading:
                prompt = f"{character_before_heading} {heading}"
            else:
                prompt = heading

            # Checking và tạo một danh sách để theo dõi heading đã được xử lý
            processed_prompts = set()

            # Checking và tạo một danh sách để theo dõi heading đã xuất hiện trong file output
            existing_prompts = set()

            # Checking nếu file output đã tồn tại, Read data từ file output để kiểm tra trùng lặp
            if output_file_exists:
                with open(output_file, "r", encoding="utf-8") as existing_output_file:
                    for line in existing_output_file:
                        parts = line.split("\t")
                        if len(parts) > 0:
                            existing_prompts.add(parts[0])

            # Checking nếu câu hỏi đã được xử lý trước đó hoặc đã xuất hiện trong file output, bỏ qua
            if prompt in processed_prompts or heading in existing_prompts:
                continue

            # Sử dụng GPT-3 để tạo nội dung
            content = generate_text(prompt, api_key, creativity)

            if content is not None:
                # Ghi dữ liệu vào file output
                file_output.write(f"{heading}\t{content}\n")

                # Thêm câu hỏi vào danh sách đã xử lý
                processed_prompts.add(prompt)

if __name__ == "__main__":
    main()
