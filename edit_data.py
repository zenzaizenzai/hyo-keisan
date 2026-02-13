import json
import os

DATA_FILE = 'data.json'

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def append_japanese_text(row_idx, col_idx, text):
    """
    指定したセルに日本語を追加します。
    """
    data = load_data()
    
    # 必要に応じて行を拡張
    while len(data) <= row_idx:
        data.append([])
    
    # 必要に応じて列を拡張
    while len(data[row_idx]) <= col_idx:
        data[row_idx].append("")
        
    current_val = data[row_idx][col_idx]
    if current_val:
        data[row_idx][col_idx] = f"{current_val} {text}"
    else:
        data[row_idx][col_idx] = text
        
    save_data(data)
    print(f"Row {row_idx+1}, Col {col_idx+1} に '{text}' を追加しました。")

if __name__ == '__main__':
    # サンプル実行: A1セルに「こんにちは」を追加
    append_japanese_text(0, 0, "こんにちは（Pythonより）")
