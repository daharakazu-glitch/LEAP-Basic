import pandas as pd
import json
import os
import re

excel_file = "見出語・用例リスト_改訂版.xlsx"
xls = pd.ExcelFile(excel_file)

with open("leap_revision/leap_app.html", "r", encoding="utf-8") as f:
    template_html = f.read()

def clean_example_text(text):
    if pd.isna(text):
        return ""
    val = str(text).strip()
    if val in ["―", "-", "－"]:
        return ""
    # Remove any line that is just a dash or empty
    lines = [line.strip() for line in val.split('\n') if line.strip() and line.strip() not in ["―", "-", "－"]]
    return "\n".join(lines)

def clean_english_example(text):
    text = clean_example_text(text)
    if not text:
        return ""
    # Remove incorrect examples in brackets if any
    text = re.sub(r'[（(［\[【]\s*[✖×].*?[）)］\]】]', '', text)
    return text.strip()

def clean_japanese_example(text):
    return clean_example_text(text)

sheets_to_process = ['Part 1', 'Part 2', 'Part 3', 'Part 4', '＋α']

output_dir = "leap_revision/generated_apps"
os.makedirs(output_dir, exist_ok=True)

apps_info = []

for sheet in sheets_to_process:
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Check/assign Week column
    if sheet == '＋α':
        df['Week'] = [f"Week {i // 50 + 1}" for i in range(len(df))]
    else:
        week_col = [c for c in df.columns if 'Week' in c]
        if not week_col:
            continue
        df['Week'] = df[week_col[0]].ffill()
        
    weeks = df['Week'].unique()
    for week in weeks:
        if pd.isna(week):
            continue
        week_df = df[df['Week'] == week]
        chapter_data = []
        for idx, row in week_df.iterrows():
            word = str(row['単語']) if pd.notna(row['単語']) else ""
            explanation = str(row['語の意味']) if pd.notna(row['語の意味']) else ""
            midashi = str(int(row['No.'])) if pd.notna(row['No.']) else str(idx)
            
            en_text = row['用例（英語）'] if '用例（英語）' in row else ""
            ja_text = row['用例（日本語）'] if '用例（日本語）' in row else ""
            
            en_clean = clean_english_example(en_text)
            ja_clean = clean_japanese_example(ja_text)
            en_lines = [line.strip() for line in en_clean.split('\n') if line.strip()]
            ja_lines = [line.strip() for line in ja_clean.split('\n') if line.strip()]
            
            if len(en_lines) > 1 and len(en_lines) == len(ja_lines):
                for i, (en_l, ja_l) in enumerate(zip(en_lines, ja_lines)):
                    en_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', en_l)
                    ja_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', ja_l)
                    en_l = re.sub(r'[あ-んア-ン一-龥]+', '', en_l).strip()
                    chapter_data.append({"id": f"{midashi}-{i+1}", "answer": word, "explanation": explanation, "en": en_l, "ja": ja_l, "target_word": word})
            else:
                en_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', en_clean)
                ja_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', ja_clean)
                en_l = re.sub(r'[あ-んア-ン一-龥]+', '', en_l).strip()
                chapter_data.append({"id": str(midashi), "answer": word, "explanation": explanation, "en": en_l, "ja": ja_l, "target_word": word})
                
        if not chapter_data:
            continue
            
        week_num = re.sub(r'[^0-9]', '', str(week))
        
        if 'Part' in sheet:
            part_num = re.sub(r'[^0-9]', '', sheet)
            file_name = f"part{part_num}_week{week_num}.html"
            title = f"LEAP改訂版学習アプリ (P{part_num} W{week_num})"
            subtitle = f"Part {part_num} Week {week_num}"
            part_display = f"Part {part_num}"
            week_display = f"Week {week_num}"
        else:
            file_name = f"{sheet}_week{week_num}.html"
            title = f"LEAP改訂版学習アプリ ({sheet} W{week_num})"
            subtitle = f"{sheet} Week {week_num}"
            part_display = sheet
            week_display = f"Week {week_num}"
            
        # Replace template headers & content
        html_content = template_html.replace("<title>LEAP改訂版学習アプリ (P1 W1)</title>", f"<title>{title}</title>")
        html_content = html_content.replace('<h2 id="app-subtitle" class="text-xl md:text-2xl font-bold text-white">Part 1 Week 1</h2>', f'<h2 id="app-subtitle" class="text-xl md:text-2xl font-bold text-white">{subtitle}</h2>')
        
        json_data = json.dumps(chapter_data, ensure_ascii=False, indent=4)
        html_content = html_content.replace("const chapterData = [];", f"const chapterData = {json_data};")
        
        with open(os.path.join(output_dir, file_name), "w", encoding="utf-8") as f:
            f.write(html_content)
 
        apps_info.append({
            "title": title,
            "subtitle": subtitle,
            "file_name": file_name,
            "part": part_display,
            "week": week_display,
            "word_count": len(chapter_data)
        })

# Sort apps_info logically: Part 1 -> Part 2 -> Part 3 -> Part 4 -> ＋α, and Week 1 -> Week 2...
def sort_key(app):
    part = app['part']
    week_str = app['week']
    week_num = int(re.sub(r'[^0-9]', '', week_str))
    
    if 'Part 1' in part:
        part_idx = 1
    elif 'Part 2' in part:
        part_idx = 2
    elif 'Part 3' in part:
        part_idx = 3
    elif 'Part 4' in part:
        part_idx = 4
    else:
        part_idx = 5 # ＋α
        
    return (part_idx, week_num)

apps_info.sort(key=sort_key)

# Update index.html dynamically with the collected apps_info
index_html_path = os.path.join(output_dir, "index.html")
if os.path.exists(index_html_path):
    with open(index_html_path, "r", encoding="utf-8") as f:
        index_content = f.read()
    
    apps_info_json = json.dumps(apps_info, ensure_ascii=False)
    index_content = re.sub(
        r'const appsInfo = \[.*?\];',
        f'const appsInfo = {apps_info_json};',
        index_content,
        flags=re.DOTALL
    )
    
    with open(index_html_path, "w", encoding="utf-8") as f:
        f.write(index_content)
            
print(f"Success! Generated {len(apps_info)} learning apps and populated the dashboard.")
