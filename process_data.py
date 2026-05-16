import pandas as pd
import os

# 1. 定義澳洲各州的固定土地面積 (平方公里)
state_areas = {
    'New South Wales': 800811,
    'Victoria': 227444,
    'Queensland': 1727200,
    'South Australia': 984242,
    'Western Australia': 2527013,
    'Tasmania': 68401,
    'Northern Territory': 1347791,
    'Australian Capital Territory': 2358
}

current_dir = os.path.dirname(os.path.abspath(__file__))
source_file_name = '310104.xlsx' 
input_path = os.path.join(current_dir, 'data', source_file_name)

map_output_path = os.path.join(current_dir, 'data', 'abs_all_years_density.csv')
pyramid_output_path = os.path.join(current_dir, 'data', 'abs_gender_pyramid.csv')

print(f"🚀 啟動終極無損掃描清洗引擎: {input_path}")

# 2. 完全不預設任何 header，直接讀取整張 Excel 作為原始矩陣
excel_file = pd.ExcelFile(input_path)
target_sheet = [s for s in excel_file.sheet_names if 'Data' in s][0]
raw_df = pd.read_excel(input_path, sheet_name=target_sheet, header=None)

# 3. 安全掃描：尋找真正包含欄位說明的「那一行」
columns_row_idx = None
for idx, row in raw_df.iterrows():
    # 🛠️ 核心修正：使用強制的 [str(x) for x in row] 列表推導式，徹底根除 TypeError！
    row_str = " ".join([str(x) for x in row])
    if 'Estimated Resident Population' in row_str and 'Persons' in row_str:
        columns_row_idx = idx
        print(f"💡 完美定位：在 Excel 的第 {idx + 1} 行成功捕獲資料標頭描述！")
        break

if columns_row_idx is None:
    print("❌ 錯誤：在 Excel 中找不到包含 'Estimated Resident Population' 的描述行。")
    exit()

# 4. 用抓到的那一行重新建立乾淨的欄位名稱
clean_columns = raw_df.iloc[columns_row_idx].astype(str).str.strip().tolist()
clean_columns[0] = 'Date'  # 強制第一欄命名為 Date

# 擷取真實資料列 (描述行之後的所有資料)
data_df = raw_df.iloc[columns_row_idx + 1:].copy()
data_df.columns = clean_columns

map_data = []
pyramid_data = []

# 5. 精準拆解與結構化數據
for index, row in data_df.iterrows():
    date_val = row['Date']
    if pd.isna(date_val): 
        continue
    date_str = str(date_val).strip()
    
    # 解析年份 (相容 '1981-06-01' 或 datetime 物件格式)
    try:
        year = int(date_str.split('-')[0])
    except: 
        continue
    
    # 篩選年中普查數據 (每年 6 月)，滑桿拉起來最順暢、網頁體積最輕量
    if '-06-' in date_str:
        for col in data_df.columns:
            if col == 'Date' or pd.isna(col):
                continue
            
            # 🛠️ 結構化分割：ABS 的欄位都是用分號 ";" 隔開的
            # 例如: ['Estimated Resident Population', 'Male', 'New South Wales', '']
            parts = [p.strip() for p in str(col).split(";")]
            
            if len(parts) >= 3:
                sex = parts[1]      # 'Male', 'Female', 或 'Persons'
                region = parts[2]   # 州名 或 'Australia'
                
                try:
                    population = int(float(row[col]))
                except:
                    continue
                
                # A. 收集地圖數據：只要總人口(Persons)且屬於八個州之一
                if sex == 'Persons' and region in state_areas:
                    area = state_areas[region]
                    map_data.append({
                        'State': region,  # 保持全稱 'New South Wales'，完美黏上你的 TopoJSON
                        'Year': year, 
                        'Population': population, 
                        'Density': round(population / area, 4)
                    })
                
                # B. 收集金字塔數據：只要全澳洲總計(Australia)的男性與女性人口消長
                elif region == 'Australia' and (sex == 'Male' or sex == 'Female'):
                    pyramid_data.append({
                        'Year': year,
                        'Sex': sex,
                        'Population': population
                    })

# 6. 驗證並匯出為網頁專用精簡 CSV
if map_data:
    pd.DataFrame(map_data).to_csv(map_output_path, index=False)
    pd.DataFrame(pyramid_data).to_csv(pyramid_output_path, index=False)
    print(f"\n✨ 【全自動資料清洗——大成功！】")
    print(f"📊 1981-2024 地圖歷史數據 (abs_all_years_density.csv) 共 {len(map_data)} 筆。")
    print(f"📊 1981-2024 男女人口結構 (abs_gender_pyramid.csv) 共 {len(pyramid_data)} 筆。")
    print(f"📂 乾淨的網頁數據庫已完美注入至你的 data/ 資料夾中！")
else:
    print("\n❌ 錯誤：欄位結構拆解未命中任何數據。")