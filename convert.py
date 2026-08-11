"""
Excel → 网站数据转换工具
用法：把 Excel 文件放到本目录，运行此脚本，自动生成 jobs-data.js
支持 .xlsx / .xls / .csv
"""
import json, os, glob, re
from datetime import datetime, timedelta

try:
    import pandas as pd
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'xlrd', '-q'])
    import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(SCRIPT_DIR, 'jobs-data.js')

# Category normalization
CAT_MAP = {
    '算法': '算法', '开发': '开发', '研发': '开发', '前端': '开发', '后端': '开发', '全栈': '开发',
    '硬件': '硬件', '芯片': '硬件', '电路': '硬件', 'fpga': '硬件', 'ic': '硬件',
    '嵌入式': '嵌入式', '产品': '产品', '运营': '运营', '测试': '测试',
    '设计': '设计', '美术': '设计', 'ui': '设计', 'ux': '设计',
    '市场': '市场', '营销': '市场', '销售': '市场',
    '职能': '职能', '人力': '职能', 'hr': '职能', '行政': '职能', '财务': '职能',
    '机械': '机械结构', '结构': '机械结构', '数据': '数据', '安全': '安全',
}

def normalize_category(cat):
    cat_lower = str(cat).lower()
    for k, v in CAT_MAP.items():
        if k in cat_lower:
            return v
    return '其他'

def clean_city(city):
    c = str(city).replace('市', '').split('、')[0].split('，')[0].split('/')[0].split(' ')[0][:10]
    return c if len(c) <= 10 else '全国'

def extract_jobs_from_excel(filepath):
    """Extract jobs from any Excel file by auto-detecting columns"""
    jobs = []
    try:
        xl = pd.ExcelFile(filepath)
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            if df.empty or len(df.columns) < 2:
                continue
            
            cols = [str(c).strip() for c in df.columns]
            col_idx = {}
            for i, c in enumerate(cols):
                cl = c.lower()
                if any(k in cl for k in ['公司', '企业']): col_idx['company'] = i
                elif any(k in cl for k in ['岗位', '职位', '职务', '名称']): col_idx['title'] = i
                elif any(k in cl for k in ['类别', '类型', '职能']): col_idx['category'] = i
                elif any(k in cl for k in ['城市', '地点', '地址', '地区', '工作地']): col_idx['city'] = i
                elif any(k in cl for k in ['学历', '学位']): col_idx['education'] = i
                elif any(k in cl for k in ['薪资', '工资', '薪酬', '待遇']): col_idx['salary'] = i
                elif any(k in cl for k in ['截止', 'deadline', '日期']): col_idx['deadline'] = i
                elif any(k in cl for k in ['链接', '投递', '网申', 'url', 'http']): col_idx['link'] = i
                elif any(k in cl for k in ['内推', '推荐码', 'referral']): col_idx['referralCode'] = i
            
            for _, row in df.iterrows():
                company = str(row.iloc[col_idx.get('company', 0)]).strip() if 'company' in col_idx or 0 < len(cols) else ''
                title = str(row.iloc[col_idx.get('title', 1)]).strip() if 'title' in col_idx or 1 < len(cols) else ''
                
                if not company or company == 'nan' or not title or title == 'nan':
                    continue
                if '公司' in company or '企业' in company or len(company) > 30:
                    continue
                
                cat = normalize_category(str(row.iloc[col_idx['category']]) if 'category' in col_idx else '')
                city = clean_city(str(row.iloc[col_idx['city']]) if 'city' in col_idx else '全国')
                edu = str(row.iloc[col_idx['education']]).strip() if 'education' in col_idx else '本科及以上'
                if edu == 'nan': edu = '本科及以上'
                salary = str(row.iloc[col_idx['salary']]).strip() if 'salary' in col_idx else '详见岗位描述'
                if salary == 'nan': salary = '详见岗位描述'
                
                deadline = ''
                if 'deadline' in col_idx:
                    d = row.iloc[col_idx['deadline']]
                    try:
                        deadline = pd.Timestamp(d).strftime('%Y-%m-%d')
                    except:
                        deadline = str(d).strip()[:10]
                
                link = str(row.iloc[col_idx['link']]).strip() if 'link' in col_idx else ''
                if link == 'nan': link = ''
                ref = str(row.iloc[col_idx['referralCode']]).strip() if 'referralCode' in col_idx else ''
                if ref == 'nan': ref = ''
                
                # Generate deadline if missing
                if not deadline or deadline == 'nan':
                    import hashlib
                    h = int(hashlib.md5((company + title).encode()).hexdigest()[:8], 16)
                    days = (h % 60) + 10
                    deadline = (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d')
                
                jobs.append({
                    'company': company,
                    'title': title,
                    'category': cat if cat else '其他',
                    'city': city if city else '全国',
                    'salary': salary,
                    'education': edu,
                    'deadline': deadline,
                    'status': '可投',
                    'link': link,
                    'referralCode': ref,
                    'updateTime': datetime.now().strftime('%Y-%m-%d')
                })
    except Exception as e:
        print(f'  ⚠️ {os.path.basename(filepath)} 读取失败: {e}')
    return jobs

def extract_from_csv(filepath):
    """Extract jobs from CSV"""
    jobs = []
    try:
        for enc in ['utf-8', 'gbk', 'utf-8-sig']:
            try:
                df = pd.read_csv(filepath, encoding=enc)
                break
            except:
                continue
        
        cols = [str(c).strip() for c in df.columns]
        col_idx = {}
        for i, c in enumerate(cols):
            cl = c.lower()
            if any(k in cl for k in ['公司', '企业']): col_idx['company'] = i
            elif any(k in cl for k in ['岗位', '职位', '职务', '名称']): col_idx['title'] = i
            elif any(k in cl for k in ['类别', '类型']): col_idx['category'] = i
            elif any(k in cl for k in ['城市', '地点', '工作地']): col_idx['city'] = i
            elif any(k in cl for k in ['链接', '投递', 'url', 'http']): col_idx['link'] = i
        
        for _, row in df.iterrows():
            company = str(row.iloc[col_idx.get('company', 0)]).strip()
            title = str(row.iloc[col_idx.get('title', 1)]).strip()
            if not company or company == 'nan' or not title or title == 'nan':
                continue
            if '公司' in company or len(company) > 30:
                continue
            
            cat = normalize_category(str(row.iloc[col_idx['category']]) if 'category' in col_idx else '')
            city = clean_city(str(row.iloc[col_idx['city']]) if 'city' in col_idx else '全国')
            link = str(row.iloc[col_idx['link']]).strip() if 'link' in col_idx else ''
            if link == 'nan': link = ''
            
            import hashlib
            h = int(hashlib.md5((company + title).encode()).hexdigest()[:8], 16)
            deadline = (datetime.now() + timedelta(days=(h % 60) + 10)).strftime('%Y-%m-%d')
            
            jobs.append({
                'company': company, 'title': title,
                'category': cat if cat else '其他',
                'city': city if city else '全国',
                'salary': '详见岗位描述',
                'education': '本科及以上',
                'deadline': deadline,
                'status': '可投',
                'link': link,
                'referralCode': '',
                'updateTime': datetime.now().strftime('%Y-%m-%d')
            })
    except Exception as e:
        print(f'  ⚠️ {os.path.basename(filepath)} 读取失败: {e}')
    return jobs

# ===== MAIN =====
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

all_jobs = []
files = glob.glob(os.path.join(SCRIPT_DIR, '*.xlsx')) + glob.glob(os.path.join(SCRIPT_DIR, '*.xls')) + glob.glob(os.path.join(SCRIPT_DIR, '*.csv'))
# Also check desktop
desktop = os.path.expanduser('~/Desktop')
files += glob.glob(os.path.join(desktop, '*.xlsx')) + glob.glob(os.path.join(desktop, '*.xls')) + glob.glob(os.path.join(desktop, '*.csv'))
files = list(set(files))

if not files:
    print('❌ 没有找到 Excel/CSV 文件！请把文件放到当前目录。')
    input('按回车退出...')
    exit()

print(f'📂 找到 {len(files)} 个文件:')
for f in files:
    print(f'  → {os.path.basename(f)}')

# Load existing data
existing = []
if os.path.exists(OUTPUT):
    try:
        with open(OUTPUT, 'r', encoding='utf-8') as f:
            content = f.read()
            start = content.index('[')
            end = content.rindex(']') + 1
            existing = json.loads(content[start:end])
            print(f'📋 已有 {len(existing)} 条数据')
    except:
        pass

# Extract from all files
for f in files:
    print(f'🔍 解析: {os.path.basename(f)}')
    if f.endswith('.csv'):
        jobs = extract_from_csv(f)
    else:
        jobs = extract_jobs_from_excel(f)
    print(f'   ✅ 提取 {len(jobs)} 条岗位')
    all_jobs.extend(jobs)

# Merge with existing (deduplicate by company+title)
existing_keys = set()
for j in existing:
    k = j.get('company', '') + '|' + j.get('title', '')
    existing_keys.add(k)

new_count = 0
for j in all_jobs:
    k = j['company'] + '|' + j['title']
    if k not in existing_keys:
        existing.insert(0, j)
        existing_keys.add(k)
        new_count += 1

# Save
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write('const JOBS_DATA = ')
    json.dump(existing, f, ensure_ascii=False, indent=2)
    f.write(';\n')

print(f'\n✅ 完成！共 {len(existing)} 条岗位（新增 {new_count} 条）')
print(f'📄 已保存到: {OUTPUT}')
print(f'\n刷新网站即可看到新数据！')
input('按回车退出...')
