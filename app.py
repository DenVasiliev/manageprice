from flask import Flask, render_template, request, send_file, flash, redirect, url_for, after_this_request
import pandas as pd
import os
import re
import uuid
import io
from werkzeug.utils import secure_filename
import traceback

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = '/tmp/uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def clean_val(val):
    if pd.isna(val): return 0.0
    s = str(val).strip().lower()
    if 'более' in s:
        nums = re.findall(r'\d+', s)
        return float(nums[0]) if nums else 10.0
    try:
        res = s.replace(',', '.').replace(' ', '').replace('\xa0', '')
        return float(res)
    except:
        return 0.0

def process_madwave(site_df, vendor_path):
    mw = pd.read_excel(vendor_path, engine='openpyxl')
    db = {}
    for _, row in mw.iterrows():
        sku = str(row.get('article', '')).strip()
        if sku and sku != 'nan':
            db[sku] = {
                'roz': clean_val(row.get('rrp', 0)),
                'price': clean_val(row.get('price', 0)),
                'stock': clean_val(row.get('stock', 0)),
            }

    stats = {'updated': 0, 'zeroed': 0}
    for idx, row in site_df.iterrows():
        sku = str(row.get('Артикул', '')).strip()
        title = str(row.get('Название товара', '')).lower()

        if sku in db:
            new_roz = db[sku]['roz']
            if new_roz > 0:
                site_df.at[idx, 'Цена продажи, без учёта скидок'] = new_roz
            site_df.at[idx, 'Закупочная цена'] = db[sku]['price']
            site_df.at[idx, 'Остаток'] = db[sku]['stock']
            stats['updated'] += 1
        elif 'madwave' in title:
            site_df.at[idx, 'Остаток'] = 0
            stats['zeroed'] += 1

    return site_df, stats

def process_scorpena(site_df, vendor_path):
    engine = 'openpyxl'
    try:
        pd.read_excel(vendor_path, engine='openpyxl', header=None, nrows=5)
    except Exception:
        engine = 'xlrd'
    
    sc = pd.read_excel(vendor_path, engine=engine)

    dealer_price_col = None
    for col in sc.columns:
        col_clean = str(col).replace(' ', '').lower()
        if 'дилерская' in col_clean and 'цена' in col_clean:
            dealer_price_col = col
            break
    if not dealer_price_col:
        dealer_price_col = sc.columns[5]

    db = {}
    for _, row in sc.iterrows():
        sku = str(row.get('Артикул', '')).strip()
        if sku and sku != 'nan':
            db[sku] = {
                'roz': clean_val(row.get('Минимальная Цена в Рекламе (МЦР)', 0)),
                'price': clean_val(row.get(dealer_price_col, 0)),
                'stock': clean_val(row.get('Количество', 0)),
            }

    stats = {'updated': 0, 'zeroed': 0, 'dealer_col': dealer_price_col}
    for idx, row in site_df.iterrows():
        sku = str(row.get('Артикул', '')).strip()
        title = str(row.get('Название товара', '')).lower()

        if sku in db:
            new_roz = db[sku]['roz']
            if new_roz > 0:
                site_df.at[idx, 'Цена продажи, без учёта скидок'] = new_roz
            site_df.at[idx, 'Закупочная цена'] = db[sku]['price']
            site_df.at[idx, 'Остаток'] = db[sku]['stock']
            stats['updated'] += 1
        elif 'scorpena' in title:
            site_df.at[idx, 'Остаток'] = 0
            stats['zeroed'] += 1

    return site_df, stats

def process_igrushka(site_df, vendor_path, deposit_col_idx):
    try:
        engine = 'openpyxl'
        try:
            pd.read_excel(vendor_path, engine='openpyxl', header=None, nrows=5)
        except Exception:
            engine = 'xlrd'

        df_raw = pd.read_excel(vendor_path, engine=engine, header=None)
        header_row = None
        for idx, row in df_raw.iterrows():
            if 'КОДЫ' in row.astype(str).values:
                header_row = idx
                break
        
        if header_row is None:
            raise ValueError("В прайсе Игрушки не найден заголовок 'КОДЫ'. Проверьте файл.")
            
        df = pd.read_excel(vendor_path, engine=engine, header=header_row)
        
        stock_col = None
        for col in df.columns:
            col_str = str(col).lower()
            if 'склад' in col_str or 'остаток' in col_str or 'наличие' in col_str or 'кол-во' in col_str:
                stock_col = col
                break
                
        db = {}
        for _, row in df.iterrows():
            sku = str(row.get('КОДЫ', '')).strip()
            if sku and sku != 'nan':
                if stock_col:
                    stock_val = row.get(stock_col, 0)
                else:
                    stock_val = row.iloc[14] if len(row) > 14 else 0
                
                stock_raw = str(stock_val).lower()
                stock = 100.0 if 'более' in stock_raw else clean_val(stock_val)
                
                price = clean_val(row.iloc[deposit_col_idx])
                
                db[sku] = {
                    'price': price,
                    'stock': stock,
                }
    except Exception as e:
        raise ValueError(f"Ошибка чтения прайса Игрушки: {str(e)}")

    stats = {'updated': 0, 'zeroed': 0}
    
    for idx, row in site_df.iterrows():
        sku = str(row.get('Артикул', '')).strip()
        
        if sku.startswith('SO-'):
            clean_sku = sku.replace('SO-', '', 1)
            
            if clean_sku in db:
                site_df.at[idx, 'Закупочная цена'] = db[clean_sku]['price']
                site_df.at[idx, 'Остаток'] = db[clean_sku]['stock']
                stats['updated'] += 1
            else:
                site_df.at[idx, 'Остаток'] = 0
                stats['zeroed'] += 1
                
    return site_df, stats

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files_to_cleanup = []
        
        @after_this_request
        def cleanup(response):
            for file_path in files_to_cleanup:
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except Exception:
                    pass
            return response
        
        vendor = request.form.get('vendor')
        deposit_tier = request.form.get('deposit_tier')
        site_file = request.files.get('site_file')
        vendor_file = request.files.get('vendor_file')

        if not site_file or not site_file.filename:
            flash('Загрузите файл выгрузки с сайта', 'error')
            return redirect(url_for('index'))
        if not vendor_file or not vendor_file.filename:
            flash('Загрузите прайс-лист поставщика', 'error')
            return redirect(url_for('index'))
        if not vendor:
            flash('Выберите поставщика', 'error')
            return redirect(url_for('index'))
        
        if vendor == 'igrushka' and not deposit_tier:
            flash('Для Игрушки необходимо выбрать размер депозита', 'error')
            return redirect(url_for('index'))

        job_id = str(uuid.uuid4())[:8]
        site_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_site.csv')
        vendor_fn = secure_filename(vendor_file.filename)
        vendor_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_{vendor_fn}')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_result.csv')

        files_to_cleanup.extend([site_path, vendor_path])

        site_file.save(site_path)
        vendor_file.save(vendor_path)

        try:
            site_df = pd.read_csv(site_path, sep=';', encoding='cp1251', engine='python')
            
            for col in ['Закупочная цена', 'Остаток', 'Цена продажи, без учёта скидок']:
                if col in site_df.columns:
                    site_df[col] = site_df[col].astype(str).apply(clean_val)

            if vendor == 'madwave':
                site_df, stats = process_madwave(site_df, vendor_path)
            elif vendor == 'scorpena':
                site_df, stats = process_scorpena(site_df, vendor_path)
            elif vendor == 'igrushka':
                site_df, stats = process_igrushka(site_df, vendor_path, int(deposit_tier))
            else:
                flash('Неизвестный поставщик', 'error')
                return redirect(url_for('index'))

            site_df.to_csv(output_path, sep=';', index=False, encoding='cp1251', decimal=',')

            return redirect(url_for('index', 
                                   download=job_id, 
                                   vendor=vendor,
                                   updated=stats['updated'],
                                   zeroed=stats['zeroed']))

        except Exception as e:
            error_msg = f'Ошибка обработки: {str(e)}'
            print(f"ERROR: {error_msg}", flush=True)
            print(traceback.format_exc(), flush=True)
            flash(error_msg, 'error')
            return redirect(url_for('index'))

    return render_template('index.html')

@app.route('/download/<job_id>')
def download_result(job_id):
    vendor = request.args.get('vendor', 'unknown')
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_result.csv')
    
    if os.path.exists(output_path):
        try:
            with open(output_path, 'rb') as f:
                file_data = f.read()
            
            os.remove(output_path)
            
            return send_file(
                io.BytesIO(file_data),
                as_attachment=True,
                download_name=f'FINAL_UPDATE_{vendor}.csv',
                mimetype='text/csv'
            )
        except Exception as e:
            print(f"ERROR при скачивании: {e}", flush=True)
            return redirect(url_for('index'))
    else:
        # Просто редиректим без flash-сообщения
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
