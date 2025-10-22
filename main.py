import os
import json
import time
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import traceback

# Load environment variables
load_dotenv()

app = Flask(__name__)
APP_VERSION = "1.0.2"

# ==================== CACHE SYSTEM ====================
data_cache = {}
cache_timestamp = {}
last_request_time = {}
REQUEST_INTERVAL = 2  # 2 giây giữa các request

def rate_limit(api_name):
    """Giới hạn tần suất request"""
    current_time = time.time()
    
    if api_name in last_request_time:
        time_since_last = current_time - last_request_time[api_name]
        if time_since_last < REQUEST_INTERVAL:
            wait_time = REQUEST_INTERVAL - time_since_last
            print(f"⏳ [RATE LIMIT] Chờ {wait_time:.1f}s cho {api_name}")
            time.sleep(wait_time)
    
    last_request_time[api_name] = time.time()

def get_cached_data(sheet_name, cache_duration=30):
    """Lấy dữ liệu có cache để giảm request"""
    current_time = time.time()
    
    # Kiểm tra cache
    if (sheet_name in data_cache and 
        sheet_name in cache_timestamp and
        current_time - cache_timestamp[sheet_name] < cache_duration):
        print(f"📦 [CACHE] Sử dụng cache cho {sheet_name}")
        return data_cache[sheet_name]
    
    # Lấy dữ liệu mới
    print(f"🔄 [CACHE] Lấy dữ liệu mới cho {sheet_name}")
    try:
        client = connect_to_sheets()
        if not client:
            return []
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        
        data = sheet.get_all_values()
        
        # Lưu cache
        data_cache[sheet_name] = data
        cache_timestamp[sheet_name] = current_time
        
        return data
        
    except Exception as e:
        print(f"❌ [CACHE] Lỗi lấy dữ liệu {sheet_name}: {e}")
        return []

def clear_cache():
    """Xóa cache (có thể gọi từ API nếu cần)"""
    data_cache.clear()
    cache_timestamp.clear()
    print("🧹 [CACHE] Đã xóa toàn bộ cache")

# ==================== GOOGLE SHEETS CONNECTION ====================
def connect_to_sheets():
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        # Ưu tiên dùng file JSON trực tiếp
        creds_file = 'credentials/service-account.json'
        if os.path.exists(creds_file):
            print("✅ Đang dùng file JSON credentials")
            creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
        else:
            # Dùng environment variables
            print("✅ Đang dùng environment variables")
            private_key = os.environ.get('PRIVATE_KEY', '')
            
            # Fix private key formatting
            private_key = private_key.replace('"', '').replace('\\n', '\n')
            
            creds_dict = {
                "type": "service_account",
                "project_id": os.environ.get('PROJECT_ID'),
                "private_key_id": os.environ.get('PRIVATE_KEY_ID'),
                "private_key": private_key,
                "client_email": os.environ.get('CLIENT_EMAIL'),
                "client_id": os.environ.get('CLIENT_ID'),
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
            
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        
        client = gspread.authorize(creds)
        print("✅ Kết nối Google Sheets thành công!")
        return client
        
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {str(e)}")
        return None

def normalize_date(date_str):
    """Chuẩn hóa định dạng ngày từ nhiều định dạng khác nhau"""
    if not date_str:
        return ""
    
    # Nếu là đối tượng datetime, chuyển thành string
    if hasattr(date_str, 'strftime'):
        return date_str.strftime("%d/%m/%Y")
    
    # Chuẩn hóa chuỗi ngày
    date_str = str(date_str).strip()
    
    # Xử lý các định dạng ngày phổ biến
    try:
        # Định dạng dd/mm/yyyy
        if '/' in date_str and len(date_str.split('/')) == 3:
            parts = date_str.split('/')
            if len(parts) == 3:
                day, month, year = parts
                # Đảm bảo đủ 2 chữ số cho ngày và tháng
                day = day.zfill(2)
                month = month.zfill(2)
                # Đảm bảo năm có 4 chữ số
                if len(year) == 2:
                    year = '20' + year
                return f"{day}/{month}/{year}"
        
        # Định dạng khác, thử parse với datetime
        for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%m/%d/%Y']:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime("%d/%m/%Y")
            except ValueError:
                continue
                
    except Exception:
        pass
    
    return date_str

# ==================== ROUTES CHÍNH ====================
@app.route('/')
def index():
    return render_template('index.html', version=APP_VERSION)

@app.route('/report')
def report():
    return render_template('report.html')

@app.route('/display')
def display():
    return render_template('display.html')

# ==================== API ENDPOINTS (OPTIMIZED) ====================
@app.route('/api/add_dulieusv', methods=['POST'])
def add_dulieusv():
    try:
        data = request.json
        client = connect_to_sheets()
        
        if not client:
            return jsonify({'error': 'Không thể kết nối Google Sheets'}), 500
        
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet_data = spreadsheet.worksheet('Data')
        sheet_listds = spreadsheet.worksheet('LISTDS')
        
        mssv = data.get('mssv', '')
        khoavien = data.get('khoavien', '')
        phonghocnhom = data.get('phonghocnhom', '')
        soluong = data.get('soluong', '')
        nguoi_nhap = data.get('nguoiNhap', '')
        
        if not all([mssv, khoavien, phonghocnhom, soluong, nguoi_nhap]):
            return jsonify({'error': 'Vui lòng nhập đầy đủ thông tin!'}), 400
        
        current_time = datetime.now()
        check_in_time = current_time.strftime("%H:%M:%S")
        check_out_time = (current_time + timedelta(minutes=90)).strftime("%H:%M:%S")
        current_date = current_time.strftime("%d/%m/%Y")
        
        phong_num = int(phonghocnhom)
        if 1 <= phong_num <= 7:
            floor_position = 'Lầu 3'
        elif 8 <= phong_num <= 14:
            floor_position = 'Lầu 4'
        else:
            floor_position = 'Tầng trệt'
        
        new_row = [
            mssv, khoavien, phonghocnhom, soluong, check_in_time,
            check_out_time, current_date, floor_position,
            f"Tháng {current_time.month} năm {current_time.year}",
            f"Phòng {phonghocnhom}", nguoi_nhap
        ]
        
        sheet_data.append_row(new_row)
        
        # Xóa cache Data vì có dữ liệu mới
        if 'Data' in data_cache:
            del data_cache['Data']
            print("🧹 [CACHE] Đã xóa cache Data do có dữ liệu mới")
        
        listds_data = sheet_listds.get_all_values()
        existing_mssvs = [row[0].replace("'", "") for row in listds_data[1:] if row]
        
        clean_mssv = mssv.replace("'", "")
        if clean_mssv not in existing_mssvs:
            sheet_listds.append_row([mssv, khoavien])
        
        return jsonify({'message': 'Dữ liệu đã được thêm thành công'})
        
    except Exception as e:
        print(f"Lỗi add_dulieusv: {e}")
        return jsonify({'error': f'Lỗi server: {str(e)}'}), 500

@app.route('/api/get_data')
def get_data():
    try:
        rate_limit('get_data')
        print("🔍 [get_data] Đang lấy dữ liệu (cached)...")
        
        # Sử dụng cache - 30 giây
        data = get_cached_data('Data', 30)
        
        if len(data) <= 1:
            return jsonify([])
            
        formatted_data = []
        for row in data[1:]:
            if len(row) >= 7:
                formatted_row = []
                for col_idx, cell in enumerate(row[:7]):
                    cell_value = str(cell) if cell is not None else ""
                    
                    # Xử lý cột ngày (G)
                    if col_idx == 6:
                        normalized = normalize_date(cell_value)
                        formatted_row.append(normalized)
                    # Xử lý cột giờ (E, F)
                    elif col_idx in [4, 5]:
                        if cell_value and ':' in cell_value:
                            formatted_row.append(cell_value)
                        else:
                            formatted_row.append("")
                    else:
                        formatted_row.append(cell_value)
                
                formatted_data.append(formatted_row)
        
        result = formatted_data[::-1]  # Reverse order
        print(f"✅ [get_data] Trả về {len(result)} bản ghi")
        return jsonify(result)
        
    except Exception as e:
        print(f"❌ LỖI get_data: {str(e)}")
        return jsonify([])

@app.route('/api/get_data1')
def get_data1():
    try:
        rate_limit('get_data1')
        print("🔍 [get_data1] Đang lấy dữ liệu (cached)...")
        
        # Sử dụng cache - 30 giây
        data = get_cached_data('Data1', 30)
        
        if len(data) < 2:
            return jsonify([])
            
        formatted_data = []
        for row in data[1:]:
            if len(row) >= 6:
                formatted_row = [str(cell) if cell is not None else "" for cell in row[:6]]
                formatted_data.append(formatted_row)
        
        print(f"✅ [get_data1] Trả về {len(formatted_data)} bản ghi")
        return jsonify(formatted_data)
        
    except Exception as e:
        print(f"❌ LỖI get_data1: {str(e)}")
        return jsonify([])

@app.route('/api/get_data_count_today')
def get_data_count_today():
    try:
        print("🔍 [get_data_count_today] Đang tính thống kê hôm nay...")
        
        client = connect_to_sheets()
        if not client:
            return jsonify({"count": 0})
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('Data')
        
        data = sheet.get_all_values()
        
        if len(data) <= 1:
            return jsonify({"count": 0})
            
        today = datetime.now().strftime("%d/%m/%Y")
        count = 0
        
        for row in data[1:]:
            if len(row) >= 11:  # Đảm bảo có đủ cột đến K
                date_cell = row[6]  # Cột G - Ngày
                normalized_date = normalize_date(date_cell)
                
                if normalized_date == today:
                    count += 1
        
        print(f"✅ [get_data_count_today] Kết quả: {count} lượt hôm nay")
        return jsonify({"count": count})
        
    except Exception as e:
        print(f"❌ [get_data_count_today] Lỗi: {e}")
        return jsonify({"count": 0})

@app.route('/api/get_data1_count_today')
def get_data1_count_today():
    try:
        print("🔍 [get_data1_count_today] Đang tính thống kê đăng ký hôm nay...")
        
        client = connect_to_sheets()
        if not client:
            return jsonify(0)
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('Data1')
        
        today = datetime.now().strftime("%d/%m/%Y")
        data = sheet.get_all_values()
        
        if len(data) <= 1:
            return jsonify(0)
            
        count = 0
        
        for row in data[1:]:
            if len(row) >= 6:
                date_cell = row[5]  # Cột F - Ngày trong Data1
                normalized_date = normalize_date(date_cell)
                
                if normalized_date == today:
                    count += 1
        
        print(f"✅ [get_data1_count_today] Kết quả: {count} lượt đăng ký hôm nay")
        return jsonify(count)
        
    except Exception as e:
        print(f"❌ [get_data1_count_today] Lỗi: {e}")
        return jsonify(0)

@app.route('/api/get_current_month_count_data')
def get_current_month_count_data():
    try:
        print("🔍 [get_current_month_count_data] Đang tính thống kê tháng...")
        
        client = connect_to_sheets()
        if not client:
            return jsonify(0)
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('Data')
        
        current_date = datetime.now()
        current_month = current_date.month
        current_year = current_date.year
        
        data = sheet.get_all_values()
        
        if len(data) <= 1:
            return jsonify(0)
            
        count = 0
        
        for row in data[1:]:
            if len(row) >= 11:
                # Kiểm tra cả cột ngày (G) và cột tháng (I)
                date_cell = row[6]  # Cột G - Ngày
                month_cell = row[8]  # Cột I - Tháng
                
                # Ưu tiên kiểm tra cột tháng trước
                if month_cell and "Tháng" in month_cell:
                    try:
                        parts = month_cell.split()
                        if len(parts) >= 4:
                            month_str = parts[1]
                            year_str = parts[3]
                            
                            month_num = int(month_str)
                            year_num = int(year_str)
                            
                            if month_num == current_month and year_num == current_year:
                                count += 1
                                continue
                    except (ValueError, IndexError):
                        pass
                
                # Nếu không có thông tin tháng, kiểm tra từ cột ngày
                normalized_date = normalize_date(date_cell)
                if normalized_date and '/' in normalized_date:
                    try:
                        day, month, year = normalized_date.split('/')
                        month_num = int(month)
                        year_num = int(year)
                        
                        if month_num == current_month and year_num == current_year:
                            count += 1
                    except (ValueError, IndexError):
                        continue
        
        print(f"✅ [get_current_month_count_data] Kết quả: {count} lượt tháng {current_month}")
        return jsonify(count)
        
    except Exception as e:
        print(f"❌ [get_current_month_count_data] Lỗi: {e}")
        return jsonify(0)

@app.route('/api/get_online_data')
def get_online_data():
    try:
        rate_limit('get_online_data')
        print("🔍 [get_online_data] Đang lấy dữ liệu (cached)...")
        
        # Sử dụng cache - 15 giây (online data cần cập nhật thường xuyên hơn)
        data = get_cached_data('Online', 15)
        
        if len(data) < 2:
            return jsonify({'headers': [], 'data': []})

        headers = data[0]
        headers = [headers[i] for i in [0, 1, 3] if i < len(headers)]

        result_data = []
        for i in range(1, min(len(data), 21)):  # Giới hạn 20 dòng
            row = data[i]
            if len(row) >= 4:
                result_data.append([
                    row[0] if len(row) > 0 else '',
                    row[1] if len(row) > 1 else '',
                    row[3] if len(row) > 3 else ''
                ])

        print(f"✅ [get_online_data] Trả về {len(result_data)} bản ghi")
        return jsonify({'headers': headers, 'data': result_data})
        
    except Exception as e:
        print(f"❌ [get_online_data] Lỗi: {e}")
        return jsonify({'headers': [], 'data': []})

@app.route('/api/search_data')
def search_data():
    try:
        keyword = request.args.get('keyword', '')
        if not keyword:
            return jsonify([])
            
        rate_limit('search_data')
        client = connect_to_sheets()
        if not client:
            return jsonify([])
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('LISTDS')
        
        data = sheet.get_all_values()
        result = []
        
        for row in data[1:]:
            if len(row) >= 2 and keyword and keyword in str(row[0]):
                result.append([row[0], row[1]])
                break
        
        return jsonify(result)
        
    except Exception as e:
        print(f"Lỗi search_data: {e}")
        return jsonify([])

@app.route('/api/get_nguoinhap_options')
def get_nguoinhap_options():
    try:
        rate_limit('get_nguoinhap_options')
        client = connect_to_sheets()
        if not client:
            return jsonify([])
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('LISTDS')
        
        data = sheet.col_values(4)
        if len(data) > 1:
            return jsonify([item for item in data[1:21] if item])
        return jsonify([])
        
    except Exception as e:
        print(f"Lỗi get_nguoinhap_options: {e}")
        return jsonify([])

@app.route('/api/delete_data1', methods=['GET'])
def delete_data1():
    try:
        index = int(request.args.get('index', 0))
        rate_limit('delete_data1')
        client = connect_to_sheets()
        if not client:
            return jsonify([])
            
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet('Data1')
        
        sheet.delete_rows(index + 2)
        
        # Xóa cache Data1
        if 'Data1' in data_cache:
            del data_cache['Data1']
            print("🧹 [CACHE] Đã xóa cache Data1 do xóa dữ liệu")
            
        return get_data1()
        
    except Exception as e:
        print(f"Lỗi delete_data1: {e}")
        return jsonify([])

# ==================== DEBUG & UTILITY ENDPOINTS ====================
@app.route('/api/debug_cache')
def debug_cache():
    """Xem trạng thái cache"""
    cache_info = {}
    current_time = time.time()
    
    for sheet_name in data_cache:
        if sheet_name in cache_timestamp:
            age = current_time - cache_timestamp[sheet_name]
            cache_info[sheet_name] = {
                'cached': True,
                'age_seconds': round(age, 1),
                'rows': len(data_cache[sheet_name])
            }
        else:
            cache_info[sheet_name] = {
                'cached': False,
                'age_seconds': None,
                'rows': 0
            }
    
    return jsonify({
        'cache_info': cache_info,
        'total_cached_sheets': len(data_cache)
    })

@app.route('/api/clear_cache')
def clear_cache_endpoint():
    """API để xóa cache thủ công"""
    clear_cache()
    return jsonify({'message': 'Cache đã được xóa'})

@app.route('/api/test_connection')
def test_connection():
    """Test kết nối cơ bản đến Google Sheets"""
    try:
        rate_limit('test_connection')
        print("=== TEST CONNECTION ===")
        client = connect_to_sheets()
        if not client:
            return jsonify({'status': 'error', 'message': 'No connection'})
        
        sheet_id = os.environ.get('SHEET_ID', '1i5N5Gdk-SqPN7Vy5IFiHiK5CTCw9WDag2EMZ1GBI8Wo')
        spreadsheet = client.open_by_key(sheet_id)
        
        # Test các sheet tồn tại
        sheets_info = []
        for sheet_name in ['Data', 'Data1', 'LISTDS', 'Online']:
            try:
                sheet = spreadsheet.worksheet(sheet_name)
                row_count = len(sheet.get_all_values())
                sheets_info.append({
                    'name': sheet_name,
                    'rows': row_count,
                    'status': 'OK'
                })
                print(f"TEST: Sheet '{sheet_name}' - {row_count} rows")
            except Exception as e:
                sheets_info.append({
                    'name': sheet_name,
                    'rows': 0,
                    'status': f'Error: {str(e)}'
                })
                print(f"TEST: Sheet '{sheet_name}' - ERROR: {e}")
        
        return jsonify({
            'status': 'success',
            'sheets': sheets_info
        })
        
    except Exception as e:
        print(f"❌ TEST CONNECTION ERROR: {e}")
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/quick_stats')
def quick_stats():
    """Thống kê nhanh tất cả"""
    try:
        rate_limit('quick_stats')
        results = {}
        
        # Lấy tất cả dữ liệu một lần
        data = get_cached_data('Data', 30)
        data1 = get_cached_data('Data1', 30)
        
        today = datetime.now().strftime("%d/%m/%Y")
        current_month = datetime.now().month
        current_year = datetime.now().year
        
        # Thống kê Data
        data_count_today = 0
        month_count = 0
        
        if len(data) > 1:
            for row in data[1:]:
                if len(row) >= 7:
                    date_cell = row[6]
                    normalized_date = normalize_date(date_cell)
                    
                    if normalized_date == today:
                        data_count_today += 1
                    
                    if normalized_date and '/' in normalized_date:
                        try:
                            day, month, year = normalized_date.split('/')
                            if int(month) == current_month and int(year) == current_year:
                                month_count += 1
                        except (ValueError, IndexError):
                            continue
        
        # Thống kê Data1
        data1_count_today = 0
        if len(data1) > 1:
            for row in data1[1:]:
                if len(row) >= 6:
                    date_cell = row[5]
                    normalized_date = normalize_date(date_cell)
                    if normalized_date == today:
                        data1_count_today += 1
        
        results = {
            'data_count_today': data_count_today,
            'data1_count_today': data1_count_today,
            'current_month_count': month_count,
            'total_data_records': len(data) - 1,
            'total_data1_records': len(data1) - 1,
            'today': today,
            'cache_status': 'using_cache'
        }
        
        print(f"✅ [quick_stats] Thống kê hoàn tất")
        return jsonify(results)
        
    except Exception as e:
        print(f"❌ [quick_stats] Lỗi: {e}")
        return jsonify({'error': str(e)})

@app.route('/api/get_all_stats')
def get_all_stats():
    """API tổng hợp thống kê cho frontend (used by /api/get_all_stats)"""
    try:
        rate_limit('get_all_stats')
        # Lấy dữ liệu cache (30s)
        data = get_cached_data('Data', 30)
        data1 = get_cached_data('Data1', 30)

        today = datetime.now().strftime("%d/%m/%Y")
        current_month = datetime.now().month
        current_year = datetime.now().year

        today_usage = 0
        month_usage = 0
        today_register = 0

        # Data (lịch sử)
        if len(data) > 1:
            for row in data[1:]:
                if len(row) >= 7:
                    normalized = normalize_date(row[6])
                    if normalized == today:
                        today_usage += 1
                    if normalized and '/' in normalized:
                        try:
                            _, m, y = normalized.split('/')
                            if int(m) == current_month and int(y) == current_year:
                                month_usage += 1
                        except Exception:
                            continue

        # Data1 (đăng ký)
        if len(data1) > 1:
            for row in data1[1:]:
                if len(row) >= 6:
                    normalized = normalize_date(row[5])
                    if normalized == today:
                        today_register += 1

        return jsonify({
            'success': True,
            'data': {
                'today_usage': today_usage,
                'month_usage': month_usage,
                'today_register': today_register
            }
        })
    except Exception as e:
        print(f"❌ [get_all_stats] Lỗi: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/get_report_data', methods=['POST'])
def get_report_data():
    try:
        data = request.json
        staff_code = data.get('staffCode', '')
        location = data.get('location', '')
        start_date_str = data.get('startDate', '')
        end_date_str = data.get('endDate', '')
        
        print(f"📊 [REPORT] Đang xử lý báo cáo: staff={staff_code}, location={location}, from={start_date_str}, to={end_date_str}")
        
        # Lấy dữ liệu từ cache
        sheet_data = get_cached_data('Data', 30)
        
        if len(sheet_data) <= 1:
            return jsonify([])
        
        # Danh sách khoa viện
        departments = [
            'Khoa Công nghệ Cơ khí',
            'Khoa Công nghệ Thông tin',
            'Khoa Công nghệ Điện',
            'Khoa Công nghệ Điện tử',
            'Khoa Công nghệ Động lực',
            'Khoa Công nghệ Nhiệt - Lạnh',
            'Khoa Công nghệ May - Thời trang',
            'Khoa Công nghệ Hóa học',
            'Khoa Ngoại ngữ',
            'Khoa Quản trị Kinh doanh',
            'Khoa Thương mại - Du lịch',
            'Khoa Kỹ thuật Xây dựng',
            'Khoa Luật',
            'Viện Tài chính - Kế toán',
            'Viện Công nghệ Sinh học và Thực phẩm',
            'Viện Khoa học Công nghệ và Quản lý Môi trường',
            'Khoa Khoa học Cơ bản'
        ]
        
        # Khởi tạo kết quả
        result = {}
        for dept in departments:
            result[dept] = {'count': 0, 'sum': 0}
        
        # Xử lý ngày tháng
        try:
            if start_date_str:
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            else:
                start_date = None
                
            if end_date_str:
                end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
            else:
                end_date = None
        except ValueError as e:
            print(f"❌ [REPORT] Lỗi định dạng ngày: {e}")
            return jsonify([])
        
        # Xử lý từng dòng dữ liệu
        for row in sheet_data[1:]:  # Bỏ qua header
            if len(row) < 11:  # Đảm bảo có đủ cột
                continue
                
            # Lấy thông tin từ các cột
            department = row[1] if len(row) > 1 else ''  # Cột B: Khoa viện
            quantity_str = row[3] if len(row) > 3 else '0'  # Cột D: Số lượng
            date_str = row[6] if len(row) > 6 else ''  # Cột G: Ngày
            row_location = row[7] if len(row) > 7 else ''  # Cột H: Vị trí
            staff_id = row[10] if len(row) > 10 else ''  # Cột K: Người nhập
            
            # Kiểm tra điều kiện lọc
            valid_staff = not staff_code or staff_id == staff_code
            valid_location = not location or row_location == location
            
            # Kiểm tra ngày
            valid_date = True
            if start_date or end_date:
                try:
                    row_date = normalize_date(date_str)
                    if '/' in row_date:
                        day, month, year = row_date.split('/')
                        row_date_obj = datetime(int(year), int(month), int(day))
                        
                        if start_date and row_date_obj < start_date:
                            valid_date = False
                        if end_date and row_date_obj > end_date:
                            valid_date = False
                except Exception as e:
                    valid_date = False
            
            # Nếu thỏa mãn tất cả điều kiện
            if valid_staff and valid_location and valid_date and department in result:
                try:
                    quantity = int(float(quantity_str)) if quantity_str else 0
                except ValueError:
                    quantity = 0
                
                result[department]['count'] += 1
                result[department]['sum'] += quantity
        
        # Chuyển đổi kết quả thành danh sách
        processed_data = []
        for dept in departments:
            processed_data.append({
                'faculty': dept,
                'count': result[dept]['count'],
                'sum': result[dept]['sum']
            })
        
        print(f"✅ [REPORT] Trả về {len(processed_data)} khoa viện")
        return jsonify(processed_data)
        
    except Exception as e:
        print(f"❌ [REPORT] Lỗi xử lý báo cáo: {e}")
        import traceback
        traceback.print_exc()
        return jsonify([])
                        
if __name__ == '__main__':
    print("Đang khởi động ứng dụng...")
    try:
        client = connect_to_sheets()
        if client:
            print("✅ Kết nối Google Sheets thành công!")
        else:
            print("❌ Không thể kết nối Google Sheets")
    except Exception as e:
        print(f"❌ Lỗi kết nối: {e}")
    
    app.run(host='0.0.0.0', port=8000, debug=True)
