from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import io
from reconciliation import reconcile_data, format_mst

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB Max Upload size
UPLOAD_FOLDER = r"C:\Users\PRECISION\.gemini\antigravity\brain\579a9696-e4ba-44bd-a3b8-b2ed9ab797aa\scratch\uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/reconcile', methods=['POST'])
def handle_reconcile():
    if 'sapo_file' not in request.files or 'supplier_file' not in request.files or 'invoice_file' not in request.files:
        return jsonify({'error': 'Vui lòng tải lên đầy đủ 3 file nguồn.'}), 400
        
    sapo_file = request.files['sapo_file']
    supplier_file = request.files['supplier_file']
    invoice_file = request.files['invoice_file']
    
    if sapo_file.filename == '' or supplier_file.filename == '' or invoice_file.filename == '':
        return jsonify({'error': 'Một hoặc nhiều file chưa được chọn.'}), 400
        
    try:
        # Read bytes
        sapo_bytes = sapo_file.read()
        supplier_bytes = supplier_file.read()
        invoice_bytes = invoice_file.read()
        
        # Save files temporarily for potential export rebuild
        sapo_path = os.path.join(UPLOAD_FOLDER, 'sapo_latest.xlsx')
        supplier_path = os.path.join(UPLOAD_FOLDER, 'supplier_latest.xlsx')
        invoice_path = os.path.join(UPLOAD_FOLDER, 'invoice_latest.xlsx')
        
        with open(sapo_path, 'wb') as f:
            f.write(sapo_bytes)
        with open(supplier_path, 'wb') as f:
            f.write(supplier_bytes)
        with open(invoice_path, 'wb') as f:
            f.write(invoice_bytes)
            
        # Reconcile
        df_orders, df_items, df_unmatched_inv = reconcile_data(sapo_path, supplier_path, invoice_path)
        
        # Check matching summary stats
        total_orders = len(df_orders)
        stats = {
            'total': total_orders,
            'khop': int((df_orders['status'] == 'Khớp').sum()),
            'lech': int((df_orders['status'] == 'Lệch').sum()),
            'chua_hd': int((df_orders['status'] == 'Chưa xuất hóa đơn').sum()),
            'thieu_mst': int((df_orders['status'] == 'Thiếu MST').sum())
        }
        
        # Convert DataFrames to dict lists
        orders_list = df_orders.to_dict(orient='records')
        items_list = df_items.to_dict(orient='records')
        unmatched_list = df_unmatched_inv.to_dict(orient='records')
        
        return jsonify({
            'success': True,
            'stats': stats,
            'orders': orders_list,
            'items': items_list,
            'unmatched_invoices': unmatched_list
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Đã xảy ra lỗi khi đối chiếu: {str(e)}'}), 500

@app.route('/export', methods=['POST'])
def handle_export():
    try:
        # Load from temp saved paths to rebuild dataframes and export beautifully
        sapo_path = os.path.join(UPLOAD_FOLDER, 'sapo_latest.xlsx')
        supplier_path = os.path.join(UPLOAD_FOLDER, 'supplier_latest.xlsx')
        invoice_path = os.path.join(UPLOAD_FOLDER, 'invoice_latest.xlsx')
        
        if not (os.path.exists(sapo_path) and os.path.exists(supplier_path) and os.path.exists(invoice_path)):
            return "Không tìm thấy dữ liệu đối chiếu trước đó. Vui lòng tải lại file.", 400
            
        df_orders, df_items, df_unmatched_inv = reconcile_data(sapo_path, supplier_path, invoice_path)
        
        # Rename columns for Excel to look user-friendly in Vietnamese
        df_orders_excel = df_orders.rename(columns={
            'order_id': 'Mã đơn nhập hàng',
            'date': 'Ngày nhập',
            'supplier_code': 'Mã NCC (Sapo)',
            'supplier_name': 'Tên NCC (Sapo)',
            'mst': 'Mã số thuế',
            'total_incl': 'Tổng tiền (Có VAT)',
            'total_excl': 'Tổng tiền (Không VAT)',
            'status': 'Trạng thái đối chiếu',
            'notes': 'Chi tiết sai lệch'
        })
        
        df_items_excel = df_items.rename(columns={
            'order_id': 'Mã đơn nhập hàng',
            'sku': 'Mã SKU Sapo',
            'name': 'Tên sản phẩm Sapo',
            'qty': 'SL nhập (Sapo)',
            'price_excl': 'Đơn giá chưa thuế (Sapo)',
            'total_excl': 'Thành tiền chưa thuế (Sapo)',
            'invoice_no': 'Số hóa đơn TCT',
            'inv_name': 'Tên hàng trên hóa đơn',
            'inv_qty': 'SL hóa đơn (TCT)',
            'inv_price': 'Đơn giá hóa đơn (TCT)',
            'inv_total': 'Thành tiền hóa đơn (TCT)',
            'status': 'Kết quả',
            'diff_qty': 'Lệch số lượng',
            'diff_amt': 'Lệch thành tiền'
        })
        
        df_unmatched_excel = df_unmatched_inv.rename(columns={
            'mst': 'Mã số thuế người bán',
            'supplier_name': 'Tên nhà cung cấp',
            'date': 'Ngày hóa đơn',
            'invoice_no': 'Số hóa đơn TCT',
            'name': 'Tên hàng hóa TCT',
            'qty': 'Số lượng',
            'price': 'Đơn giá chưa thuế',
            'total': 'Thành tiền chưa thuế'
        })
        
        # Create bytes stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write sheets
            df_orders_excel.to_excel(writer, sheet_name='Tổng quan đơn hàng', index=False)
            df_items_excel.to_excel(writer, sheet_name='Chi tiết đối chiếu', index=False)
            df_unmatched_excel.to_excel(writer, sheet_name='Hóa đơn chưa khớp Sapo', index=False)
            
            # Excel formatting
            workbook = writer.book
            
            # Formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'fg_color': '#1F4E78',
                'font_color': '#FFFFFF',
                'border': 1
            })
            
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            
            border_format = workbook.add_format({'border': 1})
            money_format = workbook.add_format({'num_format': '#,##0', 'border': 1})
            
            # Format Sheet 1: Tổng quan đơn hàng
            worksheet1 = writer.sheets['Tổng quan đơn hàng']
            worksheet1.set_row(0, 26)
            for col_idx, col in enumerate(df_orders_excel.columns):
                worksheet1.write(0, col_idx, col, header_format)
                # auto-fit width
                max_len = max(df_orders_excel[col].astype(str).map(len).max(), len(col)) + 3
                worksheet1.set_column(col_idx, col_idx, min(max_len, 40), border_format)
                
            # Column-specific money formats
            worksheet1.set_column(5, 5, 18, money_format)
            worksheet1.set_column(6, 6, 18, money_format)
            
            # Add conditional formatting for sheet 1
            worksheet1.conditional_format(1, 7, len(df_orders_excel), 7, {
                'type': 'cell', 'operator': 'equal', 'value': '"Khớp"', 'format': green_format
            })
            worksheet1.conditional_format(1, 7, len(df_orders_excel), 7, {
                'type': 'cell', 'operator': 'equal', 'value': '"Lệch"', 'format': yellow_format
            })
            worksheet1.conditional_format(1, 7, len(df_orders_excel), 7, {
                'type': 'cell', 'operator': 'equal', 'value': '"Chưa xuất hóa đơn"', 'format': red_format
            })
            worksheet1.conditional_format(1, 7, len(df_orders_excel), 7, {
                'type': 'cell', 'operator': 'equal', 'value': '"Thiếu MST"', 'format': red_format
            })
            
            # Format Sheet 2: Chi tiết đối chiếu
            worksheet2 = writer.sheets['Chi tiết đối chiếu']
            worksheet2.set_row(0, 26)
            for col_idx, col in enumerate(df_items_excel.columns):
                worksheet2.write(0, col_idx, col, header_format)
                max_len = max(df_items_excel[col].astype(str).map(len).max(), len(col)) + 3
                worksheet2.set_column(col_idx, col_idx, min(max_len, 45), border_format)
                
            # Money format columns: Sapo price, Sapo total, Inv price, Inv total, diff amount
            worksheet2.set_column(4, 4, 20, money_format)
            worksheet2.set_column(5, 5, 20, money_format)
            worksheet2.set_column(9, 9, 20, money_format)
            worksheet2.set_column(10, 10, 20, money_format)
            worksheet2.set_column(13, 13, 20, money_format)
            
            worksheet2.conditional_format(1, 11, len(df_items_excel), 11, {
                'type': 'cell', 'operator': 'contains', 'value': 'Khớp', 'format': green_format
            })
            worksheet2.conditional_format(1, 11, len(df_items_excel), 11, {
                'type': 'cell', 'operator': 'contains', 'value': 'Lệch', 'format': yellow_format
            })
            worksheet2.conditional_format(1, 11, len(df_items_excel), 11, {
                'type': 'cell', 'operator': 'contains', 'value': 'Chưa', 'format': red_format
            })
            worksheet2.conditional_format(1, 11, len(df_items_excel), 11, {
                'type': 'cell', 'operator': 'contains', 'value': 'Thiếu', 'format': red_format
            })
            
            # Format Sheet 3: Hóa đơn chưa khớp Sapo
            worksheet3 = writer.sheets['Hóa đơn chưa khớp Sapo']
            worksheet3.set_row(0, 26)
            for col_idx, col in enumerate(df_unmatched_excel.columns):
                worksheet3.write(0, col_idx, col, header_format)
                max_len = max(df_unmatched_excel[col].astype(str).map(len).max(), len(col)) + 3
                worksheet3.set_column(col_idx, col_idx, min(max_len, 45), border_format)
                
            worksheet3.set_column(6, 6, 20, money_format)
            worksheet3.set_column(7, 7, 20, money_format)
            
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name='Bao_cao_doi_chieu_hoa_don_nhap_hang.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Lỗi xuất file Excel: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
