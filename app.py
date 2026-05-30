import streamlit as st
import pandas as pd
import datetime
from docxtpl import DocxTemplate
import io
import zipfile

st.title("Phần Mềm Cấp Giấy Xác Nhận Hoàn Thành Khóa Học")

st.subheader("1. Tải lên danh sách học viên")
uploaded_file = st.file_uploader("Chọn file Excel (.xlsx, .xls) hoặc CSV (.csv)", type=["xlsx", "xls", "csv"])

st.subheader("2. Thông tin chung")
ngay_ky_input = st.date_input("Ngày ký xác nhận:", datetime.date.today())

def format_date(val):
    if pd.isna(val) or val == "" or str(val).lower() == 'nan':
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%d/%m/%Y")
    val_str = str(val).split()[0]
    try:
         dt = datetime.datetime.strptime(val_str, "%Y-%m-%d")
         return dt.strftime("%d/%m/%Y")
    except ValueError:
         pass
    return str(val)

def clean_val(val):
    """Hàm làm sạch dữ liệu: xóa chữ 'nan', loại bỏ đuôi .0 nếu bị ép kiểu số"""
    if pd.isna(val) or str(val).lower() == 'nan':
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

if uploaded_file is not None:
    try:
        # Ép kiểu dtype=str để giữ nguyên vẹn số 0 ở đầu CCCD
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, dtype=str) 
        else:
            df = pd.read_excel(uploaded_file, dtype=str)
            
        st.write("📋 **Bản xem trước dữ liệu:**")
        st.dataframe(df.head())
        
        # Chuẩn hóa tên cột trong Excel của anh để dễ so sánh (chuyển về chữ thường)
        cols = {col.strip().lower(): col for col in df.columns}
        
        # Hàm thông minh tự động dò tìm các tên cột khả thi
        def get_val(row, possible_names):
            for name in possible_names:
                if name.lower() in cols:
                    return clean_val(row[cols[name.lower()]])
            return ""

        if st.button("🚀 Bắt Đầu Tạo File Word"):
            zip_buffer = io.BytesIO()
            success_count = 0
            
            ngay = ngay_ky_input.strftime("%d")
            thang = ngay_ky_input.strftime("%m")
            nam = ngay_ky_input.strftime("%Y")
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in df.iterrows():
                    
                    # Ưu tiên tìm cột Họ và tên
                    ho_ten = get_val(row, ['Họ và tên', 'Họ tên', 'Tên học viên', 'Tên'])
                    if not ho_ten:
                        continue
                    
                    try:
                        doc = DocxTemplate("template.docx")
                    except Exception as e:
                        st.error("❌ Không tìm thấy file 'template.docx' nằm cùng thư mục!")
                        st.stop()
                        
                    # Ánh xạ dữ liệu đa năng: tìm đúng tên cột anh đang dùng
                    context = {
                        'so_hoan_thanh': get_val(row, ['Số chứng chỉ', 'Số căn cước', 'Số hoàn thành']), 
                        'ho_ten': ho_ten.upper(),
                        'ngay_sinh': format_date(get_val(row, ['Ngày sinh'])),
                        'cccd': get_val(row, ['Số căn cước', 'CCCD', 'Số CCCD', 'CMND', 'Số CMND']),
                        'hang_xe': get_val(row, ['Hạng xe', 'Hạng']),
                        'so_ngay': get_val(row, ['Số ngày học', 'Số ngày', 'Thời gian học']),
                        'tu_ngay': format_date(get_val(row, ['Từ ngày'])),
                        'den_ngay': format_date(get_val(row, ['Đến ngày'])),
                        'ngay_ky': ngay,
                        'thang': thang,
                        'nam': nam
                    }
                    
                    doc.render(context)
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    
                    file_name = f"Giay_Xac_Nhan_{ho_ten.replace(' ', '_')}.docx"
                    zip_file.writestr(file_name, doc_buffer.getvalue())
                    success_count += 1
                    
            if success_count > 0:
                st.success(f"✅ Đã tạo thành công {success_count} giấy xác nhận!")
                st.download_button(
                    label="📥 Tải File ZIP chứa toàn bộ Giấy Xác Nhận",
                    data=zip_buffer.getvalue(),
                    file_name=f"Danh_Sach_{nam}{thang}{ngay}.zip",
                    mime="application/zip"
                )
            else:
                st.error("❌ Không tìm thấy dữ liệu. Kiểm tra lại xem có cột 'Họ và tên' chưa.")
            
    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra: {e}")
