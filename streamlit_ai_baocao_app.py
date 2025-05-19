
import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import openai

st.set_page_config(page_title="AI Agent Báo Cáo Đào Tạo", layout="centered")

st.title("📊 AI Agent – Tổng Hợp Báo Cáo Đào Tạo Tự Động")
st.markdown("Tải lên 3 file Excel và 1 file Word template để tạo báo cáo tự động.")

# Upload các file
hoc_vien_file = st.file_uploader("📘 1. Danh sách học viên (HocVien.xlsx)", type=["xlsx"])
diem_danh_file = st.file_uploader("📝 2. Danh sách điểm danh (DiemDanh.xlsx)", type=["xlsx"])
ket_qua_file = st.file_uploader("📈 3. Kết quả cuối khóa (KetQua.xlsx)", type=["xlsx"])
template_file = st.file_uploader("📄 4. File mẫu báo cáo (Word Template)", type=["docx"])

# OpenAI API key
openai_api_key = st.text_input("🔑 Nhập OpenAI API Key để tạo nhận xét AI", type="password")

if st.button("🚀 Tạo báo cáo") and all([hoc_vien_file, diem_danh_file, ket_qua_file, template_file]):
    # Đọc dữ liệu
    df_hoc_vien = pd.read_excel(hoc_vien_file)
    df_diem_danh = pd.read_excel(diem_danh_file)
    df_ket_qua = pd.read_excel(ket_qua_file)

    # Xử lý dữ liệu
    total_students = len(df_hoc_vien)
    df_diem_danh['Số buổi tham gia'] = df_diem_danh.iloc[:, 1:-1].apply(lambda x: x.eq('X').sum(), axis=1)
    attendance_rate = round(df_diem_danh['Số buổi tham gia'].mean() / (df_diem_danh.shape[1] - 2) * 100, 2)
    completed_students = df_ket_qua[df_ket_qua['Tổng điểm'] >= 7]
    completion_rate = round(len(completed_students) / total_students * 100, 2)
    top_students = df_ket_qua.sort_values(by='Tổng điểm', ascending=False).head(3)

    # Nhận xét từ GPT (nếu có API)
    if openai_api_key:
        openai.api_key = openai_api_key
        prompt = f"Viết 3 dòng nhận xét tổng quan về khóa học có {total_students} học viên, tỉ lệ hoàn thành {completion_rate}%, điểm danh trung bình {attendance_rate}%, 3 học viên cao điểm nhất có điểm lần lượt là {top_students['Tổng điểm'].tolist()}."
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}]
            )
            ai_comments = response.choices[0].message['content'].split("\n")
        except:
            ai_comments = ["Không thể kết nối GPT.", "", ""]
    else:
        ai_comments = ["Chưa nhập API Key.", "", ""]

    # Tạo file Word báo cáo
    template_doc = Document(template_file)
    template_doc.add_paragraph()
    template_doc.add_paragraph(f"Khóa học: Ứng dụng AI vào công việc tại Viettel")
    template_doc.add_paragraph("Thời gian: 15–17/05/2025")
    template_doc.add_paragraph(f"Số học viên: {total_students} người")
    template_doc.add_paragraph(f"Tỷ lệ hoàn thành: {completion_rate}%")
    template_doc.add_paragraph(f"Tỷ lệ đạt loại Giỏi – Xuất sắc: {len(df_ket_qua[df_ket_qua['Xếp loại'].isin(['Giỏi', 'Xuất sắc'])]) / total_students * 100:.1f}%")

    template_doc.add_paragraph("\nDanh sách học viên tiêu biểu:")
    for _, row in top_students.iterrows():
        line = f"- {row['Họ tên']} – {row['Tổng điểm']} điểm – {row['Xếp loại']}"
        template_doc.add_paragraph(line)

    template_doc.add_paragraph("\nThống kê điểm danh:")
    template_doc.add_paragraph(f"- Trung bình mỗi học viên tham gia {attendance_rate}% số buổi")
    vang_phep = df_diem_danh['Ghi chú'].str.contains("có phép", case=False).sum()
    template_doc.add_paragraph(f"- Số trường hợp vắng mặt có phép: {vang_phep}")

    template_doc.add_paragraph("\nNhận xét tổng quan của hệ thống AI:")
    for line in ai_comments:
        template_doc.add_paragraph(f"- {line}")

    output_stream = BytesIO()
    template_doc.save(output_stream)
    output_stream.seek(0)

    st.success("🎉 Báo cáo đã được tạo thành công!")
    st.download_button(label="📥 Tải file báo cáo Word",
                       data=output_stream,
                       file_name="BaoCaoTongHop.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
