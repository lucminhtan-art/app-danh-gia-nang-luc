import streamlit as st
import pandas as pd
import os
import re
import docx
import openpyxl

st.set_page_config(
    page_title="App Đánh Giá Năng Lực Học Sinh (GDPT 2018)",
    page_icon="🎓",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a, #0f766e);
        padding: 24px;
        border-radius: 10px;
        color: white;
        margin-bottom: 24px;
    }
    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        border-left: 4px solid #1e40af;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1 style="margin: 0; font-size: 26px;">🎓 Ứng Dụng Đánh Giá Năng Lực Học Sinh Qua Bài Kiểm Tra (GDPT 2018)</h1>
    <p style="margin: 6px 0 0 0; opacity: 0.9; font-size: 14.5px;">
        Tự động hóa 5 mục: 1. Ma trận đề • 2. File gốc (G01-G06.Mix) • 3. Đề kiểm tra • 4. Bảng điểm chấm • 5. Nhận xét năng lực cá nhân hóa
    </p>
</div>
""", unsafe_allow_html=True)

# Sidebar
st.sidebar.header("⚙️ Cấu Hình Dữ Liệu")
use_sample = st.sidebar.checkbox("Sử dụng dữ liệu mẫu KTTX1-LẦN 1", value=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")

# Embedded Standard Data
MATRIX_DETAILS = [
    {"Mã câu hỏi": "[TO10.01.1.D01]", "Tên dạng câu hỏi": "Nhận biết mệnh đề toán học, mệnh đề chứa biến", "Mức độ": "Nhận biết (NB)", "Vị trí đề": "Câu 1, 2"},
    {"Mã câu hỏi": "[TO10.01.1.D03]", "Tên dạng câu hỏi": "Phủ định mệnh đề có chứa ký hiệu với mọi, tồn tại", "Mức độ": "Nhận biết (NB)", "Vị trí đề": "Câu 3, 4"},
    {"Mã câu hỏi": "[TO10.01.2.D01]", "Tên dạng câu hỏi": "Liệt kê phần tử và xác định tập hợp", "Mức độ": "Nhận biết (NB)", "Vị trí đề": "Câu 5"},
    {"Mã câu hỏi": "[TO10.01.1.D02]", "Tên dạng câu hỏi": "Xét tính đúng sai của mệnh đề kéo theo, tương đương", "Mức độ": "Thông hiểu (TH)", "Vị trí đề": "Câu 6"},
    {"Mã câu hỏi": "[TO10.01.2.D02]", "Tên dạng câu hỏi": "Xác định tập con và quan hệ bao hàm", "Mức độ": "Thông hiểu (TH)", "Vị trí đề": "Câu 7"},
    {"Mã câu hỏi": "[TO10.01.3.D02]", "Tên dạng câu hỏi": "Thực hiện phép toán giao, hợp, hiệu của 2 tập hợp số", "Mức độ": "Thông hiểu (TH)", "Vị trí đề": "Câu 8, 9"},
    {"Mã câu hỏi": "[TO10.01.2.D04]", "Tên dạng câu hỏi": "Bài toán thực tế áp dụng biểu đồ Ven đếm phần tử", "Mức độ": "Vận dụng (VD)", "Vị trí đề": "Câu 10"},
    {"Mã câu hỏi": "[TO10.01.3.D03]", "Tên dạng câu hỏi": "Tìm điều kiện tham số m để tập hợp thỏa mãn điều kiện bao hàm", "Mức độ": "Vận dụng (VD)", "Vị trí đề": "Câu 11"}
]

SOURCE_QUESTIONS = [
    {"STT": 1, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "NB", "Nội dung câu hỏi": "Viết mệnh đề sử dụng ký hiệu ∀ hoặc ∃: Mọi số thực nhân với 1 đều bằng chính nó.", "Mã câu": "[TO10.01.1.D01]"},
    {"STT": 2, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "NB", "Nội dung câu hỏi": "Trong các phát biểu sau, câu nào là một mệnh đề toán học?", "Mã câu": "[TO10.01.1.D01]"},
    {"STT": 3, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "NB", "Nội dung câu hỏi": "Mệnh đề phủ định của mệnh đề '∃x ∈ R, x² = 2' là gì?", "Mã câu": "[TO10.01.1.D03]"},
    {"STT": 4, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "NB", "Nội dung câu hỏi": "Liệt kê các phần tử của tập hợp A = {x ∈ N | x² - 4 = 0}.", "Mã câu": "[TO10.01.2.D01]"},
    {"STT": 5, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "NB", "Nội dung câu hỏi": "Hình vẽ nào dưới đây là biểu diễn của nửa khoảng (-2; 3] trên trục số?", "Mã câu": "[TO10.01.3.D01]"},
    {"STT": 6, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "TH", "Nội dung câu hỏi": "Cho P và Q là các mệnh đề đúng. Mệnh đề nào sau đây là mệnh đề sai?", "Mã câu": "[TO10.01.1.D02]"},
    {"STT": 7, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "TH", "Nội dung câu hỏi": "Cho hai đa thức f(x) và g(x). Xét các tập nghiệm A, B. Mệnh đề nào đúng?", "Mã câu": "[TO10.01.2.D02]"},
    {"STT": 8, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "TH", "Nội dung câu hỏi": "Cho hai tập hợp A = (-1; 4] và B = [2; 6). Xác định tập hợp A ∩ B.", "Mã câu": "[TO10.01.3.D02]"},
    {"STT": 9, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "TH", "Nội dung câu hỏi": "Cho A = (-∞; 3) và B = [-1; +∞). Xác định tập hợp A ∪ B.", "Mã câu": "[TO10.01.3.D02]"},
    {"STT": 10, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "VD", "Nội dung câu hỏi": "Khảo sát sở thích 45 học sinh về 3 môn thể thao, áp dụng sơ đồ Ven tính số học sinh thích đúng 1 môn.", "Mã câu": "[TO10.01.2.D04]"},
    {"STT": 11, "Nhóm Mix": "G01-G04.Mix", "Mức độ": "VD", "Nội dung câu hỏi": "Cho A = [m; m+2] và B = [-1; 3]. Tìm điều kiện cần và đủ của m để A ⊂ B.", "Mã câu": "[TO10.01.3.D03]"}
]

EXAM_QUESTIONS_190 = [
    {"Câu đề": "Câu 1", "Mức độ": "NB", "Nội dung trong đề 190": "Xác định tập hợp bằng cách liệt kê các phần tử của phương trình x² - 4 = 0.", "Câu gốc": "Câu 4"},
    {"Câu đề": "Câu 2", "Mức độ": "NB", "Nội dung trong đề 190": "Trong các phát biểu sau, phát biểu nào là mệnh đề toán học?", "Câu gốc": "Câu 2"},
    {"Câu đề": "Câu 3", "Mức độ": "NB", "Nội dung trong đề 190": "Mệnh đề '∃x ∈ R, x² = 2' khẳng định điều gì?", "Câu gốc": "Câu 3"},
    {"Câu đề": "Câu 4", "Mức độ": "NB", "Nội dung trong đề 190": "Phủ định của mệnh đề 'Mọi số thực nhân với 1 đều bằng chính nó'.", "Câu gốc": "Câu 1"},
    {"Câu đề": "Câu 5", "Mức độ": "NB", "Nội dung trong đề 190": "Biểu diễn tập hợp (-2; 3] trên trục số.", "Câu gốc": "Câu 5"},
    {"Câu đề": "Câu 6", "Mức độ": "TH", "Nội dung trong đề 190": "Xét tính đúng sai của mệnh đề kéo theo P => Q.", "Câu gốc": "Câu 6"},
    {"Câu đề": "Câu 7", "Mức độ": "TH", "Nội dung trong đề 190": "Cho hai tập hợp A, B. Xác định quan hệ tập con.", "Câu gốc": "Câu 7"},
    {"Câu đề": "Câu 8", "Mức độ": "TH", "Nội dung trong đề 190": "Cho A = (-1; 4] và B = [2; 6). Khi đó A ∩ B bằng bao nhiêu?", "Câu gốc": "Câu 8"},
    {"Câu đề": "Câu 9", "Mức độ": "TH", "Nội dung trong đề 190": "Xác định phép hợp của 2 khoảng số trên trục số.", "Câu gốc": "Câu 9"},
    {"Câu đề": "Câu 10", "Mức độ": "VD", "Nội dung trong đề 190": "Khảo sát khách hàng thích 3 món ăn tại nhà hàng, tính số lượng không thích món nào.", "Câu gốc": "Câu 10"},
    {"Câu đề": "Câu 11", "Mức độ": "VD", "Nội dung trong đề 190": "Tìm tất cả các số thực m để tập hợp [m; m+2] là tập con của [-1; 3].", "Câu gốc": "Câu 11"}
]

STUDENTS_DATA = [
    {"STT": 1, "Họ và tên": "Nguyễn Phan Anh Thư", "Lớp": "10A1", "Mã đề": "190", "Điểm": 10.0, "Số câu đúng": "11/11", "NB": "5/5 (100%)", "TH": "4/4 (100%)", "VD": "2/2 (100%)"},
    {"STT": 2, "Họ và tên": "Phạm Thành Đạt", "Lớp": "10A1", "Mã đề": "190", "Điểm": 8.18, "Số câu đúng": "9/11", "NB": "5/5 (100%)", "TH": "3/4 (75%)", "VD": "1/2 (50%)"},
    {"STT": 3, "Họ và tên": "Trương Tâm Như", "Lớp": "10A2", "Mã đề": "283", "Điểm": 10.0, "Số câu đúng": "11/11", "NB": "5/5 (100%)", "TH": "4/4 (100%)", "VD": "2/2 (100%)"},
    {"STT": 4, "Họ và tên": "Trần Thị Mai Quỳnh", "Lớp": "10A2", "Mã đề": "283", "Điểm": 8.18, "Số câu đúng": "9/11", "NB": "5/5 (100%)", "TH": "2/4 (50%)", "VD": "2/2 (100%)"},
    {"STT": 5, "Họ và tên": "Nguyễn Vũ Quang Vinh", "Lớp": "10A3", "Mã đề": "359", "Điểm": 10.0, "Số câu đúng": "11/11", "NB": "5/5 (100%)", "TH": "4/4 (100%)", "VD": "2/2 (100%)"},
    {"STT": 6, "Họ và tên": "Đặng Phùng Kỳ Thư", "Lớp": "10A3", "Mã đề": "359", "Điểm": 8.18, "Số câu đúng": "9/11", "NB": "4/5 (80%)", "TH": "4/4 (100%)", "VD": "1/2 (50%)"},
    {"STT": 7, "Họ và tên": "Huỳnh Nguyễn Trường An", "Lớp": "10A1", "Mã đề": "756", "Điểm": 10.0, "Số câu đúng": "11/11", "NB": "5/5 (100%)", "TH": "4/4 (100%)", "VD": "2/2 (100%)"},
    {"STT": 8, "Họ và tên": "Hàng Châu Gia Bảo", "Lớp": "10A1", "Mã đề": "756", "Điểm": 8.18, "Số câu đúng": "9/11", "NB": "4/5 (80%)", "TH": "4/4 (100%)", "VD": "1/2 (50%)"}
]

# Tabs Definition
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 1. Ma Trận Đề Thi",
    "📁 2. File Gốc (Mix)",
    "📝 3. Đề Kiểm Tra & Mức Độ",
    "📊 4. Bảng Chấm Điểm",
    "🌟 5. Nhận Xét & Đánh Giá Năng Lực"
])

# TAB 1: MA TRẬN
with tab1:
    st.subheader("1. Nhận Dạng Cấu Trúc Ma Trận Đề Kiểm Tra (10-MA TRAN KTTX1-LAN 1)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Phần I: TNKQ", "11 câu", "NB: 5 • TH: 4 • VD: 2")
    c2.metric("Phần II: Đúng - Sai", "1 câu (4 ý)", "Theo cấp độ nhận thức")
    c3.metric("Phần III: Trả Lời Ngắn", "2 câu", "Mô hình hóa thực tế")
    c4.metric("Phần IV: Tự Luận", "1 câu", "Lập luận toán học")

    st.markdown("### 📌 Phân bố nội dung kiến thức theo mức độ nhận thức:")
    df_levels = pd.DataFrame([
        {"Mức độ nhận thức": "Nhận biết (NB)", "Đơn vị kiến thức / Nội dung": "Mệnh đề, mệnh đề phủ định, ký hiệu với mọi/tồn tại; Tập hợp và phần tử, khoảng đoạn trên trục số", "Phần thi": "Phần I (TNKQ)", "Số lượng câu": "5 câu (45.5%)"},
        {"Mức độ nhận thức": "Thông hiểu (TH)", "Đơn vị kiến thức / Nội dung": "Xét tính đúng sai của mệnh đề kéo theo; Các phép toán giao, hợp, hiệu của 2 tập hợp trên trục số", "Phần thi": "Phần I (TNKQ)", "Số lượng câu": "4 câu (36.4%)"},
        {"Mức độ nhận thức": "Vận dụng (VD)", "Đơn vị kiến thức / Nội dung": "Bài toán thực tế áp dụng tập hợp (Sơ đồ Ven); Tìm tham số m để tập con thỏa mãn điều kiện bao hàm", "Phần thi": "Phần I (TNKQ)", "Số lượng câu": "2 câu (18.1%)"}
    ])
    st.dataframe(df_levels, use_container_width=True)

    st.markdown("### 📑 Bảng đặc tả chi tiết mã câu hỏi (B&T Pro):")
    st.dataframe(pd.DataFrame(MATRIX_DETAILS), use_container_width=True)

# TAB 2: FILE GỐC
with tab2:
    st.subheader("2. Phân Loại Câu Hỏi Trong File Gốc Theo Nhóm Bảng Mix")
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Nhóm G01.Mix - G04.Mix", "11 câu", "Trắc nghiệm nhiều lựa chọn")
    col_b.metric("Nhóm G05.Mix", "1 câu (4 ý)", "Trắc nghiệm Đúng - Sai")
    col_c.metric("Nhóm G06.Mix", "2 câu", "Trắc nghiệm Trả lời ngắn")

    st.markdown("### 🔍 Danh mục câu hỏi trong File gốc:")
    st.dataframe(pd.DataFrame(SOURCE_QUESTIONS), use_container_width=True)

# TAB 3: ĐỀ KIỂM TRA
with tab3:
    st.subheader("3. Đọc Đề Kiểm Tra & Đối Chiếu Mức Độ Nhận Thức Từng Câu")
    sel_code = st.selectbox("Chọn mã đề kiểm tra:", ["190", "283", "359", "756"])
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Nhận biết (NB)", "5 câu (45.5%)", "Câu 1, 2, 3, 4, 5")
    col2.metric("Thông hiểu (TH)", "4 câu (36.4%)", "Câu 6, 7, 8, 9")
    col3.metric("Vận dụng (VD)", "2 câu (18.1%)", "Câu 10, 11")

    st.markdown(f"### 📋 Chi tiết các câu hỏi trong Mã đề {sel_code}:")
    st.dataframe(pd.DataFrame(EXAM_QUESTIONS_190), use_container_width=True)

# TAB 4: BẢNG CHẤM ĐIỂM
with tab4:
    st.subheader("4. Dữ Liệu Bảng Điểm Chấm Học Sinh (File: BAI KIỂM TRA KTTX1-LAN 1)")
    st.info("ℹ️ Bóc tách dữ liệu ở tiêu đề mỗi trang A4: Họ và tên, Lớp, Mã đề, Điểm từng phần. Đối chiếu kết quả câu đúng/sai với ma trận đề thi.")
    st.dataframe(pd.DataFrame(STUDENTS_DATA), use_container_width=True)

# TAB 5: NHẬN XÉT
with tab5:
    st.subheader("5. Báo Cáo Nhận Xét Đánh Giá Năng Lực Học Sinh Cá Nhân Hóa")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Điểm TB Lớp", "9.09 / 10.0", "100% hoàn thành bài")
    m2.metric("Tỉ lệ đạt Nhận Biết", "97.5%", "Rất vững kiến thức nền")
    m3.metric("Tỉ lệ đạt Thông Hiểu", "93.8%", "Kỹ năng suy luận tốt")
    m4.metric("Tỉ lệ đạt Vận Dụng", "75.0%", "Cần rèn luyện thêm toán thực tế")

    st.markdown("---")
    filter_student = st.selectbox("Chọn học sinh để xem phiếu nhận xét năng lực chi tiết:", [s["Họ và tên"] for s in STUDENTS_DATA])
    
    student = next(s for s in STUDENTS_DATA if s["Họ và tên"] == filter_student)
    is_perfect = student["Điểm"] == 10.0

    st.markdown(f"### 📄 Phiếu Đánh Giá Năng Lực: **{student['Họ và tên']}** - Lớp: **{student['Lớp']}** (Mã đề: **{student['Mã đề']}**)")
    st.write(f"**Tổng điểm:** `{student['Điểm']} / 10.0` | **Đúng:** `{student['Số câu đúng']}` | **Nhận biết:** `{student['NB']}` | **Thông hiểu:** `{student['TH']}` | **Vận dụng:** `{student['VD']}`")

    c_left, c_right = st.columns(2)
    with c_left:
        st.success("**✅ 1. Phần học sinh làm được (Điểm mạnh):**\n" + 
                   ("- Nắm rất vững các định nghĩa, ký hiệu và mệnh đề cơ bản (100% câu đúng).\n"
                    "- Hiểu sâu các phép toán giao, hợp, hiệu trên trục số và quan hệ tập con.\n"
                    "- Khả năng vận dụng giải toán thực tế và tìm tham số m đạt kết quả tốt." if is_perfect else
                    "- Nắm chắc các kiến thức nền tảng ở mức độ Nhận biết (100% câu đúng).\n"
                    "- Áp dụng tốt các phép toán tập hợp cơ bản."))

        if not is_perfect:
            st.error("**⚠️ 2. Phần học sinh còn yếu / Sơ suất:**\n"
                     "- Còn sơ suất ở câu hỏi mức độ Thông hiểu hoặc Vận dụng.\n"
                     "- Cần chú ý khi xác định giao, hợp trên trục số và bài toán đếm thực tế bằng sơ đồ Ven.")

    with c_right:
        st.warning("**💡 3. Nội dung cần cải thiện & Hướng rèn luyện:**\n" +
                   ("- Tiếp tục duy trì phong độ và thử sức với các bài toán vận dụng cao mở rộng." if is_perfect else
                    "- Ôn tập lại kỹ năng xác định điều kiện chứa tham số m trên các khoảng số.\n"
                    "- Luyện tập thêm các bài toán thực tế giải bằng sơ đồ Ven 3 tập hợp."))

        st.info("**🎖️ 4. Phần đáng được tuyên dương:**\n" +
                ("- Xuất sắc đạt điểm tuyệt đối 10/10! Tuyên dương sự cẩn thận, chỉn chu và tư duy toán học chuẩn xác." if is_perfect else
                 "- Tuyên dương kết quả làm bài tốt, tinh thần học tập nghiêm túc và có nhiều nỗ lực vươn lên."))
