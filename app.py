import io
import os
from deep_translator import GoogleTranslator
import streamlit as st
from docxtpl import DocxTemplate

# 1. 动态获取当前脚本所在目录，确保在 Streamlit Cloud 云端或本地都能精准找到模板文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "船海专业申请表(1).docx")

# 网页基本属性设置
st.set_page_config(page_title="船海专业申请表填写", layout="centered")
st.title("📄 船海专业申请表自动填写系统")


# 2. 自动翻译工具函数（将中文翻译为英文）
def translate_to_en(text):
    if not text or not text.strip():
        return ""
    try:
        # 自动识别源语言并翻译为英文
        translated = GoogleTranslator(source="auto", target="en").translate(
            text
        )
        return translated
    except Exception:
        # 如果网络波动导致翻译服务暂不可用，保底返回原文本，防止程序崩溃
        return text


# 检查模板文件是否存在
if not os.path.exists(TEMPLATE_PATH):
    st.error(
        f"❌ 未找到模板文件！请确保 `船海专业申请表(1).docx` 与 `app.py` 放在同一目录下。\n当前寻找路径：`{TEMPLATE_PATH}`"
    )
else:
    st.success(
        "✅ 模板已成功加载！下方支持直接输入中文，系统将自动翻译并格式化为英文："
    )

    # 3. 前端输入表单界面
    with st.form("application_form"):
        st.subheader("1. 姓名信息 (Name)")
        col1, col2 = st.columns(2)
        with col1:
            last_name = st.text_input(
                "Family Name (姓)", placeholder="可填中文，如: 张"
            )
        with col2:
            first_name = st.text_input(
                "First Name (名)", placeholder="可填中文，如: 阳晨"
            )

        st.subheader("2. 性别与出生日期 (Sex & Date of Birth)")
        sex = st.radio("Sex (性别)", ["Male (男)", "Female (女)"], horizontal=True)

        col3, col4, col5 = st.columns(3)
        with col3:
            day_input = st.text_input(
                "Day (出生日，2位)", max_chars=2, placeholder="05"
            )
        with col4:
            month_input = st.text_input(
                "Month (出生月，2位)", max_chars=2, placeholder="08"
            )
        with col5:
            year_input = st.text_input(
                "Year (出生年，4位)", max_chars=4, placeholder="2002"
            )

        st.subheader("3. 证件与联系地址 (Passport & Address)")
        passport_no = st.text_input(
            "Passport No. (护照号码)", placeholder="例: E12345678"
        )
        address_street = st.text_input(
            "Number and street name (门牌与街道地址)",
            placeholder="可填中文，如: 中山路100号",
        )

        col6, col7 = st.columns(2)
        with col6:
            city = st.text_input("City (城市)", placeholder="可填中文，如: 南京")
        with col7:
            postal_code = st.text_input("Postal code (邮编)")

        col8, col9 = st.columns(2)
        with col8:
            phone = st.text_input(
                "Telephone (电话)", placeholder="+86-13800000000"
            )
        with col9:
            email = st.text_input("Email (邮箱)", placeholder="example@gmail.com")

        submitted = st.form_submit_button("🚀 自动翻译并导出 Word")

    # 4. 后端提交处理逻辑
    if submitted:
        if not last_name or not first_name:
            st.warning("⚠️ 请输入完整的 Family Name 和 First Name！")
        else:
            try:
                with st.spinner("正在处理文本、自动翻译并填充 Word..."):

                    # 4.1 姓名翻译与清洗：去除 - 和空格，强制只有首字母大写（例如: 阳晨 -> Yangchen）
                    last_name_en = (
                        translate_to_en(last_name)
                        .replace("-", "")
                        .replace(" ", "")
                        .capitalize()
                    )
                    first_name_en = (
                        translate_to_en(first_name)
                        .replace("-", "")
                        .replace(" ", "")
                        .capitalize()
                    )

                    # 地址与城市自动翻译为英文
                    address_street_en = translate_to_en(address_street)
                    city_en = translate_to_en(city)

                    # 4.2 性别打勾逻辑（使用标准 Unicode 打勾符号 ✓，跨设备不会变字符 P）
                    male_check = "✓" if "Male" in sex else ""
                    female_check = "✓" if "Female" in sex else ""

                    # 4.3 日期补零与拆分逻辑
                    day_str = (
                        day_input.zfill(2)
                        if day_input.isdigit()
                        else day_input.ljust(2)
                    )
                    month_str = (
                        month_input.zfill(2)
                        if month_input.isdigit()
                        else month_input.ljust(2)
                    )

                    day_1, day_2 = (
                        (day_str[0], day_str[1])
                        if len(day_str) >= 2
                        else ("", "")
                    )
                    month_1, month_2 = (
                        (month_str[0], month_str[1])
                        if len(month_str) >= 2
                        else ("", "")
                    )

                    # 年份（4位）拆分：前2位存入 year_1（如 20），后2位存入 year_2（如 02）
                    year_str = (
                        year_input.zfill(4)
                        if year_input.isdigit()
                        else year_input.ljust(4)
                    )
                    year_1 = year_str[:2]
                    year_2 = year_str[2:]

                    # 4.4 读取 Word 模板并注入数据
                    doc = DocxTemplate(TEMPLATE_PATH)

                    context = {
                        "last_name": last_name_en,
                        "first_name": first_name_en,
                        "male_check": male_check,
                        "female_check": female_check,
                        "day_1": day_1,
                        "day_2": day_2,
                        "month_1": month_1,
                        "month_2": month_2,
                        "year_1": year_1,
                        "year_2": year_2,
                        "passport_no": passport_no,
                        "address_street": address_street_en,
                        "city": city_en,
                        "postal_code": postal_code,
                        "phone": phone,
                        "email": email,
                    }

                    doc.render(context)

                    # 4.5 保存至二进制流供网页下载
                    file_stream = io.BytesIO()
                    doc.save(file_stream)
                    file_stream.seek(0)

                st.success("🎉 Word 申请表生成并翻译成功！")
                st.download_button(
                    label="📥 点击下载填好的英文申请表",
                    data=file_stream,
                    file_name=f"Application_{last_name_en}_{first_name_en}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"❌ 生成失败，错误信息: {e}")
