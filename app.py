import io
import os
from deep_translator import GoogleTranslator
from pypinyin import lazy_pinyin
import streamlit as st
from docxtpl import DocxTemplate

# 1. 动态获取当前文件目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "船海专业申请表(1).docx")

# ✨ 网页标题更新为：乌克兰留学申请表
st.set_page_config(page_title="乌克兰留学申请表", layout="centered")
st.title("📄 乌克兰留学申请表")


# 2. 姓名转拼音函数：中文转拼音、去除 - 和空格，首字母大写
def name_to_pinyin(text):
    if not text or not text.strip():
        return ""
    py_list = lazy_pinyin(text)
    combined = "".join(py_list).replace("-", "").replace(" ", "")
    return combined.capitalize()


# 3. 常规中文自动翻译函数
def translate_to_en(text):
    if not text or not text.strip():
        return ""
    try:
        translated = GoogleTranslator(source="auto", target="en").translate(
            text
        )
        return translated
    except Exception:
        return text


# 检查模板文件是否存在
if not os.path.exists(TEMPLATE_PATH):
    st.error(
        f"❌ 未找到模板文件！请确保 `船海专业申请表(1).docx` 与 `app.py` 在同一目录下。"
    )
else:
    st.success("✅ 模板加载成功，请填写下方信息：")

    # 前端输入表单
    with st.form("application_form"):
        st.subheader("1. 姓名与专业信息 (Name & Major)")
        col1, col2 = st.columns(2)
        with col1:
            last_name = st.text_input(
                "Family Name (姓)", placeholder="支持中文拼音转换，如: 张"
            )
        with col2:
            first_name = st.text_input(
                "First Name (名)", placeholder="支持中文拼音转换，如: 阳晨"
            )

        # 专业选择下拉菜单
        major_cn = st.selectbox(
            "Intended Major Field (申请专业)",
            ["机器人工程", "船舶与海洋工程"],
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
            placeholder="支持中文翻译，如: 中山路100号",
        )

        col6, col7 = st.columns(2)
        with col6:
            city = st.text_input("City (城市)", placeholder="如: 南京")
        with col7:
            postal_code = st.text_input("Postal code (邮编)")

        col8, col9 = st.columns(2)
        with col8:
            phone = st.text_input(
                "Telephone (电话)", placeholder="+86-13800000000"
            )
        with col9:
            email = st.text_input("Email (邮箱)", placeholder="example@gmail.com")

        submitted = st.form_submit_button("🚀 生成并导出申请表")

    # 后端逻辑处理
    if submitted:
        if not last_name or not first_name:
            st.warning("⚠️ 请输入完整的 Family Name 和 First Name！")
        else:
            try:
                with st.spinner("正在转换为拼音与英文并填充 Word..."):

                    # 1. 姓名转换：中文拼音化，仅首字母大写
                    last_name_en = name_to_pinyin(last_name)
                    first_name_en = name_to_pinyin(first_name)

                    # 2. 专业中文映射到指定英文
                    major_map = {
                        "机器人工程": "Automation and Robot Engineering",
                        "船舶与海洋工程": "Ships and Ocean Engineering",
                    }
                    major_en = major_map.get(
                        major_cn, "Ships and Ocean Engineering"
                    )

                    # 3. 地址与城市英译
                    address_street_en = translate_to_en(address_street)
                    city_en = translate_to_en(city)

                    # 4. 性别打勾逻辑
                    male_check = "✓" if "Male" in sex else ""
                    female_check = "✓" if "Female" in sex else ""

                    # 5. 日期补齐与拆分
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

                    year_str = (
                        year_input.zfill(4)
                        if year_input.isdigit()
                        else year_input.ljust(4)
                    )
                    year_1 = year_str[:2]
                    year_2 = year_str[2:]

                    # 6. 读取 Word 模板并写入数据
                    doc = DocxTemplate(TEMPLATE_PATH)

                    context = {
                        "last_name": last_name_en,
                        "first_name": first_name_en,
                        "major": major_en,
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

                    file_stream = io.BytesIO()
                    doc.save(file_stream)
                    file_stream.seek(0)

                st.success("🎉 Word 申请表生成成功！")
                st.download_button(
                    label="📥 点击下载填好的申请表",
                    data=file_stream,
                    file_name=f"Application_{last_name_en}_{first_name_en}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"❌ 生成失败，错误信息: {e}")
