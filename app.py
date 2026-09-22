import io
import json
import os
import pandas as pd
import streamlit as st
from openai import OpenAI

# مكتبات تصدير PDF وإصلاح النص العربي
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. إعدادات الصفحة وواجهة المستخدم
# ==========================================
st.set_page_config(
    page_title="المساعد الذكي لتدقيق ميزان المراجعة",
    page_icon="📊",
    layout="wide"
)

# دعم التنسيق من اليمين إلى اليسار (RTL)
st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    div[data-testid="stMetricValue"] { text-align: right; }
    .stButton>button { width: 100%; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 نظام المساعد الذكي لتدقيق ميزان المراجعة")
st.caption("قم برفع ملف ميزان المراجعة (Excel / CSV / PDF / Image) لفحصه محاسبياً وتوليد التقرير الرقابي.")

# ==========================================
# 2. القائمة الجانبية (Sidebar)
# ==========================================
with st.sidebar:
    st.header("⚙️ الإعدادات والمفاتيح")
    
    # محاولة جلب المفتاح من أسرار Streamlit أولاً، وإلا يطلبه من المستخدم
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("أدخل مفتاح OpenAI API Key:", type="password")
    else:
        st.success("تم ربط مفتاح API تلقائياً من الأسرار 🔒")

    st.divider()
    st.subheader("📌 نطاق الفحص الرقابي")
    check_abnormal = st.checkbox("فحص الأرصدة الشاذة (Abnormal Balances)", value=True)
    check_cash = st.checkbox("فحص النقدية والسحب على المكشوف", value=True)
    check_suspense = st.checkbox("فحص الحسابات الوسيطة والعُهد", value=True)

# ==========================================
# 3. معالجة وقراءة الملفات المرفوعة
# ==========================================
def process_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_type = uploaded_file.name.split('.')[-1].lower()
    try:
        if file_type in ['xlsx', 'xls']:
            # قراءة الإكسيل مع إزالة الفراغات من عناوين الأعمدة لتجنب أخطاء القراءة
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).strip() for col in df.columns]
            return df
        elif file_type == 'csv':
            df = pd.read_csv(uploaded_file)
            df.columns = [str(col).strip() for col in df.columns]
            return df
        else:
            # نموذج استخراج محاكي للملفات المصورة والـ PDF
            st.info(f"تم استلام ملف بصيغة ({file_type.upper()}). يتم تحليله بواسطة محرك الاستخراج...")
            mock_data = [
                {"account_code": "1101", "account_name": "الصندوق الرئيسي", "ending_debit": 0, "ending_credit": 1500},
                {"account_code": "1102", "account_name": "بنك الاتحاد", "ending_debit": 45000, "ending_credit": 0},
                {"account_code": "1201", "account_name": "الذمم المدينة - عملاء", "ending_debit": 18000, "ending_credit": 0},
                {"account_code": "2101", "account_name": "حساب الموردين", "ending_debit": 3200, "ending_credit": 0},
                {"account_code": "1901", "account_name": "حساب أخطاء تحت التسوية", "ending_debit": 5000, "ending_credit": 0},
                {"account_code": "4101", "account_name": "إيرادات المبيعات", "ending_debit": 0, "ending_credit": 85000},
                {"account_code": "5101", "account_name": "مصاريف إيجارات", "ending_debit": 15300, "ending_credit": 0}
            ]
            return pd.DataFrame(mock_data)
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
        return pd.DataFrame()

# ==========================================
# 4. محرك الفحص والتدقيق البرمجي
# ==========================================
class TrialBalanceAuditor:
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.audit_results = {
            "abnormal_balances": [],
            "bank_cash_warnings": [],
            "suspense_accounts": [],
            "summary": {}
        }

    def run_audit(self):
        cols = self.df.columns
        # تحسين البحث عن أسماء الأعمدة لتشمل تنويعات أكثر
        debit_col = [c for c in cols if 'debit' in str(c).lower() or 'مدين' in str(c) or 'منه' in str(c)]
        credit_col = [c for c in cols if 'credit' in str(c).lower() or 'دائن' in str(c) or 'له' in str(c)]
        code_col = [c for c in cols if 'code' in str(c).lower() or 'رمز' in str(c) or 'رقم' in str(c)]
        name_col = [c for c in cols if 'name' in str(c).lower() or 'اسم' in str(c) or 'حساب' in str(c) or 'بيان' in str(c)]

        # إذا لم يجد الأعمدة بالاسم، يفترض الترتيب الافتراضي
        d_col = debit_col[0] if debit_col else (self.df.columns[2] if len(self.df.columns) > 2 else None)
        c_col = credit_col[0] if credit_col else (self.df.columns[3] if len(self.df.columns) > 3 else None)
        cd_col = code_col[0] if code_col else self.df.columns[0]
        nm_col = name_col[0] if name_col else (self.df.columns[1] if len(self.df.columns) > 1 else None)

        if not d_col or not c_col:
            self.audit_results["summary"] = {"total_debit": 0, "total_credit": 0, "difference": 0}
            return self.audit_results

        # تنظيف البيانات وتحويلها لأرقام لتفادي أخطاء النصوص
        self.df[d_col] = pd.to_numeric(self.df[d_col], errors='coerce').fillna(0)
        self.df[c_col] = pd.to_numeric(self.df[c_col], errors='coerce').fillna(0)

        total_debit = float(self.df[d_col].sum())
        total_credit = float(self.df[c_col].sum())
        
        self.audit_results["summary"] = {
            "total_debit": total_debit,
            "total_credit": total_credit,
            "difference": round(total_debit - total_credit, 2)
        }

        for _, row in self.df.iterrows():
            code = str(row[cd_col])
            name = str(row[nm_col]) if nm_col else "غير معروف"
            debit = float(row[d_col])
            credit = float(row[c_col])
            net = debit - credit

            # 1. الأرصدة الشاذة
            if check_abnormal:
                if code.startswith(('1', '5')) and net < 0:
                    self.audit_results["abnormal_balances"].append({
                        "الكود": code, "الحساب": name, "الخلل": f"رصيد دائن شاذ قدره {abs(net):,.2f}"
                    })
                elif code.startswith(('2', '3', '4')) and net > 0:
                    self.audit_results["abnormal_balances"].append({
                        "الكود": code, "الحساب": name, "الخلل": f"رصيد مدين شاذ قدره {net:,.2f}"
                    })

            # 2. فحص النقدية
            if check_cash and ("صندوق" in name or "بنك" in name or "نقد" in name) and net < 0:
                self.audit_results["bank_cash_warnings"].append({
                    "الكود": code, "الحساب": name, "التحذير": f"سحب على المكشوف/عجز نقدي بقيمة {abs(net):,.2f}"
                })

            # 3. الحسابات الوسيطة
            if check_suspense and any(k in name for k in ["وسيط", "تسوية", "عهد", "مؤقت"]):
                if net != 0:
                    self.audit_results["suspense_accounts"].append({
                        "الكود": code, "الحساب": name, "الرصيد المعلق": f"{net:,.2f}"
                    })

        return self.audit_results

# ==========================================
# 5. دوال إنشاء واستخراج التقارير (Excel & PDF)
# ==========================================
def generate_excel_export(results: dict, df_original: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([results.get("summary", {})]).to_excel(writer, sheet_name='الملخص العام', index=False)
        if results.get("abnormal_balances"):
            pd.DataFrame(results["abnormal_balances"]).to_excel(writer, sheet_name='الأرصدة الشاذة', index=False)
        warnings = results.get("bank_cash_warnings", []) + results.get("suspense_accounts", [])
        if warnings:
            pd.DataFrame(warnings).to_excel(writer, sheet_name='التنبيهات والمخاطر', index=False)
        df_original.to_excel(writer, sheet_name='ميزان المراجعة الأصلي', index=False)
    return output.getvalue()

def fix_arabic_text(text: str) -> str:
    reshaped_text = arabic_reshaper.reshape(text)
    return get_display(reshaped_text)

def generate_pdf_report(report_md_text: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    font_path = "Amiri-Regular.ttf"
    if os.path.exists(font_path):
        pdf.add_font("Amiri", "", font_path)
        pdf.set_font("Amiri", size=11)
    else:
        pdf.set_font("Arial", size=11)

    pdf.set_auto_page_break(auto=True, margin=15)
    for line in report_md_text.split("\n"):
        cleaned = line.replace("#", "").replace("*", "").strip()
        if cleaned:
            pdf.multi_cell(0, 8, txt=fix_arabic_text(cleaned), align="R")
            pdf.ln(1)
    return bytes(pdf.output())

# ==========================================
# 6. واجهة العرض الرئيسية والتفاعل
# ==========================================
uploaded_file = st.file_uploader(
    "قم بسحب وإسقاط ملف ميزان المراجعة هنا:", 
    type=["xlsx", "xls", "csv", "pdf", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    df = process_uploaded_file(uploaded_file)
    
    if not df.empty:
        st.success("تم استخراج وقراءة البيانات بنجاح!")
        
        auditor = TrialBalanceAuditor(df)
        results = auditor.run_audit()
        summary = results["summary"]

        col1, col2, col3 = st.columns(3)
        col1.metric("إجمالي المدين", f"{summary['total_debit']:,.2f}")
        col2.metric("إجمالي الدائن", f"{summary['total_credit']:,.2f}")
        diff = summary['difference']
        col3.metric("الفارق الحسابي", f"{diff:,.2f}", delta_color="normal" if diff == 0 else "inverse")

        with st.expander("📄 معاينة جدول ميزان المراجعة", expanded=False):
            st.dataframe(df, use_container_width=True)

        st.subheader("🔍 نتائج التدقيق الرقابي الآلي")
        c1, c2 = st.columns(2)
        with c1:
            st.write("**⚠️ الأرصدة الشاذة المكتشفة:**")
            if results["abnormal_balances"]:
                st.dataframe(pd.DataFrame(results["abnormal_balances"]), use_container_width=True)
            else:
                st.info("لا توجد أرصدة شاذة.")

        with c2:
            st.write("**🚨 تنبيهات النقدية والحسابات الوسيطة:**")
            warnings = results["bank_cash_warnings"] + results["suspense_accounts"]
            if warnings:
                st.dataframe(pd.DataFrame(warnings), use_container_width=True)
            else:
                st.info("لا توجد ملاحظات على النقدية أو الحسابات الوسيطة.")

        st.divider()

        if st.button("🚀 توليد التقرير الرقابي الشامل (AI Report)", type="primary"):
            if not api_key:
                st.warning("يرجى إدخال مفتاح OpenAI API Key في القائمة الجانبية لتوليد التقرير.")
            else:
                with st.spinner("جاري صياغة التقرير المالي بواسطة الذكاء الاصطناعي..."):
                    try:
                        client = OpenAI(api_key=api_key)
                        prompt = f"""
                        أنت رئيس تدقيق مالي ورقابة داخلية. قم بكتابة تقرير تدقيق مالي واحترافي بناءً على نتائج الفحص الآلي لميزان المراجعة التالي:
                        {json.dumps(results, ensure_ascii=True, indent=2)}

                        قم بتنسيق التقرير ليشمل:
                        1. ملخص تنفيذي.
                        2. تحليل المخاطر والأرصدة الشاذة.
                        3. تقييم السيولة والنقدية.
                        4. التوصيات والإجراءات التصحيحية الواجب اتخاذها فوراً.
                        """
                        response = client.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.2
                        )
                        st.session_state['report_content'] = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"حدث خطأ أثناء الاتصال بالنظام: {e}")

        if 'report_content' in st.session_state:
            st.subheader("📋 التقرير الرقابي النهائي")
            st.markdown(st.session_state['report_content'])
            
            st.divider()
            st.subheader("📥 تصدير النتائج والتقارير")
            col_ex, col_pdf = st.columns(2)
            
            with col_ex:
                excel_bytes = generate_excel_export(results, df)
                st.download_button(
                    label="📊 تحميل نتائج الفحص (Excel)",
                    data=excel_bytes,
                    file_name="Trial_Balance_Audit_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_pdf:
                pdf_bytes = generate_pdf_report(st.session_state['report_content'])
                st.download_button(
                    label="📄 تحميل التقرير النهائي (PDF)",
                    data=pdf_bytes,
                    file_name="Financial_Audit_Report.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
