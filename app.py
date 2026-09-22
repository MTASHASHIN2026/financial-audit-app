import io
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

st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    div[data-testid="stMetricValue"] { text-align: right; }
    .stButton>button { width: 100%; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 نظام المساعد الذكي لتدقيق ميزان المراجعة")
st.caption("قم برفع ملف ميزان المراجعة لفحصه محاسبياً وتوليد التقرير الرقابي.")

# ==========================================
# 2. القائمة الجانبية (Sidebar)
# ==========================================
with st.sidebar:
    st.header("⚙️ الإعدادات والمفاتيح")
    
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("أدخل مفتاح OpenAI API Key:", type="password")
    else:
        st.success("تم ربط مفتاح API تلقائياً من الأسرار 🔒")

    st.divider()
    st.subheader("📌 نطاق الفحص الرقابي")
    check_abnormal = st.checkbox("فحص الأرصدة الشاذة", value=True)
    check_cash = st.checkbox("فحص النقدية", value=True)
    check_suspense = st.checkbox("فحص الحسابات الوسيطة", value=True)

# ==========================================
# 3. معالجة وقراءة الملفات (مع الباحث الذكي)
# ==========================================
def process_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_type = uploaded_file.name.split('.')[-1].lower()
    try:
        if file_type in ['xlsx', 'xls']:
            # قراءة الملف بدون عناوين مسبقة للبحث عن الصف الصحيح
            df = pd.read_excel(uploaded_file, header=None)
            
            # محرك البحث الذكي عن صف العناوين (يقرأ أول 20 صف)
            header_idx = 0
            for i in range(min(20, len(df))):
                row_text = ' '.join(str(x) for x in df.iloc[i].values).lower()
                if any(k in row_text for k in ['مدين', 'دائن', 'debit', 'credit']):
                    header_idx = i
                    break
                    
            # تعيين الصف المكتشف كعنوان وحذف ما قبله
            df.columns = df.iloc[header_idx].astype(str).str.strip()
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
            
            # حذف الأعمدة الفارغة تماماً
            df = df.dropna(axis=1, how='all')
            return df
            
        elif file_type == 'csv':
            df = pd.read_csv(uploaded_file)
            df.columns = [str(col).strip() for col in df.columns]
            return df
        else:
            return pd.DataFrame()
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
        debit_col = [c for c in cols if 'debit' in str(c).lower() or 'مدين' in str(c) or 'منه' in str(c)]
        credit_col = [c for c in cols if 'credit' in str(c).lower() or 'دائن' in str(c) or 'له' in str(c)]
        code_col = [c for c in cols if 'code' in str(c).lower() or 'رمز' in str(c) or 'رقم' in str(c)]
        name_col = [c for c in cols if 'name' in str(c).lower() or 'اسم' in str(c) or 'حساب' in str(c) or 'بيان' in str(c)]

        d_col = debit_col[0] if debit_col else (self.df.columns[2] if len(self.df.columns) > 2 else None)
        c_col = credit_col[0] if credit_col else (self.df.columns[3] if len(self.df.columns) > 3 else None)
        cd_col = code_col[0] if code_col else self.df.columns[0]
        nm_col = name_col[0] if name_col else (self.df.columns[1] if len(self.df.columns) > 1 else None)

        if not d_col or not c_col:
            self.audit_results["summary"] = {"total_debit": 0, "total_credit": 0, "difference": 0}
            return self.audit_results

        # تنظيف البيانات وتحويلها لأرقام
        self.df[d_col] = pd.to_numeric(self.df[d_col], errors='coerce').fillna(0)
        self.يظهر في ملف "خطا.png" رسالة الخطأ التالية: `حدث خطأ أثناء الاتصال بالنظام: 'ascii' codec can't encode characters in position 15-16: ordinal not in range(128)` وذلك بعد محاولة توليد التقرير الرقابي الشامل[cite: 1].

هذا الخطأ البرمجي (والذي يحدث غالباً في بيئة بايثون) ينتج عندما يحاول النظام معالجة، أو إرسال، أو طباعة نصوص تحتوي على أحرف عربية (مثل اسم الملف المرفوع "ميزان م...9-2026.xls" أو البيانات المحاسبية المستخرجة منه) باستخدام ترميز `ASCII` القديم الذي لا يدعم سوى الأحرف الإنجليزية، بدلاً من ترميز `UTF-8` العالمي[cite: 1].

لحل هذه المشكلة في الكود المصدري للتطبيق، يجب تطبيق التعديلات التالية بناءً على مكان حدوث الخطأ:

*   **تجهيز البيانات لـ API الذكاء الاصطناعي:** عند تحويل البيانات المحاسبية إلى نصوص (JSON) لإرسالها إلى نموذج الذكاء الاصطناعي، يجب التأكد من استخدام `ensure_ascii=False` داخل دالة `json.dumps()` لضمان عدم تحويل الأحرف العربية إلى رموز غير مقروءة.
*   **إعدادات الاتصال (HTTP Requests):** إذا كان التطبيق يستخدم مكتبة مثل `requests` للاتصال بالنظام الخارجي، تأكد من إضافة `charset=utf-8` إلى ترويسة الطلب: `{'Content-Type': 'application/json; charset=utf-8'}`.
*   **قراءة وكتابة الملفات:** في أي مكان يقوم فيه الكود بفتح ملفات نصية أو حفظ سجلات (Logs)، يجب تحديد الترميز صراحةً بإضافة `encoding='utf-8'` (على سبيل المثال: `open(filename, 'r', encoding='utf-8')`).
*   **متغيرات بيئة التشغيل (Environment Variables):** إذا كان التطبيق مستضافاً على خادم أو يعمل عبر سطر الأوامر، قد تحتاج إلى إجبار بايثون على استخدام UTF-8 كترميز افتراضي عن طريق ضبط المتغير `PYTHONIOENCODING=utf-8`.
