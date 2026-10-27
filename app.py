import io
import re
import unicodedata
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='تحليل دول الدراسة والتخصصات', page_icon='🎓', layout='wide')
st.title('🎓 تحليل دول الدراسة والتخصصات')
st.caption('تحليل تفاعلي لعدد السجلات حسب دولة الدراسة والتخصص، دون حذف التكرار')

COUNTRIES = {
 'الإمارات': ['جامعة عجمان','جامعة المدينة عجمان','جامعة أم القيوين','جامعة الإمارات العربية المتحدة','جامعة زايد','جامعة الشارقة','جامعة خليفة','جامعة أبوظبي','جامعة الغرير','جامعة الإمارات للطيران','كليات التقنية العليا','كلية التقنية العليا','كليات التقنية العيا','كلية الأفق الجامعية','أكاديمية شرطة دبي','الجامعة الأمريكية في الشارقة','الجامعة الأمريكية في الإمارات','جامعة الوصل','كلية الدراسات الاسلامية والعربية','أكاديمية ربدان','جامعة محمد بن زايد للعلوم الإنسانية','جامعة بوليتكنك ابوظبي','كلية المدينة الجامعية بعجمان','جامعة عجمان للعلوم والتكنولوجيا','كلية الدار الجامعية','جامعة الاتحاد','جامعة الجزيرة'],
 'مصر': ['جامعة الإسكندرية','جامعة طنطا','جامعة الزقازيق','جامعة القاهرة','جامعة الفيوم','جامعة قناة السويس','جامعة أسيوط','جامعة سوهاج','جامعة بنها','جامعة بني سويف','جامعة حلوان','جامعة عين شمس','جامعة المنصورة','جامعة 6 أكتوبر','جامعة مدينة السادات','جامعة الأزهر','الجامعة العمالية','الأكاديمية العربية للعلوم والتكنولوجيا','وزارة التعليم العالى - مصر'],
 'الأردن': ['الجامعة الهاشمية','جامعة اليرموك','جامعة فيلادلفيا','جامعة البتراء','الجامعة الأردنية','جامعة عمان الأهلية','جامعة الزيتونة الاردنية الاهلية','Al-Zaytoonah University of Jordan','جامعة العلوم التطبيقية الأهلية'],
 'سوريا':['جامعة دمشق','جامعة حلب'], 'لبنان':['جامعة بيروت العربية','الجامعة الأميركية في بيروت','الجامعة اللبنانية الدولية'],
 'البحرين':['جامعة البحرين'], 'الجزائر':['جامعة وهران'], 'المغرب':['جامعة الحسن الثاني','جامعة محمد الخامس'],
 'السودان':['جامعة الخرطوم','جامعة العلوم الطبية والتكنولوجيا','جامعة العلوم والتقانة'],
 'اليمن':['جامعة صنعاء','جامعة ناصر الأممية'], 'العراق':['جامعة بغداد'],
 'فلسطين':['جامعة القدس','جامعة النجاح الوطنية'], 'عمان':['كلية البريمي الجامعية'],
 'الكويت':['الهيئة العامة للتعليم التطبيقي والتدريب'],
 'السعودية':['دار الحديث الخيرية - مكة المكرمة'],
 'الهند':['University of Kerala','جامعة كيرلا','Cochin University of Science & Technology','P.A College of Engineering, Mangalore','جامعة عصمانيا'],
 'باكستان':['جامعة كراتشي','جامعة نيد للهندسة و التكنولوجيا'],
 'الفلبين':['جامعة ادامسون','جامعة الشرق الاقصى'],
 'أستراليا':['جامعة لا تروب'], 'المملكة المتحدة':['جامعة مانشستر'],
 'الولايات المتحدة الأمريكية':['جامعة نبراسكا'],
}

def normalize(v):
    if pd.isna(v): return ''
    s = unicodedata.normalize('NFKC', str(v)).strip().lower()
    s = re.sub(r'[\u064b-\u065f\u0670]', '', s)
    s = re.sub('[أإآ]', 'ا', s).replace('ى','ي').replace('ة','ه')
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()

LOOKUP = {normalize(name): country for country, names in COUNTRIES.items() for name in names}
ALIASES = {'كليات التقينه العليا':'الإمارات','كليات التقنيه':'الإمارات','كليات التقنيه العيا':'الإمارات','حامعه الشارقه':'الإمارات','الجماعه الامريكيه في الشارقه':'الإمارات','جامعه الامارات':'الإمارات','جامعه الاسكندريه':'مصر'}

def country_for(v):
    n = normalize(v)
    if not n or n in ('0','-','nan'): return 'غير مذكور'
    if n in LOOKUP: return LOOKUP[n]
    if n in ALIASES: return ALIASES[n]
    for name, country in sorted(LOOKUP.items(), key=lambda x:len(x[0]), reverse=True):
        if len(name) >= 13 and name in n: return country
    return 'غير مذكور'

def excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            ws.sheet_view.rightToLeft = True
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
    return output.getvalue()

file = st.file_uploader('ارفعي ملف Excel', type=['xlsx','xls'])
if file is None:
    st.info('ارفعي ملف البيانات للبدء. إذا لم يكن عمود دولة الدراسة موجوداً، سيحاول النظام تحديدها من اسم المؤسسة التعليمية.')
    st.stop()

try:
    workbook = pd.ExcelFile(file)
    sheet = st.selectbox('ورقة العمل', workbook.sheet_names)
    df = pd.read_excel(workbook, sheet_name=sheet, dtype={'الرقم الوظيفي':str})
except Exception as e:
    st.error(f'تعذر قراءة الملف: {e}')
    st.stop()

df.columns = [str(c).strip() for c in df.columns]
required = ['التخصص']
missing = [c for c in required if c not in df.columns]
if missing:
    st.error('الأعمدة المطلوبة غير موجودة: ' + '، '.join(missing))
    st.stop()

if 'دولة الدراسة' not in df.columns and 'الدولة' not in df.columns:
    if 'المؤسسة التعليمية' not in df.columns:
        st.error('يلزم وجود عمود دولة الدراسة/الدولة أو المؤسسة التعليمية.')
        st.stop()
    pos = df.columns.get_loc('المؤسسة التعليمية') + 1
    df.insert(pos, 'دولة الدراسة', df['المؤسسة التعليمية'].map(country_for))
    st.info('تم إنشاء عمود دولة الدراسة من المؤسسة التعليمية؛ المؤسسات غير المعروفة أو الفارغة صُنفت «غير مذكور» وتدخل في التحليل.')
elif 'دولة الدراسة' not in df.columns:
    df = df.rename(columns={'الدولة':'دولة الدراسة'})

for col in ['التخصص','دولة الدراسة']:
    df[col] = df[col].astype('string').str.strip()
    df[col] = df[col].replace({'': 'غير مذكور', 'nan': 'غير مذكور', 'None': 'غير مذكور', '-': 'غير مذكور', '0': 'غير مذكور', 'يحتاج مراجعة': 'غير مذكور'}).fillna('غير مذكور')

# الاحتفاظ بكل السجلات، حتى عندما لا تكون الدولة أو التخصص معروفين.
valid = df.copy()
