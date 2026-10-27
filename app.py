import io
import re
import unicodedata
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='تحليل دول الدراسة والتخصصات', page_icon='🎓', layout='wide')
st.title('🎓 تحليل دول الدراسة والتخصصات')
st.caption('تحليل تفاعلي لعدد الموظفين الفريدين حسب دولة الدراسة والتخصص')

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
    if not n or n in ('0','-','nan'): return ''
    if n in LOOKUP: return LOOKUP[n]
    if n in ALIASES: return ALIASES[n]
    for name, country in sorted(LOOKUP.items(), key=lambda x:len(x[0]), reverse=True):
        if len(name) >= 13 and name in n: return country
    return 'يحتاج مراجعة'

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
required = ['الرقم الوظيفي','التخصص']
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
    st.info('تم إنشاء عمود دولة الدراسة من المؤسسة التعليمية؛ راجعي القيم المصنفة «يحتاج مراجعة».')
elif 'دولة الدراسة' not in df.columns:
    df = df.rename(columns={'الدولة':'دولة الدراسة'})

for col in ['الرقم الوظيفي','التخصص','دولة الدراسة']:
    df[col] = df[col].astype('string').str.strip()
    df[col] = df[col].replace({'':pd.NA, 'nan':pd.NA, 'None':pd.NA, '-':pd.NA})

valid = df.dropna(subset=['الرقم الوظيفي','التخصص','دولة الدراسة']).copy()
valid = valid[valid['دولة الدراسة'] != 'يحتاج مراجعة']

st.subheader('الفلاتر')
filtered = valid.copy()
filter_cols = ['الدائرة','اسم الجهة التابعة','الجنسية','المستوى التعليمي','درجة المؤهل','دولة الدراسة','التخصص']
with st.expander('اختاري الفلاتر (تُطبّق على كل الرسوم والجداول)', expanded=True):
    cols = st.columns(3)
    for i, col in enumerate(filter_cols):
        if col not in filtered.columns: continue
        choices = sorted(valid[col].dropna().astype(str).unique().tolist())
        with cols[i % 3]:
            selected = st.multiselect(col, choices, key=f'filter_{col}')
        if selected:
            filtered = filtered[filtered[col].astype(str).isin(selected)]

if filtered.empty:
    st.warning('لا توجد بيانات مطابقة للفلاتر.')
    st.stop()

# Count each employee once per country/specialization combination.
# One employee can legitimately appear in multiple combinations.
groups = (filtered.groupby(['دولة الدراسة','التخصص'], dropna=False)['الرقم الوظيفي']
          .nunique().reset_index(name='عدد الموظفين الفريدين'))
country_totals = (filtered.groupby('دولة الدراسة')['الرقم الوظيفي'].nunique()
                  .reset_index(name='عدد الموظفين الفريدين').sort_values('عدد الموظفين الفريدين',ascending=False))
specialty_totals = (filtered.groupby('التخصص')['الرقم الوظيفي'].nunique()
                    .reset_index(name='عدد الموظفين الفريدين').sort_values('عدد الموظفين الفريدين',ascending=False))

c1,c2,c3,c4 = st.columns(4)
c1.metric('الموظفون الفريدون', filtered['الرقم الوظيفي'].nunique())
c2.metric('دول الدراسة', filtered['دولة الدراسة'].nunique())
c3.metric('التخصصات', filtered['التخصص'].nunique())
c4.metric('الجهات', filtered['اسم الجهة التابعة'].nunique() if 'اسم الجهة التابعة' in filtered else '—')

st.subheader('العلاقة بين دولة الدراسة والتخصص')
col1,col2,col3 = st.columns(3)
with col1:
    top_n = st.selectbox('أكثر التخصصات', [5,10,20,50], index=1)
with col2:
    chart_mode = st.radio('شكل الأعمدة', ['متجاورة','متراكمة'], horizontal=True)
with col3:
    chart_direction = st.selectbox('محور المقارنة', ['التخصص','دولة الدراسة'])

if chart_direction == 'التخصص':
    top = specialty_totals.head(top_n)['التخصص']
    chart_df = groups[groups['التخصص'].isin(top)]
    fig = px.bar(chart_df, x='التخصص', y='عدد الموظفين الفريدين', color='دولة الدراسة',
                 barmode='group' if chart_mode=='متجاورة' else 'stack',
                 title='عدد الموظفين حسب التخصص ودولة الدراسة',
                 hover_data=['دولة الدراسة'])
else:
    top = country_totals.head(top_n)['دولة الدراسة']
    chart_df = groups[groups['دولة الدراسة'].isin(top)]
    fig = px.bar(chart_df, x='دولة الدراسة', y='عدد الموظفين الفريدين', color='التخصص',
                 barmode='group' if chart_mode=='متجاورة' else 'stack',
                 title='عدد الموظفين حسب دولة الدراسة والتخصص')
fig.update_layout(xaxis_tickangle=-35, legend_title_text='التصنيف', height=560)
st.plotly_chart(fig, use_container_width=True)
st.caption('يُحسب الموظف مرة واحدة داخل كل مجموعة (دولة + تخصص). إذا ظهر في مجموعات متعددة، فلا يجوز جمع الأعمدة لاعتبارها إجمالي الموظفين الفريدين.')

left,right = st.columns(2)
with left:
    st.subheader('الموظفون حسب الدولة')
    st.dataframe(country_totals, use_container_width=True, hide_index=True)
with right:
    st.subheader('الموظفون حسب التخصص')
    st.dataframe(specialty_totals, use_container_width=True, hide_index=True)

st.subheader('جدول العلاقات التفصيلي')
st.dataframe(groups.sort_values('عدد الموظفين الفريدين',ascending=False), use_container_width=True, hide_index=True)

with st.expander('السجلات التي تحتاج مراجعة'):
    review = df[df['دولة الدراسة'].isna() | (df['دولة الدراسة']=='يحتاج مراجعة')]
    st.write(f'عدد السجلات: {len(review):,}')
    st.dataframe(review, use_container_width=True, hide_index=True)

st.download_button('📥 تنزيل النتائج وملف البيانات (Excel)',
    data=excel_bytes({'العلاقات':groups,'حسب الدولة':country_totals,
                      'حسب التخصص':specialty_totals,'البيانات المفلترة':filtered,
                      'البيانات المحدثة':df}),
    file_name='تحليل_دول_الدراسة_والتخصصات.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
