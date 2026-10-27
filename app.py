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
    st.info('تم إنشاء عمود دولة الدراسة من المؤسسة التعليمية؛ راجعي القيم المصنفة «يحتاج مراجعة».')
elif 'دولة الدراسة' not in df.columns:
    df = df.rename(columns={'الدولة':'دولة الدراسة'})

for col in ['التخصص','دولة الدراسة']:
    df[col] = df[col].astype('string').str.strip()
    df[col] = df[col].replace({'':pd.NA, 'nan':pd.NA, 'None':pd.NA, '-':pd.NA})

valid = df.dropna(subset=['التخصص','دولة الدراسة']).copy()
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

# احتساب الصفوف دون حذف أي تكرار.
groups = (filtered.groupby(['دولة الدراسة','التخصص'], dropna=False)
          .size().reset_index(name='عدد السجلات'))
country_totals = (filtered.groupby('دولة الدراسة').size()
                  .reset_index(name='عدد السجلات').sort_values('عدد السجلات',ascending=False))
specialty_totals = (filtered.groupby('التخصص').size()
                    .reset_index(name='عدد السجلات').sort_values('عدد السجلات',ascending=False))

c1,c2,c3,c4 = st.columns(4)
c1.metric('السجلات المحللة', len(filtered))
c2.metric('دول الدراسة', filtered['دولة الدراسة'].nunique())
c3.metric('التخصصات', filtered['التخصص'].nunique())
c4.metric('الجهات', filtered['اسم الجهة التابعة'].nunique() if 'اسم الجهة التابعة' in filtered else '—')

with st.expander('🔎 فحص التكرار (للمراجعة فقط، دون حذف أو دمج)'):
    if 'الرقم الوظيفي' in df.columns:
        ids = df['الرقم الوظيفي'].astype('string').str.strip()
        ids = ids.replace({'': pd.NA, 'nan': pd.NA, 'None': pd.NA, '0': pd.NA})
        present = ids.dropna()
        st.metric('سجلات بأرقام وظيفية متكررة (زائدة عن أول ظهور)', int(present.duplicated().sum()))
        repeated = df.loc[ids.notna() & ids.duplicated(keep=False)]
        st.dataframe(repeated, use_container_width=True, hide_index=True)
        st.caption('ظهور الرقم الوظيفي أكثر من مرة لا يعني بالضرورة وجود خطأ؛ قد يكون للموظف أكثر من مؤهل. لم يتم حذف أي سجل.')
    else:
        st.info('عمود الرقم الوظيفي غير موجود؛ التحليل مستمر باحتساب الصفوف.')

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
    fig = px.bar(chart_df, x='التخصص', y='عدد السجلات', color='دولة الدراسة',
                 barmode='group' if chart_mode=='متجاورة' else 'stack',
                 title='عدد السجلات حسب التخصص ودولة الدراسة',
                 hover_data=['دولة الدراسة'])
else:
    top = country_totals.head(top_n)['دولة الدراسة']
    chart_df = groups[groups['دولة الدراسة'].isin(top)]
    fig = px.bar(chart_df, x='دولة الدراسة', y='عدد السجلات', color='التخصص',
                 barmode='group' if chart_mode=='متجاورة' else 'stack',
                 title='عدد السجلات حسب دولة الدراسة والتخصص')
fig.update_layout(xaxis_tickangle=-35, legend_title_text='التصنيف', height=560)
st.plotly_chart(fig, use_container_width=True)
st.caption('كل صف يُحسب كسجل مستقل. لا يتم حذف التكرارات أو دمجها. السجلات ذات الدولة أو التخصص غير المحدد لا تدخل في الرسم وتظل محفوظة في الملف المحدث.')

left,right = st.columns(2)
with left:
    st.subheader('السجلات حسب الدولة')
    st.dataframe(country_totals, use_container_width=True, hide_index=True)
with right:
    st.subheader('السجلات حسب التخصص')
    st.dataframe(specialty_totals, use_container_width=True, hide_index=True)

st.subheader('جدول العلاقات التفصيلي')
st.dataframe(groups.sort_values('عدد السجلات',ascending=False), use_container_width=True, hide_index=True)

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
