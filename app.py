import os, sqlite3, secrets, hashlib, hmac, io, html
from datetime import datetime, timedelta, timezone
from urllib.parse import urlencode
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
from docx import Document
from docx.shared import Pt, Inches, RGBColor

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.getenv('ATTENDANCE_DB', os.path.join(APP_DIR, 'attendance.db'))
APP_TITLE = 'حُضور | إدارة الحضور الذكية'
TZ = timezone(timedelta(hours=2))  # عدّلها حسب توقيت بلدك
st.set_page_config(page_title=APP_TITLE, page_icon='✦', layout='wide', initial_sidebar_state='expanded')

CSS = '''
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@400;500;600;700&display=swap');
:root{--plum:#702963;--gold:#FFD700;--ink:#24152a;--muted:#796e7e;--paper:#fbf8fc;--line:#eee5f0}
html,body,[class*="css"]{font-family:'IBM Plex Sans Arabic',sans-serif;direction:rtl}
.stApp{background:radial-gradient(ellipse at 88% 0%,#f1e5f2 0,transparent 34%),linear-gradient(180deg,#fcf9fd,#f7f3f8);color:var(--ink)}
[data-testid="stSidebar"]{background:linear-gradient(165deg,#702963,#49204d);border-left:1px solid #ffffff20}
[data-testid="stSidebar"] *{color:white!important}
[data-testid="stSidebar"] [data-testid="stRadio"] label{padding:9px 12px;border-radius:12px}
.stButton>button,.stDownloadButton>button{background:#FFD700;color:#322713;border:0;border-radius:12px;font-weight:700;padding:.55rem 1.2rem;transition:all .2s;box-shadow:0 5px 14px #70296318}
.stButton>button:hover,.stDownloadButton>button:hover{background:#f4ca00;transform:translateY(-2px);border:0;color:#24152a}
.stTextInput input,.stTextArea textarea,.stSelectbox div[data-baseweb="select"]{border-radius:12px}
[data-testid="stMetric"]{background:white;padding:18px 20px;border:1px solid var(--line);border-radius:18px;box-shadow:0 8px 22px #3f21400c}
[data-testid="stMetricLabel"]{color:#796e7e}
[data-testid="stMetricValue"]{color:#702963;font-weight:700}
[data-testid="stDataFrame"]{border-radius:14px;overflow:hidden}
.block-container{padding-top:1.6rem;padding-bottom:3rem;max-width:1440px}
.hero{background:linear-gradient(115deg,#702963,#512353);padding:28px 32px;border-radius:24px;color:white;margin:0 0 20px;box-shadow:0 18px 38px #70296327;position:relative;overflow:hidden}
.hero:after{content:'✦';position:absolute;font-size:145px;opacity:.07;left:34px;top:-25px;color:#FFD700}
.hero h1{font-size:30px;margin:0 0 6px}.hero p{margin:0;color:#f0e4f1}.card{background:#fff;border:1px solid #eee5f0;border-radius:18px;padding:20px 22px;box-shadow:0 8px 24px #36213d0a;margin-bottom:14px}
.badge{display:inline-block;background:#fff4bd;color:#654e00;padding:5px 12px;border-radius:99px;font-size:13px;font-weight:700}.muted{color:#796e7e}
@media(max-width:700px){.hero{padding:22px}.hero h1{font-size:24px}}
</style>'''
st.markdown(CSS, unsafe_allow_html=True)


def now(): return datetime.now(TZ)
def stamp(dt=None): return (dt or now()).isoformat(timespec='seconds')
def parse_dt(s):
    if not s: return None
    d=datetime.fromisoformat(s)
    return d if d.tzinfo else d.replace(tzinfo=TZ)
def db():
    c=sqlite3.connect(DB_PATH, timeout=20)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA foreign_keys=ON')
    return c

def init_db():
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS groups(id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, schedule TEXT DEFAULT '', roster TEXT DEFAULT '', created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS lectures(id INTEGER PRIMARY KEY, group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE, title TEXT NOT NULL, token_hash TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL, expires_at TEXT NOT NULL, closed INTEGER NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS attendance(id INTEGER PRIMARY KEY, lecture_id INTEGER NOT NULL REFERENCES lectures(id) ON DELETE CASCADE, group_id INTEGER NOT NULL REFERENCES groups(id) ON DELETE CASCADE, student_name TEXT NOT NULL, created_at TEXT NOT NULL, UNIQUE(lecture_id, student_name));
        CREATE INDEX IF NOT EXISTS idx_att_group ON attendance(group_id);
        ''')

def secret_hash(token): return hashlib.sha256(token.encode()).hexdigest()
def valid_doctor_password(pw):
    configured = st.secrets.get('DOCTOR_PASSWORD', os.getenv('DOCTOR_PASSWORD',''))
    return bool(configured) and hmac.compare_digest(str(configured), str(pw))
def get_groups():
    with db() as c: return c.execute('SELECT * FROM groups ORDER BY name').fetchall()
def get_lectures(group_id=None):
    with db() as c:
        if group_id: return c.execute('SELECT l.*,g.name group_name,(SELECT COUNT(*) FROM attendance a WHERE a.lecture_id=l.id) count FROM lectures l JOIN groups g ON g.id=l.group_id WHERE l.group_id=? ORDER BY l.created_at DESC',(group_id,)).fetchall()
        return c.execute('SELECT l.*,g.name group_name,(SELECT COUNT(*) FROM attendance a WHERE a.lecture_id=l.id) count FROM lectures l JOIN groups g ON g.id=l.group_id ORDER BY l.created_at DESC').fetchall()
def attendance_frame(group_id=None):
    q='SELECT a.student_name AS الاسم,g.name AS المجموعة,l.title AS المحاضرة,a.created_at AS وقت_التسجيل FROM attendance a JOIN groups g ON g.id=a.group_id JOIN lectures l ON l.id=a.lecture_id'
    args=()
    if group_id: q+=' WHERE a.group_id=?';args=(group_id,)
    q+=' ORDER BY a.created_at DESC'
    with db() as c: return pd.read_sql_query(q,c,params=args)
def roster_names(text): return list(dict.fromkeys(x.strip() for x in (text or '').splitlines() if x.strip()))
def group_summary(g):
    roster=roster_names(g['roster'])
    with db() as c:
        sessions=c.execute('SELECT COUNT(*) FROM lectures WHERE group_id=?',(g['id'],)).fetchone()[0]
        present=c.execute('SELECT COUNT(*) FROM attendance WHERE group_id=?',(g['id'],)).fetchone()[0]
        people=c.execute('SELECT COUNT(DISTINCT student_name) FROM attendance WHERE group_id=?',(g['id'],)).fetchone()[0]
    denominator=len(roster)*sessions if roster and sessions else 0
    return {'sessions':sessions,'present':present,'roster':len(roster),'people':people,'rate':100*present/denominator if denominator else None}
def app_base_url():
    configured = st.secrets.get('APP_URL', os.getenv('APP_URL','')).strip().rstrip('/')
    if configured: return configured
    try:
        u=st.context.url
        if u: return u.split('?')[0].rstrip('/')
    except Exception: pass
    return ''
def attendance_link(token):
    base=app_base_url()
    return f'{base}/?attendance={token}' if base else f'?attendance={token}'

def make_docx(group, lecture, roster, checked):
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.65);sec.bottom_margin=Inches(.65)
    p=doc.add_paragraph();p.alignment=2
    r=p.add_run('تقرير الحضور والغياب');r.bold=True;r.font.size=Pt(22);r.font.color.rgb=RGBColor(112,41,99)
    for text in [f'المجموعة: {group["name"]}',f'المحاضرة: {lecture["title"]}',f'تاريخ المحاضرة: {parse_dt(lecture["created_at"]).strftime("%Y-%m-%d %H:%M")}',f'عدد الحاضرين: {len(checked)} | عدد المسجلين: {len(roster)}']:
        p=doc.add_paragraph(text);p.alignment=2
    table=doc.add_table(rows=1, cols=3);table.style='Light Shading Accent 1';table.autofit=True
    for cell,txt in zip(table.rows[0].cells,['#','اسم الطالب','الحالة']): cell.text=txt
    ordered=roster or sorted(checked,key=str.casefold)
    for i,name in enumerate(ordered,1):
        cells=table.add_row().cells;cells[0].text=str(i);cells[1].text=name;cells[2].text='حاضر' if name in checked else ('غائب' if roster else '—')
    out=io.BytesIO();doc.save(out);return out.getvalue()

def show_student_page(token):
    with db() as c: row=c.execute('SELECT l.*,g.name group_name,g.roster FROM lectures l JOIN groups g ON g.id=l.group_id WHERE token_hash=?',(secret_hash(token),)).fetchone()
    st.markdown('<div class="hero"><h1>تأكيد الحضور</h1><p>سجّل حضورك في محاضرتك الحالية.</p></div>',unsafe_allow_html=True)
    if not row:
        st.error('هذا الرابط غير صالح أو انتهت صلاحيته. اطلب رابطًا جديدًا من الدكتور.');return
    is_open=not row['closed'] and now()<parse_dt(row['expires_at'])
    remaining=max(0,int((parse_dt(row['expires_at'])-now()).total_seconds())) if is_open else 0
    st.markdown(f'<div class="card"><span class="badge">المجموعة: {html.escape(row["group_name"])}</span><h3>{html.escape(row["title"])}</h3><p class="muted">ينتهي الرابط في {parse_dt(row["expires_at"]).strftime("%H:%M")}</p></div>',unsafe_allow_html=True)
    if not is_open: st.warning('أُغلق التسجيل لهذه المحاضرة.');return
    components.html(f'''<div style="font-family:Arial;text-align:center;color:#702963;background:#fff6d0;border-radius:12px;padding:10px">⏳ الوقت المتبقي: <b id="timer"></b></div><script>let s={remaining};function tick(){{let m=Math.floor(s/60),x=s%60;document.getElementById('timer').textContent=m+':'+String(x).padStart(2,'0');if(s>0)s--;else document.getElementById('timer').textContent='انتهى الوقت'}}tick();setInterval(tick,1000)</script>''',height=55)
    roster=roster_names(row['roster'])
    with st.form('mark_attendance',clear_on_submit=False):
        name=st.text_input('الاسم الكامل',placeholder='اكتب اسمك كما هو مسجل لدى الدكتور',max_chars=100)
        if roster: st.caption('يُقبل الاسم المطابق لقائمة المجموعة فقط.')
        submitted=st.form_submit_button('تأكيد حضوري الآن',use_container_width=True)
    if submitted:
        clean=' '.join(name.split())
        if len(clean)<3: st.error('يرجى إدخال الاسم الكامل.')
        elif roster and clean.casefold() not in {x.casefold() for x in roster}: st.error('الاسم غير موجود ضمن قائمة هذه المجموعة. راجع الدكتور.')
        else:
            if roster: clean=next(x for x in roster if x.casefold()==clean.casefold())
            try:
                with db() as c: c.execute('INSERT INTO attendance(lecture_id,group_id,student_name,created_at) VALUES(?,?,?,?)',(row['id'],row['group_id'],clean,stamp()))
                st.success(f'تم تسجيل حضورك بنجاح يا {clean}.')
                st.balloons()
            except sqlite3.IntegrityError: st.info('تم تسجيل حضور هذا الاسم مسبقًا لهذه المحاضرة.')
    st.caption('يُسجّل الاسم مرة واحدة فقط لكل محاضرة. لا تغلق الصفحة قبل ظهور رسالة التأكيد.')

def login_page():
    st.markdown('<div class="hero"><h1>حُضور</h1><p>منصة أنيقة لإدارة حضور المجموعات والمحاضرات.</p></div>',unsafe_allow_html=True)
    st.markdown('<div class="card"><b>مساحة الدكتور</b><br><span class="muted">سجّل الدخول لإدارة المجموعات وإنشاء روابط الحضور.</span></div>',unsafe_allow_html=True)
    with st.form('login'):
        pw=st.text_input('كلمة مرور الدكتور',type='password')
        go=st.form_submit_button('دخول لوحة التحكم',use_container_width=True)
    if go:
        if valid_doctor_password(pw): st.session_state.doctor=True;st.rerun()
        else: st.error('كلمة المرور غير صحيحة. اضبط DOCTOR_PASSWORD في إعدادات Streamlit.')
    st.caption('لأمان أفضل استخدم كلمة مرور طويلة، ولا تشاركها مع الطلبة.')

def doctor_app():
    groups=get_groups()
    with st.sidebar:
        st.markdown('## ✦ حُضور')
        st.caption('لوحة تحكم الدكتور')
        page=st.radio('القائمة',['الرئيسية','المجموعات والطلاب','إنشاء محاضرة','سجل الحضور','الإحصائيات','إعدادات'],label_visibility='collapsed')
        st.divider()
        if groups:
            chosen=st.selectbox('المجموعة النشطة',[g['name'] for g in groups])
            group=next(g for g in groups if g['name']==chosen)
        else: group=None
        if st.button('تسجيل الخروج',use_container_width=True): st.session_state.doctor=False;st.rerun()
    st.markdown(f'<div class="hero"><h1>{page}</h1><p>إدارة سهلة وواضحة لحضور مجموعاتك — مع حفظ محلي للسجلات.</p></div>',unsafe_allow_html=True)
    if page=='الرئيسية':
        if not groups:
            st.info('ابدأ بإضافة أول مجموعة من تبويب «المجموعات والطلاب».')
        else:
            selected=group;summary=group_summary(selected)
            active=[l for l in get_lectures(selected['id']) if not l['closed'] and now()<parse_dt(l['expires_at'])]
            a,b,c,d=st.columns(4);a.metric('المجموعات',len(groups));b.metric('محاضرات المجموعة',summary['sessions']);c.metric('إجمالي تسجيلات الحضور',summary['present']);d.metric('نسبة الحضور',f'{summary["rate"]:.1f}%' if summary['rate'] is not None else '—')
            st.subheader(f'نظرة سريعة · {selected["name"]}')
            if active:
                l=active[0];st.success(f'رابط حضور مفتوح الآن: {l["title"]} — {l["count"]} تسجيل')
                url=attendance_link('')
                tokenrow=None
                with db() as conn: tokenrow=conn.execute('SELECT token_hash FROM lectures WHERE id=?',(l['id'],)).fetchone()
                # token is never stored in plaintext; existing link is intentionally not recoverable.
                st.caption('الرابط الأصلي يظهر عند إنشاء المحاضرة فقط. أنشئ رابطًا جديدًا إذا لم تحتفظ به.')
            else: st.info('لا يوجد رابط حضور نشط لهذه المجموعة حاليًا.')
            logs=attendance_frame(selected['id'])
            if not logs.empty: st.dataframe(logs.head(12),use_container_width=True,hide_index=True)
            if summary['sessions'] and summary['roster']:
                st.caption(f'القائمة: {summary["roster"]} طالبًا · المتوسط العام للمجموعة محسوب عبر المحاضرات المسجلة.')
    elif page=='المجموعات والطلاب':
        st.subheader('إضافة مجموعة')
        with st.form('new_group'):
            name=st.text_input('اسم المجموعة',placeholder='مثال: المجموعة A')
            schedule=st.text_input('موعد المحاضرة المعتاد',placeholder='الأحد 10:00 — اختياري')
            roster=st.text_area('قائمة أسماء الطلبة (اسم واحد في كل سطر)',height=150,placeholder='يمكن تركها فارغة، لكن إضافتها تمنع الأسماء غير المعروفة.')
            ok=st.form_submit_button('إضافة المجموعة')
        if ok:
            clean=name.strip()
            if len(clean)<2: st.error('اكتب اسمًا واضحًا للمجموعة.')
            else:
                try:
                    with db() as conn: conn.execute('INSERT INTO groups(name,schedule,roster,created_at) VALUES(?,?,?,?)',(clean,schedule.strip(), '\n'.join(roster_names(roster)),stamp()))
                    st.success('تمت إضافة المجموعة.');st.rerun()
                except sqlite3.IntegrityError: st.error('اسم المجموعة مستخدم من قبل.')
        st.divider();st.subheader('المجموعات الحالية')
        for g in get_groups():
            with st.expander(f'{g["name"]} · {len(roster_names(g["roster"]))} طالبًا · {g["schedule"] or "دون موعد محدد"}'):
                with st.form(f'edit_{g["id"]}'):
                    new_sched=st.text_input('موعد المحاضرة',value=g['schedule'])
                    new_roster=st.text_area('الأسماء، كل اسم بسطر',value=g['roster'],height=130)
                    save=st.form_submit_button('حفظ التعديلات')
                if save:
                    with db() as conn: conn.execute('UPDATE groups SET schedule=?,roster=? WHERE id=?',(new_sched,'\n'.join(roster_names(new_roster)),g['id']))
                    st.success('تم حفظ بيانات المجموعة.');st.rerun()
                st.caption('حذف المجموعة غير متاح من الواجهة لحماية سجل الحضور.')
    elif page=='إنشاء محاضرة':
        if not groups: st.warning('أضف مجموعة أولًا.');return
        st.markdown('<div class="card">كل محاضرة تحصل على رابط فريد عشوائي صالح لمدة ساعة واحدة. لا يمكن إعادة استخدام الاسم في المحاضرة ذاتها.</div>',unsafe_allow_html=True)
        with st.form('new_lecture'):
            gname=st.selectbox('المجموعة',[g['name'] for g in groups]);title=st.text_input('عنوان المحاضرة',value=f'محاضرة {now().strftime("%Y-%m-%d")}')
            minutes=st.select_slider('مدة الرابط بالدقائق',options=[15,30,45,60],value=60)
            submitted=st.form_submit_button('إنشاء رابط الحضور',use_container_width=True)
        if submitted:
            group=next(x for x in groups if x['name']==gname)
            token=secrets.token_urlsafe(32);created=now();expires=created+timedelta(minutes=minutes)
            with db() as conn: conn.execute('INSERT INTO lectures(group_id,title,token_hash,created_at,expires_at) VALUES(?,?,?,?,?)',(group['id'],title.strip() or 'محاضرة',secret_hash(token),stamp(created),stamp(expires)))
            st.session_state.new_token=token;st.session_state.new_lecture_title=title;st.session_state.new_group_name=gname;st.session_state.new_expiry=stamp(expires)
        if st.session_state.get('new_token'):
            token=st.session_state.new_token;link=attendance_link(token)
            st.success(f'تم إنشاء الرابط لمجموعة {st.session_state.new_group_name}. شاركه مع هذه المجموعة فقط.')
            st.code(link,language=None)
            if link.startswith('?'): st.warning('أدخل عنوان نشر التطبيق في تبويب الإعدادات ثم أنشئ رابطًا جديدًا ليظهر الرابط كاملًا.')
            components.html(f'''<div style="direction:rtl;font-family:Arial;background:white;border:1px solid #eee;border-radius:16px;padding:18px;text-align:center"><div style="color:#702963;font-size:13px">رابط المحاضرة صالح حتى {parse_dt(st.session_state.new_expiry).strftime('%H:%M')}</div><button onclick="navigator.clipboard.writeText({link!r}).then(()=>this.innerText='تم النسخ ✓')" style="margin:12px;background:#FFD700;border:0;border-radius:10px;padding:12px 22px;font-weight:bold;cursor:pointer">نسخ الرابط</button><div style="font-size:12px;color:#796e7e">أغلق التسجيل مبكرًا من سجل الحضور عند الحاجة.</div></div>''',height=115)
            st.caption('احتفظ بالرابط أو أرسله فورًا؛ الرمز السري لا يُخزّن كنص قابل للاسترجاع.')
    elif page=='سجل الحضور':
        if not groups: st.info('أضف مجموعة لتبدأ.');return
        g=group;lectures=get_lectures(g['id'])
        if not lectures: st.info('لا توجد محاضرات لهذه المجموعة بعد.');return
        lec=st.selectbox('اختر محاضرة',lectures,format_func=lambda x:f'{x["title"]} · {parse_dt(x["created_at"]).strftime("%Y-%m-%d %H:%M")} · {x["count"]} حاضر')
        roster=roster_names(g['roster'])
        with db() as conn: present={r[0] for r in conn.execute('SELECT student_name FROM attendance WHERE lecture_id=?',(lec['id'],)).fetchall()}
        opened=not lec['closed'] and now()<parse_dt(lec['expires_at'])
        x,y,z=st.columns(3);x.metric('الحاضرون',len(present));y.metric('أسماء القائمة',len(roster) if roster else 'غير محددة');z.metric('حالة الرابط','مفتوح' if opened else 'مغلق')
        if opened and st.button('إغلاق الرابط الآن',type='secondary'):
            with db() as conn: conn.execute('UPDATE lectures SET closed=1 WHERE id=?',(lec['id'],))
            st.success('أُغلق الرابط.');st.rerun()
        if roster:
            df=pd.DataFrame([{'الاسم':n,'الحالة':'حاضر' if n.casefold() in {p.casefold() for p in present} else 'غائب'} for n in roster])
            st.dataframe(df,use_container_width=True,hide_index=True)
        elif present: st.dataframe(pd.DataFrame({'الاسم':sorted(present),'الحالة':['حاضر']*len(present)}),use_container_width=True,hide_index=True)
        else: st.info('لم يسجل أحد حضوره بعد.')
        if roster or present:
            content=make_docx(g,lec,roster,present)
            st.download_button('تنزيل تقرير Word للمجموعة',data=content,file_name=f'attendance_{g["name"]}_{lec["id"]}.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            st.caption('التقرير يتضمن الأسماء والحضور والغياب فقط دون تقديرات.')
    elif page=='الإحصائيات':
        st.subheader('مؤشرات الحضور')
        group_filter=st.selectbox('نطاق الإحصائيات',['كل المجموعات']+[g['name'] for g in groups]) if groups else 'كل المجموعات'
        target=[next(g for g in groups if g['name']==group_filter)] if group_filter!='كل المجموعات' else groups
        rows=[]
        for g in target:
            sm=group_summary(g);rows.append({'المجموعة':g['name'],'عدد الطلبة بالقائمة':sm['roster'],'المحاضرات':sm['sessions'],'إجمالي تسجيلات الحضور':sm['present'],'نسبة الحضور العامة (%)':round(sm['rate'],1)})
        if rows:
            stats=pd.DataFrame(rows);a,b,c=st.columns(3)
            total_s=sum(r['المحاضرات'] for r in rows);total_present=sum(r['إجمالي تسجيلات الحضور'] for r in rows)
            denom=sum(r['عدد الطلبة بالقائمة']*r['المحاضرات'] for r in rows)
            general=100*total_present/denom if denom else None
            a.metric('نسبة الحضور العامة',f'{general:.1f}%' if general is not None else '—');b.metric('المجموعات',len(rows));c.metric('المحاضرات',total_s)
            st.dataframe(stats,use_container_width=True,hide_index=True)
            if len(rows)>1: st.plotly_chart(px.bar(stats,x='المجموعة',y='نسبة الحضور العامة (%)',color='المجموعة',title='نسبة الحضور حسب المجموعة',color_discrete_sequence=['#702963','#FFD700','#9b70a1']),use_container_width=True)
            if group_filter!='كل المجموعات':
                g=target[0]
                with db() as conn: lrows=conn.execute('SELECT l.title,l.created_at,l.id,(SELECT COUNT(*) FROM attendance a WHERE a.lecture_id=l.id) n FROM lectures l WHERE group_id=? ORDER BY created_at',(g['id'],)).fetchall()
                if lrows:
                    hist=pd.DataFrame([{'المحاضرة':r['title'],'التاريخ':parse_dt(r['created_at']).strftime('%m-%d'),'الحضور':r['n']} for r in lrows])
                    st.plotly_chart(px.line(hist,x='التاريخ',y='الحضور',markers=True,title=f'تطور الحضور · {g["name"]}',color_discrete_sequence=['#702963']),use_container_width=True)
                    if roster_names(g['roster']):
                        roster=roster_names(g['roster']); per=[]
                        with db() as conn:
                            for student in roster:
                                count=conn.execute('SELECT COUNT(*) FROM attendance WHERE group_id=? AND student_name=?',(g['id'],student)).fetchone()[0]
                                per.append({'الطالب':student,'مرات الحضور':count,'المحاضرات':len(lrows),'نسبة الحضور (%)':round(count*100/len(lrows),1) if lrows else 0})
                        with st.expander('تفصيل الحضور حسب الطالب'): st.dataframe(pd.DataFrame(per),use_container_width=True,hide_index=True)
        else: st.info('أضف مجموعة وابدأ تسجيل المحاضرات لظهور الإحصائيات.')
        st.caption('تُحسب النسبة العامة من إجمالي تسجيلات الحضور ÷ (عدد طلبة القائمة × عدد محاضرات المجموعة). إذا لم تُضف قائمة أسماء، تعرض النسبة على أساس التسجيلات فقط.')
    elif page=='إعدادات':
        st.subheader('إعدادات النشر والأمان')
        st.markdown('''<div class="card"><b>كلمة المرور:</b> عيّن متغير <code>DOCTOR_PASSWORD</code> في إعدادات Streamlit Cloud / ملف <code>.streamlit/secrets.toml</code>. لا ترفع ملف الأسرار إلى GitHub.<br><br><b>عنوان الموقع:</b> احفظ رابط تطبيق Streamlit في <code>APP_URL</code> كي تُنشأ روابط كاملة للطلبة.<br><br><b>قاعدة البيانات:</b> تحفظ البيانات في <code>attendance.db</code> على الخادم الحالي. نسخ Streamlit Community Cloud قد لا تضمن بقاء الملفات المحلية عند إعادة التشغيل/النشر؛ استخدم قاعدة بيانات خارجية دائمة قبل الاعتماد الإنتاجي.</div>''',unsafe_allow_html=True)
        st.code(f'APP_URL = "{app_base_url() or "https://your-app.streamlit.app"}"',language='toml')
        st.caption('استبدل المثال برابط تطبيقك داخل secrets.toml أو إعداد Secrets على Streamlit Cloud، ثم أعد تشغيل التطبيق.')
        with st.expander('عن حماية رابط الحضور'):
            st.write('الرابط يحتوي رمزًا عشوائيًا قويًا لا يُحفظ صريحًا في قاعدة البيانات، ينتهي تلقائيًا بعد ساعة، ويمكن إغلاقه يدويًا. التسجيل الفريد يمنع تكرار الاسم في المحاضرة نفسها. الاسم فقط لا يثبت هوية الطالب؛ فعّل قائمة المجموعة واستخدم إرسال الرابط إلى قناة خاصة لتقليل الانتحال.')


def main():
    init_db()
    params=st.query_params
    token=params.get('attendance')
    if token:
        show_student_page(str(token));return
    if not st.session_state.get('doctor',False): login_page()
    else: doctor_app()

if __name__=='__main__': main()
