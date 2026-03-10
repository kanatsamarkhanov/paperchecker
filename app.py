import streamlit as st
from docx import Document
from docx.shared import Mm, Pt
import re
import pandas as pd
from io import BytesIO

# ─── БЕТТІҢ БАПТАУЛАРЫ ───────────────────────────────────────────
st.set_page_config(
    page_title="Чекер статьи / Мақаланы тексеру",
    page_icon="📋",
    layout="wide"
)

# ─── СЕССИЯ ЖАҒДАЙЫН (STATE) ИНИЦИАЛИЗАЦИЯЛАУ ──────────────────────
if 'lang' not in st.session_state:
    st.session_state.lang = 'kz'
if 'theme' not in st.session_state:
    st.session_state.theme = 'light'

# ─── АУДАРМАЛАР (СӨЗДІК) ──────────────────────────────────────────
locales = {
    'ru': {
        'title': "📋 Автоматическая проверка статьи",
        'subtitle': "Вестник ЕНУ им. Л.Н. Гумилева · Серия: Химия / География · Шаблон 2025",
        'btn_lang': "🇰🇿 ҚАЗ",
        'btn_theme_dark': "🌙 Тёмная тема",
        'btn_theme_light': "☀️ Светлая тема",
        'upload_title': "📂 Загрузите статью в формате .docx",
        'upload_help': "Шаблон журнала Вестник ЕНУ, серия Химия/География, 2025",
        'analyzing': "Анализируем статью...",
        'res_title': "📊 Результаты проверки",
        'total': "Всего",
        'passed': "✅ Выполнено",
        'warned': "⚠️ Внимание",
        'failed': "❌ Не выпол.",
        'score': "🏆 Соответствие",
        'det_report': "### 📋 Детальный отчёт",
        'btn_csv': "⬇️ Скачать CSV",
        'btn_xls': "⬇️ Скачать Excel",
        'btn_docx': "⬇️ Word (DOCX)",
        'req_fix': "### ⚠️ Требует исправления",
        'req': "требование",
        'no_file': "👆 Загрузите .docx файл, чтобы начать проверку",

        # Критерии
        'c_vol': "Объём статьи", 'c_vol_req': "≥ 3500 слов",
        'c_ann_ru': "Аннотация (рус)", 'c_ann_req': "≤ 300 слов",
        'c_ann_kz': "Аннотация (каз)", 'c_ann_en': "Abstract (англ)", 'c_req_obl': "Обязательно",
        'c_kw': "Ключевые слова", 'c_kw_req': "3–10, разделитель «;»",
        'c_mrnti': "Код МРНТИ", 'c_orcid': "ORCID авторов", 'c_orcid_req': "Для каждого автора",
        'c_intro': "Раздел: Введение", 'c_mm': "Раздел: Материалы и методы",
        'c_res': "Раздел: Результаты", 'c_disc': "Раздел: Обсуждение", 'c_concl': "Раздел: Заключение",
        'c_contrib': "Вклад авторов", 'c_fund': "Финансирование", 'c_conflict': "Конфликт интересов",
        'c_refs': "Кол-во источников", 'c_refs_req': "≥ 25",
        'c_doi': "DOI в ссылках", 'c_apa': "Стиль цитирования", 'c_apa_req': "APA 7 (Автор, год)",

        # Технические критерии
        'c_paper': "Формат бумаги", 'c_paper_req': "A4 (210x297 мм)",
        'c_margins': "Поля (Жақтаулар)", 'c_margins_req': "Обычно 2 см",
        'c_font': "Шрифт и кегль", 'c_font_req': "Times New Roman, 11/12 pt",
        'c_tables': "Кестелер (Таблицы)", 'c_tables_req': "Наличие ссылок в тексте",
        'c_images': "Рисунки (DPI)", 'c_images_req': "Мин. 300 DPI (проверьте вручную)",

        'found': "Найдено", 'not_found': "Отсутствует", 'words': "слов",
    },
    'kz': {
        'title': "📋 Мақаланы автоматты түрде тексеру",
        'subtitle': "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы · Серия: Химия / География · 2025 үлгісі",
        'btn_lang': "🇷🇺 РУС",
        'btn_theme_dark': "🌙 Түнгі режим",
        'btn_theme_light': "☀️ Күндізгі режим",
        'upload_title': "📂 .docx форматындағы мақаланы жүктеңіз",
        'upload_help': "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы, Химия/География сериясы, 2025 үлгісі",
        'analyzing': "Мақала талдануда...",
        'res_title': "📊 Тексеру нәтижелері",
        'total': "Барлығы",
        'passed': "✅ Орындалды",
        'warned': "⚠️ Назар аударыңыз",
        'failed': "❌ Орындалмады",
        'score': "🏆 Сәйкестік",
        'det_report': "### 📋 Толық есеп",
        'btn_csv': "⬇️ CSV жүктеу",
        'btn_xls': "⬇️ Excel жүктеу",
        'btn_docx': "⬇️ Word (DOCX)",
        'req_fix': "### ⚠️ Түзетуді қажет етеді",
        'req': "талап",
        'no_file': "👆 Тексеруді бастау үшін .docx файлын жүктеңіз",

        # Критерии
        'c_vol': "Мақала көлемі", 'c_vol_req': "≥ 3500 сөз",
        'c_ann_ru': "Аңдатпа (орыс)", 'c_ann_req': "≤ 300 сөз",
        'c_ann_kz': "Аңдатпа (қаз)", 'c_ann_en': "Abstract (ағылш)", 'c_req_obl': "Міндетті түрде",
        'c_kw': "Түйінді сөздер", 'c_kw_req': "3–10, бөлгіш «;»",
        'c_mrnti': "МРНТИ коды", 'c_orcid': "Авторлардың ORCID", 'c_orcid_req': "Әр автор үшін",
        'c_intro': "Бөлім: Кіріспе (Введение)", 'c_mm': "Бөлім: Материалдар мен әдістер",
        'c_res': "Бөлім: Нәтижелер", 'c_disc': "Бөлім: Талқылау", 'c_concl': "Бөлім: Қорытынды",
        'c_contrib': "Авторлардың үлесі", 'c_fund': "Қаржыландыру", 'c_conflict': "Мүдделер қақтығысы",
        'c_refs': "Әдебиеттер саны", 'c_refs_req': "≥ 25",
        'c_doi': "Сілтемелердегі DOI", 'c_apa': "Дәйексөз келтіру стилі", 'c_apa_req': "APA 7 (Автор, жыл)",

        # Технические критерии
        'c_paper': "Қағаз форматы", 'c_paper_req': "A4 (210x297 мм)",
        'c_margins': "Жақтаулар (Поля)", 'c_margins_req': "Әдетте 2 см",
        'c_font': "Шрифт және өлшемі (Кегль)", 'c_font_req': "Times New Roman, 11/12 pt",
        'c_tables': "Кестелер", 'c_tables_req': "Мәтіндегі сілтемелер болуы",
        'c_images': "Суреттер (DPI)", 'c_images_req': "Мин. 300 DPI (қолмен тексеріңіз)",

        'found': "Табылды", 'not_found': "Жоқ", 'words': "сөз",
    }
}

l = locales[st.session_state.lang]

# ─── ИНТЕРФЕЙС ЖӘНЕ ТАҚЫРЫПТАР (CSS) ──────────────────────────────
dark_css = """
<style>
.stApp { background-color: #121212; color: #E0E0E0; }
.stMetric { background: #1E1E1E; border: 1px solid #333; color: #FFF;
            padding:12px; border-radius:10px; }
h1, h2, h3, h4, h5, h6 { color: #E0E0E0 !important; }
</style>
"""

light_css = """
<style>
.stMetric {
    background: #ffffff;
    padding:12px;
    border-radius:10px;
    box-shadow:0 2px 6px rgba(0,0,0,0.08);
}
h1, h2, h3 { color:#1a3a5c; }
</style>
"""

st.markdown(
    dark_css if st.session_state.theme == 'dark' else light_css,
    unsafe_allow_html=True
)

# ─── КНОПКАЛАРДЫ БАСҚАРУ БЛОГЫ (TOP RIGHT) ────────────────────────
col1, col2, col3 = st.columns([7, 1.5, 1.5])
with col2:
    if st.button(l['btn_lang'], use_container_width=True):
        st.session_state.lang = 'kz' if st.session_state.lang == 'ru' else 'ru'
        st.rerun()
with col3:
    theme_btn_text = (
        l['btn_theme_light'] if st.session_state.theme == 'dark'
        else l['btn_theme_dark']
    )
    if st.button(theme_btn_text, use_container_width=True):
        st.session_state.theme = (
            'light' if st.session_state.theme == 'dark' else 'dark'
        )
        st.rerun()

st.title(l['title'])
st.caption(l['subtitle'])
st.markdown("---")

# ─── ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ИМЯ+ФАМИЛИЯ ПЕРВОГО АВТОРА ──────────
def extract_first_author_name_surname(doc: Document) -> str:
    """
    Возвращает строку 'ИмяФамилия' или 'ФамилияИмя' первого автора
    (без пробела, чтобы удобно было использовать в имени файла).
    Если не получилось — 'report'.
    """
    header_text = "\n".join(p.text for p in doc.paragraphs[:5])
    header_text = re.sub(r"\^[0-9, ]+\^", " ", header_text)

    m = re.search(r"([А-ЯA-Z][а-яa-z]+)\s+([А-ЯA-Z][а-яa-z]+)", header_text)
    if not m:
        return "report"

    w1 = re.sub(r"[^А-Яа-яA-Za-z-]", "", m.group(1))
    w2 = re.sub(r"[^А-Яа-яA-Za-z-]", "", m.group(2))

    if not (w1 and w2):
        return "report"

    return f"{w1}{w2}"

# ─── АНАЛИЗ ФУНКЦИЯСЫ ─────────────────────────────────────────────
def check_article(doc, l):
    full_text = "\n".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    results = []

    def add(num, criterion, requirement, found, status):
        results.append({
            "№": num,
            "Критерий": criterion,
            "Требование": requirement,
            "Найдено": found,
            "Статус": status
        })

    # 1. Объём
    add(1, l['c_vol'], l['c_vol_req'], f"{word_count} {l['words']}",
        "✅" if word_count >= 3500 else "❌")

    # 2-4. Аннотации
    abstract_ru = re.search(
        r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін)",
        full_text,
        re.IGNORECASE | re.DOTALL
    )
    if abstract_ru:
        aw = len(abstract_ru.group(1).split())
        add(2, l['c_ann_ru'], l['c_ann_req'],
            f"{aw} {l['words']}", "✅" if aw <= 300 else "❌")
    else:
        add(2, l['c_ann_ru'], l['c_ann_req'], l['not_found'], "⚠️")

    has_kaz = "аңдатпа" in full_text.lower()
    add(3, l['c_ann_kz'], l['c_req_obl'],
        l['found'] if has_kaz else l['not_found'],
        "✅" if has_kaz else "❌")

    has_eng = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(4, l['c_ann_en'], l['c_req_obl'],
        l['found'] if has_eng else l['not_found'],
        "✅" if has_eng else "❌")

    # 5. Ключевые слова
    kw_match = re.search(
        r"(ключевые слова|keywords|түйінді сөздер)[:\s]+(.+?)(\n|$)",
        full_text, re.IGNORECASE
    )
    if kw_match:
        kw_list = [k.strip() for k in kw_match.group(2).split(";") if k.strip()]
        add(5, l['c_kw'], l['c_kw_req'], f"{len(kw_list)} {l['words']}",
            "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(5, l['c_kw'], l['c_kw_req'], l['not_found'], "⚠️")

    # 6-7. МРНТИ, ORCID
    mrnt = bool(re.search(r"МРНТИ|\d{2}\.\d{2}\.\d{2}", full_text))
    add(6, l['c_mrnti'], l['c_req_obl'],
        l['found'] if mrnt else l['not_found'],
        "✅" if mrnt else "⚠️")

    orcid_count = len(re.findall(
        r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}", full_text, re.IGNORECASE
    ))
    add(7, l['c_orcid'], l['c_orcid_req'],
        f"{orcid_count} ORCID",
        "✅" if orcid_count >= 1 else "⚠️")

    # 8-12. Разделы
    sections = [
        (8,  l['c_intro'], ["введени", "кіріспе", "introduction"]),
        (9,  l['c_mm'],    ["материал", "әдіс", "methods"]),
        (10, l['c_res'],   ["результат", "нәтиже", "results"]),
        (11, l['c_disc'],  ["обсужден", "талқылау", "discussion"]),
        (12, l['c_concl'], ["заключени", "қорытынды", "conclusion"]),
    ]
    for num, name, keys in sections:
        found = any(k in full_text.lower() for k in keys)
        add(num, name, l['c_req_obl'],
            l['found'] if found else l['not_found'],
            "✅" if found else "❌")

    # 13-15. Этика
    contrib = any(k in full_text.lower()
                  for k in ["вклад авторов", "author contribution", "авторлардың үлесі"])
    add(13, l['c_contrib'], "CRediT",
        l['found'] if contrib else l['not_found'],
        "✅" if contrib else "❌")

    fund = any(k in full_text.lower()
               for k in ["финансирован", "funding", "қаржыландыру"])
    add(14, l['c_fund'], l['c_req_obl'],
        l['found'] if fund else l['not_found'],
        "✅" if fund else "❌")

    conflict = any(k in full_text.lower()
                   for k in ["конфликт интересов",
                             "conflict of interest",
                             "мүдделер қақтығысы"])
    add(15, l['c_conflict'], l['c_req_obl'],
        l['found'] if conflict else l['not_found'],
        "✅" if conflict else "❌")

    # 16-18. Литература
    total_refs = len(re.findall(r"(?m)^\d+\.\s+", full_text))
    add(16, l['c_refs'], l['c_refs_req'],
        f"~{total_refs}", "✅" if total_refs >= 25 else "❌")

    doi_count = len(re.findall(r"https?://doi\.org/", full_text))
    add(17, l['c_doi'], l['c_req_obl'],
        f"{doi_count} DOI", "✅" if doi_count >= 5 else "⚠️")

    apa_style = bool(re.search(
        r"\([A-ZА-Я][a-zA-Zа-яА-Я]+.*?\d{4}\)", full_text
    ))
    add(18, l['c_apa'], l['c_apa_req'],
        l['found'] if apa_style else l['not_found'],
        "✅" if apa_style else "⚠️")

    # 19-20. Формат бумаги и поля
    try:
        sec = doc.sections[0]
        w_mm, h_mm = round(sec.page_width.mm), round(sec.page_height.mm)
        is_a4 = (209 <= w_mm <= 211) and (296 <= h_mm <= 298)
        add(19, l['c_paper'], l['c_paper_req'],
            f"{w_mm}x{h_mm} мм", "✅" if is_a4 else "❌")

        t, b, lf, rg = (
            round(sec.top_margin.mm),
            round(sec.bottom_margin.mm),
            round(sec.left_margin.mm),
            round(sec.right_margin.mm)
        )
        add(20, l['c_margins'], l['c_margins_req'],
            f"Л:{lf} П:{rg} В:{t} Н:{b} мм",
            "✅" if lf >= 20 else "⚠️")
    except Exception:
        add(19, l['c_paper'], l['c_paper_req'], "Қате / Ошибка", "⚠️")
        add(20, l['c_margins'], l['c_margins_req'], "Қате / Ошибка", "⚠️")

    # 21. Шрифт
    try:
        font_name = doc.styles['Normal'].font.name or "Анықталмады"
        font_size_pt = (
            doc.styles['Normal'].font.size.pt
            if doc.styles['Normal'].font.size else "Анықталмады"
        )
        is_correct_font = (
            "Times New Roman" in str(font_name)
            and font_size_pt in [11.0, 12.0]
        )
        add(21, l['c_font'], l['c_font_req'],
            f"{font_name}, {font_size_pt} pt",
            "✅" if is_correct_font else "⚠️")
    except Exception:
        add(21, l['c_font'], l['c_font_req'], "Қате / Ошибка", "⚠️")

    # 22-23. Таблицы и Рисунки
    tbl_count = len(doc.tables)
    add(22, l['c_tables'], l['c_tables_req'],
        f"{tbl_count} шт.", "✅" if tbl_count > 0 else "⚠️")

    img_count = len(doc.inline_shapes)
    add(23, l['c_images'], l['c_images_req'],
        f"{img_count} шт.", "⚠️" if img_count > 0 else "✅")

    return results, full_text

# ─── ПОСТРОЕНИЕ DOCX-ОТЧЁТА ────────────────────────────────────────
def build_docx_report(results, l, total, passed, warned, failed, score):
    buf = BytesIO()
    d = Document()
    d.add_heading(l['res_title'], level=1)

    p = d.add_paragraph()
    p.add_run(f"{l['total']}: {total}, ").bold = True
    p.add_run(f"{l['passed']}: {passed}, ")
    p.add_run(f"{l['warned']}: {warned}, ")
    p.add_run(f"{l['failed']}: {failed}, ")
    p.add_run(f"{l['score']}: {score}%")

    d.add_paragraph("")

    table = d.add_table(rows=1, cols=5)
    hdr = table.rows[0].cells
    hdr[0].text = "№"
    hdr[1].text = "Критерий"
    hdr[2].text = "Требование"
    hdr[3].text = "Найдено"
    hdr[4].text = "Статус"

    for r in results:
        row_cells = table.add_row().cells
        row_cells[0].text = str(r["№"])
        row_cells[1].text = str(r["Критерий"])
        row_cells[2].text = str(r["Требование"])
        row_cells[3].text = str(r["Найдено"])
        row_cells[4].text = str(r["Статус"])

    d.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── ФАЙЛДЫ ЖҮКТЕУ ЖӘНЕ ӨҢДЕУ ──────────────────────────────────────
uploaded_file = st.file_uploader(
    l['upload_title'], type=["docx"], help=l['upload_help']
)

if uploaded_file:
    with st.spinner(l['analyzing']):
        doc = Document(uploaded_file)
        results, full_text = check_article(doc, l)
        df = pd.DataFrame(results)
        first_author = extract_first_author_name_surname(doc)

    passed = sum(1 for r in results if r["Статус"] == "✅")
    warned = sum(1 for r in results if r["Статус"] == "⚠️")
    failed = sum(1 for r in results if r["Статус"] == "❌")
    total = len(results)
    score = int(passed / total * 100) if total > 0 else 0

    st.markdown(f"## {l['res_title']}")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(l['total'], total)
    c2.metric(l['passed'], passed)
    c3.metric(l['warned'], warned)
    c4.metric(l['failed'], failed)
    c5.metric(l['score'], f"{score}%")

    bar_color = "#4caf50" if score >= 80 else "#ffc107" if score >= 60 else "#f44336"
    bg_bar = "#2b2b2b" if st.session_state.theme == 'dark' else "#e9ecef"
    st.markdown(
        f"""
        <div style="background:{bg_bar};border-radius:10px;
                    height:28px;margin:8px 0 20px 0;">
          <div style="background:{bar_color};width:{score}%;height:28px;
                      border-radius:10px;display:flex;align-items:center;
                      justify-content:center;color:white;font-weight:bold;">
            {score}%
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    def highlight(row):
        if st.session_state.theme == 'dark':
            colors = {
                "✅": "background-color:#1b5e20",
                "⚠️": "background-color:#795548",
                "❌": "background-color:#b71c1c"
            }
        else:
            colors = {
                "✅": "background-color:#d4edda",
                "⚠️": "background-color:#fff3cd",
                "❌": "background-color:#f8d7da"
            }
        return [colors.get(row["Статус"], "")] * len(row)

    st.markdown(l['det_report'])
    st.dataframe(
        df.style.apply(highlight, axis=1),
        use_container_width=True,
        height=850
    )

    st.markdown("---")
    col_a, col_b, col_c = st.columns(3)

    base_name = f"compliance_{first_author}"

    # CSV
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    col_a.download_button(
        l['btn_csv'],
        csv_bytes,
        f"{base_name}.csv",
        "text/csv"
    )

    # Excel
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    col_b.download_button(
        l['btn_xls'],
        excel_buf.getvalue(),
        f"{base_name}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Word
    docx_bytes = build_docx_report(results, l, total, passed, warned, failed, score)
    col_c.download_button(
        l['btn_docx'],
        docx_bytes,
        f"{base_name}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    problems = [r for r in results if r["Статус"] in ("❌", "⚠️")]
    if problems:
        st.markdown(l['req_fix'])
        for p in problems:
            icon = "🔴" if p["Статус"] == "❌" else "🟡"
            st.markdown(
                f"{icon} **{p['Критерий']}** — {p['Найдено']} "
                f"*({l['req']}: {p['Требование']})*"
            )
else:
    st.info(l['no_file'])
