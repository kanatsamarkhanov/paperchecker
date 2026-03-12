import streamlit as st
from docx import Document
import re
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Чекер статьи / Мақаланы тексеру",
    page_icon="📋",
    layout="wide"
)

if "lang" not in st.session_state:
    st.session_state.lang = "kz"
if "theme" not in st.session_state:
    st.session_state.theme = "light"

# ─── LOCALES ─────────────────────────────────────────────────────
locales = {
    "ru": {
        "title":          "📋 Автоматическая проверка статьи",
        "subtitle":       "Вестник ЕНУ им. Л.Н. Гумилева · Серия: Химия / География · Шаблон 2025",
        "btn_lang":       "🇰🇿 ҚАЗ",
        "btn_theme_dark":  "🌙 Тёмная тема",
        "btn_theme_light": "☀️ Светлая тема",
        "upload_title":   "📂 Загрузите статью в формате .docx",
        "upload_help":    "Шаблон журнала Вестник ЕНУ, серия Химия/География, 2025",
        "analyzing":      "Анализируем статью...",
        "res_title":      "📊 Результаты проверки",
        "total":  "Всего",
        "passed": "✅ Выполнено",
        "warned": "⚠️ Внимание",
        "failed": "❌ Не выпол.",
        "score":  "🏆 Соответствие",
        "det_report": "### 📋 Детальный отчёт",
        "btn_csv":  "⬇️ Скачать CSV",
        "btn_xls":  "⬇️ Скачать Excel",
        "btn_docx": "⬇️ Word (DOCX)",
        "req_fix":  "### ⚠️ Требует исправления",
        "req":      "требование",
        "no_file":  "👆 Загрузите .docx файл, чтобы начать проверку",
        "c_title":     "Наименование статьи",
        "c_title_req": "Строки 3–5 документа",
        "c_lang":      "Язык статьи",
        "c_lang_req":  "По преобладанию символов",
        "c_vol":       "Объём статьи",
        "c_vol_req":   "≥3500 слов",
        "c_ann_ru":    "Аннотация (рус)",
        "c_ann_req":   "≤300 слов",
        "c_ann_kz":    "Аннотация (каз)",
        "c_ann_en":    "Abstract (англ)",
        "c_req_obl":   "Обязательно",
        "c_kw":        "Ключевые слова",
        "c_kw_req":    "3–10, разделитель «;»",
        "c_mrnti":     "Код МРНТИ / IRSTI",
        "c_orcid":     "ORCID авторов",
        "c_orcid_req": "Для каждого автора",
        "c_intro":   "§1. Введение / Introduction",
        "c_mm":      "§2. Материалы и методы",
        "c_res":     "§3. Результаты / Results",
        "c_disc":    "§4. Обсуждение / Discussion",
        "c_concl":   "§5. Заключение / Conclusion",
        "c_supp":    "§6. Вспомогательный материал",
        "c_contrib": "§7. Вклад авторов",
        "c_authinfo":"§8. Информация об авторе",
        "c_fund":    "§9. Финансирование",
        "c_ack":     "§10. Благодарности",
        "c_conflict":"§11. Конфликты интересов",
        "c_paper":       "Формат бумаги",
        "c_paper_req":   "A4 (210x297 мм)",
        "c_margins":     "Поля",
        "c_margins_req": "Все поля 20 мм",
        "c_font":        "Шрифт и кегль",
        "c_font_req":    "Times New Roman, 12 pt",
        "c_tables":      "Таблицы",
        "c_tables_req":  "Должны быть в тексте",
        "c_images":      "Рисунки (DPI)",
        "c_images_req":  "Мин. 300 DPI (вручную)",
        "c_multi_ann":     "Многоязычные аннотации",
        "c_multi_ann_req": "Ещё 2 аннотации на других языках",
        "found":     "Найдено",
        "not_found": "Отсутствует",
        "words":     "слов",
        "f_author":  "Канат Самарханов / Kanat Samarkhanov",
        "f_license": "Лицензия",
        "f_univ":    "ЕНУ им. Л.Н. Гумилева — Кафедра физической и экономической географии",
    },
    "kz": {
        "title":          "📋 Мақаланы автоматты түрде тексеру",
        "subtitle":       "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы · Серия: Химия / География · 2025 үлгісі",
        "btn_lang":       "🇷🇺 РУС",
        "btn_theme_dark":  "🌙 Түнгі режим",
        "btn_theme_light": "☀️ Күндізгі режим",
        "upload_title":   "📂 .docx форматындағы мақаланы жүктеңіз",
        "upload_help":    "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы, Химия/География сериясы, 2025 үлгісі",
        "analyzing":      "Мақала талдануда...",
        "res_title":      "📊 Тексеру нәтижелері",
        "total":  "Барлығы",
        "passed": "✅ Орындалды",
        "warned": "⚠️ Назар аударыңыз",
        "failed": "❌ Орындалмады",
        "score":  "🏆 Сәйкестік",
        "det_report": "### 📋 Толық есеп",
        "btn_csv":  "⬇️ CSV жүктеу",
        "btn_xls":  "⬇️ Excel жүктеу",
        "btn_docx": "⬇️ Word (DOCX)",
        "req_fix":  "### ⚠️ Түзетуді қажет етеді",
        "req":      "талап",
        "no_file":  "👆 Тексеруді бастау үшін .docx файлын жүктеңіз",
        "c_title":     "Мақаланың атауы",
        "c_title_req": "Құжаттың 3–5 жолдары",
        "c_lang":      "Мақала тілі",
        "c_lang_req":  "Таңбалардың басымдылығы бойынша",
        "c_vol":       "Мақала көлемі",
        "c_vol_req":   "≥3500 сөз",
        "c_ann_ru":    "Аңдатпа (орыс)",
        "c_ann_req":   "≤300 сөз",
        "c_ann_kz":    "Аңдатпа (қаз)",
        "c_ann_en":    "Abstract (ағылш)",
        "c_req_obl":   "Міндетті түрде",
        "c_kw":        "Түйінді сөздер",
        "c_kw_req":    "3–10, бөлгіш «;»",
        "c_mrnti":     "МРНТИ / IRSTI коды",
        "c_orcid":     "Авторлардың ORCID",
        "c_orcid_req": "Әр автор үшін",
        "c_intro":   "§1. Кіріспе / Introduction",
        "c_mm":      "§2. Материалдар мен әдістер",
        "c_res":     "§3. Нәтижелер / Results",
        "c_disc":    "§4. Талқылау / Discussion",
        "c_concl":   "§5. Қорытынды / Conclusion",
        "c_supp":    "§6. Қосымша материалдар",
        "c_contrib": "§7. Авторлардың үлесі",
        "c_authinfo":"§8. Автор туралы ақпарат",
        "c_fund":    "§9. Қаржыландыру",
        "c_ack":     "§10. Алғыстар",
        "c_conflict":"§11. Мүдделер қақтығысы",
        "c_paper":       "Қағаз форматы",
        "c_paper_req":   "A4 (210x297 мм)",
        "c_margins":     "Жақтаулар",
        "c_margins_req": "Барлық жақтаулар 20 мм",
        "c_font":        "Шрифт және кегль",
        "c_font_req":    "Times New Roman, 12 pt",
        "c_tables":      "Кестелер",
        "c_tables_req":  "Мәтінде болуы керек",
        "c_images":      "Суреттер (DPI)",
        "c_images_req":  "Мин. 300 DPI (қолмен тексеру)",
        "c_multi_ann":     "Көптілді аңдатпалар",
        "c_multi_ann_req": "Басқа 2 тілде аңдатпа болуы керек",
        "found":     "Табылды",
        "not_found": "Жоқ",
        "words":     "сөз",
        "f_author":  "Канат Самарханов / Kanat Samarkhanov",
        "f_license": "Лицензия",
        "f_univ":    "Л.Н. Гумилев атындағы ЕҰУ — Физикалық және экономикалық география кафедрасы",
    },
}

l = locales[st.session_state.lang]

# ─── THEME CSS ────────────────────────────────────────────────────
dark_css = '''
<style>
html, body, [class*='css'], .stApp {
    background-color: #0d1117 !important;
    color: #c9d1d9 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif !important;
}
h1,h2,h3,h4,h5,h6,
[data-testid='stMarkdownContainer'] h1,
[data-testid='stMarkdownContainer'] h2,
[data-testid='stMarkdownContainer'] h3 {
    color: #e6edf3 !important; font-weight: 600 !important;
}
p, span, label, div, li,
[data-testid='stMarkdownContainer'] p,
[data-testid='stCaptionContainer'], .stCaption { color: #c9d1d9 !important; }
[data-testid='block-container'],
[data-testid='stVerticalBlock'],
section[data-testid='stSidebar'] { background-color: #0d1117 !important; }
[data-testid='stMetric'] {
    background: #161b22 !important; border: 1px solid #30363d !important;
    border-radius: 6px !important; padding: 12px 16px !important;
}
[data-testid='stMetricValue'] { color: #e6edf3 !important; }
[data-testid='stMetricLabel'] { color: #8b949e !important; }
.stButton > button {
    background-color: #21262d !important; color: #c9d1d9 !important;
    border: 1px solid #30363d !important; border-radius: 6px !important;
}
.stButton > button:hover {
    background-color: #30363d !important; border-color: #8b949e !important; color: #e6edf3 !important;
}
[data-testid='stDownloadButton'] > button {
    background-color: #238636 !important; color: #ffffff !important;
    border: 1px solid #2ea043 !important; border-radius: 6px !important;
}
[data-testid='stDownloadButton'] > button:hover { background-color: #2ea043 !important; }
[data-testid='stFileUploader'] {
    background-color: #161b22 !important; border: 1px dashed #30363d !important; border-radius: 6px !important;
}
[data-testid='stFileUploader'] * { color: #c9d1d9 !important; }
[data-testid='stDataFrame'], .stDataFrame iframe { border: 1px solid #30363d !important; border-radius: 6px !important; }
[data-testid='stAlert'] {
    background-color: #161b22 !important; border: 1px solid #1f6feb !important;
    color: #c9d1d9 !important; border-radius: 6px !important;
}
.stSpinner > div { color: #c9d1d9 !important; }
hr { border-color: #30363d !important; }
input, textarea, select {
    background-color: #161b22 !important; color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
}
</style>
'''

light_css = '''
<style>
[data-testid='stMetric'] {
    background: #ffffff; padding: 12px; border-radius: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
}
h1,h2,h3 { color: #1a3a5c; }
[data-testid='stDownloadButton'] > button {
    background-color: #2ea043; color: #ffffff; border-radius: 6px;
}
</style>
'''

st.markdown(dark_css if st.session_state.theme == 'dark' else light_css, unsafe_allow_html=True)

# ─── TOP BUTTONS ──────────────────────────────────────────────────
col1, col2, col3 = st.columns([7, 1.5, 1.5])
with col2:
    if st.button(l["btn_lang"], use_container_width=True):
        st.session_state.lang = "kz" if st.session_state.lang == "ru" else "ru"
        st.rerun()
with col3:
    tbtn = l["btn_theme_light"] if st.session_state.theme == "dark" else l["btn_theme_dark"]
    if st.button(tbtn, use_container_width=True):
        st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"
        st.rerun()

st.title(l["title"])
st.caption(l["subtitle"])
st.markdown("---")

# ─── HELPERS ──────────────────────────────────────────────────────
_KAZ_CHARS = set("қңөұүіәғҚҢӨҰҮІӘҒ")

def detect_lang_from_text(text: str) -> str:
    latin = sum(1 for c in text if c.isalpha() and c.isascii())
    cyr   = sum(1 for c in text if "\u0400" <= c <= "\u04FF")
    kaz   = sum(1 for c in text if c in _KAZ_CHARS)
    if latin > cyr:
        return "en"
    if kaz >= 1:
        return "kz"
    return "ru"

def extract_title_and_lang(doc: Document):
    non_empty = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    candidates = non_empty[2:5] if len(non_empty) >= 3 else non_empty
    title = max(candidates, key=len) if candidates else ""
    return title, detect_lang_from_text(title)

def title_for_filename(title: str) -> str:
    if not title:
        return "report"
    import re as _re
    clean = _re.sub(r"[^\w\s-]", "", title, flags=_re.UNICODE)
    return clean.strip().replace(" ", "_")[:40] or "report"

_CONFLICT_KW = [
    "конфликт интерес",
    "conflict of interest",
    "мүдделер қақтығыс",
]

def has_conflict_section(doc: Document, full_text_low: str) -> bool:
    if any(kw in full_text_low for kw in _CONFLICT_KW):
        return True
    for par in doc.paragraphs:
        t = par.text.lower()
        if any(kw in t for kw in _CONFLICT_KW):
            return True
    return False

# ─── MAIN CHECK FUNCTION ──────────────────────────────────────────
def check_article(doc: Document, l: dict):
    import re
    full_text  = "\n".join(p.text for p in doc.paragraphs)
    word_count = len(full_text.split())
    text_low   = full_text.lower()
    results    = []
    title, main_lang = extract_title_and_lang(doc)

    def add(num, criterion, requirement, found_val, status):
        results.append({
            "№": num, "Критерий": criterion, "Требование": requirement,
            "Найдено": found_val, "Статус": status,
        })

    add(0, l["c_title"], l["c_title_req"],
        title if title else l["not_found"], "✅" if title else "⚠️")

    lang_map = {"ru": "Русский", "kz": "Қазақша", "en": "English"}
    add(1, l["c_lang"], l["c_lang_req"], lang_map.get(main_lang, main_lang), "✅")

    add(2, l["c_vol"], l["c_vol_req"], f"{word_count} {l['words']}",
        "✅" if word_count >= 3500 else "⚠️")

    abstract_ru = re.search(
        r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін|abstract)",
        full_text, re.IGNORECASE | re.DOTALL,
    )
    if abstract_ru:
        aw = len(abstract_ru.group(1).split())
        add(3, l["c_ann_ru"], l["c_ann_req"], f"{aw} {l['words']}", "✅" if aw <= 300 else "❌")
        has_ru_ann = True
    else:
        add(3, l["c_ann_ru"], l["c_ann_req"], l["not_found"], "⚠️")
        has_ru_ann = False

    has_kaz_ann = "аңдатпа" in text_low or "аннотация (қаз" in text_low
    add(4, l["c_ann_kz"], l["c_req_obl"], l["found"] if has_kaz_ann else l["not_found"], "✅" if has_kaz_ann else "❌")

    has_eng_ann = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(5, l["c_ann_en"], l["c_req_obl"], l["found"] if has_eng_ann else l["not_found"], "✅" if has_eng_ann else "❌")

    kw_match = re.search(
        r"(ключевые слова|keywords|түйінді сөздер|түйін сөздер)[:\s]+(.+?)(\n|$)",
        full_text, re.IGNORECASE,
    )
    if kw_match:
        kw_list = [k.strip() for k in kw_match.group(2).split(";") if k.strip()]
        add(6, l["c_kw"], l["c_kw_req"], f"{len(kw_list)}", "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(6, l["c_kw"], l["c_kw_req"], l["not_found"], "⚠️")

    mrnt = bool(re.search(r"МРНТИ|IRSTI|\d{2}\.\d{2}\.\d{2}", full_text))
    add(7, l["c_mrnti"], l["c_req_obl"], l["found"] if mrnt else l["not_found"], "✅" if mrnt else "⚠️")

    orcid_count = len(re.findall(r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}", full_text, re.IGNORECASE))
    add(8, l["c_orcid"], l["c_orcid_req"], f"{orcid_count} ORCID", "✅" if orcid_count >= 1 else "⚠️")

    for num, name, keys in [
        (9,  l["c_intro"], ["введение", "кіріспе", "introduction"]),
        (10, l["c_mm"],    ["материалы и методы", "материалдар мен әдістер", "materials and methods", "материал", "әдістер"]),
        (11, l["c_res"],   ["результаты", "нәтижелер", "results"]),
        (12, l["c_disc"],  ["обсуждение", "талқылау", "discussion"]),
        (13, l["c_concl"], ["заключение", "қорытынды", "conclusion"]),
    ]:
        found = any(k in text_low for k in keys)
        add(num, name, l["c_req_obl"], l["found"] if found else l["not_found"], "✅" if found else "❌")

    has_supp = any(k in text_low for k in ["вспомогательный материал", "қосымша материалдар", "supplementary materials"])
    add(14, l["c_supp"], l["c_req_obl"], l["found"] if has_supp else l["not_found"], "✅" if has_supp else "⚠️")

    contrib = any(k in text_low for k in ["вклад авторов", "author contributions", "авторлардың үлесі", "авторлық үлестер"])
    add(15, l["c_contrib"], "CRediT", l["found"] if contrib else l["not_found"], "✅" if contrib else "❌")

    authinfo = any(k in text_low for k in ["информация об авторе", "author information", "автор туралы ақпарат"])
    add(16, l["c_authinfo"], l["c_req_obl"], l["found"] if authinfo else l["not_found"], "✅" if authinfo else "⚠️")

    fund = any(k in text_low for k in ["финансирование", "funding", "қаржыландыру"])
    add(17, l["c_fund"], l["c_req_obl"], l["found"] if fund else l["not_found"], "✅" if fund else "❌")

    ack = any(k in text_low for k in ["благодарност", "acknowledgements", "acknowledgments", "алғыстар"])
    add(18, l["c_ack"], l["c_req_obl"], l["found"] if ack else l["not_found"], "✅" if ack else "⚠️")

    conflict = has_conflict_section(doc, text_low)
    add(19, l["c_conflict"], l["c_req_obl"], l["found"] if conflict else l["not_found"], "✅" if conflict else "❌")

    refs_match = None
    for p in [r"список литературы", r"references", r"әдебиет(тер)? тізімі"]:
        m = re.search(p, text_low)
        if m:
            refs_match = m
            break
    if refs_match:
        refs_text = full_text[refs_match.end():]
        ref_lines = re.findall(r"\n\s*(\d+[\.\)]|\[\d+\])\s", refs_text)
        if not ref_lines:
            ref_lines = [ln for ln in refs_text.split("\n") if len(ln.strip()) > 40]
        rc = len(ref_lines)
        add(20, "Список литературы / References", "≥10 источников", f"{rc}", "✅" if rc >= 10 else "⚠️")
    else:
        add(20, "Список литературы / References", l["c_req_obl"], l["not_found"], "❌")

    try:
        sec   = doc.sections[0]
        w_mm  = round(sec.page_width.mm)
        h_mm  = round(sec.page_height.mm)
        is_a4 = (209 <= w_mm <= 211) and (296 <= h_mm <= 298)
        add(21, l["c_paper"], l["c_paper_req"], f"{w_mm}x{h_mm} мм", "✅" if is_a4 else "❌")
        t, b   = round(sec.top_margin.mm), round(sec.bottom_margin.mm)
        lf, rg = round(sec.left_margin.mm), round(sec.right_margin.mm)
        add(22, l["c_margins"], l["c_margins_req"], f"Л:{lf} П:{rg} В:{t} Н:{b} мм",
            "✅" if (t == 20 and b == 20 and lf == 20 and rg == 20) else "❌")
    except Exception:
        add(21, l["c_paper"],   l["c_paper_req"],  "Қате/Ошибка", "⚠️")
        add(22, l["c_margins"], l["c_margins_req"], "Қате/Ошибка", "⚠️")

    try:
        fn = doc.styles["Normal"].font.name or "?"
        fs = doc.styles["Normal"].font.size.pt if doc.styles["Normal"].font.size else "?"
        ok_font = "Times New Roman" in str(fn) and fs in [11.0, 12.0]
        add(23, l["c_font"], l["c_font_req"], f"{fn}, {fs} pt", "✅" if ok_font else "⚠️")
    except Exception:
        add(23, l["c_font"], l["c_font_req"], "Қате/Ошибка", "⚠️")

    tbl_count = len(doc.tables)
    add(24, l["c_tables"], l["c_tables_req"], f"{tbl_count} шт.", "✅" if tbl_count > 0 else "⚠️")
    img_count = len(doc.inline_shapes)
    add(25, l["c_images"], l["c_images_req"], f"{img_count} шт.", "⚠️" if img_count > 0 else "✅")

    if main_lang == "ru":
        ok_multi = has_kaz_ann and has_eng_ann
    elif main_lang == "kz":
        ok_multi = has_ru_ann and has_eng_ann
    else:
        ok_multi = has_ru_ann and has_kaz_ann
    add(26, l["c_multi_ann"], l["c_multi_ann_req"], l["found"] if ok_multi else l["not_found"], "✅" if ok_multi else "❌")

    return results, full_text, title, main_lang

# ─── DOCX REPORT ─────────────────────────────────────────────────
def build_docx_report(results, l, total, passed, warned, failed, score):
    buf = BytesIO()
    d = Document()
    d.add_heading(l["res_title"], level=1)
    p = d.add_paragraph()
    p.add_run(f"{l['total']}: {total},  ").bold = True
    p.add_run(f"{l['passed']}: {passed},  {l['warned']}: {warned},  {l['failed']}: {failed},  {l['score']}: {score}%")
    d.add_paragraph("")
    tbl = d.add_table(rows=1, cols=5)
    for i, h in enumerate(["№", "Критерий", "Требование", "Найдено", "Статус"]):
        tbl.rows[0].cells[i].text = h
    for r in results:
        row = tbl.add_row().cells
        for i, key in enumerate(["№", "Критерий", "Требование", "Найдено", "Статус"]):
            row[i].text = str(r[key])
    d.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── UI ──────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(l["upload_title"], type=["docx"], help=l["upload_help"])

if uploaded_file:
    with st.spinner(l["analyzing"]):
        doc = Document(uploaded_file)
        results, full_text, title, main_lang = check_article(doc, l)
        df = pd.DataFrame(results)

    passed = sum(1 for r in results if r["Статус"] == "✅")
    warned = sum(1 for r in results if r["Статус"] == "⚠️")
    failed = sum(1 for r in results if r["Статус"] == "❌")
    total  = len(results)
    score  = int(passed / total * 100) if total > 0 else 0

    st.markdown(f"## {l['res_title']}")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(l["total"],  total)
    c2.metric(l["passed"], passed)
    c3.metric(l["warned"], warned)
    c4.metric(l["failed"], failed)
    c5.metric(l["score"],  f"{score}%")

    bar_color = "#238636" if score >= 80 else "#d29922" if score >= 60 else "#da3633"
    bg_bar    = "#161b22" if st.session_state.theme == "dark" else "#e9ecef"
    txt_color = "#e6edf3" if st.session_state.theme == "dark" else "#ffffff"
    st.markdown(
        f"""<div style="background:{bg_bar};border:1px solid #30363d;border-radius:6px;height:28px;margin:8px 0 20px 0;">
          <div style="background:{bar_color};width:{score}%;height:28px;border-radius:6px;
                      display:flex;align-items:center;justify-content:center;
                      color:{txt_color};font-weight:600;font-size:13px;">{score}%</div></div>""",
        unsafe_allow_html=True,
    )

    def highlight(row):
        if st.session_state.theme == "dark":
            colors = {
                "✅": "background-color:#1a4a1a;color:#3fb950",
                "⚠️": "background-color:#3d2e00;color:#d29922",
                "❌": "background-color:#3d0e0e;color:#f85149",
            }
        else:
            colors = {
                "✅": "background-color:#d4edda",
                "⚠️": "background-color:#fff3cd",
                "❌": "background-color:#f8d7da",
            }
        return [colors.get(row["Статус"], "")] * len(row)

    st.markdown(l["det_report"])
    st.dataframe(df.style.apply(highlight, axis=1), use_container_width=True, height=900)
    st.markdown("---")
    col_a, col_b, col_c = st.columns(3)
    base_name = f"compliance_{title_for_filename(title)}"
    col_a.download_button(
        l["btn_csv"], df.to_csv(index=False).encode("utf-8-sig"), f"{base_name}.csv", "text/csv",
    )
    xbuf = BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Report")
    col_b.download_button(
        l["btn_xls"], xbuf.getvalue(), f"{base_name}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    col_c.download_button(
        l["btn_docx"],
        build_docx_report(results, l, total, passed, warned, failed, score),
        f"{base_name}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    problems = [r for r in results if r["Статус"] in ("❌", "⚠️")]
    if problems:
        st.markdown(l["req_fix"])
        for prob in problems:
            icon = "🔴" if prob["Статус"] == "❌" else "🟡"
            st.markdown(
                f"{icon} **{prob['Критерий']}** — {prob['Найдено']} *({l['req']}: {prob['Требование']})*"
            )
else:
    st.info(l["no_file"])

# ─── FOOTER ──────────────────────────────────────────────────────
fc  = "#8b949e" if st.session_state.theme == "dark" else "#555555"
flk = "#58a6ff" if st.session_state.theme == "dark" else "#0969da"
st.markdown("---")
st.markdown(
    f"""
<div style="text-align:center;font-size:12px;color:{fc};padding:12px 0 20px 0;line-height:2.2;">
  <b style="font-size:13px;">© 2025 {l["f_author"]}</b><br>
  📧 <a href="mailto:samarkhanov_kb@enu.kz" style="color:{flk};text-decoration:none;">samarkhanov_kb@enu.kz</a>
  &nbsp;·&nbsp;
  <a href="mailto:kanat.baurzhanuly@gmail.com" style="color:{flk};text-decoration:none;">kanat.baurzhanuly@gmail.com</a><br>
  🏛️ <a href="https://fns.enu.kz/kz/page/departments/physical-and-economical-geography/faculty-members"
        target="_blank" style="color:{flk};text-decoration:none;">{l["f_univ"]}</a><br>
  📄 {l["f_license"]}:&nbsp;
  <a href="https://creativecommons.org/licenses/by/4.0/" target="_blank" style="color:{flk};text-decoration:none;">
    CC BY 4.0 — Creative Commons Attribution 4.0 International</a>
</div>""",
    unsafe_allow_html=True,
)