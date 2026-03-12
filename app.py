import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Чекер статьи / Мақаланы тексеру", page_icon="📋", layout="wide")

if "lang" not in st.session_state: st.session_state.lang = "kz"
if "theme" not in st.session_state: st.session_state.theme = "light"

locales = {
    "ru": {
        "title": "📋 Автоматическая проверка статьи",
        "subtitle": "Вестник ЕНУ им. Л.Н. Гумилева · Серия: Химия / География · Шаблон 2025",
        "btn_lang": "🇰🇿 ҚАЗ", "btn_theme_dark": "🌙 Тёмная тема", "btn_theme_light": "☀️ Светлая тема",
        "upload_title": "📂 Загрузите статью в формате .docx",
        "upload_help": "Шаблон журнала Вестник ЕНУ, серия Химия/География, 2025",
        "analyzing": "Анализируем статью...", "res_title": "📊 Результаты проверки",
        "total": "Всего", "passed": "✅ Выполнено", "warned": "⚠️ Внимание",
        "failed": "❌ Не выпол.", "score": "🏆 Соответствие",
        "det_report": "### 📋 Детальный отчёт", "img_report": "### 🖼️ Анализ рисунков",
        "img_num": "№", "img_pixels": "Пиксели", "img_size_mm": "Размер в doc",
        "img_dpi_calc": "DPI (расч.)", "img_dpi_emb": "DPI (встр.)",
        "img_format": "Формат", "img_status": "Статус",
        "btn_csv": "⬇️ Скачать CSV", "btn_xls": "⬇️ Скачать Excel", "btn_docx": "⬇️ Word (DOCX)",
        "req_fix": "### ⚠️ Требует исправления", "req": "требование",
        "no_file": "👆 Загрузите .docx файл, чтобы начать проверку",
        "c_title": "Наименование статьи", "c_title_req": "Строки 3–4 документа",
        "c_lang": "Язык статьи", "c_lang_req": "По преобладанию символов",
        "c_vol": "Объём статьи", "c_vol_req": "≥3500 слов",
        "c_ann_ru": "Аннотация (рус)", "c_ann_req": "≤300 слов",
        "c_ann_kz": "Аннотация (каз)", "c_ann_en": "Abstract (англ)", "c_req_obl": "Обязательно",
        "c_kw": "Ключевые слова", "c_kw_req": "3–10, разделитель «;»",
        "c_mrnti": "Код МРНТИ / IRSTI", "c_orcid": "ORCID авторов", "c_orcid_req": "Для каждого автора",
        "c_intro": "§1. Введение / Introduction", "c_mm": "§2. Материалы и методы",
        "c_res": "§3. Результаты / Results", "c_disc": "§4. Обсуждение / Discussion",
        "c_concl": "§5. Заключение / Conclusion", "c_supp": "§6. Вспомогательный материал",
        "c_contrib": "§7. Вклад авторов", "c_authinfo": "§8. Информация об авторе",
        "c_fund": "§9. Финансирование", "c_ack": "§10. Благодарности",
        "c_conflict": "§11. Конфликты интересов",
        "c_paper": "Формат бумаги", "c_paper_req": "A4 (210x297 мм)",
        "c_margins": "Поля", "c_margins_req": "Все поля 20 мм",
        "c_font": "Шрифт и кегль", "c_font_req": "Times New Roman, 12 pt",
        "c_tables": "Таблицы", "c_tables_req": "Должны быть в тексте",
        "c_images": "Рисунки", "c_images_req": "Мин. 300 DPI",
        "c_multi_ann": "Многоязычные аннотации", "c_multi_ann_req": "Ещё 2 аннотации на других языках",
        "found": "Найдено", "not_found": "Отсутствует", "words": "слов",
        "f_author": "Канат Самарханов / Kanat Samarkhanov", "f_license": "Лицензия",
        "f_univ": "ЕНУ им. Л.Н. Гумилева — Кафедра физической и экономической географии",
    },
    "kz": {
        "title": "📋 Мақаланы автоматты түрде тексеру",
        "subtitle": "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы · Серия: Химия / География · 2025 үлгісі",
        "btn_lang": "🇷🇺 РУС", "btn_theme_dark": "🌙 Түнгі режим", "btn_theme_light": "☀️ Күндізгі режим",
        "upload_title": "📂 .docx форматындағы мақаланы жүктеңіз",
        "upload_help": "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы, Химия/География сериясы, 2025 үлгісі",
        "analyzing": "Мақала талдануда...", "res_title": "📊 Тексеру нәтижелері",
        "total": "Барлығы", "passed": "✅ Орындалды", "warned": "⚠️ Назар аударыңыз",
        "failed": "❌ Орындалмады", "score": "🏆 Сәйкестік",
        "det_report": "### 📋 Толық есеп", "img_report": "### 🖼️ Суреттерді талдау",
        "img_num": "№", "img_pixels": "Пикселдер", "img_size_mm": "Doc өлшемі",
        "img_dpi_calc": "DPI (есепт.)", "img_dpi_emb": "DPI (енгіз.)",
        "img_format": "Формат", "img_status": "Статус",
        "btn_csv": "⬇️ CSV жүктеу", "btn_xls": "⬇️ Excel жүктеу", "btn_docx": "⬇️ Word (DOCX)",
        "req_fix": "### ⚠️ Түзетуді қажет етеді", "req": "талап",
        "no_file": "👆 Тексеруді бастау үшін .docx файлын жүктеңіз",
        "c_title": "Мақаланың атауы", "c_title_req": "Құжаттың 3–4 жолдары",
        "c_lang": "Мақала тілі", "c_lang_req": "Таңбалардың басымдылығы бойынша",
        "c_vol": "Мақала көлемі", "c_vol_req": "≥3500 сөз",
        "c_ann_ru": "Аңдатпа (орыс)", "c_ann_req": "≤300 сөз",
        "c_ann_kz": "Аңдатпа (қаз)", "c_ann_en": "Abstract (ағылш)", "c_req_obl": "Міндетті түрде",
        "c_kw": "Түйінді сөздер", "c_kw_req": "3–10, бөлгіш «;»",
        "c_mrnti": "МРНТИ / IRSTI коды", "c_orcid": "Авторлардың ORCID", "c_orcid_req": "Әр автор үшін",
        "c_intro": "§1. Кіріспе / Introduction", "c_mm": "§2. Материалдар мен әдістер",
        "c_res": "§3. Нәтижелер / Results", "c_disc": "§4. Талқылау / Discussion",
        "c_concl": "§5. Қорытынды / Conclusion", "c_supp": "§6. Қосымша материалдар",
        "c_contrib": "§7. Авторлардың үлесі", "c_authinfo": "§8. Автор туралы ақпарат",
        "c_fund": "§9. Қаржыландыру", "c_ack": "§10. Алғыстар",
        "c_conflict": "§11. Мүдделер қақтығысы",
        "c_paper": "Қағаз форматы", "c_paper_req": "A4 (210x297 мм)",
        "c_margins": "Жақтаулар", "c_margins_req": "Барлық жақтаулар 20 мм",
        "c_font": "Шрифт және кегль", "c_font_req": "Times New Roman, 12 pt",
        "c_tables": "Кестелер", "c_tables_req": "Мәтінде болуы керек",
        "c_images": "Суреттер", "c_images_req": "Мин. 300 DPI",
        "c_multi_ann": "Көптілді аңдатпалар", "c_multi_ann_req": "Басқа 2 тілде аңдатпа болуы керек",
        "found": "Табылды", "not_found": "Жоқ", "words": "сөз",
        "f_author": "Канат Самарханов / Kanat Samarkhanov", "f_license": "Лицензия",
        "f_univ": "Л.Н. Гумилев атындағы ЕҰУ — Физикалық және экономикалық география кафедрасы",
    },
}
l = locales[st.session_state.lang]

# Dark blue (GitHub-inspired) colours
_DB_BG     = "#0d1b2e"
_DB_CARD   = "#0f2340"
_DB_BORDER = "#1e3a5f"
_DB_TEXT   = "#c9d8ee"
_DB_HEAD   = "#e2edf7"
_DB_MUTED  = "#7b96b8"

dark_css = """
<style>
html, body, [class*='css'], .stApp {
    background-color: #0d1b2e !important;
    color: #c9d8ee !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif !important;
}
h1,h2,h3,h4,h5,h6,
[data-testid='stMarkdownContainer'] h1,
[data-testid='stMarkdownContainer'] h2,
[data-testid='stMarkdownContainer'] h3 { color: #e2edf7 !important; font-weight:600 !important; }
p,span,label,div,li,
[data-testid='stMarkdownContainer'] p,
[data-testid='stCaptionContainer'],.stCaption { color: #c9d8ee !important; }
[data-testid='block-container'],
[data-testid='stVerticalBlock'],
section[data-testid='stSidebar'] { background-color: #0d1b2e !important; }
[data-testid='stMetric'] {
    background: #0f2340 !important; border: 1px solid #1e3a5f !important;
    border-radius: 6px !important; padding: 12px 16px !important;
}
[data-testid='stMetricValue'] { color: #e2edf7 !important; }
[data-testid='stMetricLabel'] { color: #7b96b8 !important; }
.stButton > button {
    background-color: #0f2340 !important; color: #c9d8ee !important;
    border: 1px solid #1e3a5f !important; border-radius: 6px !important;
}
.stButton > button:hover { background-color: #1e3a5f !important; color: #e2edf7 !important; }
[data-testid='stDownloadButton'] > button {
    background-color: #238636 !important; color: #fff !important;
    border: 1px solid #2ea043 !important; border-radius: 6px !important;
}
[data-testid='stDownloadButton'] > button:hover { background-color: #2ea043 !important; }
[data-testid='stFileUploader'] {
    background-color: #0f2340 !important; border: 1px dashed #1e3a5f !important; border-radius: 6px !important;
}
[data-testid='stFileUploader'] * { color: #c9d8ee !important; }
[data-testid='stDataFrame'],.stDataFrame iframe { border: 1px solid #1e3a5f !important; border-radius: 6px !important; }
[data-testid='stAlert'] {
    background-color: #0f2340 !important; border: 1px solid #1f6feb !important;
    color: #c9d8ee !important; border-radius: 6px !important;
}
.stSpinner > div { color: #c9d8ee !important; }
hr { border-color: #1e3a5f !important; }
input,textarea,select {
    background-color: #0f2340 !important; color: #c9d8ee !important;
    border: 1px solid #1e3a5f !important;
}
</style>
"""

light_css = """
<style>
[data-testid='stMetric'] { background:#fff;padding:12px;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08); }
h1,h2,h3 { color:#1a3a5c; }
[data-testid='stDownloadButton'] > button { background-color:#2ea043;color:#fff;border-radius:6px; }
</style>
"""

st.markdown(dark_css if st.session_state.theme == "dark" else light_css, unsafe_allow_html=True)

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


_KAZ_CHARS = set("қңөұүіәғҚҢӨҰҮІӘҒ")

def detect_lang_from_text(text):
    latin = sum(1 for c in text if c.isalpha() and c.isascii())
    cyr   = sum(1 for c in text if "\u0400" <= c <= "\u04FF")
    kaz   = sum(1 for c in text if c in _KAZ_CHARS)
    if latin > cyr: return "en"
    if kaz >= 1:    return "kz"
    return "ru"

def extract_title_and_lang(doc):
    # lines 3-4 → index 2,3 among non-empty paragraphs
    non_empty  = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    candidates = non_empty[2:4] if len(non_empty) >= 3 else non_empty
    title = max(candidates, key=len) if candidates else ""
    return title, detect_lang_from_text(title)

def title_for_filename(title):
    if not title: return "report"
    clean = re.sub(r"[^\w\s-]", "", title, flags=re.UNICODE)
    return clean.strip().replace(" ", "_")[:40] or "report"

# Regex covers: конфликт/а/ы + интересов/интересам/...
_CONFLICT_RE = re.compile(
    r"конфликт\w*\s+интерес|conflict\w*\s+of\s+interest|мүдделер\s+қақтығыс",
    re.IGNORECASE,
)

def has_conflict_section(doc, full_text):
    if _CONFLICT_RE.search(full_text): return True
    for par in doc.paragraphs:
        if _CONFLICT_RE.search(par.text): return True
    return False

def extract_image_info(doc, l):
    rows, EMU = [], 914400
    for i, shape in enumerate(doc.inline_shapes):
        w_in = (shape.width  or 0) / EMU
        h_in = (shape.height or 0) / EMU
        size_mm = f"{round(w_in*25.4)}x{round(h_in*25.4)} mm" if w_in > 0 else "-"
        try:
            pic  = shape._inline.graphic.graphicData.pic
            rId  = pic.blipFill.blip.get(qn("r:embed"))
            blob = doc.part.related_parts[rId].blob
            img  = Image.open(BytesIO(blob))
            px_w, px_h = img.size
            dw = round(px_w / w_in) if w_in > 0 else None
            dh = round(px_h / h_in) if h_in > 0 else None
            dpi_calc = f"{dw}x{dh}" if dw else "-"
            emb = img.info.get("dpi")
            dpi_emb = f"{round(emb[0])}x{round(emb[1])}" if emb else "-"
            ok = isinstance(dw, int) and dw >= 300
            rows.append({l["img_num"]: i+1, l["img_pixels"]: f"{px_w}x{px_h}",
                         l["img_size_mm"]: size_mm, l["img_dpi_calc"]: dpi_calc,
                         l["img_dpi_emb"]: dpi_emb, l["img_format"]: img.format or "?",
                         l["img_status"]: "✅" if ok else "❌"})
        except Exception:
            rows.append({l["img_num"]: i+1, l["img_pixels"]: "-", l["img_size_mm"]: size_mm,
                         l["img_dpi_calc"]: "-", l["img_dpi_emb"]: "-",
                         l["img_format"]: "-", l["img_status"]: "⚠️"})
    return rows


def check_article(doc, l):
    full_text  = "\n".join(p.text for p in doc.paragraphs)
    word_count = len(full_text.split())
    text_low   = full_text.lower()
    results    = []
    title, main_lang = extract_title_and_lang(doc)

    def add(num, crit, req, found, status):
        results.append({"№": num, "Критерий": crit, "Требование": req,
                        "Найдено": found, "Статус": status})

    add(0, l["c_title"], l["c_title_req"], title or l["not_found"], "✅" if title else "⚠️")
    lang_map = {"ru": "Русский", "kz": "Қазақша", "en": "English"}
    add(1, l["c_lang"], l["c_lang_req"], lang_map.get(main_lang, main_lang), "✅")
    add(2, l["c_vol"], l["c_vol_req"], f"{word_count} {l['words']}",
        "✅" if word_count >= 3500 else "⚠️")

    m = re.search(r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін|abstract)",
                  full_text, re.IGNORECASE | re.DOTALL)
    if m:
        aw = len(m.group(1).split())
        add(3, l["c_ann_ru"], l["c_ann_req"], f"{aw} {l['words']}", "✅" if aw <= 300 else "❌")
        has_ru_ann = True
    else:
        add(3, l["c_ann_ru"], l["c_ann_req"], l["not_found"], "⚠️"); has_ru_ann = False

    has_kaz_ann = "аңдатпа" in text_low or "аннотация (қаз" in text_low
    add(4, l["c_ann_kz"], l["c_req_obl"], l["found"] if has_kaz_ann else l["not_found"],
        "✅" if has_kaz_ann else "❌")

    has_eng_ann = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(5, l["c_ann_en"], l["c_req_obl"], l["found"] if has_eng_ann else l["not_found"],
        "✅" if has_eng_ann else "❌")

    kw = re.search(r"(ключевые слова|keywords|түйінді сөздер|түйін сөздер)[:\s]+(.+?)(\n|$)",
                   full_text, re.IGNORECASE)
    if kw:
        kw_list = [k.strip() for k in kw.group(2).split(";") if k.strip()]
        add(6, l["c_kw"], l["c_kw_req"], str(len(kw_list)), "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(6, l["c_kw"], l["c_kw_req"], l["not_found"], "⚠️")

    mrnt = bool(re.search(r"МРНТИ|IRSTI|\d{2}\.\d{2}\.\d{2}", full_text))
    add(7, l["c_mrnti"], l["c_req_obl"], l["found"] if mrnt else l["not_found"], "✅" if mrnt else "⚠️")

    orcid = len(re.findall(r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}", full_text, re.IGNORECASE))
    add(8, l["c_orcid"], l["c_orcid_req"], f"{orcid} ORCID", "✅" if orcid >= 1 else "⚠️")

    for num, name, keys in [
        (9,  l["c_intro"], ["введение", "кіріспе", "introduction"]),
        (10, l["c_mm"],    ["материалы и методы", "материалдар мен әдістер",
                            "materials and methods", "материал", "әдістер"]),
        (11, l["c_res"],   ["результаты", "нәтижелер", "results"]),
        (12, l["c_disc"],  ["обсуждение", "талқылау", "discussion"]),
        (13, l["c_concl"], ["заключение", "қорытынды", "conclusion"]),
    ]:
        f = any(k in text_low for k in keys)
        add(num, name, l["c_req_obl"], l["found"] if f else l["not_found"], "✅" if f else "❌")

    for num, key, kws, st_ok in [
        (14, l["c_supp"],    ["вспомогательный материал","қосымша материалдар","supplementary materials"], "⚠️"),
        (15, l["c_contrib"], ["вклад авторов","author contributions","авторлардың үлесі","авторлық үлестер"], "❌"),
        (16, l["c_authinfo"],["информация об авторе","author information","автор туралы ақпарат"], "⚠️"),
        (17, l["c_fund"],    ["финансирование","funding","қаржыландыру"], "❌"),
        (18, l["c_ack"],     ["благодарност","acknowledgements","acknowledgments","алғыстар"], "⚠️"),
    ]:
        f = any(k in text_low for k in kws)
        req = "CRediT" if num == 15 else l["c_req_obl"]
        add(num, key, req, l["found"] if f else l["not_found"], "✅" if f else st_ok)

    conflict = has_conflict_section(doc, full_text)
    add(19, l["c_conflict"], l["c_req_obl"], l["found"] if conflict else l["not_found"],
        "✅" if conflict else "❌")

    refs_match = None
    for pat in [r"список литературы", r"references", r"әдебиет(тер)? тізімі"]:
        m2 = re.search(pat, text_low)
        if m2: refs_match = m2; break
    if refs_match:
        rt = full_text[refs_match.end():]
        rl = re.findall(r"\n\s*(\d+[.)]|\[\d+\])\s", rt)
        if not rl: rl = [ln for ln in rt.split("\n") if len(ln.strip()) > 40]
        rc = len(rl)
        add(20, "Список литературы / References", "≥10 источников",
            str(rc), "✅" if rc >= 10 else "⚠️")
    else:
        add(20, "Список литературы / References", l["c_req_obl"], l["not_found"], "❌")

    try:
        sec  = doc.sections[0]
        w_mm = round(sec.page_width.mm); h_mm = round(sec.page_height.mm)
        is_a4 = (209 <= w_mm <= 211) and (296 <= h_mm <= 298)
        add(21, l["c_paper"], l["c_paper_req"], f"{w_mm}x{h_mm} mm", "✅" if is_a4 else "❌")
        t,b,lf,rg = round(sec.top_margin.mm),round(sec.bottom_margin.mm),\
                    round(sec.left_margin.mm),round(sec.right_margin.mm)
        add(22, l["c_margins"], l["c_margins_req"], f"Л:{lf} П:{rg} В:{t} Н:{b} мм",
            "✅" if all(x==20 for x in [t,b,lf,rg]) else "❌")
    except Exception:
        add(21, l["c_paper"], l["c_paper_req"], "Error", "⚠️")
        add(22, l["c_margins"], l["c_margins_req"], "Error", "⚠️")

    try:
        fn = doc.styles["Normal"].font.name or "?"
        fs = doc.styles["Normal"].font.size.pt if doc.styles["Normal"].font.size else "?"
        ok_f = "Times New Roman" in str(fn) and fs in [11.0, 12.0]
        add(23, l["c_font"], l["c_font_req"], f"{fn}, {fs} pt", "✅" if ok_f else "⚠️")
    except Exception:
        add(23, l["c_font"], l["c_font_req"], "Error", "⚠️")

    tc = len(doc.tables)
    add(24, l["c_tables"], l["c_tables_req"], f"{tc} шт.", "✅" if tc > 0 else "⚠️")

    ic = len(doc.inline_shapes)
    msg = f"{ic} шт. — см. таблицу ниже" if ic > 0 else l["not_found"]
    add(25, l["c_images"], l["c_images_req"], msg, "⚠️" if ic > 0 else "✅")

    ok_multi = (has_kaz_ann and has_eng_ann) if main_lang == "ru" else \
               (has_ru_ann  and has_eng_ann) if main_lang == "kz" else \
               (has_ru_ann  and has_kaz_ann)
    add(26, l["c_multi_ann"], l["c_multi_ann_req"],
        l["found"] if ok_multi else l["not_found"], "✅" if ok_multi else "❌")

    return results, full_text, title, main_lang

def build_docx_report(results, l, total, passed, warned, failed, score):
    buf = BytesIO(); d = Document()
    d.add_heading(l["res_title"], level=1)
    p = d.add_paragraph()
    p.add_run(f"{l['total']}: {total},  ").bold = True
    p.add_run(f"{l['passed']}: {passed},  {l['warned']}: {warned},  "
              f"{l['failed']}: {failed},  {l['score']}: {score}%")
    d.add_paragraph("")
    tbl = d.add_table(rows=1, cols=5)
    for i, h in enumerate(["\u2116", "Критерий", "Требование", "Найдено", "Статус"]):
        tbl.rows[0].cells[i].text = h
    for r in results:
        row = tbl.add_row().cells
        for i, k in enumerate(["\u2116", "Критерий", "Требование", "Найдено", "Статус"]):
            row[i].text = str(r.get(k, r.get("\u2116", "")))
    d.save(buf); buf.seek(0); return buf.getvalue()

uploaded_file = st.file_uploader(l["upload_title"], type=["docx"], help=l["upload_help"])

if uploaded_file:
    with st.spinner(l["analyzing"]):
        doc = Document(uploaded_file)
        results, full_text, title, main_lang = check_article(doc, l)
        df = pd.DataFrame(results)
        img_rows = extract_image_info(doc, l)

    passed = sum(1 for r in results if r["Статус"] == "✅")
    warned = sum(1 for r in results if r["Статус"] == "⚠️")
    failed = sum(1 for r in results if r["Статус"] == "❌")
    total  = len(results)
    score  = int(passed / total * 100) if total > 0 else 0

    st.markdown(f"## {l['res_title']}")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric(l["total"],total); c2.metric(l["passed"],passed)
    c3.metric(l["warned"],warned); c4.metric(l["failed"],failed)
    c5.metric(l["score"],f"{score}%")

    bar_clr = "#238636" if score>=80 else "#d29922" if score>=60 else "#da3633"
    bg_bar  = _DB_CARD  if st.session_state.theme=="dark" else "#e9ecef"
    txt_clr = _DB_HEAD  if st.session_state.theme=="dark" else "#ffffff"
    st.markdown(
        f'<div style="background:{bg_bar};border:1px solid #30363d;border-radius:6px;'
        f'height:28px;margin:8px 0 20px 0;">'
        f'<div style="background:{bar_clr};width:{score}%;height:28px;border-radius:6px;'
        f'display:flex;align-items:center;justify-content:center;'
        f'color:{txt_clr};font-weight:600;font-size:13px;">{score}%</div></div>',
        unsafe_allow_html=True)

    def highlight(row):
        dk = {"✅":"background-color:#1a4a1a;color:#3fb950",
              "⚠️":"background-color:#3d2e00;color:#d29922",
              "❌":"background-color:#3d0e0e;color:#f85149"}
        lt = {"✅":"background-color:#d4edda","⚠️":"background-color:#fff3cd","❌":"background-color:#f8d7da"}
        c = dk if st.session_state.theme=="dark" else lt
        return [c.get(row["Статус"],"")] * len(row)

    st.markdown(l["det_report"])
    st.dataframe(df.style.apply(highlight, axis=1), use_container_width=True, height=900)

    if img_rows:
        st.markdown(l["img_report"])
        df_img = pd.DataFrame(img_rows)
        scol = l["img_status"]
        def hl_img(row):
            dk = {"✅":"background-color:#1a4a1a;color:#3fb950",
                  "❌":"background-color:#3d0e0e;color:#f85149",
                  "⚠️":"background-color:#3d2e00;color:#d29922"}
            lt = {"✅":"background-color:#d4edda","❌":"background-color:#f8d7da","⚠️":"background-color:#fff3cd"}
            c = dk if st.session_state.theme=="dark" else lt
            return [c.get(row[scol],"")] * len(row)
        st.dataframe(df_img.style.apply(hl_img, axis=1), use_container_width=True)

    st.markdown("---")
    ca,cb,cc = st.columns(3)
    bn = f"compliance_{title_for_filename(title)}"
    ca.download_button(l["btn_csv"], df.to_csv(index=False).encode("utf-8-sig"),
                       f"{bn}.csv","text/csv")
    xb = BytesIO()
    with pd.ExcelWriter(xb, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Report")
        if img_rows: pd.DataFrame(img_rows).to_excel(w, index=False, sheet_name="Images")
    cb.download_button(l["btn_xls"],xb.getvalue(),f"{bn}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    cc.download_button(l["btn_docx"],
        build_docx_report(results,l,total,passed,warned,failed,score),
        f"{bn}.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    problems = [r for r in results if r["Статус"] in ("❌","⚠️")]
    if problems:
        st.markdown(l["req_fix"])
        for prob in problems:
            icon = "🔴" if prob["Статус"]=="❌" else "🟡"
            st.markdown(f"{icon} **{prob['Критерий']}** — {prob['Найдено']} *({l['req']}: {prob['Требование']})*")
else:
    st.info(l["no_file"])

fc  = _DB_MUTED if st.session_state.theme=="dark" else "#555"
flk = "#58a6ff"  if st.session_state.theme=="dark" else "#0969da"
st.markdown("---")
st.markdown(
    f'<div style="text-align:center;font-size:12px;color:{fc};padding:12px 0 20px 0;line-height:2.2;">'
    f'<b style="font-size:13px;">© 2025 {l["f_author"]}</b><br>'
    f'📧 <a href="mailto:samarkhanov_kb@enu.kz" style="color:{flk};text-decoration:none;">samarkhanov_kb@enu.kz</a>'
    f'&nbsp;·&nbsp;'
    f'<a href="mailto:kanat.baurzhanuly@gmail.com" style="color:{flk};text-decoration:none;">kanat.baurzhanuly@gmail.com</a><br>'
    f'🏛️ <a href="https://fns.enu.kz/kz/page/departments/physical-and-economical-geography/faculty-members"'
    f'     target="_blank" style="color:{flk};text-decoration:none;">{l["f_univ"]}</a><br>'
    f'📄 {l["f_license"]}:&nbsp;'
    f'<a href="https://creativecommons.org/licenses/by/4.0/" target="_blank" style="color:{flk};text-decoration:none;">'
    f'CC BY 4.0 — Creative Commons Attribution 4.0 International</a>'
    f'</div>',
    unsafe_allow_html=True)
