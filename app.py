import streamlit as st
from docx import Document
import re
import pandas as pd
from io import BytesIO

# ─── PAGE CONFIG ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Чекер статьи / Мақаланы тексеру",
    page_icon="📋",
    layout="wide"
)

# ─── SESSION STATE ────────────────────────────────────────────────
if "lang" not in st.session_state:
    st.session_state.lang = "kz"
if "theme" not in st.session_state:
    st.session_state.theme = "light"

# ─── LOCALES ──────────────────────────────────────────────────────
locales = {
    "ru": {
        "title": "📋 Автоматическая проверка статьи",
        "subtitle": "Вестник ЕНУ им. Л.Н. Гумилева · Серия: Химия / География · Шаблон 2025",
        "btn_lang": "🇰🇿 ҚАЗ",
        "btn_theme_dark": "🌙 Тёмная тема",
        "btn_theme_light": "☀️ Светлая тема",
        "upload_title": "📂 Загрузите статью в формате .docx",
        "upload_help": "Шаблон журнала Вестник ЕНУ, серия Химия/География, 2025",
        "analyzing": "Анализируем статью...",
        "res_title": "📊 Результаты проверки",
        "total": "Всего",
        "passed": "✅ Выполнено",
        "warned": "⚠️ Внимание",
        "failed": "❌ Не выпол.",
        "score": "🏆 Соответствие",
        "det_report": "### 📋 Детальный отчёт",
        "btn_csv": "⬇️ Скачать CSV",
        "btn_xls": "⬇️ Скачать Excel",
        "btn_docx": "⬇️ Word (DOCX)",
        "req_fix": "### ⚠️ Требует исправления",
        "req": "требование",
        "no_file": "👆 Загрузите .docx файл, чтобы начать проверку",
        # Критерии
        "c_author": "Имя первого автора",
        "c_author_req": "Имя Фамилия из шапки статьи",
        "c_lang": "Язык статьи",
        "c_lang_req": "Определение основного языка",
        "c_vol": "Объём статьи",
        "c_vol_req": "≥3500 слов",
        "c_ann_ru": "Аннотация (рус)",
        "c_ann_req": "≤300 слов",
        "c_ann_kz": "Аннотация (каз)",
        "c_ann_en": "Abstract (англ)",
        "c_req_obl": "Обязательно",
        "c_kw": "Ключевые слова",
        "c_kw_req": "3–10, разделитель «;»",
        "c_mrnti": "Код МРНТИ / IRSTI",
        "c_orcid": "ORCID авторов",
        "c_orcid_req": "Для каждого автора",
        "c_intro": "§1. Введение / Introduction",
        "c_mm": "§2. Материалы и методы",
        "c_res": "§3. Результаты / Results",
        "c_disc": "§4. Обсуждение / Discussion",
        "c_concl": "§5. Заключение / Conclusion",
        "c_supp": "§6. Вспомогательный материал",
        "c_contrib": "§7. Вклад авторов",
        "c_authinfo": "§8. Информация об авторе",
        "c_fund": "§9. Финансирование",
        "c_ack": "§10. Благодарности",
        "c_conflict": "§11. Конфликт интересов",
        "c_refs": "§12. Список литературы / Кол-во",
        "c_refs_req": "≥25 источников",
        "c_doi": "DOI в ссылках",
        "c_apa": "Стиль цитирования",
        "c_apa_req": "APA 7 (Автор, год)",
        "c_paper": "Формат бумаги",
        "c_paper_req": "A4 (210x297 мм)",
        "c_margins": "Поля",
        "c_margins_req": "Все поля 20 мм",
        "c_font": "Шрифт и кегль",
        "c_font_req": "Times New Roman, 12 pt",
        "c_tables": "Таблицы",
        "c_tables_req": "Должны быть в тексте",
        "c_images": "Рисунки (DPI)",
        "c_images_req": "Мин. 300 DPI (вручную)",
        "c_multi_ann": "Многоязычные аннотации",
        "c_multi_ann_req": "Ещё 2 аннотации на других языках",
        "c_type": "Тип статьи и соответствие объёму",
        "c_type_req": "Обзор/мини-обзор/исследовательская",
        "c_translit": "Транслитерация источников",
        "c_translit_req": "Для рус/каз источников — латиница",
        "found": "Найдено",
        "not_found": "Отсутствует",
        "words": "слов",
    },
    "kz": {
        "title": "📋 Мақаланы автоматты түрде тексеру",
        "subtitle": "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы · Серия: Химия / География · 2025 үлгісі",
        "btn_lang": "🇷🇺 РУС",
        "btn_theme_dark": "🌙 Түнгі режим",
        "btn_theme_light": "☀️ Күндізгі режим",
        "upload_title": "📂 .docx форматындағы мақаланы жүктеңіз",
        "upload_help": "Л.Н. Гумилев атындағы ЕҰУ Хабаршысы, Химия/География сериясы, 2025 үлгісі",
        "analyzing": "Мақала талдануда...",
        "res_title": "📊 Тексеру нәтижелері",
        "total": "Барлығы",
        "passed": "✅ Орындалды",
        "warned": "⚠️ Назар аударыңыз",
        "failed": "❌ Орындалмады",
        "score": "🏆 Сәйкестік",
        "det_report": "### 📋 Толық есеп",
        "btn_csv": "⬇️ CSV жүктеу",
        "btn_xls": "⬇️ Excel жүктеу",
        "btn_docx": "⬇️ Word (DOCX)",
        "req_fix": "### ⚠️ Түзетуді қажет етеді",
        "req": "талап",
        "no_file": "👆 Тексеруді бастау үшін .docx файлын жүктеңіз",
        "c_author": "Бірінші автордың аты-жөні",
        "c_author_req": "Мақаланың тақырыбынан кейінгі аты-жөн",
        "c_lang": "Мақала тілі",
        "c_lang_req": "Негізгі тілді анықтау",
        "c_vol": "Мақала көлемі",
        "c_vol_req": "≥3500 сөз",
        "c_ann_ru": "Аңдатпа (орыс)",
        "c_ann_req": "≤300 сөз",
        "c_ann_kz": "Аңдатпа (қаз)",
        "c_ann_en": "Abstract (ағылш)",
        "c_req_obl": "Міндетті түрде",
        "c_kw": "Түйінді сөздер",
        "c_kw_req": "3–10, бөлгіш «;»",
        "c_mrnti": "МРНТИ / IRSTI коды",
        "c_orcid": "Авторлардың ORCID",
        "c_orcid_req": "Әр автор үшін",
        "c_intro": "§1. Кіріспе / Introduction",
        "c_mm": "§2. Материалдар мен әдістер",
        "c_res": "§3. Нәтижелер / Results",
        "c_disc": "§4. Талқылау / Discussion",
        "c_concl": "§5. Қорытынды / Conclusion",
        "c_supp": "§6. Қосымша материалдар",
        "c_contrib": "§7. Авторлардың үлесі",
        "c_authinfo": "§8. Автор туралы ақпарат",
        "c_fund": "§9. Қаржыландыру",
        "c_ack": "§10. Алғыстар",
        "c_conflict": "§11. Мүдделер қақтығысы",
        "c_refs": "§12. Әдебиет тізімі / Саны",
        "c_refs_req": "≥25 дереккөз",
        "c_doi": "Сілтемелердегі DOI",
        "c_apa": "Дәйексөз стилі",
        "c_apa_req": "APA 7 (Автор, жыл)",
        "c_paper": "Қағаз форматы",
        "c_paper_req": "A4 (210x297 мм)",
        "c_margins": "Жақтаулар",
        "c_margins_req": "Барлық жақтаулар 20 мм",
        "c_font": "Шрифт және кегль",
        "c_font_req": "Times New Roman, 12 pt",
        "c_tables": "Кестелер",
        "c_tables_req": "Мәтінде болуы керек",
        "c_images": "Суреттер (DPI)",
        "c_images_req": "Мин. 300 DPI (қолмен тексеру)",
        "c_multi_ann": "Көптілді аңдатпалар",
        "c_multi_ann_req": "Басқа 2 тілде аңдатпа болуы керек",
        "c_type": "Мақала түрі және сәйкестігі",
        "c_type_req": "Шолу/шағын шолу/зерттеу мақаласы",
        "c_translit": "Транслитерация (дереккөздер)",
        "c_translit_req": "Орыс/қаз дереккөздері үшін — латын",
        "found": "Табылды",
        "not_found": "Жоқ",
        "words": "сөз",
    },
}

l = locales[st.session_state.lang]

# ─── THEME CSS ─────────────────────────────────────────────────────
dark_css = """
<style>
.stApp{background-color:#121212;color:#E0E0E0;}
.stMetric{background:#1E1E1E;border:1px solid #333;color:#FFF;padding:12px;border-radius:10px;}
h1,h2,h3,h4,h5,h6{color:#E0E0E0!important;}
</style>
"""
light_css = """
<style>
.stMetric{background:#ffffff;padding:12px;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,0.08);}
h1,h2,h3{color:#1a3a5c;}
</style>
"""
st.markdown(dark_css if st.session_state.theme == "dark" else light_css, unsafe_allow_html=True)

# ─── TOP BUTTONS ───────────────────────────────────────────────────
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

# ─── HELPERS ───────────────────────────────────────────────────────
# Шаблонды/техникалық жолдарды елемейтін тізім
_SKIP_PATTERNS = re.compile(
    r"приложение|қосымша|appendix|мрнти|irsti|orcid|e-mail|аффилиация|affiliation"
    r"|affiliation|секция|section|тип статьи|мақала түрі|type of the paper"
    r"|корреспонд|correspondence|copyright|citation|цитирование|дәйексөз"
    r"|received|поступила|accepted|published|academic editor|vest_chem"
    r"|beisemb|гумилева|gumilyov|doi\.org|http|https",
    re.IGNORECASE,
)

def extract_author_and_lang(doc: Document):
    """
    Автордың аты-жөнін мақала тақырыбынан кейінгі алғашқы сәйкес жолдан алады.
    Тілді бүкіл мәтін бойынша кілт сөздермен анықтайды.
    Қайтарады: (author_str, main_lang)
    """
    # --- тіл анықтау ---
    full = "\n".join(p.text for p in doc.paragraphs).lower()
    if any(k in full for k in ["кіріспе", "қорытынды", "мақала", "нәтижелер", "аңдатпа"]):
        main_lang = "kz"
    elif any(k in full for k in ["introduction", "conclusion", "results", "abstract", "discussion"]):
        main_lang = "en"
    else:
        main_lang = "ru"

    # --- автор аты ---
    # Тек мақала шапкасынан (алғашқы 25 абзац) іздейміз,
    # Список литературы/References/Әдебиет тізімі табылса тоқтаймыз.
    author_str = ""
    header_paragraphs = []
    for p in doc.paragraphs[:25]:
        txt = p.text.strip()
        if not txt:
            continue
        if re.search(r"список литературы|references|әдебиет тізімі", txt, re.IGNORECASE):
            break
        header_paragraphs.append(txt)

    for line in header_paragraphs:
        # шаблондық / техникалық жолдарды өткізіп жіберу
        if _SKIP_PATTERNS.search(line):
            continue
        # сандармен немесе тек латын/кирилл емес символдармен басталатын жолдарды өткізу
        if re.match(r"^[\d\s\*©\^]", line):
            continue
        # мақала тақырыбы - ол ұзын болады (>8 сөз), авторлар жолы қысқа болады
        words = line.split()
        if len(words) > 12:
            continue
        # Авторлар жолы: «Аты Тегі^1^, Аты Тегі^2^» форматы
        # Кирилл немесе латын бас әрптен басталатын 2+ сөз
        cleaned = re.sub(r"\d+[\*,]?", "", line).strip()
        cleaned = re.sub(r"\^[^|]*\^", "", cleaned).strip()
        # «және», «и», «and» арқылы бөлінген авторлар
        cleaned = re.sub(r"\s+(және|и|and)\s+", ", ", cleaned, flags=re.IGNORECASE)
        parts = [p.strip() for p in cleaned.split(",") if p.strip()]
        if not parts:
            continue
        first = parts[0]
        name_words = first.split()
        # Кем дегенде 2 сөз, бас әрптен басталу керек
        if len(name_words) >= 2 and re.match(r"[А-ЯЁA-ZҒҚҢӨҰҮІӘ]", name_words[0]):
            author_str = first
            break

    if not author_str:
        author_str = "Анықталмады / Не найдено"

    return author_str, main_lang


def author_for_filename(author_str: str) -> str:
    """Файл атауы үшін: бос орындар мен арнайы таңбаларды алып тастайды."""
    if "Анықталмады" in author_str or "Не найдено" in author_str:
        return "report"
    return re.sub(r"[^А-Яа-яA-Za-zҒғҚқҢңӨөҰұҮүІіӘә-]", "", author_str.replace(" ", "_"))


# ─── MAIN CHECK FUNCTION ───────────────────────────────────────────
def check_article(doc: Document, l: dict):
    full_text = "\n".join(p.text for p in doc.paragraphs)
    word_count = len(full_text.split())
    text_low = full_text.lower()
    results = []

    author_str, main_lang = extract_author_and_lang(doc)

    def add(num, criterion, requirement, found_val, status):
        results.append({
            "№": num,
            "Критерий": criterion,
            "Требование": requirement,
            "Найдено": found_val,
            "Статус": status,
        })

    # 0. Автор аты-жөні
    is_author_found = "Анықталмады" not in author_str and "Не найдено" not in author_str
    add(0, l["c_author"], l["c_author_req"],
        author_str, "✅" if is_author_found else "⚠️")

    # 0b. Мақала тілі
    lang_map = {"ru": "Русский", "kz": "Қазақша", "en": "English"}
    add(1, l["c_lang"], l["c_lang_req"],
        lang_map.get(main_lang, main_lang), "✅")

    # 1. Объём
    add(2, l["c_vol"], l["c_vol_req"],
        f"{word_count} {l['words']}",
        "✅" if word_count >= 3500 else "⚠️")

    # 2–4. Аннотации
    abstract_ru = re.search(
        r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін|abstract)",
        full_text, re.IGNORECASE | re.DOTALL,
    )
    if abstract_ru:
        aw = len(abstract_ru.group(1).split())
        add(3, l["c_ann_ru"], l["c_ann_req"],
            f"{aw} {l['words']}", "✅" if aw <= 300 else "❌")
        has_ru_ann = True
    else:
        add(3, l["c_ann_ru"], l["c_ann_req"], l["not_found"], "⚠️")
        has_ru_ann = False

    has_kaz_ann = "аңдатпа" in text_low or "аннотация (қаз" in text_low
    add(4, l["c_ann_kz"], l["c_req_obl"],
        l["found"] if has_kaz_ann else l["not_found"],
        "✅" if has_kaz_ann else "❌")

    has_eng_ann = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(5, l["c_ann_en"], l["c_req_obl"],
        l["found"] if has_eng_ann else l["not_found"],
        "✅" if has_eng_ann else "❌")

    # 5. Ключевые слова
    kw_match = re.search(
        r"(ключевые слова|keywords|түйінді сөздер|түйін сөздер)[:\s]+(.+?)(\n|$)",
        full_text, re.IGNORECASE,
    )
    if kw_match:
        kw_list = [k.strip() for k in kw_match.group(2).split(";") if k.strip()]
        add(6, l["c_kw"], l["c_kw_req"],
            f"{len(kw_list)}",
            "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(6, l["c_kw"], l["c_kw_req"], l["not_found"], "⚠️")

    # 6-7. МРНТИ / ORCID
    mrnt = bool(re.search(r"МРНТИ|IRSTI|\d{2}\.\d{2}\.\d{2}", full_text))
    add(7, l["c_mrnti"], l["c_req_obl"],
        l["found"] if mrnt else l["not_found"],
        "✅" if mrnt else "⚠️")

    orcid_count = len(re.findall(
        r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}", full_text, re.IGNORECASE))
    add(8, l["c_orcid"], l["c_orcid_req"],
        f"{orcid_count} ORCID",
        "✅" if orcid_count >= 1 else "⚠️")

    # §1-5 основные разделы
    sections = [
        (9,  l["c_intro"], ["введение", "кіріспе", "introduction"]),
        (10, l["c_mm"],    ["материал", "әдістер", "materials and methods"]),
        (11, l["c_res"],   ["результат", "нәтижелер", "results"]),
        (12, l["c_disc"],  ["обсужден", "талдау", "талқылау", "discussion"]),
        (13, l["c_concl"], ["заключени", "қорытынды", "conclusion"]),
    ]
    for num, name, keys in sections:
        found = any(k in text_low for k in keys)
        add(num, name, l["c_req_obl"],
            l["found"] if found else l["not_found"],
            "✅" if found else "❌")

    # §6-11 доп. разделы
    has_supp = any(k in text_low for k in ["вспомогательн", "қосымша материал", "supplementary materials"])
    add(14, l["c_supp"], l["c_req_obl"],
        l["found"] if has_supp else l["not_found"],
        "✅" if has_supp else "⚠️")

    contrib = any(k in text_low for k in [
        "вклад авторов", "вклады авторов", "author contributions",
        "авторлық үлестер", "авторлардың үлесі"])
    add(15, l["c_contrib"], "CRediT",
        l["found"] if contrib else l["not_found"],
        "✅" if contrib else "❌")

    authinfo = any(k in text_low for k in [
        "информация об авторе", "автор туралы ақпарат", "author information"])
    add(16, l["c_authinfo"], l["c_req_obl"],
        l["found"] if authinfo else l["not_found"],
        "✅" if authinfo else "⚠️")

    fund = any(k in text_low for k in ["финансирован", "funding", "қаржыландыру"])
    add(17, l["c_fund"], l["c_req_obl"],
        l["found"] if fund else l["not_found"],
        "✅" if fund else "❌")

    ack = any(k in text_low for k in [
        "благодарност", "алғыстар", "acknowledgements", "acknowledgments"])
    add(18, l["c_ack"], l["c_req_obl"],
        l["found"] if ack else l["not_found"],
        "✅" if ack else "⚠️")

    # §11 Конфликт интересов — ищем по номеру секции или ключевым словам
    conflict = any(k in text_low for k in [
        "конфликт интересов", "конфликты интересов",
        "conflict of interest", "мүдделер қақтығысы"])
    add(19, l["c_conflict"], l["c_req_obl"],
        l["found"] if conflict else l["not_found"],
        "✅" if conflict else "❌")

    # §12 Список литературы / Әдебиет тізімі / References
    refs_block = ""
    m_refs = re.search(
        r"(список литературы|references|әдебиет тізімі)(.*)$",
        full_text, re.IGNORECASE | re.DOTALL,
    )
    if m_refs:
        refs_block = m_refs.group(2)

    total_refs = len(re.findall(r"(?m)^\s*\d+\.\s+\S", refs_block))
    add(20, l["c_refs"], l["c_refs_req"],
        f"{total_refs}",
        "✅" if total_refs >= 25 else "❌")

    doi_count = len(re.findall(r"https?://doi\.org/", refs_block))
    add(21, l["c_doi"], l["c_req_obl"],
        f"{doi_count} DOI",
        "✅" if doi_count >= 5 else "⚠️")

    apa_style = bool(re.search(r"\([A-ZА-ЯҒҚ][a-zA-Zа-яА-ЯҒқ]+.{0,30}?\d{4}\)", refs_block))
    add(22, l["c_apa"], l["c_apa_req"],
        l["found"] if apa_style else l["not_found"],
        "✅" if apa_style else "⚠️")

    # Техническое: бумага, поля
    try:
        sec = doc.sections[0]
        w_mm = round(sec.page_width.mm)
        h_mm = round(sec.page_height.mm)
        is_a4 = (209 <= w_mm <= 211) and (296 <= h_mm <= 298)
        add(23, l["c_paper"], l["c_paper_req"],
            f"{w_mm}x{h_mm} мм", "✅" if is_a4 else "❌")
        t = round(sec.top_margin.mm)
        b = round(sec.bottom_margin.mm)
        lf = round(sec.left_margin.mm)
        rg = round(sec.right_margin.mm)
        margins_ok = (t == 20 and b == 20 and lf == 20 and rg == 20)
        add(24, l["c_margins"], l["c_margins_req"],
            f"Л:{lf} П:{rg} В:{t} Н:{b} мм",
            "✅" if margins_ok else "❌")
    except Exception:
        add(23, l["c_paper"], l["c_paper_req"], "Қате/Ошибка", "⚠️")
        add(24, l["c_margins"], l["c_margins_req"], "Қате/Ошибка", "⚠️")

    # Шрифт
    try:
        fn = doc.styles["Normal"].font.name or "?"
        fs = (doc.styles["Normal"].font.size.pt
              if doc.styles["Normal"].font.size else "?")
        ok_font = "Times New Roman" in str(fn) and fs in [11.0, 12.0]
        add(25, l["c_font"], l["c_font_req"],
            f"{fn}, {fs} pt", "✅" if ok_font else "⚠️")
    except Exception:
        add(25, l["c_font"], l["c_font_req"], "Қате/Ошибка", "⚠️")

    # Таблицы и рисунки
    tbl_count = len(doc.tables)
    add(26, l["c_tables"], l["c_tables_req"],
        f"{tbl_count} шт.", "✅" if tbl_count > 0 else "⚠️")

    img_count = len(doc.inline_shapes)
    add(27, l["c_images"], l["c_images_req"],
        f"{img_count} шт.", "⚠️" if img_count > 0 else "✅")

    # Многоязычные аннотации
    if main_lang == "ru":
        ok_multi = has_kaz_ann and has_eng_ann
    elif main_lang == "kz":
        ok_multi = has_ru_ann and has_eng_ann
    else:
        ok_multi = has_ru_ann and has_kaz_ann
    add(28, l["c_multi_ann"], l["c_multi_ann_req"],
        l["found"] if ok_multi else l["not_found"],
        "✅" if ok_multi else "❌")

    # Тип статьи: сначала ищем ключевые слова типа в шапке (первые 15 абзацев)
    header_low = "\n".join(p.text for p in doc.paragraphs[:15]).lower()
    # Потом во всём тексте
    article_type_label = ""
    ok_type = False

    is_review = bool(re.search(r"\bобзор\b|\bшолу\b|\breview\b", header_low + text_low))
    is_mini = bool(re.search(r"мини.?обзор|шағын.?шолу|mini.?review", header_low + text_low))
    is_research = bool(re.search(
        r"зерттеу мақаласы|исследовательская статья|research article|research paper"
        r"|ғылыми мақала|article\b", header_low + text_low))

    if is_mini and 6000 <= word_count <= 10000 and total_refs >= 50:
        article_type_label = "Мини-обзор / шағын шолу (6000–10000 сөз, ≥50)"
        ok_type = True
    elif is_review and not is_mini and word_count >= 10000 and total_refs >= 100:
        article_type_label = "Обзор / шолу (≥10000 сөз, ≥100)"
        ok_type = True
    elif is_research and word_count >= 3500 and total_refs >= 25:
        article_type_label = "Зерттеу мақаласы / research article (≥3500 сөз, ≥25)"
        ok_type = True
    else:
        # Автоматты анықтау - тек көлем мен дереккөз санына сүйену
        if word_count >= 10000 and total_refs >= 100:
            article_type_label = f"Обзор (предположительно): {word_count} сөз, {total_refs} дереккөз"
            ok_type = True
        elif 6000 <= word_count <= 10000 and total_refs >= 50:
            article_type_label = f"Мини-обзор (предположительно): {word_count} сөз, {total_refs} дереккөз"
            ok_type = True
        elif word_count >= 3500 and total_refs >= 25:
            article_type_label = f"Research article (предположительно): {word_count} сөз, {total_refs} дереккөз"
            ok_type = True
        else:
            article_type_label = f"Анықталмады: {word_count} сөз, {total_refs} дереккөз"
            ok_type = False

    add(29, l["c_type"], l["c_type_req"],
        article_type_label, "✅" if ok_type else "⚠️")

    # Транслитерация (только для EN)
    if main_lang == "en":
        has_cyrillic = bool(re.search(r"[А-Яа-яЁёҒғҚқҢңӨөҰұҮүІіӘә]", refs_block))
        has_translit = bool(re.search(r"in Russian|in Kazakh|Teorija|Almaty|Astana|translit", refs_block, re.IGNORECASE))
        ok_tr = (not has_cyrillic) or has_translit
        add(30, l["c_translit"], l["c_translit_req"],
            l["found"] if ok_tr else "Кириллица без транслитерации",
            "✅" if ok_tr else "⚠️")
    else:
        add(30, l["c_translit"], l["c_translit_req"],
            "Не требуется", "✅")

    return results, full_text, author_str, main_lang


# ─── DOCX REPORT ──────────────────────────────────────────────────
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
        row[0].text = str(r["№"])
        row[1].text = str(r["Критерий"])
        row[2].text = str(r["Требование"])
        row[3].text = str(r["Найдено"])
        row[4].text = str(r["Статус"])
    d.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(l["upload_title"], type=["docx"], help=l["upload_help"])

if uploaded_file:
    with st.spinner(l["analyzing"]):
        doc = Document(uploaded_file)
        results, full_text, author_str, main_lang = check_article(doc, l)
        df = pd.DataFrame(results)

    passed = sum(1 for r in results if r["Статус"] == "✅")
    warned = sum(1 for r in results if r["Статус"] == "⚠️")
    failed = sum(1 for r in results if r["Статус"] == "❌")
    total = len(results)
    score = int(passed / total * 100) if total > 0 else 0

    st.markdown(f"## {l['res_title']}")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(l["total"], total)
    c2.metric(l["passed"], passed)
    c3.metric(l["warned"], warned)
    c4.metric(l["failed"], failed)
    c5.metric(l["score"], f"{score}%")

    bar_color = "#4caf50" if score >= 80 else "#ffc107" if score >= 60 else "#f44336"
    bg_bar = "#2b2b2b" if st.session_state.theme == "dark" else "#e9ecef"
    st.markdown(
        f"""<div style="background:{bg_bar};border-radius:10px;height:28px;margin:8px 0 20px 0;">
          <div style="background:{bar_color};width:{score}%;height:28px;border-radius:10px;
                      display:flex;align-items:center;justify-content:center;
                      color:white;font-weight:bold;">{score}%</div></div>""",
        unsafe_allow_html=True,
    )

    def highlight(row):
        c = ({"✅": "background-color:#1b5e20", "⚠️": "background-color:#795548", "❌": "background-color:#b71c1c"}
             if st.session_state.theme == "dark"
             else {"✅": "background-color:#d4edda", "⚠️": "background-color:#fff3cd", "❌": "background-color:#f8d7da"})
        return [c.get(row["Статус"], "")] * len(row)

    st.markdown(l["det_report"])
    st.dataframe(df.style.apply(highlight, axis=1), use_container_width=True, height=900)

    st.markdown("---")
    col_a, col_b, col_c = st.columns(3)
    base_name = f"compliance_{author_for_filename(author_str)}"

    col_a.download_button(l["btn_csv"],
        df.to_csv(index=False).encode("utf-8-sig"),
        f"{base_name}.csv", "text/csv")

    xbuf = BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Report")
    col_b.download_button(l["btn_xls"], xbuf.getvalue(),
        f"{base_name}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    col_c.download_button(l["btn_docx"],
        build_docx_report(results, l, total, passed, warned, failed, score),
        f"{base_name}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    problems = [r for r in results if r["Статус"] in ("❌", "⚠️")]
    if problems:
        st.markdown(l["req_fix"])
        for p in problems:
            icon = "🔴" if p["Статус"] == "❌" else "🟡"
            st.markdown(f"{icon} **{p['Критерий']}** — {p['Найдено']} *({l['req']}: {p['Требование']})*")
else:
    st.info(l["no_file"])
