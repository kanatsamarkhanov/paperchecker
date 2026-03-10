import streamlit as st
from docx import Document
from docx.shared import Mm, Pt
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
if 'lang' not in st.session_state:
    st.session_state.lang = 'kz'
if 'theme' not in st.session_state:
    st.session_state.theme = 'light'

# ─── LOCALES ──────────────────────────────────────────────────────
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
        'c_vol': "Объём статьи", 'c_vol_req': "Общий объём (слов)",
        'c_ann_ru': "Аннотация (рус)", 'c_ann_req': "≤ 300 слов",
        'c_ann_kz': "Аннотация (каз)", 'c_ann_en': "Abstract (англ)", 'c_req_obl': "Обязательно",
        'c_kw': "Ключевые слова", 'c_kw_req': "3–10, разделитель «;»",
        'c_mrnti': "Код МРНТИ", 'c_orcid': "ORCID авторов", 'c_orcid_req': "Для каждого автора",
        'c_intro': "Раздел: Введение / Introduction",
        'c_mm': "Раздел: Материалы и методы / Materials and methods",
        'c_res': "Раздел: Результаты / Results",
        'c_disc': "Раздел: Обсуждение / Discussion",
        'c_concl': "Раздел: Заключение / Conclusion",
        'c_supp': "Раздел: Вспомогательный материал / Supplementary materials",
        'c_contrib': "Вклад авторов / Author contributions",
        'c_authinfo': "Информация об авторе / Author information",
        'c_fund': "Финансирование / Funding",
        'c_ack': "Благодарности / Acknowledgements",
        'c_conflict': "Конфликт интересов / Conflicts of interest",
        'c_refs': "Кол-во источников (Список литературы)",
        'c_refs_req': "≥ 25",
        'c_doi': "DOI в ссылках",
        'c_apa': "Стиль цитирования",
        'c_apa_req': "APA 7 (Автор, год)",

        # Технические критерии
        'c_paper': "Формат бумаги",
        'c_paper_req': "A4 (210x297 мм)",
        'c_margins': "Поля (Жақтаулар)",
        'c_margins_req': "Все поля по 2 см",
        'c_font': "Шрифт и кегль",
        'c_font_req': "Times New Roman, 11/12 pt",
        'c_tables': "Таблицы (Tables)",
        'c_tables_req': "Наличие в тексте и оформление как в шаблоне",
        'c_images': "Рисунки (Figures)",
        'c_images_req': "Мин. 300 DPI (проверить вручную)",

        'c_multi_ann': "Многоязычные аннотации",
        'c_multi_ann_req': "Для основного языка нужны ещё 2 аннотации",
        'c_type': "Тип статьи и объём",
        'c_type_req': "Обзор / мини-обзор / исследовательская статья",
        'c_translit': "Транслитерация рус/каз источников в англ. статье",
        'c_translit_req': "Наличие англ. описаний для рус/каз источников",

        'found': "Найдено",
        'not_found': "Отсутствует",
        'words': "слов",
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

        'c_vol': "Мақала көлемі", 'c_vol_req': "Жалпы көлем (сөз)",
        'c_ann_ru': "Аңдатпа (орыс)", 'c_ann_req': "≤ 300 сөз",
        'c_ann_kz': "Аңдатпа (қаз)", 'c_ann_en': "Abstract (ағылш)", 'c_req_obl': "Міндетті түрде",
        'c_kw': "Түйінді сөздер", 'c_kw_req': "3–10, бөлгіш «;»",
        'c_mrnti': "МРНТИ коды", 'c_orcid': "Авторлардың ORCID", 'c_orcid_req': "Әр автор үшін",
        'c_intro': "Бөлім: Кіріспе / Introduction",
        'c_mm': "Бөлім: Материалдар мен әдістер / Materials and methods",
        'c_res': "Бөлім: Нәтижелер / Results",
        'c_disc': "Бөлім: Талқылау / Discussion",
        'c_concl': "Бөлім: Қорытынды / Conclusion",
        'c_supp': "Қосымша материалдар / Supplementary materials",
        'c_contrib': "Авторлардың үлесі / Author contributions",
        'c_authinfo': "Автор туралы ақпарат / Author information",
        'c_fund': "Қаржыландыру / Funding",
        'c_ack': "Алғыстар / Acknowledgements",
        'c_conflict': "Мүдделер қақтығысы / Conflicts of interest",
        'c_refs': "Әдебиеттер саны",
        'c_refs_req': "≥ 25",
        'c_doi': "Сілтемелердегі DOI",
        'c_apa': "Дәйексөз стилі",
        'c_apa_req': "APA 7 (Автор, жыл)",

        'c_paper': "Қағаз форматы", 'c_paper_req': "A4 (210x297 мм)",
        'c_margins': "Жақтаулар (Поля)", 'c_margins_req': "Барлық жақтаулар 2 см",
        'c_font': "Шрифт және өлшемі", 'c_font_req': "Times New Roman, 11/12 pt",
        'c_tables': "Кестелер (Tables)", 'c_tables_req': "Үлгіге сәйкес рәсімделуі керек",
        'c_images': "Суреттер (Figures)", 'c_images_req': "Мин. 300 DPI (қолмен тексеру)",

        'c_multi_ann': "Көптілді аңдатпалар",
        'c_multi_ann_req': "Негізгі тілден бөлек тағы 2 аңдатпа болуы керек",
        'c_type': "Мақала түрі және көлемі",
        'c_type_req': "Шолу / шағын шолу / зерттеу мақаласы",
        'c_translit': "Ағылш. мақалада транслитерациясы бар ма",
        'c_translit_req': "Орыс/қазақ дереккөздері үшін англ. сипаттама",

        'found': "Табылды",
        'not_found': "Жоқ",
        'words': "сөз",
    }
}

l = locales[st.session_state.lang]

# ─── THEME CSS ─────────────────────────────────────────────────────
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

# ─── TOP BUTTONS ───────────────────────────────────────────────────
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

# ─── FIRST AUTHOR NAME ─────────────────────────────────────────────
def extract_first_author_name_surname(doc: Document) -> str:
    paragraphs = []
    for p in doc.paragraphs[:20]:
        text = p.text.strip()
        if not text:
            continue
        if re.search(r"список литературы|references|әдебиет тізімі", text, re.IGNORECASE):
            break
        paragraphs.append(text)

    header_text = "\n".join(paragraphs)
    header_text = re.sub(r"\^[0-9, ]+\^", " ", header_text)

    m = re.search(r"([А-ЯA-Z][а-яa-z]+)\s+([А-ЯA-Z][а-яa-z]+)\d*", header_text)
    if not m:
        return "report"

    first_name = re.sub(r"[^А-Яа-яA-Za-z-]", "", m.group(1))
    surname = re.sub(r"[^А-Яа-яA-Za-z-]", "", m.group(2))
    if not (first_name and surname):
        return "report"

    return f"{first_name}{surname}"

# ─── MAIN CHECK FUNCTION ───────────────────────────────────────────
def check_article(doc, l):
    full_text = "\n".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    text_low = full_text.lower()
    results = []

    # язык статьи
    if "кіріспе" in text_low or "қорытынды" in text_low or "мақала" in text_low:
        main_lang = "kz"
    elif "introduction" in text_low or "conclusion" in text_low:
        main_lang = "en"
    else:
        main_lang = "ru"

    def add(num, criterion, requirement, found, status):
        results.append({
            "№": num,
            "Критерий": criterion,
            "Требование": requirement,
            "Найдено": found,
            "Статус": status
        })

    # 1. общий объём
    add(1, l['c_vol'], l['c_vol_req'], f"{word_count} {l['words']}",
        "✅" if word_count >= 3500 else "⚠️")

    # 2–4. Аннотации
    abstract_ru = re.search(
        r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін|abstract)",
        full_text,
        re.IGNORECASE | re.DOTALL
    )
    if abstract_ru:
        aw = len(abstract_ru.group(1).split())
        add(2, l['c_ann_ru'], l['c_ann_req'],
            f"{aw} {l['words']}", "✅" if aw <= 300 else "❌")
        has_ru_ann = True
    else:
        add(2, l['c_ann_ru'], l['c_ann_req'], l['not_found'], "⚠️")
        has_ru_ann = False

    has_kaz_ann = (
        "аңдатпа" in text_low
        or "аннотация (қаз" in text_low
        or "abstract (kaz" in text_low
    )
    add(3, l['c_ann_kz'], l['c_req_obl'],
        l['found'] if has_kaz_ann else l['not_found'],
        "✅" if has_kaz_ann else "❌")

    has_eng_ann = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(4, l['c_ann_en'], l['c_req_obl'],
        l['found'] if has_eng_ann else l['not_found'],
        "✅" if has_eng_ann else "❌")

    # 5. ключевые слова
    kw_match = re.search(
        r"(ключевые слова|keywords|түйінді сөздер)[:\s]+(.+?)(\n|$)",
        full_text, re.IGNORECASE
    )
    if kw_match:
        kw_list = [k.strip() for k in kw_match.group(2).split(";") if k.strip()]
        add(5, l['c_kw'], l['c_kw_req'], f"{len(kw_list)}",
            "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(5, l['c_kw'], l['c_kw_req'], l['not_found'], "⚠️")

    # 6-7. МРНТИ и ORCID
    mrnt = bool(re.search(r"МРНТИ|IRSTI|\d{2}\.\d{2}\.\d{2}", full_text))
    add(6, l['c_mrnti'], l['c_req_obl'],
        l['found'] if mrnt else l['not_found'],
        "✅" if mrnt else "⚠️")

    orcid_count = len(re.findall(
        r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}", full_text, re.IGNORECASE
    ))
    add(7, l['c_orcid'], l['c_orcid_req'],
        f"{orcid_count} ORCID",
        "✅" if orcid_count >= 1 else "⚠️")

    # 8-12. основные разделы (по шаблонам трёх языков)
    sections = [
        (8,  l['c_intro'], ["введение", "кіріспе", "introduction"]),
        (9,  l['c_mm'],    ["материал", "әдістер", "materials and methods"]),
        (10, l['c_res'],   ["результат", "нәтижелер", "results"]),
        (11, l['c_disc'],  ["обсуждени", "талдау", "талқылау", "discussion"]),
        (12, l['c_concl'], ["заключени", "қорытынды", "conclusion"]),
    ]
    for num, name, keys in sections:
        found = any(k in text_low for k in keys)
        add(num, name, l['c_req_obl'],
            l['found'] if found else l['not_found'],
            "✅" if found else "❌")

    # 13-18. разделы «supplementary, contributions, author info, funding, ack, conflicts»
    has_supp = any(k in text_low for k in ["вспомогательн", "қосымша материал", "supplementary materials"])
    add(13, l['c_supp'], l['c_req_obl'],
        l['found'] if has_supp else l['not_found'],
        "✅" if has_supp else "⚠️")

    contrib = any(k in text_low
                  for k in ["вклад авторов", "author contributions", "авторлық үлестер", "авторлардың үлесі"])
    add(14, l['c_contrib'], "CRediT",
        l['found'] if contrib else l['not_found'],
        "✅" if contrib else "❌")

    authinfo = any(k in text_low
                   for k in ["информация об авторе", "автор туралы ақпарат", "author information"])
    add(15, l['c_authinfo'], l['c_req_obl'],
        l['found'] if authinfo else l['not_found'],
        "✅" if authinfo else "⚠️")

    fund = any(k in text_low
               for k in ["финансирован", "funding", "қаржыландыру"])
    add(16, l['c_fund'], l['c_req_obl'],
        l['found'] if fund else l['not_found'],
        "✅" if fund else "❌")

    ack = any(k in text_low
              for k in ["благодарност", "алғыстар", "acknowledgements", "acknowledgments"])
    add(17, l['c_ack'], l['c_req_obl'],
        l['found'] if ack else l['not_found'],
        "✅" if ack else "⚠️")

    conflict = any(k in text_low
                   for k in ["конфликт интересов",
                             "conflict of interest",
                             "мүдделер қақтығысы"])
    add(18, l['c_conflict'], l['c_req_obl'],
        l['found'] if conflict else l['not_found'],
        "✅" if conflict else "❌")

    # 19-21. Литература, DOI, APA
    refs_block = ""
    m_refs = re.search(
        r"(список литературы|references|әдебиет тізімі)(.*)$",
        full_text,
        re.IGNORECASE | re.DOTALL
    )
    if m_refs:
        refs_block = m_refs.group(2)
    else:
        refs_block = ""

    total_refs = len(re.findall(r"(?m)^\s*\d+\.\s+", refs_block))
    add(19, l['c_refs'], l['c_refs_req'],
        f"{total_refs}", "✅" if total_refs >= 25 else "❌")

    doi_count = len(re.findall(r"https?://doi\.org/", refs_block))
    add(20, l['c_doi'], l['c_req_obl'],
        f"{doi_count} DOI", "✅" if doi_count >= 5 else "⚠️")

    apa_style = bool(re.search(
        r"\([A-ZА-Я][a-zA-Zа-яА-Я]+.*?\d{4}\)", refs_block
    ))
    add(21, l['c_apa'], l['c_apa_req'],
        l['found'] if apa_style else l['not_found'],
        "✅" if apa_style else "⚠️")

    # 22-23. Формат и поля
    try:
        sec = doc.sections[0]
        w_mm, h_mm = round(sec.page_width.mm), round(sec.page_height.mm)
        is_a4 = (209 <= w_mm <= 211) and (296 <= h_mm <= 298)
        add(22, l['c_paper'], l['c_paper_req'],
            f"{w_mm}x{h_mm} мм", "✅" if is_a4 else "❌")

        t = round(sec.top_margin.mm)
        b = round(sec.bottom_margin.mm)
        lf = round(sec.left_margin.mm)
        rg = round(sec.right_margin.mm)
        margins_ok = (t == 20 and b == 20 and lf == 20 and rg == 20)
        add(23, l['c_margins'], l['c_margins_req'],
            f"Л:{lf} П:{rg} В:{t} Н:{b} мм",
            "✅" if margins_ok else "❌")
    except Exception:
        add(22, l['c_paper'], l['c_paper_req'], "Қате / Ошибка", "⚠️")
        add(23, l['c_margins'], l['c_margins_req'], "Қате / Ошибка", "⚠️")

    # 24. Шрифт
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
        add(24, l['c_font'], l['c_font_req'],
            f"{font_name}, {font_size_pt} pt",
            "✅" if is_correct_font else "⚠️")
    except Exception:
        add(24, l['c_font'], l['c_font_req'], "Қате / Ошибка", "⚠️")

    # 25-26. Таблицы и рисунки
    tbl_count = len(doc.tables)
    add(25, l['c_tables'], l['c_tables_req'],
        f"{tbl_count} шт.", "✅" if tbl_count > 0 else "⚠️")

    img_count = len(doc.inline_shapes)
    add(26, l['c_images'], l['c_images_req'],
        f"{img_count} шт.", "⚠️" if img_count > 0 else "✅")

    # 27. Многоязычные аннотации
    if main_lang == "ru":
        ok_multi = has_kaz_ann and has_eng_ann
    elif main_lang == "kz":
        ok_multi = has_ru_ann and has_eng_ann
    else:  # en
        ok_multi = has_ru_ann and has_kaz_ann

    add(27, l['c_multi_ann'], l['c_multi_ann_req'],
        l['found'] if ok_multi else l['not_found'],
        "✅" if ok_multi else "❌")

    # 28. Тип статьи (обзор / мини-шолу / research article) + объём/литература
    article_type = ""
    ok_type = False
    if re.search(r"обзор|шолу|review", text_low) and word_count >= 10000 and total_refs >= 100:
        article_type = "Обзор / шолу (≥10000 слов и ≥100 источников)"
        ok_type = True
    elif re.search(r"мини обзор|мини-шолу|mini-review|mini review", text_low) and 6000 <= word_count <= 10000 and total_refs >= 50:
        article_type = "Мини-обзор / шағын шолу (6000–10000 слов, ≥50 источников)"
        ok_type = True
    elif re.search(r"зерттеу мақаласы|исследовательская статья|research article|research paper", text_low) and word_count >= 3500 and total_refs >= 25:
        article_type = "Зерттеу мақаласы / research article (≥3500 слов, ≥25 источников)"
        ok_type = True
    else:
        article_type = "Тип не распознан или объём/литература не соответствуют"

    add(28, l['c_type'], l['c_type_req'],
        article_type, "✅" if ok_type else "⚠️")

    # 29. Транслитерация рус/каз источников для англ. статьи
    if main_lang == "en":
        # эвристика: наличие латинизированных русских названий в скобках + русских букв
        has_cyrillic_in_refs = bool(re.search(r"[А-Яа-яЁё]", refs_block))
        has_translit_pattern = bool(re.search(
            r"\([Tt]eorija|in Russian|in Kazakh|Almaty|Astana", refs_block))
        ok_translit = (not has_cyrillic_in_refs) or has_translit_pattern
        add(29, l['c_translit'], l['c_translit_req'],
            "есть пример translit" if ok_translit else "возможны кириллические ссылки без translit",
            "✅" if ok_translit else "⚠️")
    else:
        add(29, l['c_translit'], l['c_translit_req'],
            "Не требуется (статья не на англ.)", "✅")

    return results, full_text

# ─── DOCX REPORT ──────────────────────────────────────────────────
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

# ─── UI ───────────────────────────────────────────────────────────
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

    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    col_a.download_button(
        l['btn_csv'],
        csv_bytes,
        f"{base_name}.csv",
        "text/csv"
    )

    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    col_b.download_button(
        l['btn_xls'],
        excel_buf.getvalue(),
        f"{base_name}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
