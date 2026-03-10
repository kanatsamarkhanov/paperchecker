import streamlit as st
from docx import Document
import re
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Чекер статьи — Вестник ЕНУ",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
<style>
.stMetric {background:white;padding:12px;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,0.08);}
h1 {color:#1a3a5c;}
</style>
""", unsafe_allow_html=True)

st.title("📋 Автоматическая проверка статьи")
st.caption("Вестник ЕНУ им. Л.Н. Гумилева · Серия: Химия / География · Шаблон 2025")
st.markdown("---")


def check_article(doc):
    full_text = "\n".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    results = []

    def add(num, criterion, requirement, found, status):
        results.append({"№": num, "Критерий": criterion,
                         "Требование": requirement, "Найдено": found, "Статус": status})

    add(1, "Объём статьи", "≥ 3500 слов", f"{word_count} слов",
        "✅" if word_count >= 3500 else "❌")

    abstract_ru = re.search(r"аннотация[:\s]+(.{50,}?)(?=ключевые|keywords|түйін)",
                             full_text, re.IGNORECASE | re.DOTALL)
    if abstract_ru:
        aw = len(abstract_ru.group(1).split())
        add(2, "Объём аннотации (рус.)", "≤ 300 слов", f"{aw} слов",
            "✅" if aw <= 300 else "❌")
    else:
        add(2, "Аннотация (рус.)", "≤ 300 слов", "Не найдена", "⚠️")

    has_kaz = "аңдатпа" in full_text.lower()
    add(3, "Аннотация (каз.)", "Обязательно",
        "Есть" if has_kaz else "Отсутствует", "✅" if has_kaz else "❌")

    has_eng = bool(re.search(r"\babstract\b", full_text, re.IGNORECASE))
    add(4, "Abstract (англ.)", "Обязательно",
        "Есть" if has_eng else "Отсутствует", "✅" if has_eng else "❌")

    kw_match = re.search(r"(ключевые слова|keywords)[:\s]+(.+?)(\n|$)",
                          full_text, re.IGNORECASE)
    if kw_match:
        kw_list = [k.strip() for k in kw_match.group(2).split(";") if k.strip()]
        add(5, "Ключевые слова", "3–10, разделитель «;»", f"{len(kw_list)} слов",
            "✅" if 3 <= len(kw_list) <= 10 else "❌")
    else:
        add(5, "Ключевые слова", "3–10, разделитель «;»", "Не найдены", "⚠️")

    mrnt = bool(re.search(r"МРНТИ|\d{2}\.\d{2}\.\d{2}", full_text))
    add(6, "Код МРНТИ", "Обязателен",
        "Найден" if mrnt else "Отсутствует", "✅" if mrnt else "⚠️")

    orcid_count = len(re.findall(r"orcid\.org/\d{4}-\d{4}-\d{4}-\d{4}",
                                   full_text, re.IGNORECASE))
    add(7, "ORCID авторов", "Для каждого автора", f"{orcid_count} ORCID",
        "✅" if orcid_count >= 1 else "⚠️")

    sections = [
        (8,  "Введение",           "Введени"),
        (9,  "Материалы и методы", "Материал"),
        (10, "Результаты",         "Результат"),
        (11, "Обсуждение",         "Обсужден"),
        (12, "Заключение",         "Заключени"),
    ]
    for num, name, key in sections:
        found = key.lower() in full_text.lower()
        add(num, f"Раздел: {name}", "Обязателен",
            "Найден" if found else "Отсутствует", "✅" if found else "❌")

    contrib = ("вклад авторов" in full_text.lower() or
               "author contribution" in full_text.lower())
    add(13, "Вклад авторов", "CRediT-формат",
        "Найден" if contrib else "Отсутствует", "✅" if contrib else "❌")

    fund = "финансирован" in full_text.lower() or "funding" in full_text.lower()
    add(14, "Финансирование", "Обязательно",
        "Найдено" if fund else "Отсутствует", "✅" if fund else "❌")

    conflict = ("конфликт интересов" in full_text.lower() or
                "conflict of interest" in full_text.lower())
    add(15, "Конфликт интересов", "Явное заявление",
        "Найден" if conflict else "Отсутствует", "✅" if conflict else "❌")

    total_refs = len(re.findall(r"(?m)^\d+\.\s+", full_text))
    add(16, "Количество источников", "≥ 25 (research paper)",
        f"~{total_refs} источников", "✅" if total_refs >= 25 else "❌")

    doi_count = len(re.findall(r"https?://doi\.org/", full_text))
    add(17, "DOI в ссылках", "При наличии у источников", f"{doi_count} DOI",
        "✅" if doi_count >= 5 else "⚠️")

    apa_style = bool(re.search(r"\([A-ZА-Я][a-zA-Zа-яА-Я]+.*?\d{4}\)", full_text))
    add(18, "Стиль цитирования", "APA 7 (Автор, год)",
        "APA обнаружен" if apa_style else "Стиль не определён",
        "✅" if apa_style else "⚠️")

    suppl = ("вспомогательный материал" in full_text.lower() or
             "supplementary" in full_text.lower())
    add(19, "Вспомогательный материал", "Указан или отмечено «нет»",
        "Найден" if suppl else "Отсутствует", "⚠️")

    return results, full_text


uploaded_file = st.file_uploader(
    "📂 Загрузите статью в формате .docx",
    type=["docx"],
    help="Шаблон журнала Вестник ЕНУ, серия Химия/География, 2025"
)

if uploaded_file:
    with st.spinner("Анализируем статью..."):
        doc = Document(uploaded_file)
        results, full_text = check_article(doc)
        df = pd.DataFrame(results)

    passed = sum(1 for r in results if r["Статус"] == "✅")
    warned  = sum(1 for r in results if r["Статус"] == "⚠️")
    failed  = sum(1 for r in results if r["Статус"] == "❌")
    total   = len(results)
    score   = int(passed / total * 100)

    st.markdown("## 📊 Результаты проверки")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Всего",        total)
    c2.metric("✅ Выполнено", passed)
    c3.metric("⚠️ Внимание",  warned)
    c4.metric("❌ Не выпол.", failed)
    c5.metric("🏆 Соответствие", f"{score}%")

    color = "#198754" if score >= 80 else "#ffc107" if score >= 60 else "#dc3545"
    st.markdown(f"""
    <div style="background:#e9ecef;border-radius:10px;height:28px;margin:8px 0 20px 0;">
      <div style="background:{color};width:{score}%;height:28px;border-radius:10px;
                  display:flex;align-items:center;justify-content:center;
                  color:white;font-weight:bold;">{score}%</div>
    </div>""", unsafe_allow_html=True)

    def highlight(row):
        c = {"✅": "background-color:#d4edda",
             "⚠️": "background-color:#fff3cd",
             "❌": "background-color:#f8d7da"}
        return [c.get(row["Статус"], "")] * len(row)

    st.markdown("### 📋 Детальный отчёт")
    st.dataframe(df.style.apply(highlight, axis=1),
                 use_container_width=True, height=620)

    st.markdown("---")
    col_a, col_b = st.columns(2)
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    col_a.download_button("⬇️ Скачать CSV", csv_bytes,
                           "compliance_report.csv", "text/csv")
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Отчёт")
    col_b.download_button("⬇️ Скачать Excel", excel_buf.getvalue(),
                           "compliance_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    problems = [r for r in results if r["Статус"] in ("❌", "⚠️")]
    if problems:
        st.markdown("### ⚠️ Требует исправления")
        for p in problems:
            icon = "🔴" if p["Статус"] == "❌" else "🟡"
            st.markdown(
                f"{icon} **{p['Критерий']}** — {p['Найдено']}"
                f"  *(требование: {p['Требование']})*"
            )
else:
    st.info("👆 Загрузите .docx файл, чтобы начать проверку")
    st.markdown("""
### Что проверяется автоматически:
| Блок | Критерии |
|------|----------|
| 📝 Текст | Объём статьи (≥3500 слов) |
| 🌐 Мультиязычность | Аннотации рус / каз / англ, ключевые слова |
| 📌 Метаданные | МРНТИ, ORCID |
| 🗂 Структура | Введение, М&М, Результаты, Обсуждение, Заключение |
| 📚 Литература | Количество источников (≥25), DOI, стиль APA-7 |
| ✍️ Доп. разделы | Вклад авторов, финансирование, конфликт интересов |
""")
