# Импорт необходимых библиотек
import tkinter as tk  # Для создания графического интерфейса
from tkinter import filedialog, messagebox, ttk  # Диалоговые окна, сообщения, виджеты
import threading  # Для многопоточной работы
import os  # Для работы с файловой системой
import pandas as pd  # Для работы с табличными данными
import re  # Для регулярных выражений
from collections import Counter  # Для подсчета частот

# Пользовательский словарь для исключения
EXCLUDED_WORDS = {
    # Предлоги
    "в", "на", "с", "по", "из", "у", "к", "от", "до", "за", "о", "об", "со", "изо",
    # Союзы
    "и", "а", "но", "или", "либо", "что", "чтобы", "как", "потому", "также",
    # ГОСТы (шаблоны)
    "гост", "р", "стандарт", "iso", "ту", "ост", "снип", "сп", "гн", "рд", "санпин"
}

# Загрузка модели spaCy для русского языка
try:
    import spacy
    from spacy import displacy # Для визуализации синтаксических деревьев

    nlp = spacy.load("ru_core_news_sm")
except Exception:
    raise SystemExit(
        "⚠️ Модель spaCy 'ru_core_news_sm' не найдена. Установите командой:\n"
        "    pip install -U spacy && python -m spacy download ru_core_news_sm"
    )


# ---------------------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ --------------------------- #

# Очистка текста
def clean_links(text: str) -> str:
    text = str(text)
    text = re.sub(r"\(?\b(?:https?://|www\.)\S+\b\)?", "", text) # Удаление URL
    text = re.sub(r"=HYPERLINK\(\"[^\"]+\",\"([^\"]+)\"\)", r"\1", text) # Удаление гиперссылок Excel
    return re.sub(r"\(\s*\)", "", text).strip() # Удаление пустых скобок


# Удаление содержимого в скобках
def remove_brackets(text: str) -> str:
    text = str(text)
    text = re.sub(r"\([^)]*\)", "", text) # Удаление круглых скобок
    text = re.sub(r"\[[^\]]*\]", "", text) # Квадратных скобок
    text = re.sub(r"[\(\)\[\]]", "", text) # Оставшихся скобок
    return re.sub(r"\s+", " ", text).strip() # Нормализация пробелов


# Возвращает статистику текста: количество слов (без пунктуации) и количество символов
def basic_stats(text: str):
    doc = nlp(text)
    tokens = [t for t in doc if not t.is_punct and not t.is_space]
    return len(tokens), len(text)


# Извлекает из текстов n-граммы (биграммы/триграммы)
def extract_ngrams(texts, n=2, top_k=3):
    counts = Counter()
    examples = {}  # Будем хранить примеры предложений для каждой n-граммы

    for txt in texts:
        doc = nlp(txt)
        # Получаем леммы и исходные токены
        lemmas = [t.lemma_.lower() for t in doc if t.is_alpha and t.lemma_.lower() not in EXCLUDED_WORDS]
        tokens = [t.text.lower() for t in doc if t.is_alpha and t.lemma_.lower() not in EXCLUDED_WORDS]

        # Собираем n-граммы
        for i in range(len(lemmas) - n + 1):
            ngram_lemmas = tuple(lemmas[i:i + n])
            ngram_text = tuple(tokens[i:i + n])

            # Проверяем, что ни одно слово в n-грамме не входит в исключенные
            if not any(word in EXCLUDED_WORDS for word in ngram_lemmas):
                counts[ngram_lemmas] += 1
                # Сохраняем пример
                if ngram_lemmas not in examples:
                    examples[ngram_lemmas] = ' '.join(ngram_text)

    # Возвращаем n-граммы в лемматизированной форме с частотами и примерами в исходной форме
    return [(gram, cnt, examples[gram]) for gram, cnt in counts.most_common(top_k)]


# Проверяет, является ли текст полноценным предложением. Возвращает True,
# если в тексте ≥3 значимых слов и хотя бы один глагол
def is_real_sentence(text: str) -> bool:
    doc = nlp(text)
    tokens = [t for t in doc if not t.is_punct and not t.is_space]
    return len(tokens) >= 3 and any(t.pos_ == "VERB" for t in tokens)


# Создает сигнатуру ROOT + метки зависимостей его прямых потомков,
# упорядоченные по порядку в предложении
def clause_dep_signature(root_token):
    sig = [root_token.dep_] # Начинаем с корневого элемента

    # Берём только прямых потомков (children), без пунктуации
    children = [tok for tok in root_token.children
                if not tok.is_punct and not tok.is_space]
    children.sort(key=lambda t: t.i)

    # Исключаем некоторые типы зависимостей
    excluded_deps = {"det", "case", "punct"}
    for child in children:
        if child.dep_ not in excluded_deps:
            sig.append(child.dep_)
    return tuple(sig)


# Удаление ГОСТов
def remove_gost_phrases(text: str) -> str:
    """Удаляет конструкции типа ГОСТ Р 57528 – 2016 и слово ГОСТ."""
    pattern = r"\bгост(?:\s*[а-яa-zA-Z]*\s*\d{4,6}(?:\s*[–-]\s*\d{2,4})?)?\b"
    return re.sub(pattern, "", text, flags=re.IGNORECASE).strip()


# Топ синтаксических структур
def top_syntactic_structures(sentences, top_n: int = 3):
    """Анализирует предложения и находит top_n самых частых синтаксических структур
    Для каждой структуры сохраняет пример предложения и частоту встречаемости"""
    counter: Counter[tuple[str, ...]] = Counter()
    examples: dict[tuple[str, ...], tuple[str, "spacy.tokens.Doc"]] = {}

    for sent in sentences:
        sent = remove_gost_phrases(sent)
        doc = nlp(sent)
        # Находим корневые элементы предложений
        for root in (t for t in doc if t.dep_ == "ROOT" and not t.is_punct):
            sig = clause_dep_signature(root)
            if not sig or len(sig) < 2: # Пропускаем слишком короткие структуры
                continue
            counter[sig] += 1
            examples.setdefault(sig, (sent, doc))  # Сохраняем первый пример

    return [(sig, freq, *examples[sig]) for sig, freq in counter.most_common(top_n)]


# Преобразование кортежа с синтаксической структурой в строку
def sig_to_str(sig: tuple[str, ...]) -> str:
    return "+".join(sig)


# Находит первое предложение, содержащее указанную n-грамму (по леммам слов)
def find_sentence_with_ngram(sentences, ngram_lemmas):
    n = len(ngram_lemmas)
    for sent in sentences:
        sent = remove_gost_phrases(sent)
        # Ищем по леммам, но возвращаем предложение в исходной форме
        lemmas = [t.lemma_.lower() for t in nlp(sent) if t.is_alpha and t.lemma_.lower() not in EXCLUDED_WORDS]
        for i in range(len(lemmas) - n + 1):
            if tuple(lemmas[i:i + n]) == ngram_lemmas:
                return sent
    return None


# ---------------------- GUI ------------------------------------------------ #

# Главный класс приложения с графическим интерфейсом
class NLPApp(tk.Tk):
    # Инициализирует главное окно приложения и его состояние
    def __init__(self):
        super().__init__()
        self.title("Анализатор русских текстов (spaCy)")
        self.geometry("1100x800")
        self.resizable(False, False)

        # состояние
        self.file_path = None # Путь к загруженному файлу
        self.cleaned_df = None # Очищенные данные
        self.report_txt = "" # Текст отчета
        self.dep_htmls = [] # Список HTML-деревьев зависимостей

        self._build_ui() # Построение интерфейса

    # Создает все элементы пользовательского интерфейса
    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=10)

        # Кнопки управления
        self.btn_load = tk.Button(top, text="📂 Загрузить Excel", command=self.load_file)
        self.btn_load.pack(side=tk.LEFT, padx=5)

        self.btn_run = tk.Button(top, text="🔍 Запустить анализ", state=tk.DISABLED,
                                 command=lambda: threading.Thread(target=self.analyze, daemon=True).start())
        self.btn_run.pack(side=tk.LEFT, padx=5)

        self.btn_clear = tk.Button(top, text="🧹 Очистить", command=self.clear_output)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        self.btn_save_clean = tk.Button(top, text="💾 Сохранить очищенные", state=tk.DISABLED, command=self.save_cleaned)
        self.btn_save_clean.pack(side=tk.LEFT, padx=5)

        self.btn_save_report = tk.Button(top, text="📝 Сохранить отчёт", state=tk.DISABLED, command=self.save_report)
        self.btn_save_report.pack(side=tk.LEFT, padx=5)

        self.btn_save_trees = tk.Button(top, text="🌳 Сохранить деревья", state=tk.DISABLED, command=self.save_trees)
        self.btn_save_trees.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(top, orient="horizontal", length=240, mode="determinate")
        self.progress.pack(side=tk.RIGHT, padx=5)

        self.out = tk.Text(self, wrap=tk.WORD, state=tk.DISABLED)
        self.out.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))

    # Выводит сообщение в текстовое поле интерфейса
    def log(self, msg):
        self.out.config(state=tk.NORMAL)
        self.out.insert(tk.END, msg)
        self.out.see(tk.END)
        self.out.config(state=tk.DISABLED)

    # Обновляет значение прогресс-бара
    def set_progress(self, val):
        self.progress["value"] = val
        self.update_idletasks()

    # Открывает диалог выбора Excel-файла для анализа
    def load_file(self):
        path = filedialog.askopenfilename(title="Выберите Excel-файл", filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            self.btn_run.config(state=tk.NORMAL)
            self.log(f"Файл выбран: {os.path.basename(path)}\n")

    # Основная функция анализа
    def analyze(self):
        if not self.file_path:
            return
        try:
            # Загрузка и первичная обработка данных
            self.set_progress(0)
            self.log("Чтение файла...\n")
            df = pd.read_excel(self.file_path).dropna(how="all")
            df = df.replace(['-', '—', '–'], pd.NA).dropna(subset=["Контекст (рус)"])

            # Очистка текста
            self.log("Очистка...\n")
            df["Контекст (рус)"] = (df["Контекст (рус)"]
                                    .apply(clean_links)
                                    .apply(remove_brackets)
                                    .str.replace('\n', ' ', regex=False))
            df = df[df["Контекст (рус)"].str.strip().ne('')]
            df = df[df["Контекст (рус)"].apply(is_real_sentence)]
            self.set_progress(15)

            # Cтатистика текста
            self.log("Расчёт статистики...\n")
            df["num_words"], df["num_chars"] = zip(*df["Контекст (рус)"].map(basic_stats))
            rows = len(df)
            w_sum, c_sum = df["num_words"].sum(), df["num_chars"].sum()
            w_avg, c_avg = df["num_words"].mean(), df["num_chars"].mean()
            self.set_progress(30)

            # Топ-3 биграмм и триграмм
            self.log("Извлечение n-грамм...\n")
            top_bigrams = extract_ngrams(df["Контекст (рус)"], 2, 3)
            top_trigrams = extract_ngrams(df["Контекст (рус)"], 3, 3)
            self.set_progress(50)

            # Анализ синтаксических структур
            self.log("Поиск синтаксических структур...\n")
            sentences = df["Контекст (рус)"].tolist()
            top_structs = top_syntactic_structures(sentences, 3)
            self.set_progress(65)

            # Подготовка данных для отчета
            self.dep_htmls.clear() # Очистка предыдущих визуализаций
            rep = [
                f"Файл: {os.path.basename(self.file_path)}",
                f"Всего строк: {rows}",
                f"Всего слов: {w_sum}",
                f"Всего символов: {c_sum}",
                f"Среднее число слов в строке: {w_avg:.0f}",
                f"Среднее число символов в строке: {c_avg:.0f}",
            ]

            # Раздел отчета по биграммам
            rep.append("\nТоп биграмм:")
            for idx, (bg, cnt, example) in enumerate(top_bigrams, 1):
                rep.append(f"  {idx}. {' '.join(bg)} — {cnt} (пример: {example})")
                # Поиск полного предложения с этой биграммой
                sent = find_sentence_with_ngram(sentences, bg)
                if sent:
                    doc = nlp(sent)
                    label = f"bigram_{idx}"
                    # Генерация HTML для визуализации зависимостей
                    self.dep_htmls.append((label, displacy.render(doc, style="dep", page=True)))
                    rep.append(f"     ▶ полный пример: {sent}")

            # Раздел отчета по триграммам
            rep.append("\nТоп триграмм:")
            for idx, (tg, cnt, example) in enumerate(top_trigrams, 1):
                rep.append(f"  {idx}. {' '.join(tg)} — {cnt} (пример: {example})")
                sent = find_sentence_with_ngram(sentences, tg)
                if sent:
                    doc = nlp(sent)
                    label = f"trigram_{idx}"
                    self.dep_htmls.append((label, displacy.render(doc, style="dep", page=True)))
                    rep.append(f"     ▶ полный пример: {sent}")

            # Раздел отчета по синтаксическим структурам
            if top_structs:
                rep.append("\nТоп синтаксических структур:")
                for idx, (struct, freq, sent, doc) in enumerate(top_structs, 1):
                    rep.append(f"\n#{idx} (встречается {freq}×): {sig_to_str(struct)}")
                    rep.append(sent)
                    rep.append("Подробный разбор:")
                    # Добавление морфологического разбора для каждого токена в предложении
                    for tok in doc:
                        if tok.is_punct:
                            continue
                        rep.append(f"{tok.text:<15} {tok.pos_:<5} {tok.dep_:<10} → {tok.head.text}")

                    label = f"struct_{idx}"
                    self.dep_htmls.append((label, displacy.render(doc, style="dep", page=True)))
            else:
                rep.append("\nСинтаксические структуры не найдены.")

            # Вывод результатов
            self.report_txt = "\n".join(rep)
            self.log("\nАнализ завершён.\n")
            self.log(self.report_txt + "\n")

            # Обновление интерфейса
            self.cleaned_df = df.drop(columns=["num_words", "num_chars"], errors="ignore")
            # Активация кнопок сохранения
            self.btn_save_clean.config(state=tk.NORMAL)
            self.btn_save_report.config(state=tk.NORMAL)
            if self.dep_htmls:
                self.btn_save_trees.config(state=tk.NORMAL)
            self.set_progress(100) # Завершение прогресса

        except Exception as exc:
            # Обработка ошибок с выводом сообщения
            messagebox.showerror("Ошибка", str(exc))
            self.log(f"Ошибка: {exc}\n")
            self.set_progress(0)

    # Очищает все поля вывода и сбрасывает состояние программы
    def clear_output(self):
        self.out.config(state=tk.NORMAL)
        self.out.delete("1.0", tk.END)
        self.out.config(state=tk.DISABLED)
        self.set_progress(0)
        self.cleaned_df = None
        self.report_txt = ""
        self.dep_htmls = []
        for btn in (self.btn_save_clean, self.btn_save_report, self.btn_save_trees):
            btn.config(state=tk.DISABLED)

    # Сохраняет очищенные данные в новый Excel-файл
    def save_cleaned(self):
        if self.cleaned_df is None:
            return
        path = filedialog.asksaveasfilename(title="Сохранить очищенные данные", defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.cleaned_df.to_excel(path, index=False)
            messagebox.showinfo("Сохранено", f"Очищенные данные → {path}")

    # Сохраняет текстовый отчет в файл
    def save_report(self):
        if not self.report_txt:
            return
        path = filedialog.asksaveasfilename(title="Сохранить отчёт", defaultextension=".txt",
                                            filetypes=[("Text", "*.txt")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.report_txt)
            messagebox.showinfo("Сохранено", f"Отчёт → {path}")

    # Сохраняет визуализации синтаксических деревьев в HTML-файлы
    def save_trees(self):
        if not self.dep_htmls:
            return
        dir_path = filedialog.askdirectory(title="Выберите папку для HTML-деревьев")
        if not dir_path:
            return
        for label, html in self.dep_htmls:
            file_path = os.path.join(dir_path, f"{label}.html")
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(html)
        messagebox.showinfo("Сохранено", f"Сохранено {len(self.dep_htmls)} HTML-файлов в {dir_path}")


if __name__ == "__main__":
    NLPApp().mainloop()