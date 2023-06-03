
import difflib
import tkinter as tk
from tkinter import filedialog
import os
import string
import re
from difflib import SequenceMatcher
import docx2txt
import win32com.client as win32
def preprocess_text(text):
    # Преобразование текста к нижнему регистру
    text = text.lower()

    # Удаление знаков препинания
    text = text.translate(str.maketrans('', '', string.punctuation))

    # Удаление лишних пробелов
    text = re.sub(r'\s+', ' ', text)

    # Удаление цитирования
    text = re.sub(r'\[.*?\]', '', text)

    # Удаление библиографии
    text = re.sub(r'references?(\n|$).*', '', text, flags=re.IGNORECASE)

    return text

def preprocess_text1(text):
    # Преобразование текста к нижнему регистру
    text = text.lower()

    # Удаление знаков препинания
    text = text.translate(str.maketrans('', '', string.punctuation))

    # Удаление лишних пробелов
    text = re.sub(r'\s+', ' ', text)

    return text
def read_docx(file_path):
    text = docx2txt.process(file_path)
    return text


def read_doc(file_path):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    return text


def read_text_file(file_path):
    with open(file_path, 'r') as file:
        text = file.read()
    return text


def read_document(file_path):
    if file_path.endswith('.docx'):
        return read_docx(file_path)
    elif file_path.endswith('.doc'):
        return read_doc(file_path)
    elif file_path.endswith('.txt'):
        return read_text_file(file_path)
    else:
        return ''


def calculate_similarity(text1, text2):
    # Предварительная обработка текстов
    text1 = preprocess_text(text1)
    text2 = preprocess_text(text2)

    # Разбиваем тексты на множества уникальных слов
    set1 = set(text1.split())
    set2 = set(text2.split())

    # Вычисление схожести текстов с помощью множественного коэффициента Жаккара
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    similarity = intersection / union if union != 0 else 0

    return similarity


def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[('All Supported Formats', '*.txt;*.doc;*.docx'), ('Text Files', '*.txt'),
                   ('Word Documents', '*.doc;*.docx')])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(tk.END, file_path)


def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(tk.END, folder_path)



def compare_documents():
    selected_file = file_entry.get()
    directory = folder_entry.get()

    if selected_file and directory:
        original_text = read_document(selected_file)
        remaining_text = preprocess_text(original_text)

        similarities = []

        for filename in os.listdir(directory):
            if filename.endswith('.txt') or filename.endswith('.doc') or filename.endswith('.docx'):
                doc_path = os.path.join(directory, filename)
                doc_text = read_document(doc_path)
                doc_text1 = read_document(doc_path)

                similarity = calculate_similarity(original_text, doc_text)
                similarities.append((filename, similarity))

                # Вычисление различий между строками
                diff = difflib.SequenceMatcher(None, remaining_text, preprocess_text(doc_text1))
                matches = diff.get_matching_blocks()

                # Удаление совпадающего текста из переменной
                for match in matches:
                    start = match.a
                    end = match.a + match.size
                    remaining_text = remaining_text[:start] + remaining_text[end:]

        similarities.sort(key=lambda x: x[1], reverse=True)

        max_similarity = max(similarities, key=lambda x: x[1])[1]

        results_text.delete(1.0, tk.END)
        percent_unique = 100 - (len(preprocess_text1(remaining_text)) / len(preprocess_text1(original_text)) * 100)
        results_text.insert(tk.END, f'Общий процент плагиата: {percent_unique:.2f}%\n-----\n')

        # results_text.insert(tk.END, f'Максимальное сходство: {max_similarity:.2%}\n-----\n')

        for item in similarities:
            results_text.insert(tk.END, f'{item[0]}: {item[1]:.2%}\n')

        # results_text.insert(tk.END, f'Всего сравниваемых документов: {len(similarities)}\n')





# Создание графического интерфейса
window = tk.Tk()
window.title('Антиплагиат')
window.geometry('400x300')

file_label = tk.Label(window, text='Выберите файл:')
file_label.pack()

file_entry = tk.Entry(window, width=40)
file_entry.pack()

file_button = tk.Button(window, text='Обзор', command=select_file)
file_button.pack()

folder_label = tk.Label(window, text='Выберите папку:')
folder_label.pack()

folder_entry = tk.Entry(window, width=40)
folder_entry.pack()

folder_button = tk.Button(window, text='Обзор', command=select_folder)
folder_button.pack()

compare_button = tk.Button(window, text='Сравнить', command=compare_documents)
compare_button.pack()

results_label = tk.Label(window, text='Результаты:')
results_label.pack()

results_text = tk.Text(window, height=10)
results_text.pack()

window.mainloop()
