import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter


def merge_pdfs_gui():
    root = tk.Tk()
    root.withdraw()

    # выбираем файлы
    file_paths = filedialog.askopenfilenames(
        title="Выберите PDF-файлы",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_paths:
        print("Файлы не выбраны.")
        return

    writer = PdfWriter()

    for path in file_paths:
        print(f"Добавляю: {path}")
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)

    # выбираем место сохранения
    save_path = filedialog.asksaveasfilename(
        title="Сохранить как",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile="merged.pdf"
    )

    if not save_path:
        print("Сохранение отменено.")
        return

    with open(save_path, "wb") as f:
        writer.write(f)

    print(f"Готово! Сохранено как: {save_path}")
    messagebox.showinfo("Готово", f"Файл создан:\n{save_path}")


if __name__ == "__main__":
    merge_pdfs_gui()