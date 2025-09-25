import os
import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, font

# Đường dẫn Tesseract (chỉnh theo máy bạn)
pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"

# Đường dẫn Poppler (tự động theo thư mục project)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
POPPLER_PATH = os.path.join(BASE_DIR, "poppler-24.08.0", "Library", "bin")

# ----------------------------
# Xử lý ảnh trước khi OCR
# ----------------------------
def preprocess_image(pil_img):
    """Tiền xử lý ảnh để OCR chính xác hơn"""
    img = np.array(pil_img)
    img = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    # Adaptive threshold
    img = cv2.adaptiveThreshold(
        img, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35, 11
    )

    # Giảm nhiễu
    img = cv2.medianBlur(img, 3)

    return img

# ----------------------------
# OCR PDF
# ----------------------------
def ocr_pdf(pdf_path):
    images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    text_all = ""
    for i, img in enumerate(images):
        pre_img = preprocess_image(img)
        text = pytesseract.image_to_string(pre_img, lang="vie+eng")
        text_all += f"\n--- Page {i+1} ---\n{text}"
    return text_all

# ----------------------------
# GUI functions
# ----------------------------
def open_pdf():
    file_path = filedialog.askopenfilename(
        title="Chọn file PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not file_path:
        return

    try:
        text = ocr_pdf(file_path)
        filename = os.path.basename(file_path)
        file_list.insert(tk.END, filename)
        text_results[filename] = text
        messagebox.showinfo("Thành công", f"OCR thành công: {filename}")
    except Exception as e:
        print("OCR error:", e)
        messagebox.showerror("Lỗi", f"Không thể OCR file:\n{e}")

def on_file_select(event):
    selection = file_list.curselection()
    if not selection:
        return
    filename = file_list.get(selection[0])
    text_box.delete(1.0, tk.END)
    text_box.insert(tk.END, text_results[filename])

def save_text():
    selection = file_list.curselection()
    if not selection:
        messagebox.showwarning("Chưa chọn file", "Vui lòng chọn file đã OCR để lưu.")
        return
    filename = file_list.get(selection[0])
    text = text_results.get(filename, "")
    if not text:
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        initialfile=filename.replace(".pdf", ".txt")
    )
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(text)
        messagebox.showinfo("Đã lưu", f"Lưu thành công: {save_path}")

# ----------------------------
# GUI Layout
# ----------------------------
root = tk.Tk()
root.title("OCR PDF Tool (Phát triển bởi Nhật Long)")

# Frame trái: danh sách file
frame_left = tk.Frame(root)
frame_left.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.Y)

btn_open = tk.Button(frame_left, text="📂 Mở file PDF", command=open_pdf)
btn_open.pack(pady=5)

btn_save = tk.Button(frame_left, text="💾 Lưu text ra file", command=save_text)
btn_save.pack(pady=5)

file_list = tk.Listbox(frame_left, width=40, height=20)
file_list.pack(pady=5, fill=tk.Y)
file_list.bind("<<ListboxSelect>>", on_file_select)

# Frame phải: hiển thị text OCR
frame_right = tk.Frame(root)
frame_right.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

# Font Times New Roman
custom_font = font.Font(family="Times New Roman", size=12)

text_box = scrolledtext.ScrolledText(
    frame_right,
    wrap=tk.WORD,
    width=80,
    height=30,
    font=custom_font
)
text_box.pack(fill=tk.BOTH, expand=True)

# Bộ nhớ text OCR theo filename
text_results = {}

root.mainloop()
