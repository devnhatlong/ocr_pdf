import os
import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, font, ttk

# Đường dẫn Tesseract (chỉnh theo máy bạn)
pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"

# Đường dẫn Poppler (tự động theo thư mục project)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
POPPLER_PATH = os.path.join(BASE_DIR, "poppler-24.08.0", "Library", "bin")

# Tessdata directories
TESSDATA_BEST_DIR = os.path.join(BASE_DIR, "tessdata_best")
TESSDATA_DEFAULT_DIR = os.path.join(BASE_DIR, "Tesseract-OCR", "tessdata")

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
    images = convert_from_path(pdf_path, dpi=360, poppler_path=POPPLER_PATH)
    text_all = ""

    # Decide which tessdata to use. If user enabled tessdata_best and it exists, prefer it.
    tessdata_dir = TESSDATA_DEFAULT_DIR
    try:
        if use_tessdata_best.get() and os.path.isdir(TESSDATA_BEST_DIR):
            tessdata_dir = TESSDATA_BEST_DIR
    except NameError:
        # If GUI variables aren't defined (calling from non-GUI code), keep default
        pass

    # Set TESSDATA_PREFIX so Tesseract finds models in chosen folder
    os.environ["TESSDATA_PREFIX"] = tessdata_dir

    # language selection (fallback to vie+eng)
    try:
        lang = lang_var.get().strip() or "vie+eng"
    except NameError:
        lang = "vie+eng"

    for i, img in enumerate(images):
        pre_img = preprocess_image(img)
        text = pytesseract.image_to_string(pre_img, lang=lang)
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
        # insert into treeview
        try:
            file_tree.insert('', 'end', values=(filename,))
        except NameError:
            # fallback to legacy listbox if tree not available
            file_list.insert(tk.END, filename)
        text_results[filename] = text
        messagebox.showinfo("Thành công", f"OCR thành công: {filename}")
    except Exception as e:
        print("OCR error:", e)
        messagebox.showerror("Lỗi", f"Không thể OCR file:\n{e}")

def on_file_select(event):
    # support Treeview selection and fallback to Listbox
    try:
        selection = file_tree.selection()
        if not selection:
            return
        filename = file_tree.item(selection[0])['values'][0]
    except NameError:
        selection = file_list.curselection()
        if not selection:
            return
        filename = file_list.get(selection[0])
    text_box.delete(1.0, tk.END)
    text_box.insert(tk.END, text_results[filename])

def save_text():
    try:
        selection = file_tree.selection()
        if not selection:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn file đã OCR để lưu.")
            return
        filename = file_tree.item(selection[0])['values'][0]
    except NameError:
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

# Set application icon from image/ocr_logo_app.png if available
try:
    icon_path = os.path.join(BASE_DIR, "image", "ocr_logo_app.png")
    if os.path.exists(icon_path):
        photo = tk.PhotoImage(file=icon_path)
        # iconphoto works cross-platform for PNG
        root.iconphoto(False, photo)
    else:
        # Fallback: if .ico exists in project root, use it
        ico_fallback = os.path.join(BASE_DIR, "ocr_logo_app.ico")
        if os.path.exists(ico_fallback):
            root.iconbitmap(ico_fallback)
except Exception:
    # If icon loading fails, continue without crashing the app
    pass

# Use ttk for a more modern look
style = ttk.Style()
try:
    style.theme_use('clam')
except Exception:
    pass

# Frame trái: danh sách file
frame_left = ttk.Frame(root)
frame_left.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.Y)

btn_open = ttk.Button(frame_left, text="📂 Mở file PDF", command=open_pdf)
btn_open.pack(pady=5, fill='x')

btn_save = ttk.Button(frame_left, text="💾 Lưu text ra file", command=save_text)
btn_save.pack(pady=5, fill='x')

# Treeview for file list (modern)
file_tree = ttk.Treeview(frame_left, columns=("filename",), show='headings', height=12)
file_tree.heading('filename', text='Files')
file_tree.column('filename', width=300)
file_tree.pack(pady=5, fill=tk.Y)
file_tree.bind('<<TreeviewSelect>>', on_file_select)

# Checkbox to enable using tessdata_best (if available)
use_tessdata_best = tk.BooleanVar(value=os.path.isdir(TESSDATA_BEST_DIR))
chk_tessbest = tk.Checkbutton(
    frame_left,
    text="Sử dụng tessdata_best (mô hình chất lượng cao)",
    variable=use_tessdata_best
)
chk_tessbest.pack(pady=5, anchor="w")

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
