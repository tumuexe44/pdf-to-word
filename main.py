# Kütüphaneler
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document
from PIL import Image, ImageTk
import time
# PDF to Word Dönüşümü
def pdf_to_word(pdf_file, word_file):
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()

# Word to PDF Dönüşümü
def word_to_pdf(word_file, pdf_file):
    doc = Document(word_file)
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter

    c.setFont("Helvetica", 12)
    y_position = height - 40  # Başlangıç yüksekliği

    for para in doc.paragraphs:
        c.drawString(40, y_position, para.text)
        y_position -= 15  # satır yüksekliği

        # Sayfa sonuna geldiğinde yeni bir sayfa oluştur
        if y_position < 40:
            c.showPage()
            c.setFont("Helvetica", 12)
            y_position = height - 40
    c.save()

# Tkinter Arayüzü
def open_file():
    file_path = filedialog.askopenfilename()
    return file_path

def pdf_to_word_click():
    pdf_file = open_file()
    if pdf_file:
        word_file = pdf_file.replace('.pdf', '.docx')
        pdf_to_word(pdf_file, word_file)
        messagebox.showinfo("Başarılı", f"{word_file} oluşturuldu!")

def word_to_pdf_click():
    word_file = open_file()
    if word_file:
        pdf_file = word_file.replace('.docx', '.pdf')
        word_to_pdf(word_file, pdf_file)
        messagebox.showinfo('Başarılı', f"{pdf_file} oluşturuldu!")

# Tkinter Penceresi
window = tk.Tk()
window.title("PDF - Word Dönüştürücü")
window.geometry("400x300")  # Pencere boyutu ayarlandı
window.configure(bg='#3d2b1f')

# İkonları yükle ve boyutlandır
def load_image(file_path, size=(32,32)):
    try:
        img = Image.open(file_path)
        img = img.resize(size, Image.Resampling.LANCZOS)  # Resampling modülü kullanılarak
        return ImageTk.PhotoImage(img)
    except FileNotFoundError:
        print(f"Dosya bulunamadı: {file_path}")
        return None

pdf_icon = load_image('pdf_icon.png', (32, 32))  # Boyutu 32x32 piksel olarak ayarladık
word_icon = load_image('word_icon.png', (32, 32))  # Boyutu 32x32 piksel olarak ayarladık
logo_image = load_image('logo.png', (200, 100))   # Logo boyutunu ayarladık

# Butonları oluştur
def create_button(parent, text, command, icon=None):
    return tk.Button(parent, text=text, command=command,
                     image=icon, compound='left' if icon else 'none',
                     bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'),
                     relief='raised', bd=4, width=180, height=50)

pdf_to_word_button = create_button(window, "PDF to Word", pdf_to_word_click, pdf_icon)
word_to_pdf_button = create_button(window, "Word to PDF", word_to_pdf_click, word_icon)

pdf_to_word_button.pack(pady=10)
word_to_pdf_button.pack(pady=10)

# Logo ekle
if logo_image:
    logo_label = tk.Label(window, image=logo_image, bg='#f0f8ff')
    logo_label.pack(pady=10)

# İlerleme çubuğu
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(window, orient='horizontal', length=300, mode='determinate', variable=progress_var)
progress_bar.pack(pady=20)

# Yüzdeyi Gösterecek Label

progress_label = tk.Label(window,text= "0%")
progress_label.pack(pady=10)

# Yükleme Çubuğu ve Yüzdeyi Simüle Etme
def simulate_processing():
    for i in range(101):
        progress_var.set(i)
        progress_label.config(text=f"{i}%") #Yüzdeyi göster
        if i < 100:
            progress_bar['style'] = 'red.Horizontal.TProgressbar' # Tamamlanmamışsa kırmızı
        else:
            progress_bar['style'] = 'green.Horizontal.TProgressbar' # Tamamlandığında yeşil
        window.update_idletasks()
        time.sleep(0.05) # Simülasyon için bekleme süresi
# Tema ve Stil
style = ttk.Style()
style.configure("TButton", font=("Arial", 12, "bold"), padding=10)

# Stilleri Ayarlama
style = ttk.Style()
style.configure("red.Horizontal.TProgressbar", troughcolor='white', background='red')
style.configure("green.Horizontal.TProgressbar", troughcolor='white', background='green')

# Pencereyi başlat
window.mainloop()
