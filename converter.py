import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # Adicionado ttk
import pdfkit
from pdf2docx import Converter
from fpdf import FPDF
import pytesseract
from PIL import Image
from docx import Document
import threading

# Defina o caminho para o Tesseract OCR (se necessário)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Defina o caminho para o wkhtmltopdf
caminho_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# Configurar o pdfkit
config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)

# Função para exibir a tela de loading
def mostrar_loading(janela_principal):
    loading_window = tk.Toplevel(janela_principal)
    loading_window.title("Carregando...")
    loading_window.geometry("300x100")
    loading_window.transient(janela_principal)  # Faz a janela de loading ficar na frente da principal

    tk.Label(loading_window, text="Processando, aguarde...", font=("Arial", 12)).pack(pady=20)
    progress = ttk.Progressbar(loading_window, mode="indeterminate")
    progress.pack(pady=10)
    progress.start()

    return loading_window

# Função para fechar a tela de loading
def fechar_loading(loading_window):
    loading_window.destroy()

# Funções de Conversão
def html_to_pdf(input_file, output_file):
    try:
        pdfkit.from_file(input_file, output_file, configuration=config)
        messagebox.showinfo("Sucesso", f"Arquivo '{output_file}' gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def pdf_to_word(input_file, output_file):
    try:
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        messagebox.showinfo("Sucesso", f"Arquivo '{output_file}' gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def txt_to_pdf(input_file, output_file):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        with open(input_file, "r", encoding="utf-8") as file:
            for line in file:
                pdf.cell(200, 10, line, ln=True)
        pdf.output(output_file)
        messagebox.showinfo("Sucesso", f"Arquivo '{output_file}' gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def txt_to_docx(input_file, output_file):
    try:
        doc = Document()
        with open(input_file, "r", encoding="utf-8") as file:
            for line in file:
                doc.add_paragraph(line)
        doc.save(output_file)
        messagebox.showinfo("Sucesso", f"Arquivo '{output_file}' gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def image_to_text(input_file):
    try:
        text = pytesseract.image_to_string(Image.open(input_file))
        messagebox.showinfo("Texto Extraído", text)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao extrair texto: {e}")

# Função para escolher arquivo e converter
def escolher_arquivo(tipo):
    file_path = filedialog.askopenfilename()
    if not file_path:
        return

    output_path = filedialog.asksaveasfilename(defaultextension=".pdf" if tipo in ["html_to_pdf", "txt_to_pdf"] else ".docx")
    if not output_path:
        return

    # Mostra a tela de loading
    loading_window = mostrar_loading(root)

    # Executa a conversão em uma thread separada
    def executar_conversao():
        try:
            if tipo == "html_to_pdf":
                html_to_pdf(file_path, output_path)
            elif tipo == "pdf_to_word":
                pdf_to_word(file_path, output_path)
            elif tipo == "txt_to_pdf":
                txt_to_pdf(file_path, output_path)
            elif tipo == "txt_to_docx":
                txt_to_docx(file_path, output_path)
            elif tipo == "image_to_text":
                image_to_text(file_path)
                return

            messagebox.showinfo("Sucesso", f"Conversão concluída! Arquivo salvo em:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        finally:
            # Fecha a tela de loading após a conclusão
            fechar_loading(loading_window)

    # Inicia a thread de conversão
    threading.Thread(target=executar_conversao).start()

# Criando a Interface Gráfica
root = tk.Tk()
root.title("Conversor de Documentos")
root.geometry("400x300")

# Adicionando o ícone da aplicação
root.iconbitmap("convert.ico")  # Substitua "icone.ico" pelo caminho do seu ícone

tk.Label(root, text="Escolha um tipo de conversão:", font=("Arial", 12)).pack(pady=10)

tk.Button(root, text="HTML → PDF", command=lambda: escolher_arquivo("html_to_pdf")).pack(pady=5)
tk.Button(root, text="PDF → Word", command=lambda: escolher_arquivo("pdf_to_word")).pack(pady=5)
tk.Button(root, text="TXT → PDF", command=lambda: escolher_arquivo("txt_to_pdf")).pack(pady=5)
tk.Button(root, text="TXT → Word", command=lambda: escolher_arquivo("txt_to_docx")).pack(pady=5)
tk.Button(root, text="Imagem → Texto (OCR)", command=lambda: escolher_arquivo("image_to_text")).pack(pady=5)

# Adicionando o texto "Feito por WinmeiSoftware" na parte inferior
tk.Label(root, text="Feito por WimeiSoftware", font=("Arial", 10), fg="gray").pack(side=tk.BOTTOM, pady=10)

root.mainloop()