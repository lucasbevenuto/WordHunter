import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
from docx import Document
import re
import os
import win32com.client  # Para arquivos .doc

# Função para extrair texto de PDF
def extrair_texto_pdf(caminho):
    try:
        with pdfplumber.open(caminho) as pdf:
            texto = "".join(pagina.extract_text() for pagina in pdf.pages if pagina.extract_text())
        return texto
    except Exception as e:
        return f"Erro ao ler PDF: {str(e)}"

# Função para extrair texto de DOCX
def extrair_texto_docx(caminho):
    try:
        doc = Document(caminho)
        texto = "\n".join(paragrafo.text for paragrafo in doc.paragraphs if paragrafo.text.strip())
        return texto
    except Exception as e:
        return f"Erro ao ler DOCX: {str(e)}"

# Função para extrair texto de DOC (usando Word via pywin32)
def extrair_texto_doc(caminho):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho)
        texto = doc.Content.Text
        doc.Close()
        word.Quit()
        return texto
    except Exception as e:
        return f"Erro ao ler DOC: {str(e)}"

# Função para processar um único arquivo com base nas regras
def processar_arquivo(caminho, regras):
    if not os.path.exists(caminho):
        return {"Erro": "Arquivo não encontrado"}

    extensao = caminho.lower().split('.')[-1]
    if extensao == "pdf":
        texto = extrair_texto_pdf(caminho)
    elif extensao == "docx":
        texto = extrair_texto_docx(caminho)
    elif extensao == "doc":
        texto = extrair_texto_doc(caminho)
    else:
        return {"Erro": "Formato não suportado (use PDF, DOCX ou DOC)"}

    if "Erro" in texto:
        return {"Erro": texto}

    resultado = {}
    for regra in regras:
        regra = regra.strip()
        if regra == "comodato":  # Busca exata por "comodato" (case sensitive)
            if re.search(r"comodato", texto):
                resultado["comodato"] = "Encontrado"
            else:
                resultado["comodato"] = "Não encontrado"
                
    # Aqui você pode adicionar outras regras, se desejar
    
        elif regra.lower() == "data":
            match = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
            resultado["Data"] = match.group(1) if match else "Não encontrado"
        # Exemplo: busca por índice (simples palavra após "Índice")
        elif regra.lower() == "índice":
            match = re.search(r"índice (\w+)", texto, re.IGNORECASE)
            resultado["Índice"] = match.group(1) if match else "Não encontrado"

    return resultado

# Função para processar múltiplos arquivos
def processar_varios():
    caminhos = filedialog.askopenfilenames(filetypes=[
        ("PDF files", "*.pdf"),
        ("Word files", "*.docx;*.doc"),
        ("All files", "*.*")
    ])
    if not caminhos:
        return
    
    regras = regras_entry.get().split(",")
    if not regras or regras == [""]:
        messagebox.showwarning("Atenção", "Insira pelo menos uma regra (ex.: comodato, Data, Índice)")
        return
    
    resultado_text.delete(1.0, tk.END)
    for caminho in caminhos:
        resultado = processar_arquivo(caminho, regras)
        resultado_text.insert(tk.END, f"Arquivo: {os.path.basename(caminho)}\n")
        for chave, valor in resultado.items():
            resultado_text.insert(tk.END, f"{chave}: {valor}\n")
        resultado_text.insert(tk.END, "-" * 40 + "\n")  # Separador entre arquivos

# Configuração da interface gráfica
janela = tk.Tk()
janela.title("Boot de Processamento - Múltiplos Arquivos")
janela.geometry("400x400")  # Aumentei o tamanho para caber mais resultados

tk.Label(janela, text="Regras (separadas por vírgula, ex.: comodato, Data, Índice):").pack(pady=5)
regras_entry = tk.Entry(janela, width=40)
regras_entry.pack(pady=5)
regras_entry.insert(0, "comodato, Data, Índice")  # Ajustei para "comodato" exato

tk.Button(janela, text="Selecionar Vários Arquivos e Processar", command=processar_varios).pack(pady=10)

resultado_text = tk.Text(janela, height=20, width=50)  # Aumentei a altura
resultado_text.pack(pady=10)

janela.mainloop()