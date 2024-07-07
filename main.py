import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import datetime
from docx2pdf import convert
import os

from ScriptHipo import hipo_creator
from ScriptContrato import contrato_creator
from ScriptProcuracao import proc_creator

local_dir = os.getcwd()

def create_document(nome_doc_entry, nome_entry, std_civil_entry, ocupacao_entry, cpf_entry, rg_entry, endereco_entry, local_entry, dia_entry, mes_entry, ano_entry, choice_var, valor_mensal_entry, valor_total_entry, valor_assinatura_entry):
    nome_doc = nome_doc_entry.get()
    nome = nome_entry.get()
    std_civil = std_civil_entry.get()
    ocupacao = ocupacao_entry.get()
    cpf = cpf_entry.get() 
    cpfVerified = False   
    rg = rg_entry.get()
    endereco = endereco_entry.get()
    local = local_entry.get()

    dia = dia_entry.get()
    mes = mes_entry.get()
    ano = ano_entry.get()

    choice = choice_var.get()

    valor_mensal = valor_mensal_entry.get()
    valor_total = valor_total_entry.get()  
    valor_assinatura = valor_assinatura_entry.get()

    def valida_cpf(cpf: str) -> bool:
        cpf = ''.join(filter(str.isdigit, cpf))
        
        if len(cpf) != 11:
            return False

        if cpf == cpf[0] * len(cpf):
            return False

        soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
        primeiro_digito = (soma * 10 % 11) % 10

        soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
        segundo_digito = (soma * 10 % 11) % 10

        return cpf[-2:] == f'{primeiro_digito}{segundo_digito}'


    if nome_doc and nome and std_civil and ocupacao and valida_cpf(cpf) == True and rg and endereco and local and dia and mes and ano and choice and valor_mensal and valor_total and valor_assinatura:
        if choice == 1:
            proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        elif choice == 2:
            hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        elif choice == 3:
            proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
            hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
            contrato_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano, valor_mensal, valor_total, valor_assinatura)
        else:
            messagebox.showerror("Erro", "Escolha inválida!")
    elif(not valida_cpf(cpf)):
        messagebox.showerror("Erro", "CPF Inválido")
    else:
        messagebox.showerror("Erro", "Preencha todos os campos!")


def convert_to_pdf(docx_path):
    try:
        convert(docx_path)
        messagebox.showinfo("Sucesso", f"{docx_path} convertido para PDF com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter para PDF: {str(e)}")

def create_gui():
    root = tk.Tk()
    root.title("Criação de Documentos")

    mainframe = ttk.Frame(root, padding="20")
    mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)

    ttk.Label(mainframe, text="Nome do Documento:").grid(column=1, row=1, sticky=tk.W)
    nome_doc_entry = ttk.Entry(mainframe, width=40)
    nome_doc_entry.grid(column=2, row=1, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Nome da Parte:").grid(column=1, row=2, sticky=tk.W)
    nome_entry = ttk.Entry(mainframe, width=40)
    nome_entry.grid(column=2, row=2, sticky=(tk.W, tk.E))


    ttk.Label(mainframe, text="Estado Civil:").grid(column=1, row=3, sticky=tk.W)
    std_civil_entry = ttk.Entry(mainframe, width=40)
    std_civil_entry.grid(column=2, row=3, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Ocupação:").grid(column=1, row=4, sticky=tk.W)
    ocupacao_entry = ttk.Entry(mainframe, width=40)
    ocupacao_entry.grid(column=2, row=4, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="CPF:").grid(column=1, row=5, sticky=tk.W)
    cpf_entry = ttk.Entry(mainframe, width=40)
    cpf_entry.grid(column=2, row=5, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="RG:").grid(column=1, row=6, sticky=tk.W)
    rg_entry = ttk.Entry(mainframe, width=40)
    rg_entry.grid(column=2, row=6, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Endereço:").grid(column=1, row=7, sticky=tk.W)
    endereco_entry = ttk.Entry(mainframe, width=40)
    endereco_entry.grid(column=2, row=7, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Local da Assinatura:").grid(column=1, row=8, sticky=tk.W)
    local_entry = ttk.Entry(mainframe, width=40)
    local_entry.grid(column=2, row=8, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Dia:").grid(column=1, row=9, sticky=tk.W)
    dia_entry = ttk.Entry(mainframe, width=40)
    dia_entry.grid(column=2, row=9, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Mês:").grid(column=1, row=10, sticky=tk.W)
    mes_entry = ttk.Entry(mainframe, width=40)
    mes_entry.grid(column=2, row=10, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Ano:").grid(column=1, row=11, sticky=tk.W)
    ano_entry = ttk.Entry(mainframe, width=40)
    ano_entry.grid(column=2, row=11, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor Mensal das Parcelas:").grid(column=1, row=12, sticky=tk.W)
    valor_mensal_entry = ttk.Entry(mainframe, width=40)
    valor_mensal_entry.grid(column=2, row=12, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor Total:").grid(column=1, row=13, sticky=tk.W)
    valor_total_entry = ttk.Entry(mainframe, width=40)
    valor_total_entry.grid(column=2, row=13, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor na assinatura deste:").grid(column=1, row=14, sticky=tk.W)
    valor_assinatura_entry = ttk.Entry(mainframe, width=40)
    valor_assinatura_entry.grid(column=2, row=14, sticky=(tk.W, tk.E))

    choice_var = tk.IntVar()
    ttk.Radiobutton(mainframe, text="Procuração", variable=choice_var, value=1).grid(column=1, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Declaração de Hipossuficiência", variable=choice_var, value=2).grid(column=2, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Contrato", variable=choice_var, value=4).grid(column=3, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Todos", variable=choice_var, value=3).grid(column=1, row=16, sticky=tk.W)

    create_button = ttk.Button(mainframe, text="Criar Documento", command=lambda: create_document(nome_doc_entry, nome_entry, std_civil_entry, ocupacao_entry, cpf_entry, rg_entry, endereco_entry, local_entry, dia_entry, mes_entry, ano_entry, choice_var, valor_mensal_entry, valor_total_entry, valor_assinatura_entry))
    create_button.grid(column=2, row=16, sticky=(tk.W, tk.E))

    def convert_to_pdf_wrapper():
        docx_path = fr'{local_dir}/proc_{nome_doc_entry.get()}.docx'  # alterar caminho do arquivo, dependendo do dispositivo
        if choice_var.get() == 1:
            docx_path = fr'{local_dir}/proc_{nome_doc_entry.get()}.docx'  # alterar caminho do arquivo, dependendo do dispositivo
        elif choice_var.get() == 2:
            docx_path = fr'{local_dir}/hipo_{nome_doc_entry.get()}.docx'  # alterar caminho do arquivo, dependendo do dispositivo
        elif choice_var.get() == 3:
            proc_path = fr'{local_dir}\proc_{nome_doc_entry.get()}.docx'
            hipo_path = fr'{local_dir}\hipo_{nome_doc_entry.get()}.docx'
            contrato_path = fr'{local_dir}\contrato_{nome_doc_entry.get()}.docx'
            convert_to_pdf(proc_path)
            convert_to_pdf(hipo_path)
            convert_to_pdf(contrato_path)
            return
        elif choice_var.get() == 4:
            docx_path = fr'{local_dir}/contrato_{nome_doc_entry.get()}.docx'
            return
        convert_to_pdf(docx_path)

    convert_button = ttk.Button(mainframe, text="Converter para PDF", command=convert_to_pdf_wrapper)
    convert_button.grid(column=2, row=17, sticky=(tk.W, tk.E))

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

create_gui()
