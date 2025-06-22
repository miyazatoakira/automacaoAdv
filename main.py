import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import sys
import os
import re
import docx2pdf

# Se utilizar a biblioteca num2words (instalar via: pip install num2words)
from num2words import num2words

from ScriptHipo import hipo_creator
from ScriptContrato import contrato_creator
from ScriptProcuracao import proc_creator
from rg_processor import extract_rg_data

#################################################
# 1) Função para identificar o caminho base
#    (caso esteja empacotado em PyInstaller)
#################################################
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS  # Diretório temporário do executável PyInstaller
    else:
        return os.path.dirname(os.path.abspath(__file__))

############################################################
# 2) (Opcional) Função para obter o caminho de um arquivo
############################################################
def resource_path(filename):
    base_path = get_base_path()
    return os.path.join(base_path, filename)

########################################
# 3) Converte Float → Formato BR + Extenso
########################################
def formata_valor_e_extenso(valor_float):
    """
    Recebe um float e retorna (valor_br, valor_extenso).
    Exemplo: 1000.0 -> ("1.000,00", "mil")
    """
    valor_float = round(valor_float, 2)
    # Formato brasileiro: 1.000,00
    valor_br = f"{valor_float:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    # Converte para texto por extenso (ex.: "mil")
    valor_ext = num2words(valor_float, lang='pt_BR').lower()
    return valor_br, valor_ext

##################################################
# 4) parse_valor: extrai a parte numérica ou
#    coloca "XXXXXXXXXX" se não houver dígitos
##################################################
def parse_valor(valor_str, entry_widget):
    """
    - Se achar parte numérica, reescreve o campo só com essa parte formatada (ex.: "1.000,00").
    - Se NÃO achar nenhum dígito, substitui o campo por "XXXXXXXXXX" e retorna None.
    Retorna:
      - float, se encontrou parte numérica
      - None, se não encontrou dígitos
    """
    # Regex para achar primeira sequência de dígitos, pontos ou vírgulas
    match = re.search(r"[\d\.,]+", valor_str)
    if not match:
        # Não encontrou parte numérica => Placeholder
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, "XXXXXXXXXX")
        return None

    numeric_part = match.group(0)  # ex. "1.000,00"
    # Remove pontos, troca vírgula por ponto
    temp = numeric_part.replace('.', '').replace(',', '.')

    try:
        valor_float = float(temp)
    except ValueError:
        # Se a conversão falhar, é muito improvável, mas
        # caso digitar algo "1..000" => cairia aqui. Então definimos placeholder
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, "XXXXXXXXXX")
        return None

    # Reescreve o campo com o valor formatado (ex.: "1.000,00")
    valor_float = round(valor_float, 2)
    valor_br = f"{valor_float:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, valor_br)

    return valor_float

##############################################
# 5) Função principal para criar o documento(s)
##############################################
def create_document(nome_doc_entry,
                    nome_entry,
                    std_civil_entry,
                    ocupacao_entry,
                    cpf_entry,
                    rg_entry,
                    endereco_entry,
                    local_entry,
                    dia_entry,
                    mes_entry,
                    ano_entry,
                    choice_var,
                    valor_mensal_entry,
                    valor_total_entry,
                    valor_assinatura_entry):
    """
    Lê dados do usuário, valida e chama os scripts correspondentes.
    """
    # Campos de texto
    nome_doc = nome_doc_entry.get()
    nome = nome_entry.get()
    std_civil = std_civil_entry.get()
    ocupacao = ocupacao_entry.get()
    cpf = cpf_entry.get()
    rg = rg_entry.get()
    endereco = endereco_entry.get()
    local = local_entry.get()
    dia = dia_entry.get()
    mes = mes_entry.get()
    ano = ano_entry.get()
    choice = choice_var.get()

    vm_str = valor_mensal_entry.get()
    vt_str = valor_total_entry.get()
    va_str = valor_assinatura_entry.get()

    # Validação de campos básicos
    if not (nome_doc and nome and std_civil and ocupacao and rg and endereco
            and local and dia and mes and ano and choice):
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
        return

    # Se for Contrato (3) ou Todos (4), precisa obrigatoriamente de 3 valores
    if choice in (3, 4):
        if not vm_str or not vt_str or not va_str:
            messagebox.showerror("Erro", "Preencha todos os valores do contrato!")
            return

    # Validação do CPF
    def valida_cpf(cpf_):
        cpf_ = ''.join(filter(str.isdigit, cpf_))
        if len(cpf_) != 11:
            return False
        if cpf_ == cpf_[0] * len(cpf_):
            return False
        soma = sum(int(cpf_[i]) * (10 - i) for i in range(9))
        primeiro_digito = (soma * 10 % 11) % 10
        soma = sum(int(cpf_[i]) * (11 - i) for i in range(10))
        segundo_digito = (soma * 10 % 11) % 10
        return cpf_[-2:] == f'{primeiro_digito}{segundo_digito}'

    if not valida_cpf(cpf):
        messagebox.showerror("Erro", "CPF Inválido!")
        return

    # Se for criar Contrato ou Todos: parse dos 3 valores
    if choice in (3, 4):
        vm_float = parse_valor(vm_str, valor_mensal_entry)
        vt_float = parse_valor(vt_str, valor_total_entry)
        va_float = parse_valor(va_str, valor_assinatura_entry)

        # Se não achou parte numérica em algum, vm_float/ vt_float/ va_float será None
        # Então passamos "XXXXXXXXXX" no doc
        if vm_float is not None:
            vm_br, vm_ext = formata_valor_e_extenso(vm_float)
            valor_mensal_final = f"{vm_br} ({vm_ext} reais)"
        else:
            valor_mensal_final = "XXXXXXXXXX"

        if vt_float is not None:
            vt_br, vt_ext = formata_valor_e_extenso(vt_float)
            valor_total_final = f"{vt_br} ({vt_ext} reais)"
        else:
            valor_total_final = "XXXXXXXXXX"

        if va_float is not None:
            va_br, va_ext = formata_valor_e_extenso(va_float)
            valor_assinatura_final = f"{va_br} ({va_ext} reais)"
        else:
            valor_assinatura_final = "XXXXXXXXXX"
    else:
        # Se não for gerar contrato, não precisamos de valores
        valor_mensal_final = None
        valor_total_final = None
        valor_assinatura_final = None

    # Chama os scripts de acordo com a escolha
    if choice == 1:
        proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
    elif choice == 2:
        hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
    elif choice == 3:
        contrato_creator(
            nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco,
            local, dia, mes, ano,
            valor_mensal_final, valor_total_final, valor_assinatura_final
        )
    elif choice == 4:
        proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        contrato_creator(
            nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco,
            local, dia, mes, ano,
            valor_mensal_final, valor_total_final, valor_assinatura_final
        )
    else:
        messagebox.showerror("Erro", "Escolha inválida!")
        return

    messagebox.showinfo("Sucesso", "Documento(s) criado(s) com sucesso!")

########################################
# 6) Converter DOCX em PDF (subpastas)
########################################
def convert_to_pdf(docx_path):
    try:
        docx2pdf.convert(docx_path)
        messagebox.showinfo("Sucesso", f"{docx_path} convertido para PDF com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter para PDF: {str(e)}")

##################################################
# 7) Wrapper p/ converter cada escolha em PDF
#    usando as subpastas Procuracoes, Hipo, Contratos
##################################################
def convert_to_pdf_wrapper(choice_var, nome_doc_entry):
    base_dir = get_base_path()
    nome_doc = nome_doc_entry.get()
    choice = choice_var.get()

    if choice == 1:
        docx_path = os.path.join(base_dir, "Procuracoes", f"procuracao_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 2:
        docx_path = os.path.join(base_dir, "Hipo", f"hipo_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 3:
        docx_path = os.path.join(base_dir, "Contratos", f"contrato_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 4:
        proc_path = os.path.join(base_dir, "Procuracoes", f"procuracao_{nome_doc}.docx")
        hipo_path = os.path.join(base_dir, "Hipo", f"hipo_{nome_doc}.docx")
        contrato_path = os.path.join(base_dir, "Contratos", f"contrato_{nome_doc}.docx")
        convert_to_pdf(proc_path)
        convert_to_pdf(hipo_path)
        convert_to_pdf(contrato_path)
    else:
        messagebox.showerror("Erro", "Escolha inválida para conversão em PDF!")

#####################################
# 7.1) Carregar dados a partir do RG
#####################################
def load_rg_file(nome_entry, cpf_entry, rg_entry):
    file_path = filedialog.askopenfilename(
        title="Selecione imagem ou PDF do RG",
        filetypes=[
            ("Imagens", "*.png *.jpg *.jpeg"),
            ("PDF", "*.pdf"),
            ("Todos", "*.png *.jpg *.jpeg *.pdf"),
        ],
    )
    if not file_path:
        return
    try:
        nome, cpf, rg = extract_rg_data(file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao extrair dados do RG: {e}")
        return

    if nome:
        nome_entry.delete(0, tk.END)
        nome_entry.insert(0, nome)
    if cpf:
        cpf_entry.delete(0, tk.END)
        cpf_entry.insert(0, cpf)
    if rg:
        rg_entry.delete(0, tk.END)
        rg_entry.insert(0, rg)

#############################
# 8) Interface Principal (GUI)
#############################
def create_gui():
    root = tk.Tk()
    root.title("Criação de Documentos")

    mainframe = ttk.Frame(root, padding="20")
    mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)

    # ====== Campos do Formulário ======
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
    ttk.Button(
        mainframe,
        text="Carregar RG",
        command=lambda: load_rg_file(nome_entry, cpf_entry, rg_entry),
    ).grid(column=3, row=6, sticky=tk.W)

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

    # ====== Valores Contrato (ou Todos) ======
    ttk.Label(mainframe, text="Valor Mensal das Parcelas:").grid(column=1, row=12, sticky=tk.W)
    valor_mensal_entry = ttk.Entry(mainframe, width=40)
    valor_mensal_entry.grid(column=2, row=12, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor Total:").grid(column=1, row=13, sticky=tk.W)
    valor_total_entry = ttk.Entry(mainframe, width=40)
    valor_total_entry.grid(column=2, row=13, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor na assinatura deste:").grid(column=1, row=14, sticky=tk.W)
    valor_assinatura_entry = ttk.Entry(mainframe, width=40)
    valor_assinatura_entry.grid(column=2, row=14, sticky=(tk.W, tk.E))

    # ====== Radiobuttons ======
    choice_var = tk.IntVar()
    ttk.Radiobutton(mainframe, text="Procuração", variable=choice_var, value=1).grid(column=1, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Declaração de Hipossuficiência", variable=choice_var, value=2).grid(column=2, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Contrato", variable=choice_var, value=3).grid(column=3, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Todos", variable=choice_var, value=4).grid(column=1, row=16, sticky=tk.W)

    # ====== Botões ======
    create_button = ttk.Button(
        mainframe,
        text="Criar Documento",
        command=lambda: create_document(
            nome_doc_entry,
            nome_entry,
            std_civil_entry,
            ocupacao_entry,
            cpf_entry,
            rg_entry,
            endereco_entry,
            local_entry,
            dia_entry,
            mes_entry,
            ano_entry,
            choice_var,
            valor_mensal_entry,
            valor_total_entry,
            valor_assinatura_entry
        )
    )
    create_button.grid(column=2, row=16, sticky=(tk.W, tk.E))

    convert_button = ttk.Button(
        mainframe,
        text="Converter para PDF",
        command=lambda: convert_to_pdf_wrapper(choice_var, nome_doc_entry)
    )
    convert_button.grid(column=2, row=17, sticky=(tk.W, tk.E))

    # Padding
    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
	create_gui()
