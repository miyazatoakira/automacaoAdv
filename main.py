import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import sys
import os
import docx2pdf
from num2words import num2words  # Instale via: pip install num2words

from ScriptHipo import hipo_creator
from ScriptContrato import contrato_creator
from ScriptProcuracao import proc_creator

#################################################
# Função para identificar o caminho base
# (caso esteja empacotado em PyInstaller)
#################################################
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS  # Diretório temporário do executável PyInstaller
    else:
        return os.path.dirname(os.path.abspath(__file__))

##########################################################
# Se precisar de um caminho de arquivo específico
##########################################################
def resource_path(filename):
    base_path = get_base_path()
    return os.path.join(base_path, filename)

########################################
# Converte Float → Formato BR + Extenso
########################################
def formata_valor_e_extenso(valor_float):
    """
    Recebe um float e retorna (valor_br, valor_ext).
    Exemplo: 1000.0 -> ("1.000,00", "mil")
    """
    # Arredondamos a 2 casas decimais
    valor_float = round(valor_float, 2)

    # Formato brasileiro: 1.000,00
    valor_br = f"{valor_float:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

    # Converte para texto por extenso em português
    valor_ext = num2words(valor_float, lang='pt_BR').lower()  # ex.: "mil"

    return valor_br, valor_ext

##################################################
# Tenta extrair apenas a parte numérica do input
# e converte para float, ignorando o texto.
##################################################
def parse_valor_remove_trailing_text(valor_str):
    """
    Exemplo de comportamento:
      - "1.000,00 mil" => pega "1.000,00", ignora " mil"
      - "2500 e bla"  => pega "2500", ignora " e bla"
      - "mil e quinhentos" => ERRO, pois não há dígito
    Caso não encontre dígitos, lança ValueError.
    """
    n = len(valor_str)
    i = 0

    # 1) Ignora tudo até achar o 1º dígito ou '.' ou ','
    while i < n and not (valor_str[i].isdigit() or valor_str[i] in ('.', ',')):
        i += 1
    # Se i == n, não há parte numérica alguma
    if i == n:
        raise ValueError("Nenhum valor numérico encontrado.")

    start = i
    # 2) Coletar dígitos, '.' ou ',' até achar algo que não seja
    while i < n and (valor_str[i].isdigit() or valor_str[i] in ('.', ',')):
        i += 1

    # numeric_part => substring que supostamente é o valor
    numeric_part = valor_str[start:i].strip()
    # Se ficou vazio
    if not numeric_part:
        raise ValueError("Nenhum dígito encontrado.")

    # 3) Converte p/ float:
    #    - Remove '.' => ''
    #    - Troca ',' => '.'
    temp = numeric_part.replace('.', '').replace(',', '.')
    valor_float = float(temp)  # Se falhar, gera ValueError

    return valor_float

##############################################
# Função principal p/ criar documento(s)
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

    # Textos do Contrato (opcionais p/ Procuração/Hipo)
    vm_str = valor_mensal_entry.get()
    vt_str = valor_total_entry.get()
    va_str = valor_assinatura_entry.get()

    # Verificação de preenchimento mínimo
    if not (nome_doc and nome and std_civil and ocupacao and rg and endereco
            and local and dia and mes and ano and choice):
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
        return

    # Se for Contrato (3) ou Todos (4), exigir valores
    if choice in (3, 4):
        if not vm_str or not vt_str or not va_str:
            messagebox.showerror("Erro", "Preencha os 3 valores para o contrato!")
            return

    # Validação CPF
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

    # Se for criar Contrato ou Todos, vamos parsear e formatar
    # todos os 3 valores
    if choice in (3, 4):
        try:
            vm_float = parse_valor_remove_trailing_text(vm_str)
            vt_float = parse_valor_remove_trailing_text(vt_str)
            va_float = parse_valor_remove_trailing_text(va_str)
        except ValueError as e:
            messagebox.showerror(
                "Erro no Valor",
                f"Não foi possível interpretar valor numérico:\n{str(e)}"
            )
            return

        # Formata e sobrescreve no Entry
        vm_br, vm_ext = formata_valor_e_extenso(vm_float)
        vm_final = f"{vm_br} ({vm_ext} reais)"
        valor_mensal_entry.delete(0, tk.END)
        valor_mensal_entry.insert(0, vm_final)

        vt_br, vt_ext = formata_valor_e_extenso(vt_float)
        vt_final = f"{vt_br} ({vt_ext} reais)"
        valor_total_entry.delete(0, tk.END)
        valor_total_entry.insert(0, vt_final)

        va_br, va_ext = formata_valor_e_extenso(va_float)
        va_final = f"{va_br} ({va_ext} reais)"
        valor_assinatura_entry.delete(0, tk.END)
        valor_assinatura_entry.insert(0, va_final)

        # Agora, passamos essas strings formatadas para o script
        vm_for_script = vm_final
        vt_for_script = vt_final
        va_for_script = va_final
    else:
        # Não é Contrato => não precisamos parsear valores
        vm_for_script = None
        vt_for_script = None
        va_for_script = None

    # Chama os scripts conforme a escolha
    if choice == 1:
        # Procuracao
        proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
    elif choice == 2:
        # Hipossuficiencia
        hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
    elif choice == 3:
        # Contrato
        contrato_creator(
            nome_doc, nome, std_civil, ocupacao, cpf,
            rg, endereco, local, dia, mes, ano,
            vm_for_script, vt_for_script, va_for_script
        )
    elif choice == 4:
        # Todos
        proc_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        hipo_creator(nome_doc, nome, std_civil, ocupacao, cpf, rg, endereco, local, dia, mes, ano)
        contrato_creator(
            nome_doc, nome, std_civil, ocupacao, cpf,
            rg, endereco, local, dia, mes, ano,
            vm_for_script, vt_for_script, va_for_script
        )
    else:
        messagebox.showerror("Erro", "Escolha inválida!")
        return

    messagebox.showinfo("Sucesso", "Documento(s) criado(s) com sucesso!")

########################################
# Converter DOCX em PDF nas subpastas
########################################
def convert_to_pdf(docx_path):
    try:
        docx2pdf.convert(docx_path)
        messagebox.showinfo("Sucesso", f"{docx_path} convertido para PDF com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter para PDF: {str(e)}")

##################################################
# Para converter cada escolha em PDF
# usando subpastas: "Procuracoes", "Hipo", "Contratos"
##################################################
def convert_to_pdf_wrapper(choice_var, nome_doc_entry):
    base_dir = get_base_path()
    nome_doc = nome_doc_entry.get()
    choice = choice_var.get()

    if choice == 1:
        # Procuracao => subpasta Procuracoes
        docx_path = os.path.join(base_dir, "Procuracoes", f"procuracao_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 2:
        # Hipo => subpasta Hipo
        docx_path = os.path.join(base_dir, "Hipo", f"hipo_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 3:
        # Contrato => subpasta Contratos
        docx_path = os.path.join(base_dir, "Contratos", f"contrato_{nome_doc}.docx")
        convert_to_pdf(docx_path)
    elif choice == 4:
        # Todos => Procuracao + Hipo + Contrato
        proc_path = os.path.join(base_dir, "Procuracoes", f"procuracao_{nome_doc}.docx")
        hipo_path = os.path.join(base_dir, "Hipo", f"hipo_{nome_doc}.docx")
        contrato_path = os.path.join(base_dir, "Contratos", f"contrato_{nome_doc}.docx")
        convert_to_pdf(proc_path)
        convert_to_pdf(hipo_path)
        convert_to_pdf(contrato_path)
    else:
        messagebox.showerror("Erro", "Escolha inválida para conversão em PDF!")

#############################
# Interface Principal (GUI)
#############################
def create_gui():
    root = tk.Tk()
    root.title("Criação de Documentos")

    mainframe = ttk.Frame(root, padding="20")
    mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)

    # ============= Formulário =============
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

    # Valores para Contrato (ou "Todos")
    ttk.Label(mainframe, text="Valor Mensal das Parcelas:").grid(column=1, row=12, sticky=tk.W)
    valor_mensal_entry = ttk.Entry(mainframe, width=40)
    valor_mensal_entry.grid(column=2, row=12, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor Total:").grid(column=1, row=13, sticky=tk.W)
    valor_total_entry = ttk.Entry(mainframe, width=40)
    valor_total_entry.grid(column=2, row=13, sticky=(tk.W, tk.E))

    ttk.Label(mainframe, text="Valor na assinatura deste:").grid(column=1, row=14, sticky=tk.W)
    valor_assinatura_entry = ttk.Entry(mainframe, width=40)
    valor_assinatura_entry.grid(column=2, row=14, sticky=(tk.W, tk.E))

    # ============= Radiobuttons Escolha =============
    choice_var = tk.IntVar()
    ttk.Radiobutton(mainframe, text="Procuração", variable=choice_var, value=1).grid(column=1, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Declaração de Hipossuficiência", variable=choice_var, value=2).grid(column=2, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Contrato", variable=choice_var, value=3).grid(column=3, row=15, sticky=tk.W)
    ttk.Radiobutton(mainframe, text="Todos", variable=choice_var, value=4).grid(column=1, row=16, sticky=tk.W)

    # ============= Botões =============
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

    # Padding em todos widgets
    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
