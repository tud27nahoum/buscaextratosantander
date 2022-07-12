import time
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import filedialog
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from pathlib import Path
import babel.numbers


# configurando navegador

chr_options = Options()
chr_options.add_experimental_option("detach", True)

# abrindo navegador no banco

navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chr_options)
navegador.get("https://www.santander.com.br/")


class Extrato:
    def __init__(self):
        self.data_ini = None
        self.data_fi = None
        self.descricao = None
        self.valor = None

    def baixar_extrato(self):

        # buscando as datas no tcalendar
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()

        data_ini = datetime.strptime(f'{data_inicial}', "%d/%m/%Y").date()
        data_fi = datetime.strptime(f'{data_final}', "%d/%m/%Y").date()

        # Calculando a diferença entre a data inicial e o dia de hoje
        hoje = datetime.today().date()
        diferenca = (hoje - data_ini).days

        # Calculando a quantidade de dias
        quantidade = (data_fi - data_ini).days

        # baixando o extrato apenas se o número de dias for menor ou igual a 90 e se a data inicial for de 90 dias até o dia de hoje.
        if quantidade <= 90 and diferenca <= 90:
            navegador.find_element(By.CLASS_NAME, 'form-control').send_keys('extrato conta corrente', Keys.ENTER)
            navegador.find_element(By.ID, 'operativa_seleccionada').click()
            navegador.find_element(By.ID, 'daterange').click()

            # preenchendo a datga inicial
            navegador.find_element(By.NAME, 'daterangepicker_start').click()
            navegador.find_element(By.NAME, 'daterangepicker_start').send_keys(data_inicial)

            # preenchendo a data final
            navegador.find_element(By.NAME, 'daterangepicker_end').click()
            navegador.find_element(By.NAME, 'daterangepicker_end').send_keys(data_final)

            navegador.find_element(By.CLASS_NAME, 'btn-success').click()

            time.sleep(10)
            navegador.find_element(By.CLASS_NAME, 'icon-excel').click()
            time.sleep(5)

        else:
            label_extrato[
                'text'] = "O extrato deve conter no máximo 90 dias e referente aos últimos 3 meses! Favor alterar as datas."

        navegador.quit()

    def pesquisar_arquivo(self):

        # Abrindo a tabela excel de extrato baixada no banco
        global extrato_df
        caminho = Path('C:/Users/Eric Nahoum/Downloads/')
        extrato_df = pd.read_excel(f'{caminho}/planilhaExtrato.xls', skiprows=5, usecols=[0, 1, 4, 5, 6], header=[0])
        extrato_df = extrato_df.dropna(thresh=4)

        # Atribuindo as variáveis os campos de busca
        descricao = input_descr.get().upper()
        valor = input_valor.get()

        # Atribuindo zero a campos nulos
        extrato_df.loc[extrato_df['Crédito (R$) '].isnull(), 'Crédito (R$) '] = '0,0'
        extrato_df.loc[extrato_df['Débito (R$) '].isnull(), 'Débito (R$) '] = '0,0'
        extrato_df = extrato_df.iloc[1:, :]
        # Apagando linhas com valores nulos
        extrato_df = extrato_df.dropna(how="any", axis=0)

        # Se campo de busca de valor e descrição estiverem preenchidos
        if valor and descricao:
            # pesquisando nos campos da tabela
            busca_df = extrato_df[extrato_df["Débito (R$) "].str.contains(valor)]
            busca1_df = extrato_df[extrato_df["Crédito (R$) "].str.contains(valor)]
            busca2_df = extrato_df[extrato_df['Descrição '].str.contains(descricao)]
            # somando as tabelas
            frames = [busca_df, busca1_df, busca2_df]
            extrato_df = pd.concat(frames)
            # Eliminando linas duplicadas
            extrato_df = extrato_df.drop_duplicates()
            label_resultado_busca['text'] = 'Busca realizada com sucesso!'

        # Se campo de busca de valor estiver preenchido
        elif valor:
            # pesquisando nos campos da tabela
            busca_df = extrato_df[extrato_df["Débito (R$) "].str.contains(valor)]
            busca1_df = extrato_df[extrato_df["Crédito (R$) "].str.contains(valor)]
            frames = [busca_df, busca1_df]
            extrato_df = pd.concat(frames)
            label_resultado_busca['text'] = 'Busca realizada com sucesso!'

        # Se campo de busca de descrição estiver preenchido
        elif descricao:
            # pesquisando nos campos da tabela
            extrato_df = extrato_df[extrato_df['Descrição '].str.contains(descricao)]
            label_resultado_busca['text'] = 'Busca realizada com sucesso!'

        # Se nenhum campo de busca de valor e descrição estiver preenchido
        else:
            label_resultado_busca['text'] = 'Digite ao menos uma descrição ou valor para busca no extrato!'

        # Transformando os campos em numéricos
        extrato_df["Débito (R$) "] = extrato_df["Débito (R$) "].str.replace(',', '.')
        extrato_df["Crédito (R$) "] = extrato_df["Crédito (R$) "].str.replace(',', '.')
        # extrato_df["Data "] = pd.to_datetime(extrato_df["Data "], errors="coerce")
        extrato_df["Crédito (R$) "] = pd.to_numeric(extrato_df["Crédito (R$) "], errors="coerce")
        extrato_df["Débito (R$) "] = pd.to_numeric(extrato_df["Débito (R$) "], errors="coerce")

    def salvar_busca(self):

        # Salvar planilha "extrato_df"  em "buscaextrato.xlsx"
        caminho_arquivo = filedialog.asksaveasfilename(initialfile="buscaextrato.xlsx", title = "Escolha o nome do arquivo que deseja salvar",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
        with pd.ExcelWriter(caminho_arquivo, mode="w", engine="openpyxl") as writer:
            extrato_df.to_excel(writer, index=False)



        label_exportar_busca['text'] = 'Arquivo salvo com sucesso!'


class Banco:

    def __init__(self):
        self.cpf = None
        self.senha = None
        self.codigo = None

    def entrar_conta(self):

        codigo = self.codigo.get()
        navegador.find_element(By.ID, 'qrcode-sdk').send_keys(codigo, Keys.ENTER)
        time.sleep(1)
        try:
            elemento = WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.ID, "botaoCiente")))
            elemento.click()
        except:
            pass
        janelacod.destroy()

    def fazer_login(self):

        # atribuindo as variáveis cpf e senha os campos do formulário
        cpf = input_cpf.get()
        senha = input_senha.get()

        # Se CPF e Senha preenchidos
        if cpf and senha:
            label_fazer_login['text'] = "Login efetuado com sucesso! Favor digitar o código gerado pelo ID Santander"
        else:
            label_fazer_login['text'] = "Favor preencher CPF e senha corretamente!"

        # preenchendo cpf e senha no navegador
        navegador.find_element(By.ID, 'loginIB').send_keys(cpf, Keys.ENTER)
        elemento = WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.ID, "senha")))
        time.sleep(1)
        elemento.click()
        navegador.find_element(By.ID, 'senha').send_keys(senha, Keys.ENTER)
        time.sleep(1)

        # abrindo janela para solicitar o codigo id santander
        global janelacod
        janelacod = tk.Toplevel(janela)
        janelacod.grab_set()

        janelacod.title('Código QR code ID Santander')
        janelacod.rowconfigure([0, 1, 2, 3], weight=1)
        janelacod.columnconfigure([0, 1, 2], weight=1)

        label_cod = tk.Label(janelacod, text="")
        label_cod.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

        label_cod = tk.Label(janelacod,
                             text="Leia o código de QR code do navegador, gere este no ID Santander - Internet Banking:")
        label_cod.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

        input_cod = tk.Entry(janelacod, show='*')
        input_cod.grid(row=2, column=1, padx=10, pady=10, sticky='nse', columnspan=1)

        label_cod = tk.Label(janelacod, text="")
        label_cod.grid(row=3, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

        self.codigo = input_cod

        botao_entrar = tk.Button(janelacod, text="Entrar", command=santander.entrar_conta)
        botao_entrar.grid(row=2, column=2, padx=10, pady=10, sticky='nse', columnspan=1)

        janelacod.mainloop()


# Instanciando objetos banco e extrato
santander = Banco()
busca_extrato = Extrato()

# abrir janela principal do programa
global janela
janela = tk.Tk()

janela.title('Ferramenta de busca em extrato Santander')
janela.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], weight=1)
janela.columnconfigure([0, 1, 2], weight=1)

label_conta = tk.Label(text="Entrando em sua conta:", borderwidth=2, relief='solid')
label_conta.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_cpf = tk.Label(text="CPF:", anchor='w')
label_cpf.grid(row=1, column=0, padx=10, pady=1, sticky='nswe', columnspan=1)

input_cpf = tk.Entry()
input_cpf.grid(row=1, column=0, padx=45, pady=1, sticky='nsew', columnspan=1)

label_senha = tk.Label(text="Senha:", anchor='w')
label_senha.grid(row=1, column=1, padx=10, pady=1, sticky='nswe', columnspan=1)

input_senha = tk.Entry(show='*')
input_senha.grid(row=1, column=1, padx=55, pady=1, sticky='nsew')

botao_fazer_login = tk.Button(text="Login", command=santander.fazer_login)
botao_fazer_login.grid(row=1, column=2, padx=10, pady=1, sticky='nsew')

label_fazer_login = tk.Label(text="")
label_fazer_login.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_periodo = tk.Label(text="Selecione o período desejado - máximo de 90 dias - Apenas os últimos 3 meses", borderwidth=2, relief='solid')
label_periodo.grid(row=3, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_datainicial = tk.Label(text="Data Inicial:", anchor='w')
label_datafinal = tk.Label(text="Data Final:", anchor='w')
label_datainicial.grid(row=4, column=0, padx=10, pady=10, sticky='nsew', columnspan=1)
label_datafinal.grid(row=4, column=1, padx=10, pady=10, sticky='nsew', columnspan=1)

calendario_datainicial = DateEntry(year=2022, locale='pt_br')
calendario_datafinal = DateEntry(year=2022, locale='pt_br')
calendario_datainicial.grid(row=4, column=0, padx=85, pady=10,  sticky='nsew')
calendario_datafinal.grid(row=4, column=1, padx=85, pady=10, sticky='nsew')

botao_extrato = tk.Button(text="Extrato", command=busca_extrato.baixar_extrato)
botao_extrato.grid(row=4, column=2, padx=10, pady=10, sticky='nsew')

label_extrato = tk.Label(text='')
label_extrato.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_descr = tk.Label(text="Descrição:", anchor='w')
label_descr.grid(row=6, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

input_descr = tk.Entry()
input_descr.grid(row=6, column=0, padx=85, pady=10, sticky='nsew', columnspan=1)

label_valor = tk.Label(text="Valor:", anchor='w')
label_valor.grid(row=6, column=1, padx=10, pady=10, sticky='nswe', columnspan=1)

input_valor = tk.Entry()
input_valor.grid(row=6, column=1, padx=85, pady=10, sticky='nsew', columnspan=1)

botao_selecionararquivo = tk.Button(text="Pesquisar no extrato", command=busca_extrato.pesquisar_arquivo)
botao_selecionararquivo.grid(row=6, column=2, padx=10, pady=10, sticky='nsew')

label_resultado_busca = tk.Label(text='')
label_resultado_busca.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

botao_exportar_busca = tk.Button(text='Salvar busca em excel', command=busca_extrato.salvar_busca)
botao_exportar_busca.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

label_exportar_busca = tk.Label(text="", anchor='w')
label_exportar_busca.grid(row=8, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text='Fechar', command=janela.destroy)
botao_fechar.grid(row=8, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()