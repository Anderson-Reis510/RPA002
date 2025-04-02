import pyautogui
import subprocess
import time
import pyperclip 
import sys
import logging
import pygetwindow as gw
from datetime import datetime, date
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, simpledialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import fitz
import shutil
import threading
from PIL import Image
import os
import pytesseract
import re
import pyodbc
import threading
from collections import defaultdict

pasta_pedidos = "C:/2 - Boletos"

def iniciar_robo():
    # Tempo
    time_busca = 0.2
    time_login = 20
    time_short = 1
    time_medium = 3
    time_maximizar = 10
    time_long = 8

    senha_topmanager = "123456"
    tempo_espera_max = 30

    coordenadas = {
        "Movimentos": (147, 34),
        "Vendas": (166, 171),
        "Titulo Venda": (371, 486),
        #Maximiza aba
        "Filtro": (211, 78),
        "Empresa":(196, 93),
        "Agent. Cobrador": (744, 91),
        "Rest. Titulo": (944, 92),
        "Periodo Emi": (1410, 93),
        # F5
        "N. Pedido Base": (40, 114),
        "Pos. Acrobat": (104, 38),
        "Fechar janela navegador": (1900, 19),
        "Fechar Janela": (1892, 13),
        }

    pedido_x, pedido_y = coordenadas["N. Pedido Base"]
    incremento_y = 16

    def contar_pedidos():
        total_pedidos = 0
        pedidos_encontrados = []  
        temp_y = pedido_y
        last_pedido = None  

        while True:
            pyautogui.moveTo(pedido_x, temp_y)
            pyautogui.doubleClick()
            time.sleep(time_short)

            pyautogui.hotkey("ctrl", "c")
            time.sleep(time_short)
            pedido_atual = pyperclip.paste().strip()

            if pedido_atual:
                if pedido_atual == last_pedido:
                    # Se o pedido atual for o mesmo que o último, significa que chegamos ao fim
                    break
                pedidos_encontrados.append(pedido_atual)  # Armazena o pedido
                last_pedido = pedido_atual  # Atualiza o último pedido
                total_pedidos += 1  # Conta um novo pedido
            else:
                # Se a célula estiver vazia, interrompe
                break

            temp_y += incremento_y  # Move para a próxima linha
            pyautogui.press("down")
            time.sleep(time_short)

        return total_pedidos, pedidos_encontrados

    def baixar_boleto(pedido_x, pedido_y):
        try:
            time.sleep(time_short)
            pyautogui.click(button="right") 
            time.sleep(time_short)

            for _ in range(6):
                pyautogui.press("down")
                time.sleep(0.2)  

            pyautogui.press("enter")  
            time.sleep(time_long)

            pyautogui.moveTo(coordenadas['Pos. Acrobat'])
            pyautogui.click()
            time.sleep(time_long)
            
            pyautogui.hotkey("win", "up")
            time.sleep(time_medium)
            
            pyautogui.hotkey("ctrl", "shift", "s")
            time.sleep(time_medium)

            # Salvar PDF
            pyautogui.moveTo(1297, 797)  
            pyautogui.click()
            time.sleep(time_medium)
            pyautogui.moveTo(571, 48)
            pyautogui.click()
            time.sleep(time_medium)
            preencher_campo("C:/2 - Boletos")
            time.sleep(time_medium)
            
            pyautogui.press("enter")
            time.sleep(time_short)
            pyautogui.moveTo(781, 508)
            pyautogui.click()
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Fechar janela navegador'])
            pyautogui.click()
            time.sleep(time_medium)
            pyautogui.moveTo(coordenadas["Fechar Janela"])
            pyautogui.click()
            time.sleep(time_medium)
            logging.info("PDF baixado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao baixar PDF: {e}")
            
    def preencher_campo(texto):
        pyautogui.press("backspace")
        pyautogui.write(texto)
        
    #Função para maximizar uma janela específica ou clicar no botão de maximizar
    def maximizar_janela(titulo_parcial):
        try:
            time.sleep(time_maximizar)
            janelas = gw.getWindowsWithTitle(titulo_parcial)
            if janelas:
                janela = janelas[0]
                if not janela.isMaximized:
                    janela.maximize()
                logging.info(f"Janela '{titulo_parcial}' maximizada.")
            time.sleep(time_short)
        except Exception as e:
            logging.error(f"Erro ao maximizar a janela {titulo_parcial}: {e}")
        
    #Entrar e logar no TopManager
    try:
        if not gw.getWindowsWithTitle("TopManager"):
            pyautogui.hotkey("ctrl", "shift", "t")
            time.sleep(5)
            logging.info("Atalho enviado para abrir o TopManager.")

            # Espera até o TopManager abrir ou o tempo máximo ser atingido
            tempo_inicial = time.time()
            while not gw.getWindowsWithTitle("TopManager"):
                if time.time() - tempo_inicial > tempo_espera_max:
                    logging.error("O TopManager não abriu dentro do tempo esperado.")
                    raise Exception("Tempo limite excedido ao abrir o TopManager.")
                time.sleep(1)  
        logging.info("TopManager detectado. Tentando login.")

        # Login no sistema
        time.sleep(5)
        pyautogui.moveTo(398, 346)
        time.sleep(time_medium)
        pyautogui.click()
        pyautogui.write(senha_topmanager)
        pyautogui.press("enter")
        time.sleep(time_login)

        if gw.getWindowsWithTitle("TopManager (Licenciado para Stik)"):
            logging.info("Login no TopManager realizado com sucesso.")
            maximizar_janela("TopManager")
            
        else:
            raise Exception("Falha no login no TopManager.")
            
        #Acessar a aba dos Documentos Eletrônicos
        try:
            pyautogui.moveTo(coordenadas['Movimentos'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas["Vendas"])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Titulo Venda'])
            pyautogui.click()
            time.sleep(5)
            pyautogui.moveTo(731, 84)
            pyautogui.click()
            time.sleep(time_short)
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Filtro'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Empresa'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("2")
            time.sleep(time_short)
            pyautogui.moveTo(coordenadas['Agent. Cobrador'])
            time.sleep(time_short)
            pyautogui.click()
            
            preencher_campo("53")
            time.sleep(time_short)
            pyautogui.press("enter")   
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Rest. Titulo'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("em")
            time.sleep(time_short)
            pyautogui.press("enter")
            time.sleep(time_medium)
            pyautogui.moveTo(coordenadas['Periodo Emi'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("h")
            time.sleep(time_short)
            pyautogui.press("enter")
            time.sleep(time_short)
            """pyautogui.press("tab")
            preencher_campo("07/02/2025")
            time.sleep(time_medium)
            pyautogui.press("tab")
            preencher_campo("07/02/2025")
            time.sleep(time_medium)
            pyautogui.press("enter")
            time.sleep(time_medium)"""
            
            pyautogui.press('F5')
            time.sleep(time_long)
            logging.info("Acesso à aba de Titulo de Vendas.")
            
        except Exception as e:
            logging.error(f"Erro ao acessar aba de Titulo de Vendas: {e}")
            raise Exception("Falha ao acessar a aba de Títutlos de Vendas")
            
           
        def baixar_boletos(pedidos_encontrados):
            temp_y = pedido_y  
            for pedido in pedidos_encontrados:
                pyautogui.moveTo(pedido_x, temp_y)
                pyautogui.doubleClick()
                time.sleep(time_busca)

                baixar_boleto(pedido_x, temp_y)  

                temp_y += incremento_y  

        logging.info("Contando quantidade de pedidos disponíveis...")
        numero_pedidos, pedidos_encontrados = contar_pedidos()

        if numero_pedidos == 0:
            logging.info("Nenhum pedido disponível para baixar.")
        else:
            logging.info(f"Total de pedidos encontrados: {numero_pedidos}")

            # Iniciar o processo de download dos boletos
            logging.info("Iniciando download dos boletos...")
            baixar_boletos(pedidos_encontrados)
        
        try:
            time.sleep(time_medium)
            pyautogui.moveTo(1898, 14)
            pyautogui.click()
            logging.info("TopManager fechado com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao fechar o TopManager: {e}")

    except Exception as e:
        logging.error(f"Erro no processo: {e}")

def Iniciar_GUI():
    pytesseract.pytesseract.tesseract_cmd = r"C://Users//fnojosa//AppData//Local//Programs//Tesseract-OCR//tesseract.exe"

    pasta_pedidos = "C:/2 - Boletos"

    def formatar_data(event):
        texto = re.sub(r'\D', '', event.widget.get())

        # Adiciona barras conforme a digitação
        if len(texto) > 2:
            texto = f"{texto[:2]}/{texto[2:]}"
        if len(texto) > 5:
            texto = f"{texto[:5]}/{texto[5:]}"

        texto = texto[:10]  

        # Valida e corrige a data
        if len(texto) == 10:
            try:
                data = datetime.strptime(texto, '%d/%m/%Y') 
            except ValueError:
                dia, mes, ano = map(int, texto.split('/'))
                try:
                    data = datetime(ano, mes, min(dia, 28))  
                    texto = data.strftime('%d/%m/%Y')
                except ValueError:
                    texto = ''  # Reseta caso ainda seja inválido

        event.widget.delete(0, tk.END)
        event.widget.insert(0, texto)

    # Função para conectar e buscar dados SQL
    def obter_dados_sql(data_inicial, data_final):
        server = '45.235.240.135'
        database = 'Stik'
        username = 'ti'
        password = 'Stik0123'
        conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"

        consulta_sql = f"""
            Select Distinct
            Titulo = Une.SgUne + '.' + Rcd.NrRcd + '/' + CONVERT(varchar, Rct.NrRctOrd)
            , Cliente = Cli.NmCli
            , CNPJCPF = Case When Pes.TpPes = 1
                                Then
                                    Convert(VarChar, Case When Pes.NrPesCpj is not Null 
                                    Then	Substring(Replace(Space(14 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 1, 2) + '.' + 
                                        Substring(Replace(Space(14 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 3, 3) + '.' +
                                        Substring(Replace(Space(14 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 6, 3) + '/' +
                                        Substring(Replace(Space(14 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 9, 4) + '-' +
                                        Substring(Replace(Space(14 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 13, 2)
                                    Else '** *** *** **** **' End)
                                Else
                                    Convert(VarChar, Case When Pes.NrPesCpj is not Null 
                                    Then	Substring(Replace(Space(11 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 1, 3) + '.' + 
                                        Substring(Replace(Space(11 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 4, 3) + '.' +
                                        Substring(Replace(Space(11 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 7, 3) + '-' +
                                        Substring(Replace(Space(11 - Len(Convert(VarChar, Pes.NrPesCpj))), ' ', '0') + Convert(VarChar, Pes.NrPesCpj), 10, 2) 
                                    Else '*** *** *** **' End)
                                End
            , Cobrador = Agc.NmAgc
            , DtEmissao = Rcm.DtRcmMov
            , Email = Mct.NrMctEnd
            From TbRcd Rcd With (nolock)
            join TbRct Rct (nolock) on Rct.CdRcd = Rcd.CdRcd
            join TbRcm Rcm (nolock) on Rcm.CdRct = Rct.CdRct and Rcm.CdTop = 49
            join TbUne Une (nolock) on Une.CdUne = Rcd.CdUne
            join TbCli Cli (nolock) on Cli.CdCli = Rcd.CdCli
            left join TbPes Pes (nolock) on Pes.CdPes = Cli.CdPes
            join TbMct Mct (nolock) on Mct.CdPes = Pes.CdPes and  Mct.CdTmc = 31
            left join TbAgc Agc (nolock) on Agc.CdAgc = Rct.CdAgc
            WHERE (Rcm.DtRcmMov BETWEEN '{data_inicial}' AND '{data_final}')
            and Agc.CdAgc = 53 
        """
        try:
            conexao = pyodbc.connect(conn_str)
            cursor = conexao.cursor()
            cursor.execute(consulta_sql)

            print(f"Data inicial convertida para SQL: {data_inicial}")
            print(f"Data final convertida para SQL: {data_final}")

            colunas = [desc[0] for desc in cursor.description]
            dados = [list(linha) for linha in cursor.fetchall()]

            # Formata a coluna DtEmissao, se existir
            if "DtEmissao" in colunas:
                indice_data = colunas.index("DtEmissao")
                for linha in dados:
                    valor_data = linha[indice_data]
                    
                    # Se vier None, não há o que formatar
                    if valor_data is None:
                        continue
                    
                    # Se for do tipo datetime ou date, formata diretamente
                    if isinstance(valor_data, (datetime, date)):
                        linha[indice_data] = valor_data.strftime("%d/%m/%Y")
                    else:
                        # Caso seja string, tenta converter
                        try:
                            dt = datetime.strptime(valor_data, "%Y-%m-%d")
                            linha[indice_data] = dt.strftime("%d/%m/%Y")
                        except Exception:
                            # Se não conseguir converter, deixa como está
                            pass

            cursor.close()
            conexao.close()

            return colunas, dados
        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Erro ao buscar dados: {e}")
            return [], []

    # Função que extrai o número do documento do PDF
    def extrair_numero_documento(pdf):
        try:
            doc = fitz.open(pdf)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            texto = pytesseract.image_to_string(img)
            doc.close()

            match = re.search(r"Nr\.\s*do\s*Documento.*?(\d{7,10})", texto, re.IGNORECASE | re.DOTALL)
            if match:
                num_str = match.group(1)
                # Verifica se num_str contém apenas dígitos
                if num_str.isdigit():
                    return int(num_str)
                else:
                    return None
            else:
                return None
        except Exception as e:
            print(f"Erro ao extrair número do documento: {e}")
            return None
        
    # Função para renomear os boletos na pasta
    def renomear_boletos_pasta(pasta):
        documentos_parcelas = {}
        arquivos_nao_renomeados = [arquivo for arquivo in os.listdir(pasta) if arquivo.lower().endswith('.pdf') and '-' not in arquivo]

        if not arquivos_nao_renomeados:
            print("Nenhum pedido não renomeado encontrado.")
            return

        for arquivo in arquivos_nao_renomeados:
            caminho_pdf = os.path.join(pasta, arquivo)
            try:
                doc = fitz.open(caminho_pdf)
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                texto = pytesseract.image_to_string(img)
                print(f"Texto extraído de {arquivo}:\n{texto}\n")
                doc.close()

                match = re.search(r"Nr\.\s*do\s*Documento.*?(\d{7,10})", texto, re.IGNORECASE | re.DOTALL)

                if match:
                    document_number = match.group(1)
                    base_number = document_number[:-1]
                    parcela = document_number[-1]  # Usa o último dígito como parcela

                    novo_nome = f"Stik.{base_number}-{parcela}"
                    novo_caminho = os.path.join(pasta, novo_nome + ".pdf") 
                    
                    if os.path.exists(novo_caminho):
                        print(f"Arquivo {novo_nome} já existe. Pulando a renomeação de {arquivo}.")
                    else:
                        os.rename(caminho_pdf, novo_caminho)
                        print(f"Renomeado {arquivo} para {novo_nome}.pdf")
                else:
                    print(f"Número de documento não encontrado em {arquivo}")
                    
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

    def agrupar_boletos_por_documento(pasta_boletos):
        arquivos_por_documento = defaultdict(list)  

        arquivos_pasta = [arquivo for arquivo in os.listdir(pasta_boletos) if arquivo.endswith(".pdf")]
        print("Arquivos encontrados:", arquivos_pasta)

        for arquivo in arquivos_pasta:
            try:
                numero_documento = arquivo.split('/')[0].replace("Stik", "")  
                arquivos_por_documento[numero_documento].append(os.path.join(pasta_boletos, arquivo))
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

        return arquivos_por_documento

    arquivos_por_documento = agrupar_boletos_por_documento(pasta_pedidos)

    # Imprimir os agrupamentos
    for numero_documento, boletos in arquivos_por_documento.items():
        print(f"Documentos {numero_documento}: {boletos}")

    def agrupar_boletos_por_documento():
        boletos_por_documento = {}
        arquivos_renomeados = []

        renomear_boletos_pasta(pasta_pedidos)

        boletos_renomeados = [arquivo for arquivo in os.listdir(pasta_pedidos) if arquivo.lower().endswith('.pdf') and '-' in arquivo]

        if not boletos_renomeados:
            messagebox.showwarning("Erro", "Nenhum boleto encontrado para enviar.")
            return []

        for arquivo in boletos_renomeados:
            numero_documento = arquivo.split("-")[0]
            if numero_documento not in boletos_por_documento:
                boletos_por_documento[numero_documento] = []
            boletos_por_documento[numero_documento].append(arquivo)
            arquivos_renomeados.append(arquivo)  # Adiciona o arquivo renomeado na lista

        return arquivos_renomeados

    def marcar_boleto_enviado(boleto):
        novo_nome = boleto.replace(".pdf", "_enviado.pdf")
        os.rename(os.path.join(pasta_pedidos, boleto), os.path.join(pasta_pedidos, novo_nome))
        print(f"Boleto {boleto} marcado como enviado.")

    def centralizar_janela(root):
        root.update()
        largura = root.winfo_width()
        altura = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (largura // 2)
        y = (root.winfo_screenheight() // 2) - (altura // 2)
        root.geometry(f"{largura}x{altura}+{x}+{y}")

    root = ctk.CTk()
    root.title("Envio de Títulos")
    root.geometry("1400x750")
    root.configure(bg="#FFFFFF")

    frame = ctk.CTkFrame(root, fg_color="#FFFFFF")
    frame.pack(padx=10, pady=10, fill="both", expand=True)

    frame_button = ctk.CTkFrame(frame, fg_color="#FFFFFF")
    frame_button.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="ew")

    # Estilos de cores
    cor1 = "#D91656"
    cor2 = "#03346E"

    style = ttk.Style()
    style.configure("Treeview.Heading", 
                    font=('Helvetica', 10, 'bold'), 
                    background=cor1, 
                    foreground=cor2)
    style.configure("Treeview", 
                    background="#F9F9F9", 
                    rowheight=30, 
                    font=('Helvetica', 10))

    # Criando o Treeview com as colunas
    treeview = ttk.Treeview(frame, columns=('Título', 'Cliente', 'CNPJ/CPF', 'Cobrador', 'DtEmissao','Email', 'Arquivos'), show='headings')

    # Definindo os títulos das colunas
    treeview.heading('Título', text='Título')
    treeview.heading('Cliente', text='Cliente')
    treeview.heading('CNPJ/CPF', text='CNPJ/CPF')
    treeview.heading('Cobrador', text='Cobrador')
    treeview.heading('DtEmissao', text='DtEmissao')
    treeview.heading('Email', text='Email')
    treeview.heading('Arquivos', text='Arquivos')

    # Ajustando a exibição das colunas
    treeview.column('Título', width=10)
    treeview.column('Cliente', width=250) 
    treeview.column('CNPJ/CPF', width=150)
    treeview.column('Cobrador', width=120)
    treeview.column('DtEmissao', width=90)
    treeview.column('Email', width=150)
    treeview.column('Arquivos', width=200)

    # Exibe o Treeview
    treeview.grid(row=0, column=0, sticky="nsew")

    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    frame_button.columnconfigure(0, weight=0)
    frame_button.columnconfigure(1, weight=0)
    frame_button.columnconfigure(2, weight=0)
    frame_button.columnconfigure(3, weight=0)

    button_consultar = ctk.CTkButton(frame_button, text="Consulta", fg_color="#03346E")
    button_consultar.grid(row=0, column=0, padx=10, pady=5, sticky="w")

    data_inicial = ctk.CTkEntry(frame_button, width=80, placeholder_text="Data Inicial")
    data_inicial.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    data_inicial.bind("<KeyRelease>", formatar_data)

    data_final = ctk.CTkEntry(frame_button, width=80, placeholder_text="Data Final")
    data_final.grid(row=0, column=2, padx=5, pady=5, sticky="w")

    data_final.bind("<KeyRelease>", formatar_data)

    def on_double_click(event):
        item_id = treeview.selection()[0]  
        valores = treeview.item(item_id, "values")
        
        if valores:
            caminho_boleto = valores[6]
            
            if caminho_boleto and caminho_boleto != "Sem Arquivos":
                caminho_completo = os.path.join(pasta_pedidos, caminho_boleto)
                
                if os.path.exists(caminho_completo):
                    subprocess.run(["start", "", caminho_completo], shell=True)
                else:
                    messagebox.showerror("Erro", f"O arquivo {caminho_boleto} não foi encontrado!")
            else:
                messagebox.showwarning("Aviso", "Nenhum boleto associado a este pedido!")

    def on_consultar():
        data_inicial_val = data_inicial.get().strip()
        data_final_val = data_final.get().strip()

        try:
            datetime.strptime(data_inicial_val, '%d/%m/%Y')
            datetime.strptime(data_final_val, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Erro", "As datas inseridas estão no formato incorreto.")
            return

        data_inicial_formatted = datetime.strptime(data_inicial_val, "%d/%m/%Y").strftime("%Y%m%d")
        data_final_formatted = datetime.strptime(data_final_val, "%d/%m/%Y").strftime("%Y%m%d")

        colunas, dados = obter_dados_sql(data_inicial_formatted, data_final_formatted)

        if not colunas:
            messagebox.showerror("Erro", "Nenhum dado encontrado!")
            return

        if isinstance(treeview, ttk.Treeview):
            atualizar_treeview(treeview, colunas, dados)
        else:
            messagebox.showerror("Erro", "O componente treeview não é válido.")

    def atualizar_treeview(treeview, colunas, dados):
        for item in treeview.get_children():
            treeview.delete(item)

        if 'Arquivos' not in colunas:
            colunas.append('Arquivos')

        treeview["columns"] = colunas

        for col in colunas:
            treeview.heading(col, text=col)
            treeview.column(col, width=200, anchor="w")

        arquivos_por_titulo = {}
        # Filtra apenas arquivos PDF
        arquivos_pasta = [arquivo for arquivo in os.listdir(pasta_pedidos) if arquivo.lower().endswith('.pdf')]
        
        for arquivo in arquivos_pasta:
            # Verifica se o arquivo possui o padrão esperado
            if not arquivo.startswith("Stik.") or '-' not in arquivo:
                print(f"Padrão de arquivo inválido para {arquivo}")
                continue
            try:
                parte_inicial = arquivo.split('-')[0].replace("Stik.", "")
                numero_documento = str(int(parte_inicial))
                parcela = arquivo.split('-')[1].split('.')[0] 
                chave = f"{numero_documento}/{parcela}"  # Gera a chave no formato Título/Parcela
                arquivos_por_titulo[chave] = arquivo  # Salva somente um arquivo por chave
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")
                continue

        for linha in dados:
            titulo = linha[0]
            try:
                numero_titulo = str(int(titulo.split('.')[1].split('/')[0]))
                parcela_titulo = titulo.split('/')[-1]  # Obtém a parcela do título
                chave_titulo = f"{numero_titulo}/{parcela_titulo}"  # Chave de busca
            except Exception as e:
                print(f"Erro ao extrair número do título {titulo}: {e}")
                chave_titulo = ""

            arquivo = arquivos_por_titulo.get(chave_titulo, "Sem Arquivos")
            linha.append(arquivo)  # Adiciona apenas um arquivo correspondente

            treeview.insert('', 'end', values=linha)

    def on_item_click(event):
        item_id = treeview.selection()[0]
        valores = treeview.item(item_id, "values")

        if valores:
            for i, valor in enumerate(valores):
                print(f"Coluna {treeview['columns'][i]}: {valor}")

    def enviar_email_com_anexo(boletos_selecionados, pasta_pedidos, data_inicial_sql, data_final_sql):
        selecao = treeview.selection()
        
        if not selecao:
            messagebox.showerror("Erro", "Nenhum pedido selecionado!")
            return
        
        pedidos_dict = {}  # Dicionário para agrupar títulos pelo número base do pedido
        
        for item_id in selecao:
            valores = treeview.item(item_id, "values")
            
            if not valores or len(valores) < 6:
                messagebox.showerror("Erro", "Dados do pedido estão incompletos!")
                continue
            
            numero_documento = valores[0]  
            numero_base = numero_documento.rsplit("/", 1)[0]  # Remove a parte da parcela → "Stik.15385"
            cnpj_cpf = valores[2]  
            cobrador = valores[3]  
            email = valores[5]  
            arquivos_str = valores[6]

            if not cnpj_cpf or cnpj_cpf.strip() == "none":
                messagebox.showerror("Erro", f"Nenhum CPF ou CNPJ foi informado para o pedido {numero_documento}")
                continue

            if not cobrador or cobrador.strip().lower() == "none":
                messagebox.showerror("Erro:" f"Nenhum Agente Cobrador foi informado para o pedido {numero_documento}")        

            if not email or email.strip().lower() == "none":
                messagebox.showerror("Erro", f"Nenhum e-mail informado para o pedido {numero_documento}")
                continue

            if arquivos_str and arquivos_str != "Sem Arquivos":
                boletos_relacionados = [os.path.join(pasta_pedidos, nome.strip()) for nome in arquivos_str.split(",")]
            else:
                messagebox.showerror("Erro", f"Nenhum boleto encontrado para o pedido {numero_documento}!")
                continue
            
            if numero_base in pedidos_dict:
                pedidos_dict[numero_base]["boletos"].extend(boletos_relacionados)
            else:
                pedidos_dict[numero_base] = {
                    "email": email,
                    "boletos": boletos_relacionados
                }

        if not pedidos_dict:
            messagebox.showerror("Erro", "Nenhum pedido válido foi encontrado para envio de e-mail!")
            return
        
        lista_emails = "\n".join([f"Pedido {pedido} → {dados['email']}" for pedido, dados in pedidos_dict.items()])
        resposta = messagebox.askyesno("Confirmação", f"Os seguintes pedidos serão enviados:\n\n{lista_emails}\n\nDeseja continuar?")
        if not resposta:
            messagebox.showinfo("Cancelado", "O envio dos e-mails foi cancelado.")
            return

        remetente = "comunicacao@stik.com.br"
        senha = "Mailstk400"

        for numero_base, dados in pedidos_dict.items():
            email_destino = dados["email"]
            boletos_relacionados = list(set(dados["boletos"]))  # Remove duplicatas
            
            try:
                msg = MIMEMultipart()
                msg["From"] = remetente
                msg["To"] = email_destino
                msg["Subject"] = f"Boletos - {numero_base} / Stik Elásticos"

                corpo = f"""Prezado(a),\n
    Segue em anexo a 2ª via dos boletos e complementares referentes ao pedido {numero_base}.\n
    Por gentileza, conferir título e confirmar o recebimento.\n\nAtenciosamente,\n"""
                msg.attach(MIMEText(corpo, "plain"))

                for boleto in boletos_relacionados:
                    try:
                        with open(boleto, "rb") as arquivo_pdf:
                            parte_pdf = MIMEBase("application", "octet-stream")
                            parte_pdf.set_payload(arquivo_pdf.read())
                            encoders.encode_base64(parte_pdf)
                            parte_pdf.add_header("Content-Disposition", f"attachment; filename={os.path.basename(boleto)}")
                            msg.attach(parte_pdf)
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao anexar o arquivo {os.path.basename(boleto)} para o pedido {numero_base}: {str(e)}")
                        continue

                with smtplib.SMTP("smtp.stik.com.br", 587) as servidor:
                    servidor.starttls()
                    servidor.login(remetente, senha)
                    servidor.sendmail(remetente, email_destino, msg.as_string())

                messagebox.showinfo("Sucesso", f"E-mail enviado com sucesso para {email_destino} (Pedido {numero_base})")

            except smtplib.SMTPException as e:
                messagebox.showerror("Erro SMTP", f"Erro ao conectar ao servidor de e-mail para o pedido {numero_base}: {str(e)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao enviar e-mail para o pedido {numero_base}:{str(e)}")

    # Função para obter boletos relacionados ao pedido
    def obter_boletos_selecionados():
        boletos = []
        selecao = treeview.selection()

        if not selecao:
            return []

        for item_id in selecao:
            valores = treeview.item(item_id, "values")
            if len(valores) >= 6:
                arquivos = valores[5]
                if arquivos and arquivos != "Sem Arquivos":
                    boletos.extend(arquivos.split(", "))

        return boletos

    def enviar_pedidos_relacionados(titulo):
        arquivos_pasta = [arquivo for arquivo in os.listdir(pasta_pedidos) if arquivo.lower().endswith('.pdf')]
        arquivos_por_titulo = {}

        # Agrupar arquivos por número de documento
        for arquivo in arquivos_pasta:
            numero_documento_arquivo = extrair_numero_documento(os.path.join(pasta_pedidos, arquivo))
            if numero_documento_arquivo:
                arquivos_por_titulo.setdefault(numero_documento_arquivo, []).append(arquivo)
        
        arquivos = arquivos_por_titulo.get(titulo, [])
        
        if arquivos:
            # Chama a função de envio de e-mail com os boletos relacionados
            enviar_email_com_anexo(arquivos, pasta_pedidos, None, None)
        else:
            messagebox.showwarning("Aviso", f"Nenhum PDF encontrado para o título: {titulo}")

    # Pega as datas do campo de entrada e formata para o SQL
    def obter_datas():
        data_inicial_val = data_inicial.get().strip()
        data_final_val = data_final.get().strip()

        try:
            data_inicial_sql = datetime.strptime(data_inicial_val, "%d/%m/%Y").strftime("%Y%m%d")
            data_final_sql = datetime.strptime(data_final_val, "%d/%m/%Y").strftime("%Y%m%d")
            return data_inicial_sql, data_final_sql
        except ValueError:
            messagebox.showerror("Erro", "As datas estão no formato incorreto.")
            return None, None

    # Obtém as datas e envia o email se estiverem corretas
    def enviar_pedidos():
        data_inicial_sql, data_final_sql = obter_datas()
        
        if data_inicial_sql and data_final_sql:  # Verifica se as datas são válidas
            enviar_email_com_anexo(
                obter_boletos_selecionados(), 
                pasta_pedidos, 
                data_inicial_sql, 
                data_final_sql
            )

    button_enviar_pedidos = ctk.CTkButton(frame_button, text="Enviar Pedidos", fg_color="#03346E", command=enviar_pedidos)
    button_enviar_pedidos.grid(row=0, column=3, padx=5, pady=5, sticky="w")

    scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=treeview.yview)
    treeview.configure(yscrollcommand=scrollbar_y.set)
    scrollbar_y.grid(row=0,column=1, sticky="ns")
    
    button_consultar.configure(command=on_consultar)
    root.after(100, renomear_boletos_pasta(pasta_pedidos))
    treeview.bind("<ButtonRelease-1>", on_item_click)
    treeview.bind("<Double-1>", on_double_click)
    centralizar_janela(root)
    root.mainloop()

def main():
    iniciar_robo()
    
    time.sleep(1)
    
    pedidos = [
        arquivo
        for arquivo in os.listdir(pasta_pedidos)
        if arquivo.lower().endswith('.pdf') and '_enviado' not in arquivo
    ]
    
    if pedidos:
        print(f"Foram encontrados {len(pedidos)} pedidos. Iniciando a GUI...")
        Iniciar_GUI()
    else:
        messagebox.showinfo("Aviso", "Nenhum pedido encontrado para enviar.")

if __name__ == "__main__":
    main()