# -*- coding: utf-8 -*-

import os
import sys
import time
import warnings
from rich import print
from CMD import CMD
from rich.text import Text
from rich.style import Style
from rich.prompt import Prompt
from rich.console import Console
from rich.markdown import Markdown
from matplotlib import pyplot as plt

from database import consultaDB, PATH_DB

plt.close('all')
warnings.filterwarnings("ignore")

cmd = CMD()

db = consultaDB()
if not os.path.exists(PATH_DB):
    db.connect_db()
    db.create_database()
    db.close_db()

console_version = '1'

if console_version == '1':
    
    console = Console(width=167)
    os.system('mode 167, 57')
   
    print("[bold yellow]----------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    print("[bold yellow] [bold green]CMD[bold yellow] - Programa de Cálculo da Máxima Demanda Disponibilizada em Baixa Tensão ([bright_cyan]VERSÃO 0.1.2 - BETA[bold yellow])")
    print("[bold yellow] Desenvolvido pela ERA Energy Research and Analytics e pela CPFL Energia no âmbito do Projeto de PDI ANEEL")
    print("[bold yellow] PA3091 FERRAMENTA DE GESTÃO AUTOMÁTICA DA CAPACIDADE DE ACOMODAÇÃO (HOSTING CAPACITY) DE GERAÇÃO DISTRIBUÍDA")
    print("[bold yellow] Para mais informações sobre sua utilização, melhorias e correções contactar Heliton de Oliveira Vilibor (RESN)")
    print("[bold yellow]----------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    
    time.sleep(0.1)

elif console_version == '2':
    
    console = Console(width=167)
    os.system('mode 167, 57')
    
    console.print(r"[bold yellow]----------------------------------------------------------------------------------------------------------------------------------------------------------------------")    
    console.print(r"[bold green]   ___  _                  [bold red] ____ _____                                                                                                                                 ")
    console.print(r"[bold green]  / _ \| |__  _ __ __ _ ___[bold red]| __ )_   _| [bold yellow] ObrasBT - Programa de Emissão de Orçamentos de Conexão de Obras em Baixa Tensão ([bright_cyan]VERSÃO 0.1.2 - BETA[bold yellow])")
    console.print(r"[bold green] | | | | '_ \| '__/ _` / __[bold red]|  _ \ | |   [bold yellow] Desenvolvido pela ERA Energy Research and Analytics e pela CPFL Energia no âmbito do Projeto de PDI ANEEL")
    console.print(r"[bold green] | |_| | |_) | | | (_| \__ \\[bold red] |_) || |   [bold yellow] PA3091 FERRAMENTA DE GESTÃO AUTOMÁTICA DA CAPACIDADE DE ACOMODAÇÃO (HOSTING CAPACITY) DE GERAÇÃO DISTRIBUÍDA")
    console.print(r"[bold green]  \___/|_.__/|_|  \__,_|___/[bold red]____/ |_|   [bold yellow] Para mais informações sobre sua utilização, melhorias e correções contactar Heliton de Oliveira Vilibor (RESN)")
    console.print(r"[bold yellow]                                                                                                                                                                      ")
    console.print(r"[bold yellow]----------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    
    time.sleep(0.1)

linha_laranja = Style(color='dark_orange')
style_bold_red = Style(color='red', bold=True)
style_bold_green = Style(color='green', bold=True)

if 'dss' not in os.listdir(rf"{cmd.BASE_FOLDER}"):
    os.mkdir(rf"{cmd.BASE_FOLDER}\dss")
    cmd.logger.info('Pasta "dss" criada com sucesso!')

if 'Documentos Emitidos' not in os.listdir(rf"{cmd.BASE_FOLDER}"):
    os.mkdir(rf"{cmd.BASE_FOLDER}\Documentos Emitidos")
    cmd.logger.info('Pasta "Documentos Emitidos" criada com sucesso!')

cmd.logger.info("Iniciando conexão ao banco de dados CLTP...")
cmd.get_cltp()
if cmd.cur_cltp is not None:
    cmd.logger.info("Conexão ao banco de dados CLTP estabelecida com sucesso!")
else:
    cmd.logger.error("NÃO FOI POSSÍVEL ESTABELECER CONEXÃO AO BANCO DE DADOS CLTP!")

cmd.logger.info("Iniciando conexão ao banco de dados PPROJ...")
cmd.get_pproj()
if cmd.cur_pproj is not None:
    cmd.logger.info("Conexão ao banco de dados PPROJ estabelecida com sucesso!")
else:
    cmd.logger.error("NÃO FOI POSSÍVEL ESTABELECER CONEXÃO AO BANCO DE DADOS PPROJ!")
    sys.exit()

while True:
    
    modo = Prompt.ask('''[bright_cyan]\n Digite a opção desejada:[bright_cyan]\n
[bold green]1   [grey89]Configurações Iniciais
[bold green]2   [grey89]Importar e Preparar a Rede Secundária Extraída do GIS
[bold green]3   [grey89]Calcular a Máxima Demanda Disponibilizada (MDD) no Ponto de Conexão e o Índice de Proporcionalização da Obra (k) via Simulação de Fluxo de Potência
[bright_cyan]-------------------------------
[dark_orange]H   [dark orange]Ajuda / Glossário de Termos [bold bright_red][Em Construção][dark_orange]
[dark_orange]S   [dark orange]Sobre o Programa CMD [bold bright_red][Em Construção][dark_orange]
[dark_orange]Q   [dark orange]Para Encerrar o Programa[dark_orange]\n\n''')

    # [bold green]4   [grey89]Calcular a Participação Financeira do Consumidor (PFC) e o Encargo de Responsabilidade da Distribuidora (ERD)
    # [bold green]5   [grey89]Emitir Memória de Cálculo
    # [bold green]    [gray58]Opções Futuras A - Exemplo: Extrair os dados da rede secundária do GIS
    # [bold green]    [gray58]Opções Futuras B - Exemplo: Converter os dados da rede secundária extraída para o formato OpenDSS
    # [bold green]    [gray58]Opções Futuras C - Exemplo: Propor obra de reforço para atendimento da solicitação automaticamente
    # [bold green]    [gray58]Opções Futuras D - Exemplo: Emitir documento XYZ
    
    if modo == '1':
        print("")
        try:
            cmd.identify_activity()
        except:
            print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
    elif modo == '2':
        print("")
        try:
            cmd.get_dss_file()
        except:
            print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
    elif modo == '3':
        print("")
        try:
            start_time = time.time()
            cmd.run_mdd()
            try:
                cmd.logger.info("Registrando estudo no banco de dados...")
                db.connect_db()
                db.insert_record_estudo(cmd, start_time)
                db.close_db()
                cmd.logger.info("Estudo registrado no banco de dados com sucesso!")
                cmd.logger.info("Fim do estudo!")
            except:
                db.close_db()
                print("[bold bright_red]NÃO FOI POSSÍVEL REGISTRAR O ESTUDO NO BANCO DE DADOS. FAVOR CONTACTAR RESPONSÁVEL.")
        except:
            print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
    # elif modo == '4':
    #     print("")
    #     try:
    #         cmd.run_pfc_erd()
    #     except:
    #         print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
    # elif modo == '5':
    #     print("")
    #     try:
    #         cmd.create_memoria_calculo()
    #     except:
    #         print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
#    elif modo == '9':
#        print("")
#        try:
#            obt.function()
#        except:
#            print("[bold bright_red]NÃO FOI POSSÍVEL EXECUTAR A TAREFA. FAVOR CONTACTAR RESPONSÁVEL.")
    elif modo in ['h', 'H']:
        print("")
        console.rule("[dark_orange]Ajuda / Glossário de Termos / Referências", style=linha_laranja)
        print("")
        print(" [bold yellow]1 - [bold green]Configurações Iniciais: [white]descrição em construção...")
        print(" [bold yellow]2 - [bold green]Importar Rede Secundária Extraída do GIS: [white]descrição em construção...")
        print(" [bold yellow]3 - [bold green]Calcular a Máxima Demanda Disponibilizada (MDD) no Ponto de Conexão e o Índice de Proporcionalização da Obra (k) via Simulação de Fluxo de Potência: [white]descrição em construção...")
        #print(" [bold yellow]4 - [bold green]Calcular a Participação Financeira do Consumidor (PFC) e o Encargo de Responsabilidade da Distribuidora (ERD): [white]descrição em construção...")
        #print(" [bold yellow]5 - [bold green]Emitir Memória de Cálculo: [white]descrição em construção...")
        print("")
        print("[bold red] GLOSSÁRIO DE TERMOS")
        print(Markdown('- **CTO:** Custo Total da Obra é ...[continuar]'))
        print(Markdown('- **CTOp:** Custo Total da Obra Proporcionalizado é ...[continuar]'))
        print(Markdown('- **ERD:** Encargo de Responsabilidade da Distribuidora é ...[continuar]'))
        print(Markdown('- **MDD:** Máxima Demanda Disponibilizada é a máxima potência que os equipamentos elétricos e instalações do consumidor podem demandar do sistema elétrico da distribuidora em seu ponto de conexão antes que os limites mínimos de qualidade do produto especificados na regulação sejam atingidos.'))
        print(Markdown('- **OC:**  Orçamento de Conexão é o conjunto de documentos...[continuar]'))
        print(Markdown('- **PC:**  Ponto de Conexão é o conjunto de materiais e equipamentos que se destina a estabelecer a conexão entre as instalações da distribuidora e do consumidor e demais usuários. (REN 1000)'))
        print(Markdown('- **PFC:** Participação Financeira do Consumidor é ...[continuar]'))
        print("")
        print("[bold red] REFERÊNCIAS")        
        print(Markdown('- **GED 3793:** Utilização de Fator de Carga e Demanda Típicos'))
        print(Markdown('- **GED 150488:** Cálculo do PFC Encargo de Responsabilidade da Distribuição e Fator de Proporcionalidade'))
        print("")
        console.rule("", style=linha_laranja)
    elif modo in ['s', 'S']:
        print("")
        console.rule("[dark_orange]Sobre o programa CMD", style=linha_laranja)
        print("")
        print(Text("### CMD - PROGRAMA DE CÁLCULO DA MÁXIMA DEMANDA DISPONIBILIZADA EM BAIXA TENSÃO ###", justify='center', style=style_bold_red))
        print("[bold white]VERSÃO [bright_cyan]0.1.2 (BETA)[bold white] de [bright_cyan]16 de outubro de 2025")
        print("")
        print("[bold white]Este programa foi desenvolvido pela ERA Energy Research and Analytics e pela CPFL Energia no âmbito do Programa de Pesquisa, Desenvolvimento & Inovação da Agência Nacional de Energia Elétrica (ANEEL) Projeto PA3091 - Ferramenta de Gestão Automática da Capacidade de Acomodação (Hosting Capacity) de Geração Distribuída, iniciado em 2 de outubro de 2023 com finalização prevista para 1° de setembro de 2027.")
        print("")
        print("[bold red]### EQUIPE PARTICIPANTE DO PROJETO ###")
        print("[bold green]Gerente: [bold white]Heliton de Oliveira Vilibor (CPFL Energia)")
        print("[bold green]Coordenador: [bold white]Walmir de Freitas Filho (UNICAMP)")
        print("[bold green]Inovação: [bold white]Camila de Freitas Albertin (CPFL Energia)")
        print("[bold green]ERA: [bold white]Ricardo Torquato Borges, José Carlos Garcia Andrade, Pedro Augusto Viana Pato, Rodrigo Santos Bonadia, Tiago de Moraes Barbosa, Tiago Rodarte Ricciardi, Vinicius Carnelossi da Cunha")
        print("")
        third_part_list = [
            '[bold yellow] AltDSS-Python [white](https://dss-extensions.org/AltDSS-Python/)',
            '[bold yellow]cx_Oracle [white](https://oracle.github.io/python-cx_Oracle/)',
            '[bold yellow]matplotlib [white](https://matplotlib.org/)',
            '[bold yellow]NumPy [white](https://numpy.org/)',
            #'[bold yellow]oracledb [white](https://https://oracle.github.io/python-oracledb/)',
            '[bold yellow]pandas [white](https://pandas.pydata.org/)',
            '[bold yellow]python-docx-template [white](https://docxtpl.readthedocs.io)',
            '[bold yellow]pywin32 [white](https://github.com/mhammond/pywin32)',
            '[bold yellow]rich [white](https://rich.readthedocs.io)'
            ]
        print(f"[bold white]O programa CMD foi escrito em linguagem Python e utilizou as seguintes bibliotecas de terceiros: {', '.join(third_part_list)}.")
        print("")
        console.rule("", style=linha_laranja)
    elif modo in ['q', 'Q']:
        print("")
        print("[bold bright_cyan] Até a próxima!")
        break
    else:
        print("")
        print('[bold bright_red]Opção inválida!')

