# -*- coding: utf-8 -*-

VMAX = 1.05
VMIN = 0.93
IMAX = 1.00
SMAX = 1.00

USER_CLTP = 'MicroGD'
PASSWORD_CLTP = 'b691#Rd3c0e'
USER_PPROJ = 'aplicacao_hc_pnt'
PASSWORD_PPROJ = 'LnAadFKP'

import warnings
warnings.filterwarnings("ignore")

import os
import time
import shutil
import locale
import logging
import zipfile
#import oracledb
import cx_Oracle
import numpy as np
import pandas as pd
import datetime as dt
import networkx as nx
from altdss import altdss
from win32com import client
from rich.prompt import Prompt
from docxtpl import DocxTemplate
from rich.console import Console
from rich.logging import RichHandler
from matplotlib import pyplot as plt
from dss import SolveModes, ControlModes, DSSSaveFlags

VERSION = '0.1.3 - RC'
VERSION_DATA = '4 de novembro de 2025'

DB_REMOTO = True

class CMD():
    
    def __init__(self):
    
        self.BASE_FOLDER = r'C:\CMD'
        
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        logging.basicConfig(
            format='[%(asctime)s.%(msecs)03d] %(levelname)-8s %(message)s',
            datefmt='%d/%m/%Y %H:%M:%S'
        )
        
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        
        ### De acordo com a GED 4319: Ramal de Ligação - Montagem (https://sites.cpfl.com.br/documentos-tecnicos/GED-4319.pdf)
        self.dict_padrao_ramal_127_220 = {
            'A1' : '1P10(A10)',
            'A2' : '1P16(A16)',
            'B1' : '2P16(A16)',
            'B2' : '2P25(A25)',
            'C1' : '3P10(A10)',
            'C2' : '3P16(A16)',
            'C3' : '3P25(A25)',
            'C4' : '3P35(A35)',
            'C5' : '3P50(A50)',
            'C6' : '3P70(A70)'
        }
        
        self.dict_padrao_ramal_220_380 = {
            'A3' : '1P10(A10)',
            'A4' : '1P16(A16)',
            'B3a (15 < C ≤ 20 kW)' : '2P10(A10)',
            'B3b (20 < C ≤ 25 kW)' : '2P16(A16)',
            'C7' : '3P10(A10)',
            'C8' : '3P16(A16)',
            'C9' : '3P25(A25)',
            'C10' : '3P35(A35)',
            'C11' : '3P35(A35)'
        }
        
        self.dict_linecodes = {
            '1P10(A10)' : {'nphases': 1, 'r1' : 3.7524600029, 'x1' : 0.9116349816, 'r0' : 3.7524600029, 'x0' : 0.9116349816, 'normaps' :	65},
            '1P16(A16)' : {'nphases': 1, 'r1' : 2.4344499111, 'x1' : 0.8530510068, 'r0' : 2.4344499111, 'x0' : 0.8530510068, 'normaps' :	65},
            '2P10(A10)' : {'nphases': 2, 'r1' : 3.5003540516, 'x1' : 0.1149839982, 'r0' : 4.2280411720, 'x0' : 2.5151879787, 'normaps' :	311},
            '2P16(A16)' : {'nphases': 2, 'r1' : 2.1716587543, 'x1' : 0.2170950025, 'r0' : 2.9251470566, 'x0' : 2.1441090107, 'normaps' :	65},
            '2P25(A25)' : {'nphases': 2, 'r1' : 1.3646429777, 'x1' : 0.1135369986, 'r0' : 2.4357099533, 'x0' : 1.8574759960, 'normaps' :	100},
            '3P10(A10)' : {'nphases': 3, 'r1' : 3.5003540516, 'x1' : 0.1122030020, 'r0' : 4.2378230095, 'x0' : 2.5173320770, 'normaps' :	100},
            '3P16(A16)' : {'nphases': 3, 'r1' : 2.1716587543, 'x1' : 0.2143439949, 'r0' : 2.9372129440, 'x0' : 2.1432290077, 'normaps' :	65},
            '3P25(A25)' : {'nphases': 3, 'r1' : 1.3646429777, 'x1' : 0.1108369976, 'r0' : 2.4526889324, 'x0' : 1.8497240543, 'normaps' :	100},
            '3P35(A35)' : {'nphases': 3, 'r1' : 0.9865019917, 'x1' : 0.1084029973, 'r0' : 2.0914440155, 'x0' : 1.5169969797, 'normaps' :	129},
            '3P50(A50)' : {'nphases': 3, 'r1' : 0.7288320065, 'x1' : 0.1056289971, 'r0' : 1.7709170580, 'x0' : 1.2131340504, 'normaps' :	168},
            '3P70(A70)' : {'nphases': 3, 'r1' : 0.5038059950, 'x1' : 0.1041790023, 'r0' : 1.3816900253, 'x0' : 0.9110010266, 'normaps' :	227},
            '3P120(A120)' : {'nphases': 3, 'r1' : 0.28750041127204895, 'x1' : 0.10087999701499939, 'r0' : 0.8778550028800964, 'x0' : 0.6327559947967529, 'normaps' :	311}
        }
        
        self.FDs = [0.6, 0.35, 0.28, 0.3, 0.32, 0.42, 0.29, 0.27, 1, 0.38, 0.28, 0.23, 0.51, 0.2, 0.39, 0.34, 0.32, 0.53, 0.55, 0.42, 1, 0.32, 0.51, 0.2, 0.28]
        self.atividades = ['Bar', 'Beneficiamento de cereais', 'Carpintaria', 'Fábrica de bebidas', 'Fábrica de calçados', 'Fábrica de plásticos', 'Fábrica de roupas', 'Hotel', 'Iluminação (festiva, ornamental, etc.)', 'Laticínio', 'Oficina mecânica', 'Padaria', 'Posto de gasolina', 'Residência', 'Restaurante', 'Serraria', 'Semáforo', 'Sorveteria', 'Supermercado', 'Comércio, Serviço e outras atividades', 'Iluminação Pública', 'Industrial', 'Poder Público', 'Residencial', 'Rural']
        
        self.dss = altdss
        
        self.kiter = 0
        
        self.voltage_violation = False
        self.current_violation = False
        self.transformer_violation = False
        
        self.V_ARRAY = []
        self.I_ARRAY = []
        self.S_ARRAY = []
        
        self.v_array = []
        self.i_array = []
        self.s_array = []
        
        self.l_array = []
        
        self.dssfile = None
        
        self.today = dt.datetime.today().strftime("%A, %d de %B de %Y")
        self.today = self.today[0].upper() + self.today[1:]
    
    def _acess_pproj(self, host, port, service):
    
        os.environ['PATH'] = os.environ['PATH'] + \
            ';C:/41/sqldeveloper/instantclient'
        strCon = f'{USER_PPROJ}/{PASSWORD_PPROJ}@//{host}:{port}/{service}'
        conn = cx_Oracle.connect(strCon, encoding='UTF-8', nencoding='UTF-8')
        
        return conn
    
    def get_pproj(self):
    
        host = '192.168.35.221'
        port = '1535'
        service = 'ppartp.cpfl.com.br'
        
        try:
            self.con_pproj = self._acess_pproj(host, port, service)
            self.cur_pproj = self.con_pproj.cursor()
        except:
            self.con_pproj = None
            self.cur_pproj = None
    
    def _acess_cltp(self, host, port, service):
    
        os.environ['PATH'] = os.environ['PATH'] + \
            ';C:/41/sqldeveloper/instantclient'
        strCon = f'{USER_CLTP}/{PASSWORD_CLTP}@//{host}:{port}/{service}'
        conn = cx_Oracle.connect(strCon, encoding='UTF-8', nencoding='UTF-8')
        #conn = oracledb.connect(strCon)
        
        return conn
    
    def get_cltp(self):
    
        host = 'dbscan-CLTP.cpfl.com.br'
        port = '1521'
        service = 'pdb_CLTP.cpfl.com.br'
        
        try:
            self.con_cltp = self._acess_cltp(host, port, service)
            self.cur_cltp = self.con_cltp.cursor()
        except:
            self.con_cltp = None
            self.cur_cltp = None
    
    def identify_activity(self):
    
        res = Prompt.ask(' [bright_cyan]Digite o número da Nota de Serviço')# "Nota CCS"
       
        self.nota = res.zfill(12)
        
        print("")
        self.logger.info(f"Localizando informações da Nota de Serviço {self.nota} no banco de dados CLTP...")
        
        self.dados_nota = self.cur_cltp.execute(f'''
             SELECT
                 COD_INSTALACAO, DESC_TEXTO_NOTA
             FROM
                 DWCCS.DWF_NOTA_SERVICO
             WHERE
                 COD_NOTA_SERVICO = '{self.nota}'
        ''').fetchone()
        
        self.cod_emp = self.cur_cltp.execute(f'''
            SELECT
                COD_EMPRESA
            FROM
                DWCCS.DWD_ESTACAO_AVANCADA
            WHERE
                COD_ESTACAO_AVANCADA IN (
                    SELECT
                        COD_ESTACAO_AVANCADA
                    FROM
                        DWCCS.DWF_NOTA_SERVICO
                    WHERE
                        COD_NOTA_SERVICO IN (
                            '{self.nota}'
                        )
                )
        ''').fetchone()[0]
        
        if self.cod_emp in ['CPFL']:
            self.disco = 'CPFL Paulista'
            self.empresa = 'Paulista'
        elif self.cod_emp in ['PIRA']:
            self.disco = 'CPFL Piratininga'
            self.empresa = 'Piratininga'
        elif self.cod_emp in ['D003', 'D004', 'D005', 'D006', 'D007']:
            self.disco = 'CPFL Santa Cruz'
            self.empresa = 'Santa Cruz'
        elif self.cod_emp in ['D008', 'D009']:
            self.disco = 'CPFL RGE'
            self.empresa = 'RGE'
        else:
            self.empresa = None
        
        self.tipo_atividade = self.dados_nota[1]
        
        if self.tipo_atividade is None:
            self.logger.warning(f"ATENÇÃO: Nota de Serviço {self.nota} NÃO foi localizado no banco de dados CLTP. Verifique o número informado e tente novamente.")
            return
        
        if self.empresa is None:
            self.dados_pproj = self.cur_pproj.execute(f'''
                 SELECT
                     CD_EMPRESA, CD_UC, VL_DEMANDA, VL_POTENCIA
                 FROM
                     PPART.PROJETO
                 WHERE
                     NR_NOTA_SERVICO = '{self.nota}'
            ''').fetchone()
            if self.dados_pproj is not None:
                self.empresa_pproj = self.dados_pproj[0]
                self.uc_pproj = self.dados_pproj[1]
                self.VL_DEMANDA = self.dados_pproj[2]
                self.VL_POTENCIA = self.dados_pproj[3]
                if self.empresa_pproj == 1:
                    self.disco = 'CPFL Paulista'
                    self.empresa = 'Paulista'
                elif self.empresa_pproj == 2:
                    self.disco = 'CPFL Piratininga'
                    self.empresa = 'Piratininga'
                elif self.empresa_pproj == 13:
                    self.disco = 'CPFL Santa Cruz'
                    self.empresa = 'Santa Cruz'
                elif self.empresa_pproj == 18:
                    self.disco = 'CPFL RGE'
                    self.empresa = 'RGE'
        
        self.logger.info(f"Informações da Nota de Serviço {self.nota} localizadas com sucesso!")
        
        if self.empresa is not None:
            self.logger.info(f"A distribuidora foi automaticamente identificada com base na Nota de Serviço ({self.empresa})...")
        else:
            loop = True
            print("")
            res = Prompt.ask(' [bright_cyan]Selecione a distribuidora:\n\n [1] [cyan]Paulista\n[bright_cyan] [2] [dark_orange]Piratininga\n[bright_cyan] [3] [green]Santa Cruz\n[bright_cyan] [4] [red]RGE\n\n')
            while loop:
                if res in ['1', '2', '3', '4']:
                    if res == '1':
                        self.disco = 'CPFL Paulista'
                        self.empresa = 'Paulista'
                    elif res == '2':
                        self.disco = 'CPFL Piratininga'
                        self.empresa = 'Piratininga'
                    elif res == '3':
                        self.disco = 'CPFL Santa Cruz'
                        self.empresa = 'Santa Cruz'
                    elif res == '4':
                        self.disco = 'CPFL RGE'
                        self.empresa = 'RGE'
                    print("")
                    self.logger.info(f"A distribuidora selecionada foi a {self.empresa}...")
                    #self.logger.info(f"A distribuidora escolhida foi a {self.disco}...Iniciando a conexão ao banco de dados GIS da {self.disco} (SW_GIS)...")
                    #self.get_swgis(self.disco)
                    #if self.cur_swgis is not None:
                        #self.logger.info("Banco de dados SW_GIS conectado com sucesso!")
                    #else:
                        #self.logger.warning("Não foi possível estabelecer conexão com o banco de dados SW_GIS!")
                    loop = False
                else:
                    print("")
                    res = Prompt.ask(' [bright_cyan]Seleção inválida! Selecione a distribuidora:\n\n [1] [cyan]CPFL Paulista\n[bright_cyan] [2] [dark_orange]CPFL Piratininga\n[bright_cyan] [3] [green]CPFL Santa Cruz\n[bright_cyan] [4] [red]RGE\n\n')
                    loop = True
        
        if (self.dados_nota[0] == ' ') and (self.tipo_atividade != 'Reforma e Adequação - Aum. de Carga Edif'):
            self.ligacao_nova = True
            self.logger.info(f"Tipo de Atividade: {self.tipo_atividade}")
            self.logger.info(f"ATENÇÃO! UC não encontrada. Trata-se de LIGAÇÃO NOVA. Demanda Existente (DE) = 0 kW")
            self.DE = 0
            self.uc = None
        else:
            self.ligacao_nova = False
            if self.dados_nota[0] == ' ':
                self.logger.info(f"Tipo de Atividade: {self.tipo_atividade}")
            else:
                self.logger.info(f"Unidade Consumidora: {self.dados_nota[0]}, Tipo de Atividade: {self.tipo_atividade}")
                self.uc = self.dados_nota[0]
            print("")
            self.Demanda_Vigente_input = Prompt.ask(' [bright_cyan]Informe a Demanda Existente (DE), em kW')
            self.DE = float(self.Demanda_Vigente_input.replace(',', '.'))
            print("")
            if self.dados_nota[0] == ' ':
                self.logger.info(f"A demanda existente (DE, i.e., a demanda atual) é de {locale.format_string('%.2f kW', self.DE)}...")
            else:
                self.logger.info(f"A demanda existente (DE, i.e., a demanda atual) da UC {self.dados_nota[0]} é de {locale.format_string('%.2f kW', self.DE)}...")
        
        self.logger.info(f'Criando a pasta da nota {self.nota} no diretório "Documentos Emitidos"')
        if self.nota not in os.listdir(rf"{self.BASE_FOLDER}\Documentos Emitidos"):
            os.mkdir(rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}")
            self.logger.info(f'Pasta da nota {self.nota} criada com sucesso!')
        else:
            self.logger.warning(f'Pasta da nota {self.nota} já existente!')
        
        self.logger.info("Configurações iniciais finalizadas com sucesso!")
    
    def identify_critical_path(self):
    
        self.logger.info("Identificando o caminho crítico...")
        target = self.interest_bus[0].split('.')[0]
        source = self.raiz
        edges = [item for item in nx.all_simple_edge_paths(self.F, source, target, cutoff=None)][0]
        self.critical_path = [self.F.edges[item] for item in edges]
        lcp = len(self.critical_path)
        if lcp == 1:
            self.logger.info(f"Caminho crítico identificado com sucesso: {lcp} vão entre o ponto de entrega da carga de estudo e o transformador de distribuição.")
        elif len(self.critical_path) > 1:
            self.logger.info(f"Caminho crítico identificado com sucesso: {lcp} vãos entre o ponto de entrega da carga de estudo e o transformador de distribuição.")
    
    def build_graph(self):
    
        self.logger.info("Montando o grafo da rede secundária...")
        
        self.G = nx.Graph()
        
        for line in self.dss.Line:
            
            self.G.add_edges_from([(
                line.Bus1.split('.')[0],
                line.Bus2.split('.')[0],
                {'Name'   : line.Name,
                 'Length' : line.Length,
                 'Units'  : line.Units_str})])
        
        self.logger.info("Grafo da rede secundária montado com sucesso!")
        
        self.logger.info("Orientando o grafo da rede secundária...")
        
        self.raiz = self.dss.Transformer[0].Buses[-1].split('.')[0]
        
        self.F = nx.bfs_tree(self.G, self.raiz)
        
        ### Copia atributos das arestas
        for e in self.F.edges:
            self.F.add_edges_from([(e[0], e[1], self.G.edges[e])])
        for n in self.F.nodes:
            self.F.add_nodes_from([(n, self.G.nodes[n])])
        
        self.logger.info("Grafo da rede secundária orientado com sucesso!")
        
        self.identify_critical_path ()
    
    def get_dss_file(self):
    
        self.novo_trafo = False 
        
        files = os.listdir(rf"{self.BASE_FOLDER}\dss")
        modtime = [time.localtime(os.path.getmtime(rf"{self.BASE_FOLDER}\dss\{item}")) for item in files]
        fs = pd.Series(modtime, index=files).sort_values(ascending=False)
        fs = fs.apply(lambda x: time.strftime("%A, %d de %B de %Y %H:%M:%S", x))
        files = fs.index.to_list()[0:10]
        
        fcount = 1
        filesstr = ''
        for file in files:
            filesstr = filesstr + f'[bold yellow] [{fcount}] [bold green]{file} [grey78]{fs.loc[file]}\n'
            fcount += 1
        
        filesstr = filesstr + f'[bold yellow] [A] [bold green]Quero digitar o nome do arquivo...\n' + f'[bold yellow] [B] [bold green]Atendimento via trafo exclusivo (sem rede secundária existente)...\n'
        
        res = Prompt.ask(f' [bright_cyan]Selecione o arquivo exportado do GIS-D após o projeto da obra:\n\n{filesstr}[bright_cyan]\n')
        
        if res in ["a", "A"]:
            loop = True
            while loop:
                print("")
                choosen_file = Prompt.ask(f' [bright_cyan]Digite o nome do arquivo, incluindo a sua extensão no formato <nome_arquivo.ext>')
                if choosen_file not in files:
                    choosen_file = Prompt.ask(f'\n [red]ATENÇÃO! O arquivo {choosen_file} não foi localizado na pasta dss!\n\n[bright_cyan] Por favor confira o nome e digite novamente o nome do arquivo, incluindo a sua extensão no formato <nome_arquivo.ext>')
                    if choosen_file in files:
                        print("")
                        self.logger.info(f"Arquivo {choosen_file} encontrado!")
                        loop = False                         
                else:
                    print("")
                    self.logger.info(f"Arquivo {choosen_file} encontrado!")
                    loop = False
        
        elif res in ["b", "B"]:
            self.novo_trafo = True
            loop = True
            while loop:
                print("")
                
                new_tr_string = f'''
[bold yellow] [1] [bold red]10 kVA (Monofásico)
[bold yellow] [2] [bold green]15 kVA (Monofásico)
[bold yellow] [3] [bold green]25 kVA (Monofásico)
[bold yellow] [4] [bold red]50 kVA (Monofásico)

[bold yellow] [5] [bold green]15 kVA (Trifásico)
[bold yellow] [6] [bold green]30 kVA (Trifásico)
[bold yellow] [7] [bold green]45 kVA (Trifásico)
[bold yellow] [8] [bold green]75 kVA (Trifásico)
[bold yellow] [9] [bold green]112,5 kVA (Trifásico
[bold yellow] [10] [bold green]150 kVA (Trifásico)
[bold yellow] [11] [bold green]225 kVA (Trifásico)
[bold yellow] [12] [bold green]300 kVA (Trifásico)\n'''
                
                res = Prompt.ask(f' [bright_cyan]Escolha a potência (kva) do transformador de distribuição que será instalado para atendimento exclusivo à nova ligação:\n{new_tr_string}[brigh_cyan]\n')
                
                if res not in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']:
                
                    pass    
                    #print("")    
                    #res = Prompt.ask(f' [bright_cyan]Escolha a potência (kva) do transformador de distribuição que será instalado para atendimento exclusivo à nova ligação:\n\n{new_tr_string}[brigh_cyan]\n')
                
                else:
                    pots = [10, 15, 25, 50, 15, 30, 45, 75, 112.5, 150, 225, 300]
                    self.novo_tr_kva = pots[int(res)-1]
                    #print("")
                    
                    if res in ['1', '4']:
                        print("")
                        res = Prompt.ask(f' [bright_cyan]Atenção! O transformador selecionado ({self.novo_tr_kva} kVA Monofásico) é de uso restrito. Tem certeza que deseja continuar com essa escolha? [bold green ][S] Sim [bright_cyan]ou [bold red][N] Não\n\n')
                        
                        if res in ['s', 'S']:
                            print("")
                            self.logger.info(f"Será instalado um novo transformador de distribuição de {self.novo_tr_kva} kVA para atendimento exclusivo à nova ligação!")
                            loop = False
                        elif res in ['n', 'N']:
                            continue
                        else:
                            continue
                    else:
                        print("")
                        self.logger.info(f"Será instalado um novo transformador de distribuição de {self.novo_tr_kva} kVA para atendimento exclusivo à nova ligação!")
                        loop = False
        
        else:
            choosen_file = files[int(res)-1]
            print("")
            
        if self.novo_trafo:
            pass
            
        else:
            self.dssfile = rf"{self.BASE_FOLDER}\dss\{choosen_file}"
            
            self.logger.info(f"O caminho completo do arquivo dss selecionado é: {self.dssfile}")
            
            self.dss("Clear")
            self.dss(f"Redirect {self.dssfile}")
            ret = self.get_interest_bus()
            
            if ret == 'error':
                return
            
            self.flags_save = DSSSaveFlags.CalcVoltageBases | DSSSaveFlags.ExcludeDefault | DSSSaveFlags.SingleFile | DSSSaveFlags.KeepOrder | DSSSaveFlags.SetVoltageBases | DSSSaveFlags.IsOpen | DSSSaveFlags.ToString
            self.circuit_string = self.dss.to_dss_python().Circuits.Save('', self.flags_save)
        
        self.logger.info("Importação e preparação da rede secundária extraída do GIS finalizada com sucesso!")
    
    def change_service_cable(self, bitola):
    
        self.logger.info("Iniciando a troca do Ramal de Ligação...")
        if bitola.lower() in self.dss.LineCode.Name:
            self.dss.Line[self.ramal_ligacao].LineCode.Name = bitola.lower()
        else:
            nphases = self.dict_linecodes[bitola]['nphases']
            r1 = self.dict_linecodes[bitola]['r1']
            x1 = self.dict_linecodes[bitola]['x1']
            r0 = self.dict_linecodes[bitola]['r0']
            x0 = self.dict_linecodes[bitola]['x0']
            normaps = self.dict_linecodes[bitola]['normaps']
            self.dss(f'new linecode.{bitola.lower()},nphases={nphases},r1={r1},x1={x1},r0={r0},x0={x0},normamps={normaps},unit=km')
            self.dss.Line[self.ramal_ligacao].LineCode = self.dss.LineCode[bitola.lower()]
        self.logger.info("Troca do Ramal de Ligação finalizada com sucesso!")
    
    def avalia_ramal_coletivo(self):
    
        print("")    
        self.logger.info(f'Ramal de Ligação Aéreo com Atendimento Agrupado: será avaliado o dimensionamento do ramal conforme GEDs 119 e 4319...')
    
        ramais = [
            '10 mm² - Quadruplex',
            '16 mm² - Quadruplex',
            '25 mm² - Quadruplex',
            '35 mm² - Quadruplex',
            '50 mm² - Quadruplex',
            '70 mm² - Quadruplex',
            '120 mm² - Quadruplex',
            '2x70 mm² - Quadruplex',
            '2x120 mm² - Quadruplex'
            ]
        
        rcount = 1
        ramaisstr = ''
        for ramal in ramais:
            ramaisstr = ramaisstr + f'[bold green] [{rcount}] [bold yellow]{ramal}\n'
            rcount += 1
        
        print("")
        res = Prompt.ask(f' [bright_cyan]Selecione o "Ramal de Ligação" (GED 13) futura da UC, ou seja, após o aumento de carga solicitado:\n\n{ramaisstr}\n[bright_cyan]')            
        
        choosen_ramal = ramais[int(res)-1]
        print("")
        self.logger.info(f"Foi selecionado o Ramal de Ligação {choosen_ramal}...")
        
        fases_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()[0])
        fases_norma = 3
        
        bitola_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper().split('P')[-1].split('(')[0])
        bitola_norma = int(choosen_ramal.split(' ')[0].replace('2x', ''))
        qtde = 2 if '2x' in choosen_ramal else 1
        
        self.ramal_modificado = False
        if (fases_existente < fases_norma) or (bitola_existente < bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) inferior ao informado ({choosen_ramal}) e será substituído para o cálculo da MDD...')
            self.change_service_cable(f"3P{bitola_norma}(A{bitola_norma})")
            self.ramal_modificado = True
        elif (fases_existente > fases_norma) or (bitola_existente > bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente na base GIS ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) é superior ao informado ({choosen_ramal}): será mantido ramal da base GIS para o cálculo da MDD...')
        else:
            self.logger.info(f'Ramal de Ligação existente igual ao informado ({choosen_ramal}) e será mantido para o cálculo da MDD...')
    
    def avalia_ramal_subterraneo(self):
    
        print("")
        self.logger.info(f'Ramal de Ligação Subterrâneo: CE será transferida para o PC do ramal subterrâneo à rede secundária (GED 10126)...')
        self.logger.info(f'Ramal de Ligação Subterrâneo: o ramal subterrâneo será ignorado nos cálculos...')
        
        for load in self.interest_loads:
            if self.dss.Load[load].Bus1.split('.')[0] == self.dss.Line[self.ramal_ligacao].Bus2.split('.')[0]:
                temp = self.dss.Load[load].Bus1.split('.')
                temp[0] = self.dss.Line[self.ramal_ligacao].Bus1.split('.')[0]
                self.dss.Load[load].Bus1 = '.'.join(temp)
            elif self.dss.Load[load].Bus1.split('.')[0] == self.dss.Line[self.ramal_ligacao].Bus1.split('.')[0]:
                temp = self.dss.Load[load].Bus1.split('.')
                temp[0] = self.dss.Line[self.ramal_ligacao].Bus2.split('.')[0]
                self.dss.Load[load].Bus1 = '.'.join(temp)
            else:
                self.logger.error("ERRO! Não foi possível transferir a CE para o PC do ramal subterrâneo à rede secundária!")
        
        self.dss.Line[self.ramal_ligacao].Enabled = False
        
        self.logger.info(f'Ramal de Ligação Subterrâneo: CE transferida para o PC do ramal subterrâneo à rede secundária com sucesso!')
    
    def avalia_ramal_individual(self):
    
        print("")    
        self.logger.info(f'Ramal de Ligação Aéreo com Atendimento Individual: será avaliado o dimensionamento do ramal conforme GEDs 13 e 4319...')
        
        if (self.tensao_secundaria == 0.220) and (self.delta_wye):
            if self.kva0 <= 23:
                categoria_correta = 'C1'
                ramal_correto = '3P10(A10)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 30:
                categoria_correta = 'C2'
                ramal_correto = '3P16(A16)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 38:
                categoria_correta = 'C3'
                ramal_correto = '3P25(A25)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 47:
                categoria_correta = 'C4'
                ramal_correto = '3P35(A35)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 57:
                categoria_correta = 'C5'
                ramal_correto = '3P50(A50)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 76:
                categoria_correta = 'C6'
                ramal_correto = '3P70(A70)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            else:
                self.logger.warning(f"Demanda aparente da CE: ({locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 127/220 V); NÃO FOI POSSÍVEL IDENTIFICAR A CATEGORIA/RAMAL DE LIGAÇÃO PADRONIZADOS.")
        elif (self.tensao_secundaria == 0.380) and (self.delta_wye):
            if self.kva0 <= 26:
                categoria_correta = 'C7'
                ramal_correto = '3P10(A10)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 40:
                categoria_correta = 'C8'
                ramal_correto = '3P16(A16)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 52:
                categoria_correta = 'C9'
                ramal_correto = '3P25(A25)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 66:
                categoria_correta = 'C10'
                ramal_correto = '3P35(A35)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            elif self.kva0 <= 82:
                categoria_correta = 'C11'
                ramal_correto = '3P35(A35)'
                self.logger.info(f"Demanda aparente da CE: {locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V; Categoria de Ligação: {categoria_correta}/{ramal_correto}.")
            else:
                self.logger.warning(f"Demanda aparente da CE: ({locale.format_string('%.2f kVA', self.kva0)}; Tensão de atendimento: 220/380 V); NÃO FOI POSSÍVEL IDENTIFICAR A CATEGORIA/RAMAL DE LIGAÇÃO PADRONIZADOS.")
        
        if self.delta_wye and self.tensao_secundaria == 0.220:
            categorias = list(self.dict_padrao_ramal_127_220.keys())
            self.dict_padrao_ramal = self.dict_padrao_ramal_127_220
        elif self.delta_wye and self.tensao_secundaria == 0.380:
            categorias = list(self.dict_padrao_ramal_220_380.keys())
            self.dict_padrao_ramal = self.dict_padrao_ramal_220_380
        else:
            self.logger.error(f"ERRO! Ferramenta não prevê transformadores com conexão diferente de Delta/Estrela ou Tensões Secundárias de {1e3*self.tensao_secundaria} (i.e., diferentes de 220 ou de 380 V). Favor contactar responsável.")
            return 'error'
        
        ccount = 1
        categoriasstr = ''
        for categoria in categorias:
            categoriasstr = categoriasstr + f'[bold green] [{ccount}] [bold yellow]{categoria}\n'
            ccount += 1
        
        print("")
        res = Prompt.ask(f' [bright_cyan]Selecione a "Categoria de Ligação" (GED 13) futura da UC, ou seja, após o aumento de carga solicitado:\n\n{categoriasstr}\n[bright_cyan]')
        
        print("")
        choosen_categoria = categorias[int(res)-1]
        self.Categoria_Ligacao = choosen_categoria
        if self.Categoria_Ligacao == categoria_correta:
            self.logger.info(f"Foi selecionada a Categoria de Ligação {choosen_categoria}, que deve ser atendida com Ramal de Ligação {self.dict_padrao_ramal[self.Categoria_Ligacao]} conforme a GED 4319...")
        else:
            self.logger.error(f"ERRO! A demanda aparente solicitada ({locale.format_string('%.2f kVA', self.kva0)}) é incompatível com a categoria de ligação selecionada ({choosen_categoria})!")
            self.logger.error(f"Inicie novamente o processo de importar e preparar a rede secundária extraída do GIS.")
            return 'error'
        
        fases_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()[0])
        fases_norma = int(self.dict_padrao_ramal[self.Categoria_Ligacao][0])
        
        bitola_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper().split('P')[-1].split('(')[0])
        bitola_norma = int(self.dict_padrao_ramal[self.Categoria_Ligacao].split('P')[-1].split('(')[0])
        
        self.ramal_modificado = False
        if (fases_existente < fases_norma) or (bitola_existente < bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) inferior ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}) e será substituído para o cálculo da MDD...')
            self.change_service_cable(self.dict_padrao_ramal[self.Categoria_Ligacao])
            self.ramal_modificado = True
        elif (fases_existente > fases_norma) or (bitola_existente > bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente na base GIS ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) é superior ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}): será mantido ramal da base GIS para o cálculo da MDD...')
        else:
            self.logger.info(f'Ramal de Ligação existente igual ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}) e será mantido para o cálculo da MDD...')
        
        return 'ok'
    
    def avalia_condicoes_rede(self):
    
        self.tensao_secundaria = self.dss.Transformer[0].kVs[-1]
        
        self.delta_wye = False
        if (self.dss.Transformer[0].Conns[0].name == 'delta') & (self.dss.Transformer[0].Conns[1].name == 'wye'):
            self.delta_wye = True
        
        temp = dict(zip(['Fase A', 'Fase B', 'Fase C', 'Neutro'], (100*np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1])/(self.dss.Transformer[0].kVAs[0]/(np.sqrt(3)*self.dss.Transformer[0].kVs[-1]))).tolist()))
        
        self.carga_fase_A = temp['Fase A']
        self.carga_fase_B = temp['Fase B']
        self.carga_fase_C = temp['Fase C']
        self.corrente_neutro = temp['Neutro']
        
        self.logger.info("Carregamento do Transformador...")
        self.logger.info(f"...Fase A: {locale.format_string('%.2f', self.carga_fase_A)}%")
        self.logger.info(f"...Fase B: {locale.format_string('%.2f', self.carga_fase_B)}%")
        self.logger.info(f"...Fase C: {locale.format_string('%.2f', self.carga_fase_C)}%")
        self.logger.info(f"Corrente de Neutro: {locale.format_string('%.2f', self.corrente_neutro)}% da corrente nominal do transformador")
        
        amps = np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1][0:3])
        self.ids = np.max(amps)/np.mean(amps)
        self.idi = np.min(amps)/np.mean(amps)
        
        self.logger.info(f"Índice de Desbalanceamento Superior: {locale.format_string('%.3f', self.ids)}...")
        self.logger.info(f"Índice de Desbalanceamento Inferior: {locale.format_string('%.3f', self.idi)}...")
        
        if (self.ids > 1.3) | (self.idi < 0.7):
            self.logger.warning("ATENÇÃO! Foi encontrado Índice de Desbalanceamento fora da faixa na qual o transformador é considerado minimamente balanceado (0.7 a 1.3, conforme GED 13285)!")
            self.logger.warning("O estudo prosseguirá normalmente, porém recomenda-se avaliar possível rebalanceamento da rede secundária.")
        
        desequilibrio_tensao = pd.DataFrame([(bus.Name, bus.NumNodes, 100*bus.SeqVoltages[0]/bus.SeqVoltages[1], 100*bus.SeqVoltages[2]/bus.SeqVoltages[1]) for bus in self.dss.Bus], columns=['Barra', 'NumNodes', 'CUF0', 'CUF2']).set_index('Barra')
        temp = desequilibrio_tensao[desequilibrio_tensao['NumNodes'] == 3]
        max_volt_unb_0 = temp['CUF0'].max()
        max_volt_unb_2 = temp['CUF2'].max()
        #if (max_volt_unb_0 > 3) | (max_volt_unb_2 > 3):
        if (max_volt_unb_2 > 3):
            b = self.dss.Bus[temp['CUF2'].idxmax()]
            self.logger.warning("ATENÇÃO! Foi encontrado desequilíbrio de tensão superior a 3% na rede secundária em estudo.")
            self.logger.warning(f"...Tensões no nó {temp['CUF2'].idxmax()}...")
            self.logger.warning(f"......Fase A: {np.round(b.VMagAngle[0],3)}∠{np.round(b.VMagAngle[1],3)}° Volts")
            self.logger.warning(f"......Fase B: {np.round(b.VMagAngle[2],3)}∠{np.round(b.VMagAngle[3],3)}° Volts")
            self.logger.warning(f"......Fase C: {np.round(b.VMagAngle[4],3)}∠{np.round(b.VMagAngle[5],3)}° Volts")
            # if temp['CUF0'].idxmax() != temp['CUF2'].idxmax():
                # b = self.dss.Bus[temp['CUF0'].idxmax()]
                # self.logger.warning(f"...Tensões no nó {temp['CUF0'].idxmax()}...")
                # self.logger.warning(f"......Fase A: {np.round(b.VMagAngle[0],3)}∠{np.round(b.VMagAngle[1],3)}° Volts")
                # self.logger.warning(f"......Fase B: {np.round(b.VMagAngle[2],3)}∠{np.round(b.VMagAngle[3],3)}° Volts")
                # self.logger.warning(f"......Fase C: {np.round(b.VMagAngle[4],3)}∠{np.round(b.VMagAngle[5],3)}° Volts")    
            self.logger.warning("O estudo prosseguirá normalmente, porém recomenda-se avaliar possível rebalanceamento da rede secundária.")
    
    def get_interest_bus(self):
    
        self.kiter -= 1
        self.simulate()
        
        self.avalia_condicoes_rede()
        
        #######################################################################
        
        self.logger.info("Identificando a Carga de Estudo (CE)...")
        
        self.interest_loads = [item for item in self.dss.Load.Name if 'ce' in item]
        
        if len(self.interest_loads) == 0:
            self.logger.error("ATENÇÃO! Verificar a inclusão da carga de estudo no OpenDSS.")
            return 'error'
        #elif len(self.interest_loads) > 1:
            #self.logger.error("ATENÇÃO! Verificar múltiplas cargas de estudo no OpenDSS.")
        else:
            self.logger.info("Carga de Estudo (CE) identificada com sucesso!")
            temp = ', '.join(self.interest_loads)
            self.logger.info(temp)
        
        self.interest_bus = [self.dss.Load[item].Bus1 for item in self.interest_loads]
        
        if len(list(set([item.split('.')[0] for item in self.interest_bus]))) > 1:
            self.logger.error("ERRO! O arquivo carregado contém mais de uma carga de estudo! Por favor verifique o estudo no GIS e tente novamente.")
            return 'error'
        
        #######################################################################
        
        self.kW0 = [self.dss.Load[item].kW for item in self.dss.Load.Name if 'ce' in item]
        self.DTS = np.sum(self.kW0)
        self.logger.info(f"Identificada a Demanda Total Solicitada (DTS) pelo consumidor: {locale.format_string('%.2f kW', self.DTS)}...")
        self.l_array.append(np.sum(self.kW0))
        
        self.kva0 = np.sum([np.sqrt(self.dss.Load[item].kW**2 + self.dss.Load[item].kvar**2) for item in self.dss.Load.Name if 'ce' in item])
        
        #######################################################################
        
        self.build_graph()
        
        self.logger.info("Identificando o Ramal de Ligação/Conexão da Carga de Estudo...")
        self.ramal_ligacao = self.critical_path[-1]['Name']
        self.ramal_existente = self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()
        self.logger.info(f"Ramal de Ligação/Conexão da Carga de Estudo identificado com sucesso! Arranjo: {self.ramal_existente}")
        
        #######################################################################
        
        self.Categoria_Ligacao = 'N/A'
        
        if self.tipo_atividade in ['Reforma e Adequação - Aum. de Carga Edif', 'Ligação Nova Edificio - Coletivo', 'Ligação Nova BT -  Medição Agrupada']:
            print("")
            res = Prompt.ask(f' [bright_cyan]O ramal de ligação/conexão será: [green][A] AÉREO [bright_cyan]ou [red][S] SUBTERRÂNEO[bright_cyan]')
            loop = True
            while loop:
                if res in ['a', 'A']:
                    loop = False
                    self.ramal_aereo_subterraneo = 'Aéreo'
                    self.avalia_ramal_coletivo()
                elif res in ['s', 'A']:
                    loop = False
                    self.ramal_aereo_subterraneo = 'Subterrâneo'
                    self.avalia_ramal_subterraneo()
                else:
                    print("")
                    res = Prompt.ask(f' [red]Opção Inválida! [bright_cyan]O ramal de ligação/conexão será: [green][A] AÉREO [bright_cyan]ou [red][S] SUBTERRÂNEO[bright_cyan]')
        elif self.tipo_atividade in ['Ligação Nova BT - Entrada Subterranea']:
            self.ramal_aereo_subterraneo = 'Subterrâneo'
            self.avalia_ramal_subterraneo()
        else:
            self.ramal_aereo_subterraneo = 'Aéreo'
            ret = self.avalia_ramal_individual()
            if ret == 'error':
                return 'error'
        
        #######################################################################
        
        return 'ok'
    
    def get_violations(self):
    
        volts = np.concatenate([self.dss.NodeVMagPUByPhase(1), self.dss.NodeVMagPUByPhase(2), self.dss.NodeVMagPUByPhase(3)])
        temp = list(self.dss.NodeNamesByPhase(1)) + list(self.dss.NodeNamesByPhase(2)) + list(self.dss.NodeNamesByPhase(3))
        bus = [item.split('.')[0] for item in temp]
        phase = [item.split('.')[1] for item in temp]
        
        v = pd.DataFrame(zip(bus, phase, volts), columns=['Barra', 'Fase', 'V (pu)']).set_index(['Barra', 'Fase']).unstack()['V (pu)']
        
        v_point = v.min().min()
        self.v_array.append(v_point)
        
        self.V_ARRAY.append(v.unstack())
        
        ########################
        
        amps = []
        ampsph = []
        #for line in self.dss.Line.Name:
        for line in self.critical_path:
            amps.append((line['Name'], np.max(np.abs(self.dss.Line[line['Name']].Currents()))/self.dss.Line[line['Name']].NormAmps))
            ampsph.append((line['Name'], np.abs(self.dss.Line[line['Name']].Currents()).reshape(2,-1)[0, :]/self.dss.Line[line['Name']].NormAmps))
        
        amps = pd.DataFrame(amps, columns=['Linha', 'Carregamento (pu)']).set_index('Linha')
        ampsph = pd.DataFrame(ampsph, columns=['Linha', 'Carregamento (pu)']).set_index('Linha')
        ampsph = pd.DataFrame(ampsph['Carregamento (pu)'].to_list(), index=ampsph.index, columns=['1', '2', '3'])
        
        i_point = amps['Carregamento (pu)'].max()
        self.i_array.append(i_point)
        
        self.I_ARRAY.append(ampsph.unstack())
        
        ########################
        
        # POTENCIA ATIVA
        #s_point = np.abs(self.dss.TotalPower())/self.dss.Transformer[0].kVAs[0]
        #self.s_array.append(s_point)
        
        # CORRENTE
        s_point = np.max(np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1]))/(self.dss.Transformer[0].kVAs[0]/(np.sqrt(3)*self.dss.Transformer[0].kVs[-1]))
        self.s_array.append(s_point)
        
        self.S_ARRAY.append(np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1])/(self.dss.Transformer[0].kVAs[0]/(np.sqrt(3)*self.dss.Transformer[0].kVs[-1])))
        
        ########################
        
        self.voltage_violation = True if v_point < VMIN else False
        self.current_violation = True if i_point > IMAX else False
        self.transformer_violation = True if s_point > SMAX else False
        
        self.violation = self.voltage_violation | self.current_violation | self.transformer_violation
        
        if self.violation:
            
            self.fator_limitante = []
            
            if self.voltage_violation:
                self.fator_limitante.append("Tensão")
            if self.current_violation:
                self.fator_limitante.append("Carregamento de Vão")
            if self.transformer_violation:
                self.fator_limitante.append("Carregamento do Transformador")
                
            if len(self.fator_limitante) == 1:
                self.fator_limitante = self.fator_limitante[0]
            elif len(self.fator_limitante) > 1:
                self.fator_limitante = '; '.join(self.fator_limitante)
            
            self.volts = v
            self.amps = amps
    
    def simulate(self):
    
        self.dss.Solution.Mode = SolveModes.SnapShot
        self.dss.Solution.ControlMode = ControlModes.Static
        self.dss.Solution.Solve()
        self.kiter += 1
    
    def increase_load(self):
    
        l_point = 0
        
        for load in self.interest_loads:
            
            self.dss.Load[load].kW += 1
            l_point += self.dss.Load[load].kW
            
        self.l_array.append(l_point)
    
    def set_load(self, M):
    
        for load in self.interest_loads:
            
            self.dss.Load[load].kW = M/len(self.interest_loads)
            
        self.l_array.append(M)
    
    def binary_search(self, LIM_INF, LIM_SUP):
    
        while (LIM_SUP - LIM_INF) > 1e-3:
            
            M = (LIM_INF + LIM_SUP)/2
            
            self.set_load(M)
            self.simulate()
            self.get_violations()
            
            if self.violation:
                LIM_SUP = M
            else:
                LIM_INF = M
        
        return (LIM_INF + LIM_SUP)/2
    
    def find_mdd(self):
    
        self.get_violations()
        
        if self.violation:
            temp = []
            if self.voltage_violation:
                temp.append("Tensão")
            if self.current_violation:
                temp.append("Carregamento de Vão")
            if self.transformer_violation:
                temp.append("Carregamento do Transformador")
            self.logger.error(f"ERRO! O carregamento inicial da rede secundária já apresenta as violações a seguir: {temp}")
            return 'error'
        
        while not self.violation:
            
            self.increase_load()
            self.simulate()
            self.get_violations()
        
        LIM_INF = self.l_array[-2]
        LIM_SUP = self.l_array[-1]
        
        self.MDD = self.binary_search(LIM_INF, LIM_SUP)
        
        return 'ok'
    
    def plot_increase(self):
    
        TEMPV = pd.concat(self.V_ARRAY, axis=1)
        TEMPV.columns = self.l_array
        TEMPV = 100 * (TEMPV - 1)
        TEMPV[TEMPV > 0] = 0
        
        tempV = pd.Series(self.v_array, index=self.l_array).sort_index()
        tempI = 100*pd.Series(self.i_array, index=self.l_array).sort_index()
        tempS = 100*pd.Series(self.s_array, index=self.l_array).sort_index()
        
        #tempV = pd.concat([tempV, pd.Series([np.nan], index=[self.MDD])]).sort_index().interpolate()
        #tempI = pd.concat([tempI, pd.Series([np.nan], index=[self.MDD])]).sort_index().interpolate()
        
        self.fig1 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig1.gca()
        #tempV.plot(ax=ax, marker='.', ls='-', color='C0')
        #tempV.plot(ax=ax, ls='-', color='C0')
        TEMPV.T['1'].plot(ax=ax, ls='-', color='k', legend=False)
        TEMPV.T['2'].plot(ax=ax, ls='-', color='b', legend=False)
        TEMPV.T['3'].plot(ax=ax, ls='-', color='darkgray', legend=False)
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Demanda do Ponto de Conexão (kW)')
        ax.set_ylabel('Queda de Tensão nos Pontos de Entrega (%)')
        ax.fill_between(xlims, -100*(9.99-1), 100*(VMIN-1), color='mistyrose')
        ax.fill_between(xlims, 100*(VMIN-1), 9e9, color='honeydew')
        #ax.axhline(y=VMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], 100*(self.v_array[-1]-1)], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], self.v_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)
        ax.set_title("")
        
        TEMPV = pd.concat(self.V_ARRAY, axis=1)
        TEMPV.columns = self.l_array
        TEMPV = 100 * (TEMPV - 1)
        TEMPV[TEMPV < 0] = 0
        
        self.fig2 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig2.gca()
        TEMPV.T['1'].plot(ax=ax, ls='-', color='k', legend=False)
        TEMPV.T['2'].plot(ax=ax, ls='-', color='b', legend=False)
        TEMPV.T['3'].plot(ax=ax, ls='-', color='darkgray', legend=False)
        ylims = ax.get_ylim()
        ax.set_xlabel('Demanda do Ponto de Conexão (kW)')
        ax.set_ylabel('Elevação de Tensão nos Pontos de Entrega (%)')
        ax.fill_between(xlims, -100*(9.99-1), 100*(VMIN-1), color='mistyrose')
        ax.fill_between(xlims, 100*(VMIN-1), 9e9, color='honeydew')
        #ax.axhline(y=VMIN, color='r', ls='--', lw=1)
        kkk = 100*(self.v_array[-1]-1)
        if kkk < 0:
            kkk = 0
        ax.plot([self.MDD, self.MDD], [ylims[0], kkk], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], self.v_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)
        ax.set_title("")
        
        TEMPI = pd.concat(self.I_ARRAY, axis=1)
        TEMPI.columns = self.l_array
        TEMPI = 100 * (TEMPI)
        
        self.fig3 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig3.gca()
        #tempI.plot(ax=ax, marker='.', ls='-', color='C0')
        #tempI.plot(ax=ax, ls='-', color='C0')
        TEMPI.T['1'].plot(ax=ax, ls='-', color='k', legend=False)
        TEMPI.T['2'].plot(ax=ax, ls='-', color='b', legend=False)
        TEMPI.T['3'].plot(ax=ax, ls='-', color='darkgray', legend=False)
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Demanda do Ponto de Conexão (kW)')
        ax.set_ylabel('Carregamento dos Segmentos (%)')
        ax.fill_between(xlims, 100*IMAX, 999.00, color='mistyrose')
        ax.fill_between(xlims, 0, 100*IMAX, color='honeydew')
        #ax.axhline(y=100*IMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], 100*self.i_array[-1]], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], 100*self.i_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)
        ax.set_title("")
        
        TEMPS = 100*pd.DataFrame(self.S_ARRAY, columns=['1', '2', '3', '4'], index=self.l_array).sort_index()
        
        self.fig4 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig4.gca()
        #tempS.plot(ax=ax, marker='.', ls='-', color='C0')
        #tempS.plot(ax=ax, ls='-', color='C0')
        TEMPS['1'].plot(ax=ax, ls='-', color='k', legend=False)
        TEMPS['2'].plot(ax=ax, ls='-', color='b', legend=False)
        TEMPS['3'].plot(ax=ax, ls='-', color='darkgray', legend=False)
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Demanda do Ponto de Conexão (kW)')
        ax.set_ylabel('Carregamento do Transformador (%)')
        ax.fill_between(xlims, 100*SMAX, 999.00, color='mistyrose')
        ax.fill_between(xlims, 0, 100*SMAX, color='honeydew')
        #ax.axhline(y=100*IMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], 100*self.s_array[-1]], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], 100*self.s_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)   
        ax.set_title("")
    
    def plot_save(self):
    
        self.fig1.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_1.png', format='png', dpi=300)
        
        self.fig3.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_3.png', format='png', dpi=300)
        self.fig4.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_4.png', format='png', dpi=300)
    
    def prepare_form_filling(self):
    
        # temp_pf = list(set([self.dss.Load[item].PF for item in self.interest_loads]))
        # if len(temp_pf) == 1:
        #     fp_str = locale.format_string('%.2f', temp_pf[0])
        # else:
        #     self.logger.warning("ATENÇÃO! Verificar no arquivo .dss o fator de potência da carga de estudo.")
        #     fp_str = '0,92'
        
        #self.PFC = (self.CTO - self.CRC) - self.ERD
        
        variables = {
            "DATA_EXT_FULL" : self.today,
            
            "NOTA_SERV" : self.nota,
            "DISCO" : self.disco,
            "TIP_ATV" : self.tipo_atividade,
            "V_PRI" : int(1e3*float(self.dss.Transformer[0].kVs[0])),
            "V_SEC" : int(1e3*float(self.dss.Transformer[0].kVs[1])),
            
            "DEM_ATUAL" : locale.format_string('%.2f', self.DE),
            "DEM_SOLICIT" : locale.format_string('%.2f', self.DTS),
            "DEM_AUMENTO" : locale.format_string('%.2f', self.DTS-self.DE),
            "CAT_LIGACAO" : self.Categoria_Ligacao,
            
            "MDD" : locale.format_string('%.2f', self.MDD),
            "QTOBS" : locale.format_string('%.2f', 100*(1-self.volts.min().min())),
            "ETOBS" : locale.format_string('%.2f', np.abs(100*(self.volts.max().max()-1))),
            "CARVAOOBS" : locale.format_string('%.2f', 100*self.amps.max().max()),
            "CARTROBS" : locale.format_string('%.2f', 100*self.s_array[-1]),
            "FAT_LIM" : self.fator_limitante,
            #"FP" : fp_str,
            "PROPORC" : locale.format_string('%.2f', (100 * self.K_prop)),
            
            "OBS" : f"Cálculo realizado pelo usuário {os.environ['USERNAME']} utilizando a ferramenta CMD - Cálculo de Máxima Demanda, Versão {VERSION} empregando o método de simulação de fluxo de potência para cálculo da máxima demanda disponibilizada."
            
            # "CRGVIG" : locale.format_string('%.2f', self.DE),
            # "CRGDEC" : locale.format_string('%.2f', self.DTS),
            # "AUMCRG" : locale.format_string('%.2f', self.DTS-self.DE),
            # "FDT" : locale.format_string('%.2f', self.FD),
            # "DEMERD" : locale.format_string('%.2f', self.Demanda_ERD),
            # "GRUTAR" : self.grupo_tar,
            # "CTO" : locale.currency(self.CTO),
            # "CTOP" : locale.currency(self.CTOp),
            # "ERD" : locale.currency(self.ERD),
            # "CRC" : locale.currency(self.CRC),
            # "PFC" : locale.currency(self.PFC)
            }
        
        return variables
    
    def make_report(self, variables, input_file, output_file):
    
        input_file_path = rf'{input_file}'
        output_file_path = rf'{output_file}'
        
        template_document = DocxTemplate(input_file_path)
        template_document.render(variables, autoescape=True)
        template_document.save(output_file_path)
    
    def run_mdd(self):
    
        if self.novo_trafo:
            
            self.logger.info(f"Iniciando o cálculo da MDD para ligação nova com transformador exclusivo...")
            self.MDD = self.novo_tr_kva
            self.logger.info(f"A máxima demanda disponibilizada (MDD) no ponto de conexão (PC) é de {locale.format_string('%.2f kW', self.MDD)} (i.e., a capacidade do novo transformador que será instalado).")
            
            self.logger.info(f'O fator limitante da MDD é: "Carregamento do Transformador"')
            
            self.logger.info(f"Iniciando o cálculo do Índice de Proporcionalização da Obra (K)...")
            self.K_prop = (self.DTS - self.DE) / (self.MDD - self.DE)
            self.logger.info(f"O Índice de Proporcionalização da Obra (K) é de {locale.format_string('%.2f', 100*self.K_prop)}%...")
            
            self.create_memoria_calculo()
            
            self.logger.info("Máxima Demanda Disponibilizada (MDD) no Ponto de Conexão e Índice de Proporcionalização da Obra (K) calculados com sucesso!")
            self.logger.info(f"...MDD = {locale.format_string('%.2f kW.', self.MDD)}")
            self.logger.info(f"...K = {locale.format_string('%.2f', 100*self.K_prop)}%")
            
        else:
            if self.dssfile is None:
                self.logger.warning("ATENÇÃO! SELECIONE O ARQUIVO DSS ANTES DE REALIZAR O CÁLCULO DA MDD!")
                return    
            
            self.logger.info(f"Iniciando o cálculo da MDD via simulação de fluxo de potência...")
            
            self.logger.info("Calculando a MDD...")
            ret = self.find_mdd()
            if ret == 'error':
                return
                
            self.logger.info("MDD calculada com sucesso!")
            
            try:
                self.logger.info(f"A máxima demanda disponibilizada (MDD) no ponto de conexão (PC) da UC {self.uc} é de {locale.format_string('%.2f kW.', self.MDD)}")
            except:
                self.logger.info(f"A máxima demanda disponibilizada (MDD) no ponto de conexão (PC) é de {locale.format_string('%.2f kW.', self.MDD)}")
            
            self.queda_tensao = 100*(1-self.volts.min().min())
            self.elevacao_tensao = np.abs(100*(self.volts.max().max()-1))
            self.carregamento_max_vao = 100*self.amps.max().max()
            self.carregamento_max_trafo = 100*self.s_array[-1]
            
            self.logger.info(f"Foram realizadas {self.kiter} simulações de fluxo de potência até a obtenção da MDD.")
            self.logger.info(f"Valores observados:")
            self.logger.info(f"...Subtensão (Queda de Tensão Máxima): {locale.format_string('%.2f', self.queda_tensao)}% (Limite Considerado: {locale.format_string('%.2f', 100-100*VMIN)}%);")
            self.logger.info(f"...Sobretensão (Elevação de Tensão Máxima): {locale.format_string('%.2f', self.elevacao_tensao)}% (Limite Considerado: {locale.format_string('%.2f', 100*VMAX-100)}%);")
            self.logger.info(f"...Pior Carregamento (Máximo Carregamento de Vão): {locale.format_string('%.2f', self.carregamento_max_vao)}% (Limite Considerado: {locale.format_string('%.2f', 100*IMAX)}%);")
            self.logger.info(f"...Carregamento do Transformador: {locale.format_string('%.2f', self.carregamento_max_trafo)}% (Limite Considerado: {locale.format_string('%.2f', 100*SMAX)}%);")
            # self.logger.info(f"{self.l_array}")
            self.logger.info(f"O fator limitante da MDD é: {self.fator_limitante}")
            
            self.logger.info(f"Iniciando o cálculo do Índice de Proporcionalização da Obra (K)...")
            self.K_prop = (self.DTS - self.DE) / (self.MDD - self.DE)
            self.logger.info(f"O Índice de Proporcionalização da Obra (K) é de {locale.format_string('%.2f', 100*self.K_prop)}%...")
            
            self.logger.info("Elaborando as figuras com os resultados das simulações (memória de cálculo)...")
            self.plot_increase()
            self.logger.info("Figuras elaboradas com sucesso!")
            self.logger.info("Gravando as figuras...")
            self.plot_save()
            self.logger.info("Figuras salvas com sucesso!")
            
            # print("")
            # res = Prompt.ask(f' [bright_cyan]Deseja visualizar as figuras elaboradas? [green][S] SIM [bright_cyan]ou [red][N] NÃO[bright_cyan]')
            
            # loop = True
            
            # while loop:
            
            #     if res in ['s', 'S']:
            #         loop = False
            #         plt.show()
            #     elif res in ['n', 'N']:
            #         loop = False
            #         plt.close('all')
            #     else:
            #         res = Prompt.ask(f' [red]Opção Inválida! [bright_cyan]Deseja visualizar as figuras elaboradas? [green][S] SIM [bright_cyan]ou [red][N] NÃO[bright_cyan]')
            
            # print("")
            
            self.create_memoria_calculo()
            
            self.logger.info("Máxima Demanda Disponibilizada (MDD) no Ponto de Conexão e Índice de Proporcionalização da Obra (K) calculados com sucesso!")
            self.logger.info(f"...MDD = {locale.format_string('%.2f kW.', self.MDD)}")
            self.logger.info(f"...K = {locale.format_string('%.2f', 100*self.K_prop)}%")
    
    def run_pfc_erd(self):
    
        ##### Inputs
        
        # self.K_ERD - Planilha REHs
        # self.CTO = self.CME + self.CMNE + self.CMO - CONFIRMAR, Fonte
        
        # self.Demanda_ERD vs. self.Demanda_Considerada
        
        self.Carga_Instalada_input = Prompt.ask(' [bright_cyan]Informe a Carga Instalada Vigente (antes do aumento), em kW')
        self.Carga_Instalada_Atual = float(self.Carga_Instalada_input)
        
        self.Carga_Futura_input = Prompt.ask(' [bright_cyan]Informe a Carga Instalada Futura (após o aumento), em kW')
        self.Carga_Instalada_Futura = float(self.Carga_Futura_input)
        
        self.Demanda_Vigente_input = Prompt.ask(' [bright_cyan]Informe a Demanda Existente (DE), em kW')
        self.DE = float(self.Demanda_Vigente_input)
        
        self.Demanda_Solicitada_input = Prompt.ask(' [bright_cyan]Informe a Demanda Total Solicitada (DTS), em kW')
        self.DTS = float(self.Demanda_Solicitada_input)
        
        print("")
        self.logger.info(f"O aumento de demanda considerado é de {locale.format_string('%.2f kW', self.DTS-self.DE)}...")
        
        print("")
        self.CTO_input = Prompt.ask(' [bright_cyan]Informe o Custo Total da Obra')
        self.CTO = float(self.CTO_input)
        
        print("")
        self.logger.info(f"Realizando os cálculos de proporcionalização e demais custos da obra (distribuidora: {self.empresa})...")
        
        self.TABELA_K = pd.read_excel(rf"{self.BASE_FOLDER}\Fator_K_ERD.xlsx", sheet_name=f"{self.empresa}")
        
        temp = self.TABELA_K.set_index('SUBGRUPO TARIFÁRIO').loc['FATOR DE CÁLCULO DO ERD (K)'].dropna()
        del temp['REH']
        del temp['Data']
        
        grupos = temp.index.tolist()
        
        gcount = 1
        gruposstr = ''
        for grupo in grupos:
            gruposstr = gruposstr + f'[bold green] [{gcount}] [bold yellow]{grupo}\n'
            gcount += 1
        
        print("")
        res = Prompt.ask(f' [bright_cyan]Selecione o grupo tarifário da Unidade Consumidora (UC):\n\n{gruposstr}[bright_cyan]\n')
        
        print("")
        choosen_grupo = grupos[int(res)-1]
        self.grupo_tar = choosen_grupo
        self.K_ERD = self.TABELA_K.set_index('SUBGRUPO TARIFÁRIO').loc['FATOR DE CÁLCULO DO ERD (K)', choosen_grupo]
        self.logger.info(f"Foi selecionado o grupo tarifário {choosen_grupo} - fator de cálculo do ERD (art. 109 da REN nº 1.000/2021) k = {locale.format_string('%.2f R$/kW.', self.K_ERD)} ({self.disco})...")
        
        atvcount = 1
        atividadesstr = ''
        for atividade in self.atividades:
            atividadesstr = atividadesstr + f'[bold green] [{atvcount}] [bold yellow]{atividade}\n'
            atvcount += 1
        
        print("")
        res= Prompt.ask(f' [bright_cyan]Selecione o tipo de atividade da Unidade Consumidora (UC):\n\n{atividadesstr}[bright_cyan]\n')
        
        self.FD = dict(zip(self.atividades, self.FDs))[self.atividades[int(res)-1]]
        self.Demanda_ERD = (self.Carga_Instalada_Futura - self.Carga_Instalada_Atual) * self.FD
        
        print("")
        self.logger.info(f"Foi selecionada a atividade {self.atividades[int(res)-1]} - fator de demanda típico (GED 3793) FD = {locale.format_string('%.2f', self.FD)}...")
        
        self.K_prop = (self.DTS - self.DE) / (self.MDD - self.DE)
        
        self.CTOp = self.K_prop * self.CTO #(self.CME + self.CMNE + self.CMO)
        
        self.ERD = self.Demanda_ERD * self.K_ERD
        
        if self.CTOp > self.ERD:
            self.PFC = self.CTOp - self.ERD
        else:
            self.PFC = 0
        
        self.CRC = self.CTO - self.CTOp
        
        print("")
        self.logger.info("O resultado dos cálculos é:")
        self.logger.info(f"...Demanda Existente (DE): {locale.format_string('%.2f kW', self.DE)}...")
        self.logger.info(f"...Demanda Total Solicitada (DTS): {locale.format_string('%.2f kW', self.DTS)}...")
        self.logger.info(f"...Máxima Demanda Disponibilizada (MDD): {locale.format_string('%.2f kW', self.MDD)}...")
        self.logger.info(f"...Proporcionalidade da Obra (K): {locale.format_string('%.2f', 100*self.K_prop)}%...")
        self.logger.info(f"...Custo Total da Obra (CTO): {locale.currency(self.CTO)}...")
        self.logger.info(f"...Custo Total da Obra Proporcionalizado (CTOp): {locale.currency(self.CTOp)}...")
        self.logger.info(f"...Encargo sob Responsabilidade da Distribuidora (ERD): {locale.currency(self.ERD)}...")
        self.logger.info(f"...Custo de Reserva de Capacidade (CRC): {locale.currency(self.CRC)}...")
        self.logger.info(f"...Participação Financeira do Consumidor (PFC): {locale.currency(self.PFC)}...")
        
        print("")
        self.logger.info("Participação Financeira do Consumidor (PFC) e demais custos calculados com sucesso!")
    
    def zip_directory(self, folder_path, zip_path):
    
        with zipfile.ZipFile(zip_path, mode='w') as zipf:
            len_dir_path = len(folder_path)
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, file_path[len_dir_path:])
    
    def create_memoria_calculo(self):
    
        self.logger.info("Emitindo a Memória de Cálculo...")
        variables = self.prepare_form_filling()
        docx_outp_str = rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}_1.docx"
        self.make_report(variables, rf"{self.BASE_FOLDER}\TEMPLATE_MEMORIA_CALCULO.docx", docx_outp_str)
        
        self.logger.info('Iniciando a substituição das figuras...')
        
        src_file = rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}_1.docx"
        out_dir = rf"{self.BASE_FOLDER}\temporary\run"
        des_file = rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}.docx"
        
        with zipfile.ZipFile(src_file, 'r') as zip_ref:
            zip_ref.extractall(out_dir)
        
        imgrep_path_1 = rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_1.png'
        # imgrep_path_2 = rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_2.png'
        imgrep_path_3 = rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_3.png'
        imgrep_path_4 = rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_4.png'
        
        media_path = os.path.join(out_dir, "word", "media")
        img_file_1 = os.path.join(media_path, "image5.png")
        # img_file_2 = os.path.join(media_path, "image6.png")
        img_file_3 = os.path.join(media_path, "image7.png")
        img_file_4 = os.path.join(media_path, "image8.png")
        
        os.remove(img_file_1)
        # os.remove(img_file_2)
        os.remove(img_file_3)
        os.remove(img_file_4)
        
        shutil.copy(imgrep_path_1, img_file_1)
        # shutil.copy(imgrep_path_2, img_file_2)
        shutil.copy(imgrep_path_3, img_file_3)
        shutil.copy(imgrep_path_4, img_file_4)
        
        self.zip_directory(out_dir, des_file)
        
        os.remove(rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}_1.docx")
        
        # os.remove(imgrep_path_1)
        # os.remove(imgrep_path_2)
        # os.remove(imgrep_path_3)
        # os.remove(imgrep_path_4)
        
        shutil.rmtree(rf'{self.BASE_FOLDER}\temporary')
        
        self.logger.info('Figuras substituídas!')
        
        word_app = client.DispatchEx("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(f"{des_file.replace('/', '\\')}")
        
        if self.disco == 'CPFL Paulista':
            logos_apagar = ['cpfl_pira', 'cpfl_scj', 'cpfl_rge']
        elif self.disco == 'CPFL Piratininga':
            logos_apagar = ['cpfl_psta', 'cpfl_scj', 'cpfl_rge']
        elif self.disco == 'CPFL Santa Cruz':
            logos_apagar = ['cpfl_psta', 'cpfl_pira', 'cpfl_rge']
        elif self.disco == 'CPFL RGE':
            logos_apagar = ['cpfl_psta', 'cpfl_pira', 'cpfl_scj']
        
        # Frente
        k_logo = 4
        while k_logo > 1:
            for shape in doc.Shapes:
                if shape.AlternativeText in logos_apagar:
                    shape.Delete()
                    k_logo -= 1
        
        doc.Fields.Update()
        doc.Save()
        doc.SaveAs(rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}.pdf", FileFormat=17)
        doc.Close()
        word_app.Quit()
        
        self.logger.info("Memória de Cálculo emitida com sucesso!")
