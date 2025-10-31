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
import locale
import logging
#import oracledb
import cx_Oracle
import numpy as np
import pandas as pd
import datetime as dt
import networkx as nx
from altdss import altdss
from rich.prompt import Prompt
from docxtpl import DocxTemplate
from rich.console import Console
from rich.logging import RichHandler
from matplotlib import pyplot as plt
from dss import SolveModes, ControlModes

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
        self.dict_padrao_ramal = {
            'A1' : '1P10(A10)',
            'A2' : '1P16(A16)',
            'A3' : '1P10(A10)',
            'A4' : '1P16(A16)',
            'B1' : '2P16(A16)',
            'B2' : '2P25(A25)',
            'B3a (15 < C ≤ 20 kW)' : '2P10(A10)',
            'B3b (20 < C ≤ 25 kW)' : '2P16(A16)',
            'C1' : '3P10(A10)',
            'C2' : '3P16(A16)',
            'C3' : '3P25(A25)',
            'C4' : '3P35(A35)',
            'C5' : '3P50(A50)',
            'C6' : '3P70(A70)',
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
            '3P70(A70)' : {'nphases': 3, 'r1' : 0.5038059950, 'x1' : 0.1041790023, 'r0' : 1.3816900253, 'x0' : 0.9110010266, 'normaps' :	227}
        }
        
        self.FDs = [0.6, 0.35, 0.28, 0.3, 0.32, 0.42, 0.29, 0.27, 1, 0.38, 0.28, 0.23, 0.51, 0.2, 0.39, 0.34, 0.32, 0.53, 0.55, 0.42, 1, 0.32, 0.51, 0.2, 0.28]
        self.atividades = ['Bar', 'Beneficiamento de cereais', 'Carpintaria', 'Fábrica de bebidas', 'Fábrica de calçados', 'Fábrica de plásticos', 'Fábrica de roupas', 'Hotel', 'Iluminação (festiva, ornamental, etc.)', 'Laticínio', 'Oficina mecânica', 'Padaria', 'Posto de gasolina', 'Residência', 'Restaurante', 'Serraria', 'Semáforo', 'Sorveteria', 'Supermercado', 'Comércio, Serviço e outras atividades', 'Iluminação Pública', 'Industrial', 'Poder Público', 'Residencial', 'Rural']
        
        self.dss = altdss
        
        self.kiter = 0
        
        self.voltage_violation = False
        self.current_violation = False
        self.transformer_violation = False
        
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
        
        if self.dados_nota[0] == ' ':
            self.logger.info(f"Tipo de Atividade: {self.tipo_atividade}")
            self.logger.info(f"ATENÇÃO! UC não encontrada. Trata-se de LIGAÇÃO NOVA. Demanda Existente (DE) = 0 kW")
            self.DE = 0
        else:
            self.logger.info(f"Unidade Consumidora: {self.dados_nota[0]}, Tipo de Atividade: {self.tipo_atividade}")
            self.uc = self.dados_nota[0]
            print("")
            self.Demanda_Vigente_input = Prompt.ask(' [bright_cyan]Informe a Demanda Existente (DE), em kW')
            self.DE = float(self.Demanda_Vigente_input.replace(',', '.'))
            print("")
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
    
         files = os.listdir(rf"{self.BASE_FOLDER}\dss")
         modtime = [time.localtime(os.path.getmtime(rf"{self.BASE_FOLDER}\dss\{item}")) for item in files]
         fs = pd.Series(modtime, index=files).sort_values(ascending=False)
         fs = fs.apply(lambda x: time.strftime("%A, %d de %B de %Y %H:%M:%S", x))
         files = fs.index.to_list()[0:10]
         
         fcount = 1
         filesstr = ''
         for file in files:
             filesstr = filesstr + f'[bright_magenta] [{fcount}] [bold green]{file} [grey78]{fs.loc[file]}\n'
             fcount += 1
         
         filesstr = filesstr + f'[bright_magenta] [0] [bold green]Quero digitar o nome do arquivo...\n'
         
         res = Prompt.ask(f' [bright_cyan]Selecione o arquivo exportado do GIS-D após o projeto da obra:\n\n{filesstr}[bright_cyan]\n')
         
         if res == "0":
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
         else:
             choosen_file = files[int(res)-1]
             print("")
             
         self.dssfile = rf"{self.BASE_FOLDER}\dss\{choosen_file}"
         
         self.logger.info(f"O caminho completo do arquivo dss selecionado é: {self.dssfile}")
         
         self.dss("Clear")
         self.dss(f"Redirect {self.dssfile}")
         self.get_interest_bus()
         
         self.logger.info("Importação e preparação da rede secundária extraída do GIS finalizada com sucesso!")
    
    def change_service_cable(self, bitola):
    
        self.logger.info("Iniciando a troca do Ramal de Ligação...")
        if bitola in self.dss.LineCode.Name:
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
    
    def get_interest_bus(self):
    
        self.kiter -= 1
        self.simulate()
        
        temp = dict(zip(['Fase A', 'Fase B', 'Fase C', 'Neutro'], (100*np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1])/(self.dss.Transformer[0].kVAs[0]/(np.sqrt(3)*self.dss.Transformer[0].kVs[-1]))).tolist()))
        self.logger.info(f"Carregamento do Transformador...")
        self.logger.info(f"...Fase A: {locale.format_string('%.2f', temp['Fase A'])}%")
        self.logger.info(f"...Fase B: {locale.format_string('%.2f', temp['Fase B'])}%")
        self.logger.info(f"...Fase C: {locale.format_string('%.2f', temp['Fase C'])}%")
        self.logger.info(f"Corrente de Neutro: {locale.format_string('%.2f', temp['Neutro'])}% da corrente nominal do transformador")
        
        self.logger.info("Identificando a Carga de Estudo (CE)...")
        
        self.interest_loads = [item for item in self.dss.Load.Name if 'ce' in item]
        
        if len(self.interest_loads) == 0:
            self.logger.error("ATENÇÃO! Verificar a inclusão da carga de estudo no OpenDSS.")
        #elif len(self.interest_loads) > 1:
            #self.logger.error("ATENÇÃO! Verificar múltiplas cargas de estudo no OpenDSS.")
        else:
            self.logger.info("Carga de Estudo (CE) identificada com sucesso!")
            temp = ', '.join(self.interest_loads)
            self.logger.info(temp)
        
        self.interest_bus = [self.dss.Load[item].Bus1 for item in self.interest_loads]
        self.kW0 = [self.dss.Load[item].kW for item in self.dss.Load.Name if 'ce' in item]
        self.DTS = np.sum(self.kW0)
        self.logger.info(f"Identificada a Demanda Total Solicitada (DTS) pelo consumidor: {locale.format_string('%.2f kW', self.DTS)}...")
        self.l_array.append(np.sum(self.kW0))
        
        self.logger.info("Identificando o Ramal de Ligação/Conexão da Carga de Estudo...")
        temp = [item for item in self.dss.Line.Name if self.dss.Line[item].Bus2.split('.')[0] == self.interest_bus[0].split('.')[0]]
        if len(temp) == 1:
            self.ramal_ligacao = temp[0]
        self.logger.info(f"Ramal de Ligação/Conexão da Carga de Estudo identificado com sucesso! Arranjo: {self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}")
        
        categorias = list(self.dict_padrao_ramal.keys())
        
        ccount = 1
        categoriasstr = ''
        for categoria in categorias:
            categoriasstr = categoriasstr + f'[bright_magenta] [{ccount}] [bold green]{categoria}\n'
            ccount += 1
        
        print("")
        res = Prompt.ask(f' [bright_cyan]Selecione a "Categoria de Ligação" (GED 13) futura da UC, ou seja, após o aumento de carga solicitado:\n\n{categoriasstr}\n[bright_cyan]')
        
        print("")
        choosen_categoria = categorias[int(res)-1]
        self.Categoria_Ligacao = choosen_categoria
        self.logger.info(f"Foi selecionada a Categoria de Ligação {choosen_categoria}, que deve ser atendida com Ramal de Ligação {self.dict_padrao_ramal[self.Categoria_Ligacao]} conforme a GED 4319...")
        
        fases_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()[0])
        fases_norma = int(self.dict_padrao_ramal[self.Categoria_Ligacao][0])
        
        bitola_existente = int(self.dss.Line[self.ramal_ligacao].LineCode.Name.upper().split('P')[-1].split('(')[0])
        bitola_norma = int(self.dict_padrao_ramal[self.Categoria_Ligacao].split('P')[-1].split('(')[0])
        
        if (fases_existente < fases_norma) or (bitola_existente < bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) inferior ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}) e será substituído para o cálculo da MDD...')
            self.change_service_cable(self.dict_padrao_ramal[self.Categoria_Ligacao])
        elif (fases_existente > fases_norma) or (bitola_existente > bitola_norma):
            self.logger.warning(f'Ramal de Ligação existente na base GIS ({self.dss.Line[self.ramal_ligacao].LineCode.Name.upper()}) é superior ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}): será mantido ramal da base GIS para o cálculo da MDD...')
        else:
            self.logger.info(f'Ramal de Ligação existente igual ao especificado na GED 4319 ({self.dict_padrao_ramal[self.Categoria_Ligacao]}) e será mantido para o cálculo da MDD...')
    
    def get_violations(self):
    
        volts = np.concatenate([self.dss.NodeVMagPUByPhase(1), self.dss.NodeVMagPUByPhase(2), self.dss.NodeVMagPUByPhase(3)])
        temp = list(self.dss.NodeNamesByPhase(1)) + list(self.dss.NodeNamesByPhase(2)) + list(self.dss.NodeNamesByPhase(3))
        bus = [item.split('.')[0] for item in temp]
        phase = [item.split('.')[1] for item in temp]
        
        v = pd.DataFrame(zip(bus, phase, volts), columns=['Barra', 'Fase', 'V (pu)']).set_index(['Barra', 'Fase']).unstack()['V (pu)']
        
        v_point = v.min().min()
        self.v_array.append(v_point)
        
        ########################
        
        amps = []
        #for line in self.dss.Line.Name:
        for line in self.critical_path:
            amps.append((line['Name'], np.max(np.abs(self.dss.Line[line['Name']].Currents()))/self.dss.Line[line['Name']].NormAmps))
        
        amps = pd.DataFrame(amps, columns=['Linha', 'Carregamento (pu)']).set_index('Linha')
        
        i_point = amps['Carregamento (pu)'].max()
        self.i_array.append(i_point)
        
        ########################
        
        # POTENCIA ATIVA
        #s_point = np.abs(self.dss.TotalPower())/self.dss.Transformer[0].kVAs[0]
        #self.s_array.append(s_point)
        
        # CORRENTE
        s_point = np.max(np.abs(self.dss.Transformer.Currents().reshape(2,-1)[-1]))/(self.dss.Transformer[0].kVAs[0]/(np.sqrt(3)*self.dss.Transformer[0].kVs[-1]))
        self.s_array.append(s_point)
        
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
        
        while not self.violation:
            
            self.increase_load()
            self.simulate()
            self.get_violations()
        
        LIM_INF = self.l_array[-2]
        LIM_SUP = self.l_array[-1]
        
        self.MDD = self.binary_search(LIM_INF, LIM_SUP)
    
    def plot_increase(self):
    
        tempV = pd.Series(self.v_array, index=self.l_array).sort_index()
        tempI = 100*pd.Series(self.i_array, index=self.l_array).sort_index()
        tempS = 100*pd.Series(self.s_array, index=self.l_array).sort_index()
        
        #tempV = pd.concat([tempV, pd.Series([np.nan], index=[self.MDD])]).sort_index().interpolate()
        #tempI = pd.concat([tempI, pd.Series([np.nan], index=[self.MDD])]).sort_index().interpolate()
        
        self.fig1 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig1.gca()
        #tempV.plot(ax=ax, marker='.', ls='-', color='C0')
        tempV.plot(ax=ax, ls='-', color='C0')
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Carregamento do Ponto de Conexão (kW)')
        ax.set_ylabel('Tensão Mínima dos Pontos de Entrega (pu)')
        ax.fill_between(xlims, VMIN, 9.99, color='honeydew')
        ax.fill_between(xlims, 0, VMIN, color='mistyrose')
        #ax.axhline(y=VMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], self.v_array[-1]], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], self.v_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)
        ax.set_title("Tensão")
        
        self.fig2 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig2.gca()
        #tempI.plot(ax=ax, marker='.', ls='-', color='C0')
        tempI.plot(ax=ax, ls='-', color='C0')
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Carregamento do Ponto de Conexão (kW)')
        ax.set_ylabel('Carregamento Máximo dos Vãos (%)')
        ax.fill_between(xlims, 100*IMAX, 999.00, color='mistyrose')
        ax.fill_between(xlims, 0, 100*IMAX, color='honeydew')
        #ax.axhline(y=100*IMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], 100*self.i_array[-1]], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], 100*self.i_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)
        ax.set_title("Corrente nos Condutores")
        
        self.fig3 = plt.figure(figsize=(6.53, 3.50), constrained_layout=True)
        ax = self.fig3.gca()
        #tempS.plot(ax=ax, marker='.', ls='-', color='C0')
        tempS.plot(ax=ax, ls='-', color='C0')
        ylims = ax.get_ylim()
        xlims = ax.get_xlim()
        ax.set_xlabel('Carregamento do Ponto de Conexão (kW)')
        ax.set_ylabel('Carregamento do Transformador (%)')
        ax.fill_between(xlims, 100*SMAX, 999.00, color='mistyrose')
        ax.fill_between(xlims, 0, 100*SMAX, color='honeydew')
        #ax.axhline(y=100*IMIN, color='r', ls='--', lw=1)
        ax.plot([self.MDD, self.MDD], [ylims[0], 100*self.s_array[-1]], color='r', ls='--', lw=1, marker='x')
        #ax.plot([np.sum(self.kW0), np.sum(self.kW0)], [ylims[0], 100*self.s_array[0]], color='r', ls='--', lw=1, marker='x')
        ax.set_ylim(ylims)   
        ax.set_title("Carregamento do Transformador de Distribuição")
    
    def plot_save(self):
    
        self.fig1.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_1.png', format='png', dpi=300)
        self.fig2.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_2.png', format='png', dpi=300)
        self.fig3.savefig(rf'{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Figura_3.png', format='png', dpi=300)
    
    def prepare_form_filling(self):
    
        temp_pf = list(set([self.dss.Load[item].PF for item in self.interest_loads]))
        if len(temp_pf) == 1:
            fp_str = locale.format_string('%.2f', temp_pf[0])
        else:
            self.logger.warning("ATENÇÃO! Verificar no arquivo .dss o fator de potência da carga de estudo.")
            fp_str = '0,92'
        
        #self.PFC = (self.CTO - self.CRC) - self.ERD
        
        variables = {
            "DATA_EXT_FULL" : self.today, 
            #"NOTA_SAP_Z1" : '',
            "DISCO" : self.disco,
            #"V_PRI" : '',
            #"V_SEC" : '',
            
            #"DEM_ATUAL" : '',
            "DEM_SOLICIT" : locale.format_string('%.2f', np.sum(self.kW0)),
            #"DEM_AUMENTO" : '',
            "MDD" : locale.format_string('%.2f', self.MDD),
            "FP" : fp_str,
            "PROPORC" : locale.format_string('%.2f', (100 * self.K_prop)),
            "CRGVIG" : locale.format_string('%.2f', self.DE),
            "CRGDEC" : locale.format_string('%.2f', self.DTS),
            "AUMCRG" : locale.format_string('%.2f', self.DTS-self.DE),
            "FDT" : locale.format_string('%.2f', self.FD),
            "DEMERD" : locale.format_string('%.2f', self.Demanda_ERD),
            "GRUTAR" : self.grupo_tar,
            "CTO" : locale.currency(self.CTO),
            "CTOP" : locale.currency(self.CTOp),
            "ERD" : locale.currency(self.ERD),
            "CRC" : locale.currency(self.CRC),
            "PFC" : locale.currency(self.PFC)
            }
        
        return variables
    
    def make_report(self, variables, input_file, output_file):
    
        input_file_path = rf'{input_file}'
        output_file_path = rf'{output_file}'
        
        template_document = DocxTemplate(input_file_path)
        template_document.render(variables, autoescape=True)
        template_document.save(output_file_path)
    
    def run_mdd(self):
    
        if self.dssfile is None:
            self.logger.warning("ATENÇÃO! SELECIONE O ARQUIVO DSS ANTES DE REALIZAR O CÁLCULO DA MDD!")
            return    
        
        self.logger.info(f"Iniciando o cálculo da MDD via simulação de fluxo de potência...")
        
        self.logger.info("Calculando a MDD...")
        self.find_mdd()
        self.logger.info("MDD calculada com sucesso!")
        
        try:
            self.logger.info(f"A máxima demanda disponibilizada (MDD) no ponto de conexão (PC) da UC {self.uc} é de {locale.format_string('%.2f kW.', self.MDD)}")
        except:
            self.logger.info(f"A máxima demanda disponibilizada (MDD) no ponto de conexão (PC) é de {locale.format_string('%.2f kW.', self.MDD)}")
        
        self.logger.info(f"Foram realizadas {self.kiter} simulações de fluxo de potência até a obtenção da MDD.")
        self.logger.info(f"Valores observados:")
        self.logger.info(f"...Subtensão (Queda de Tensão Máxima): {locale.format_string('%.2f', 100*(1-self.volts.min().min()))}% (Limite Considerado: {locale.format_string('%.2f', 100-100*VMIN)}%);")
        self.logger.info(f"...Sobretensão (Elevação de Tensão Máxima): {locale.format_string('%.2f', np.abs(100*(self.volts.max().max()-1)))}% (Limite Considerado: {locale.format_string('%.2f', 100*VMAX-100)}%);")
        self.logger.info(f"...Pior Carregamento (Máximo Carregamento de Vão): {locale.format_string('%.2f', 100*self.amps.max().max())}% (Limite Considerado: {locale.format_string('%.2f', 100*IMAX)}%);")
        self.logger.info(f"...Carregamento do Transformador: {locale.format_string('%.2f', 100*self.s_array[-1])}% (Limite Considerado: {locale.format_string('%.2f', 100*SMAX)}%);")
        # self.logger.info(f"{self.l_array}")
        self.logger.info(f"O fator limitante da MDD é: {self.fator_limitante}")
        
        self.logger.info(f"Iniciando o cálculo do Fator de Proporcionalidade da Obra (K)...")
        self.K_prop = (self.DTS - self.DE) / (self.MDD - self.DE)
        self.logger.info(f"O Fator de Proporcionalidade da Obra (K) é de {locale.format_string('%.2f', 100*self.K_prop)}%...")
        
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
        self.logger.info("Máxima Demanda Disponibilizada (MDD) no Ponto de Conexão e Fator de Proporcionalidade da Obra (K) calculados com sucesso!")
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
    
    def create_memoria_calculo(self):
    
        self.logger.info("Emitindo a Memória de Cálculo...")
        variables = self.prepare_form_filling()
        self.make_report(variables, r"{self.BASE_FOLDER}\TEMPLATE_MEMORIA_CALCULO.docx", rf"{self.BASE_FOLDER}\Documentos Emitidos\{self.nota}\Memória de Cálculo {self.nota}.docx")
        self.logger.info("Memória de Cálculo emitida com sucesso!")
