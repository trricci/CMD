import os
import time
import sqlite3
import subprocess
import datetime as dt
from datetime import datetime

from CMD import VERSION, DB_REMOTO

if DB_REMOTO:
    PATH_DB = f'//pflsp-ap1/aplicativos/GIS/Cálculos/CMD/BD/log_cmd_{dt.datetime.now().strftime("%Y%m")}.db'
else:
    PATH_DB = f'C:/CMD/log_cmd_{dt.datetime.now().strftime("%Y%m")}.db'

class consultaDB():
    
    def __init__(self):
    
        self.conn = None
        self.cursor = None
    
    def connect_db(self):
    
        self.conn = sqlite3.connect(PATH_DB)
        self.cursor = self.conn.cursor()
    
    def close_db(self):
    
        if self.conn:
            self.conn.close()
    
    def create_database(self):
    
        # conn = sqlite3.connect(PATH_DB)
        # cursor = conn.cursor()
        
        # self.cursor.executescript('''
        #     CREATE TABLE IF NOT EXISTS ULTIMO_ESTUDO(
        #         CONEXAO                                 INTEGER,
        #         ESTIMADO                                INTEGER
        #     );
        # 
        #     CREATE INDEX IDX_ULTIMO_ESTUDO_CONEXAO ON ULTIMO_ESTUDO (CONEXAO);
        #     CREATE INDEX IDX_ULTIMO_ESTUDO_ESTIMADO ON ULTIMO_ESTUDO (ESTIMADO);
        # ''')
        # self.conn.commit()
        
        self.cursor.executescript('''
            CREATE TABLE IF NOT EXISTS ESTUDOS(
                CMD_VERSION                             TEXT,
                USER                                    TEXT,
                DATA                                    TEXT,
                HORA                                    TEXT,
                DATA_HORA                               TEXT,
                EMPRESA                                 TEXT,
                NOTA_SERVICO                            INTEGER,
                UC                                      INTEGER,
                LIGACAO_NOVA                            BOOL,
                TIPO_ATIVIDADE                          TEXT,
                CATEGORIA_LIGACAO                       TEXT,
                DEMANDA_EXISTENTE                       REAL,
                DEMANDA_TOTAL_SOLICITADA                REAL,
                DSS_FILE_PATH                           TEXT,
                RAMAL_EXISTENTE                         TEXT,
                RAMAL_MODIFICADO                        BOOL,
                DSS_STRING                              TEXT,
                TRAFO_DELTA_ESTRELA                     BOOL,
                TENSAO_SECUNDARIA                       INTEGER,
                CARGA_FASE_A                            REAL,
                CARGA_FASE_B                            REAL,
                CARGA_FASE_C                            REAL,
                CORRENTE_NEUTRO                         REAL,
                QUEDA_TENSAO                            REAL,
                ELEVACAO_TENSAO                         REAL,
                CARREGAMENTO_MAX_VAO                    REAL,
                CARREGAMENTO_MAX_TRAFO                  REAL,
                FATOR_LIMITANTE                         TEXT,
                MAXIMA_DEMANDA_DISPONIBILIZADA          REAL,
                INDICE_PROPORCIONALIDADE                REAL,
                ITERACOES_CALCULO_MDD                   INTEGER,
                DURACAO_ESTUDO                          REAL
            );
            
            CREATE INDEX IDX_ESTUDOS_NOTAS ON ESTUDOS (NOTA_SERVICO);
            
        ''')
        self.conn.commit()
    
    @staticmethod
    def change_permissions():
    
        try:
            subprocess.check_output(f'icacls.exe {PATH_DB} /inheritance:d /grant:r *S-1-5-11:(OI)(IO) /C', stderr=subprocess.STDOUT)
            subprocess.check_output(f'icacls.exe {PATH_DB} /deny *S-1-5-11:(DE)', stderr=subprocess.STDOUT)
        except:
            pass
    
    def insert_record_estudo(self, cmd, start_time):
    
        ##### Registra o estudo no banco de dados. ##### 
        
        n = dt.datetime.now()
        
        with self.conn:
            self.cursor.execute(f'''
                INSERT INTO ESTUDOS VALUES (
                    '{VERSION}',
                    '{os.environ['USERNAME']}',
                    '{n.strftime("%d/%m/%Y")}',
                    '{n.strftime("%H:%M:%S.%f")}',
                    '{n.strftime("%d/%m/%Y %H:%M:%S.%f")}',
                    '{cmd.empresa}',
                    '{cmd.nota}',
                    '{cmd.uc}',
                    '{cmd.ligacao_nova}',
                    '{cmd.tipo_atividade}',
                    '{cmd.Categoria_Ligacao}',
                    '{cmd.DE}',
                    '{cmd.DTS}',
                    '{cmd.dssfile}',
                    '{cmd.ramal_existente}',
                    '{cmd.ramal_modificado}',
                    '{cmd.circuit_string}',
                    '{cmd.delta_wye}',
                    '{1e3*cmd.tensao_secundaria}',
                    '{cmd.carga_fase_A}',
                    '{cmd.carga_fase_B}',
                    '{cmd.carga_fase_C}',
                    '{cmd.corrente_neutro}',
                    '{cmd.queda_tensao}',
                    '{cmd.elevacao_tensao}',
                    '{cmd.carregamento_max_vao}',
                    '{cmd.carregamento_max_trafo}',
                    '{cmd.fator_limitante}',
                    '{cmd.MDD}',
                    '{100*cmd.K_prop}',
                    '{cmd.kiter}',
                    '{time.time() - start_time}'
                )''')
            self.conn.commit()









if __name__ == '__main__':
    
    db = consultaDB()
    db.connect_db()
    db.create_database()
    db.close_db()
    # db.change_permissions()

