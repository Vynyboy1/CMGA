# %%
import os
import importlib.util
import subprocess
import sys
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta

def check_install_import(packages):
    """Verifica, instala (se necessário) e importa automaticamente os pacotes."""
    imported_modules = {}

    for package in packages:
        if importlib.util.find_spec(package) is None:
            print(f"Pacote '{package}' não encontrado. Instalando...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        
        # Importação dinâmica do pacote
        imported_modules[package] = importlib.import_module(package)
        print(f"Pacote '{package}' importado com sucesso.")    

#Definir função para obter caminho do diretório de trabalho]

def Checa_Caminho(caminho: str) -> str:   
    import getpass
    usuario = getpass.getuser()
    if not os.path.exists(caminho) and  usuario == 'vgcs4':
        caminho_alternativo = caminho.replace("CENG/", "CENG/CENG/")  
        return caminho_alternativo          
    return caminho 


def get_script_path(caminho):
    try:
        # Se o script estiver rodando como um arquivo normal (.py)
        return Path(caminho).resolve().parent
    except NameError:
        # Se estiver rodando em Jupyter Notebook
        return Path(os.getcwd()).resolve()

def get_last_day_of_month(year, month):
    """Retorna o último dia do mês para um determinado ano e mês."""
    # Ajusta para o primeiro dia do próximo mês e subtrai 1 dia
    first_day_next_month = datetime(year, month, 1) + relativedelta(months=1)
    last_day = first_day_next_month - relativedelta(days=1)
    return last_day.date()

def Criar_planilha_UN(Planilha):
    # Cria pasta do dia
    data_execucao = datetime.now().strftime("%Y-%m-%d")
    current_file_path = str(get_script_path(__file__))
    file_path = current_file_path.replace('\\', '/')    
    CaminhoCENG = file_path[:file_path.find('OneDrive - Energisa') + len('OneDrive - Energisa')+1]
    caminho_base = CaminhoCENG + f'03- DADOS/Base_Falhas/1.Transformador/Info_Faltando/{Periodo_Corte[0:4]}' 
    pasta_base = fr"{caminho_base}/{data_execucao}"
    os.makedirs(pasta_base, exist_ok=True)

    # Lista de empresas únicas
    Lista_UN = Planilha['EMPRESA'].unique().tolist()

    for empresa in Lista_UN:
        nome_empresa = str(empresa).replace("/", "_").replace("\\", "_")
       # caminho = os.path.join({CaminhoCENG}03- DADOS/Base_Falhas/1.Transformador/Base_Atualizada_Trafo/Substituicoes_{Periodo_Corte[0:4]}.xlsx)   
        caminho = os.path.join(pasta_base, f"Correção {nome_empresa}.xlsx")
        Planilha[Planilha['EMPRESA'] == empresa].to_excel(caminho, index=False)

        print(f"Planilha criada: {caminho}")

def saudacao(nome):
    return f'oláa {nome}'





