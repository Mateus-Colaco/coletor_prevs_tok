from concurrent.futures import ThreadPoolExecutor, as_completed
import processamento 
from requests import post, Response
from dotenv import load_dotenv
from datetime import datetime
import os, warnings, glob, tarfile

import visualizacoes
warnings.filterwarnings('ignore')
MAIN_PATH = os.path.dirname(os.path.abspath(__file__))
if os.getcwd() != MAIN_PATH: os.chdir(MAIN_PATH)

def data_do_prevs(nome_prevs: str, modelo: str):
    data = datetime.now().strftime('%Y%m%d_')
    data_prevs = nome_prevs.replace(modelo, '').replace(data, '')[:-4]
    data_prevs = data_prevs[:-1]
    nova_data_formato = data_prevs.replace('_', '')
    return data_prevs, nova_data_formato

def arruma_nome_prevs(nome_prevs: str, data_dia: str, modelo: str) -> str:
    nome_prevs = nome_prevs.replace('_sem_vies', '').replace('_com_vies', '')
    nome_prevs = nome_prevs.replace(f'{data_dia}_prevs_', '')
    data_prevs, nova_data_formato = data_do_prevs(nome_prevs, modelo)
    nome_prevs = nome_prevs.replace(data_prevs, nova_data_formato)
    nome_prevs = nome_prevs.replace("_", "-")
    nome_prevs = nome_prevs.replace(f'-{modelo.replace("_", "-")}', f'-prevs-{modelo}')
    return nome_prevs


def baixa_prevs(url: str, endpoint_base: str, modelo: str, data_dia: str, data_pasta: str) -> None:
    endpoint, arquivo = gera_url(endpoint_base, modelo, data_dia, data_pasta)
    resposta = post(url, data={'t': os.getenv('TOKEN'), 'p': endpoint}, verify=False, timeout=120)
    salva(resposta, f'Prevs/{data_pasta}/{data_dia}/{arquivo}', data_pasta, data_dia)


def config_pasta_diaria(caminho: str, data_pasta: str, data_dia: str) -> tuple[str, str]:
    """
    Cria diretorios, para organizacao dos arquivos, se necessario
    """
    caminho = caminho.replace('PREVS_', 'zip/PREVS_')
    diretorio = f'Prevs/{data_pasta}/{data_dia}'
    os.makedirs(f'{diretorio}/zip', exist_ok=True)
    return caminho, diretorio


def descompacta_targz(caminho: str, diretorio: str) -> None:
    tar = tarfile.open(caminho)
    tar.extractall(diretorio)
    tar.close()


def executa_paralelo(execucoes: list[tuple[callable, tuple]]) -> None:
    with ThreadPoolExecutor() as executor:
        submits = [executor.submit(func, *args) for func, args in execucoes]
        as_completed(submits)
        

def gera_url(endpoint_base: str, modelo: str, data_dia: str, data_pasta: str) -> tuple[str, str]:
    """
    Padroniza a url para uso da API da Tempo OK
    """
    prevs = f'PREVS_{modelo}_{data_dia}.tar.gz'
    if modelo.endswith('av_vaz'): modelo = f'{modelo.replace("av_vaz", "_estat")}/{modelo}'
    return f'{endpoint_base}/{modelo}/{data_pasta}/{prevs}', prevs


def inputs_baixa_prevs(modelo: str) -> tuple[str, str, str, str, str]:
    return (URL, ENDPOINT_BASE, modelo, DATA_DIA, DATA_PASTA)


def necessario_baixar(data_pasta: str, data_dia: str, modelos: list[str]) -> bool:
    diretorio = f'Prevs/{data_pasta}/{data_dia}'
    return not all([os.path.exists(f'{diretorio}/{modelo}') for modelo in modelos])


def novo_nome_prevs(prevs: str) -> str:
    """
    Padroniza os nomes para uso no prospec
    """
    modelo = prevs.split('/')[3]
    data_dia = prevs.split('/')[2]
    nome_prevs = arruma_nome_prevs(prevs.split('/')[-1], data_dia, modelo)
    pasta_prevs = '/'.join(prevs.split('/')[:-1])
    return f'{pasta_prevs}/{nome_prevs}'


def renomeia_prevs(data_pasta: str, data_dia: str) -> None:
    caminho = f'Prevs/{data_pasta}/{data_dia}'
    prevs = [x.replace('\\', '/') for x in glob.glob(f'{caminho}/*/*.rv*')]
    [os.rename(p, novo_nome_prevs(p)) for p in prevs]


def salva(resposta: Response, caminho: str, data_pasta: str, data_dia: str) -> None:
    caminho, diretorio = config_pasta_diaria(caminho, data_pasta, data_dia)
    if resposta.status_code == 200:
        with open(caminho, 'wb') as f: f.write(resposta.content)
    descompacta_targz(caminho, diretorio)


def main():
    if necessario_baixar(DATA_PASTA, DATA_DIA, MODELOS):
        execucoes = [(baixa_prevs, inputs_baixa_prevs(modelo)) for modelo in MODELOS]
        executa_paralelo(execucoes)
        renomeia_prevs(DATA_PASTA, DATA_DIA)


if __name__ == '__main__': 
    
    load_dotenv()
    DATA = datetime.now()
    DATA_PASTA = DATA.strftime('%Y-%m')
    DATA_DIA = DATA.strftime('%Y%m%d')
    ENDPOINT_BASE = os.getenv('ENDPOINT_BASE')
    MODELOS = os.getenv('MODELOS').split(',')
    URL = os.getenv('URL')
    main()
    os.chdir(f'Prevs/{DATA_PASTA}/{DATA_DIA}')
    dirs = [x for x in os.listdir(os.getcwd()) if os.path.isdir(x) if x!='zip']
    processamento.leitor_prevs(dirs)
    visualizacoes.ler_enas()