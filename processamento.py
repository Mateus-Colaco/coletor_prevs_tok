import win32com.client as win32
import openpyxl as opyxl
import pandas as pd
import regex as re
import glob, os, shutil

def abre_fecha_excel(caminho_excel):
    excelApp = win32.gencache.EnsureDispatch('Excel.Application')
    excelApp.DisplayAlerts = False
    excelApp.DisplayStatusBar = False
    excelApp.ScreenUpdating = False
    excelApp.Visible = False
    wbwin = excelApp.Workbooks.Open(caminho_excel, UpdateLinks=0, ReadOnly=0)
    wbwin.SaveAs(caminho_excel)
    wbwin.Close()
    excelApp.Application.Quit()


def ajusta_colunas(caminho_excel):
    workbook = opyxl.load_workbook(caminho_excel)
    workbook['Cenário_BASE'].column_dimensions['A'].width = 6.71428571428569  # 47px
    workbook['Cenário_BASE'].column_dimensions['B'].width = 5.7142857142857  # 40px
    workbook['Cenário_BASE'].column_dimensions['C'].width = 10.7142857142857  # 75px
    workbook['Cenário_BASE'].column_dimensions['D'].width = 10.7142857142857
    workbook['Cenário_BASE'].column_dimensions['E'].width = 10.7142857142857
    workbook['Cenário_BASE'].column_dimensions['F'].width = 10.7142857142857
    workbook['Cenário_BASE'].column_dimensions['G'].width = 10.7142857142857
    workbook['Cenário_BASE'].column_dimensions['H'].width = 10.7142857142857
    workbook.move_sheet('Cenário_BASE', -8)
    workbook.save(caminho_excel)


def carrega_df_p_excel(caminho_prev, caminho_excel):
    arquivo_rv = open(caminho_prev, mode='r').read()
    linhas_rv = re.findall(
        '(?x)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)(?:\s)+(\d+)\n?',
        arquivo_rv)

    tabela_prev = pd.DataFrame([list(map(int, linha)) for linha in linhas_rv])
    writer = pd.ExcelWriter(caminho_excel, engine='openpyxl', mode='a')
    tabela_prev.to_excel(
        writer,
        sheet_name='BASE_novo',
        header=False,
        index=False
    )
    writer.close()
    workbook = opyxl.load_workbook(caminho_excel)
    ws_names = [ws.title for ws in workbook.worksheets]
    if 'Cenário_BASE' in ws_names:
        workbook.remove(workbook['Cenário_BASE'])
        try:
            workbook.remove(workbook['Cenário_BASE1'])
            workbook.remove(workbook['Cenário_BASE2'])
            workbook.remove(workbook['Cenário_BASE3'])
            workbook.remove(workbook['Cenário_BASE4'])
        except:
            None
    cen_base = workbook["BASE_novo"]
    cen_base.title = 'Cenário_BASE'
    workbook['Bacias_m3_s']['J10'] = 1
    workbook['Bacias_m3_s']['J220'] = 1
    workbook['Bacias_m3_s']['J290'] = 1
    workbook['Bacias_m3_s']['J329'] = 1
    workbook.save(caminho_excel)


def copia_montador_ref(nome_montador):
    pasta_ref = r'Z:\02 TECNICO\INTELIGENCIA\RENATO\PMO\2024\_Montadores_referencia_SERGIO'
    src_montador = os.path.join(pasta_ref, nome_montador)
    dst_montador = nome_montador
    shutil.copy2(src_montador, dst_montador)


def leitor_prevs(dirs):    
    for diretorio in dirs:
        os.chdir(diretorio)
        caminhos_prevs = glob.glob('*.rv[0-5]')
        padrao_prev = r'20(2\d)([0-1][0-9])-prevs.*.(rv\d)'
        meses = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
        for caminho_prev in caminhos_prevs:
            prev_grupos = re.search(padrao_prev, caminho_prev, flags=re.I)
            ano_prev = int(prev_grupos.group(1))
            mes_prev = int(prev_grupos.group(2))
            rev_prev = str(prev_grupos.group(3)).upper()
            mes = meses[mes_prev - 1]
            excel_RVX = ("20" + str(ano_prev) + "__" + str(mes) + "__" + rev_prev + "__Montador.xlsx")
            copia_montador_ref(excel_RVX)
            caminho_excel = os.getcwd() + os.sep + excel_RVX
            caminho_prev = os.getcwd() + os.sep + prev_grupos.group(0)
            carrega_df_p_excel(caminho_prev, caminho_excel)
            ajusta_colunas(caminho_excel)
            abre_fecha_excel(caminho_excel)
        os.chdir("..")
