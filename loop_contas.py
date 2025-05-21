# ================================================== #

# ~~ Bibliotecas.
import xlwings as xw
import pandas as pd
import os

# ================================================== #

# ~~ Caminho da planilha.
caminho = os.path.dirname(os.path.abspath(__file__))
caminho = caminho + r"\agrupamento_fd32.xlsx"
planilha = xw.Book(caminho)
sheet1 = planilha.sheets["Sheet1"]
sheet2 = planilha.sheets["Planilha1"]
df = pd.read_excel(caminho, sheet_name = "Planilha1", dtype = {"MATRIZ": str})

# ================================================== #

# ~~ Execução em loop.
for linha in range (2, 999999):
    if sheet1.range(f"B{linha}").value is None:
        break
    raiz = str(sheet1.range(f"B{linha}").value + "0001")
    linha_df = df.index[df['MATRIZ'] == raiz].tolist()
    if linha_df:
        linha_df = int(linha_df[0])
        linha_df = linha_df + 2
        código_erp = str(sheet2.range(f"C{linha_df}").value)
        sheet1.range(f"G{linha}").value = código_erp
    else:
        código_erp = "Matriz não encontrada."
        sheet1.range(f"G{linha}").value = código_erp

# ================================================== #
