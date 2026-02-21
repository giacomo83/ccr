import pandas as pd
import numpy as np
import os

# =========================
# 1️⃣ PARAMETRI # AGO: TOTALI PER COMPETENZA IN CSV 
# =========================
FILE_CSV = r"C:\Users\Giacomo\Desktop\PRIMO_LAVORO\analisi_dati\sara2025.csv"
OUTPUT_DIR = "output"
OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, "dashboard_mastrini_finale.xlsx")

SEPARATORE = ";"
DECIMALE = ","
MIGLIAIA = "."
ENCODING = "latin1"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# 2️⃣ LETTURA CSV
# =========================
df = pd.read_csv(
    FILE_CSV,
    sep=SEPARATORE,
    decimal=DECIMALE,
    thousands=MIGLIAIA,
    encoding=ENCODING
)

# =========================
# 3️⃣ COLONNE UTILI
# =========================
df = df[[
    "Codice sottoconto",
    "Descrizione sottoconto",
    "Mese",
    "Imponibile movimento DARE",
    "Imponibile movimento AVERE"
]]

# =========================
# 4️⃣ NORMALIZZAZIONE
# =========================
df["Mese"] = df["Mese"].astype(int)
df["Imponibile movimento DARE"] = pd.to_numeric(df["Imponibile movimento DARE"], errors="coerce").fillna(0)
df["Imponibile movimento AVERE"] = pd.to_numeric(df["Imponibile movimento AVERE"], errors="coerce").fillna(0)

# =========================
# 5️⃣ CLASSIFICAZIONE COSTI / RICAVI
# =========================
df["Tipo"] = np.where(
    df["Codice sottoconto"].astype(str).str.startswith("7"),
    "COSTO",
    np.where(
        df["Codice sottoconto"].astype(str).str.startswith("5"),
        "RICAVO",
        "ALTRO"
    )
)

df["Importo"] = np.where(
    df["Tipo"] == "COSTO",
    df["Imponibile movimento DARE"],
    df["Imponibile movimento AVERE"]
)

df = df[df["Tipo"].isin(["COSTO", "RICAVO"])]

# =========================
# 6️⃣ ALERT COSTI/RICAVI ±2σ
# =========================
mensile = df.groupby(["Codice sottoconto", "Descrizione sottoconto", "Tipo", "Mese"], as_index=False)["Importo"].sum()

statistiche = mensile.groupby(["Codice sottoconto", "Descrizione sottoconto", "Tipo"]).agg(
    media_mensile=("Importo", "mean"),
    std_mensile=("Importo", "std")
).reset_index()

statistiche["soglia_alta"] = statistiche["media_mensile"] + 2 * statistiche["std_mensile"]
statistiche["soglia_bassa"] = statistiche["media_mensile"] - 2 * statistiche["std_mensile"]

controllo = mensile.merge(statistiche, on=["Codice sottoconto", "Descrizione sottoconto", "Tipo"], how="left")


controllo["motivo_alert"] = np.select(
    [
        (controllo["Tipo"] == "COSTO") & (controllo["std_mensile"] > 0) & (controllo["Importo"] > controllo["soglia_alta"]),
        (controllo["Tipo"] == "RICAVO") & (controllo["std_mensile"] > 0) & (controllo["Importo"] < controllo["soglia_bassa"])
    ],
    ["Costo sopra la normalità", "Ricavo sotto la normalità"],
    default=""
)

# =========================
# 7️⃣ CALCOLO EBITDA MENSILE
# =========================
esclusi_ebitda = ["506", "706", "7070", "7071", "7072", "7073", "7074", "7075", "709", "508", "71", "51", 
                  "5018", "5019", "5020", "5025", "509", "5014", "7014", "7015", "5015","7016", "5017"]
df_operativo = df[~df["Codice sottoconto"].astype(str).str.startswith(tuple(esclusi_ebitda))]

mensile_totale = df_operativo.groupby(["Tipo", "Mese"], as_index=False)["Importo"].sum()
ricavi = mensile_totale[mensile_totale["Tipo"]=="RICAVO"].set_index("Mese")["Importo"]
costi = mensile_totale[mensile_totale["Tipo"]=="COSTO"].set_index("Mese")["Importo"]

ebitda = pd.DataFrame({"Mese": range(1,13)})
ebitda["Ricavi"] = [ricavi.get(m,0) for m in ebitda["Mese"]]
ebitda["Costi"] = [costi.get(m,0) for m in ebitda["Mese"]]
ebitda["EBITDA"] = ebitda["Ricavi"] - ebitda["Costi"]

# =========================
# 8️⃣ FLUSSO DI CASSA
# =========================
df_cassa = pd.read_csv(
    FILE_CSV,
    sep=SEPARATORE,
    decimal=DECIMALE,
    thousands=MIGLIAIA,
    encoding=ENCODING,
    dtype={"Codice sottoconto": str}
)

df_cassa = df_cassa[["Mese","Codice sottoconto","Imponibile movimento DARE","Imponibile movimento AVERE"]]
df_cassa["Imponibile movimento DARE"] = pd.to_numeric(df_cassa["Imponibile movimento DARE"], errors="coerce").fillna(0)
df_cassa["Imponibile movimento AVERE"] = pd.to_numeric(df_cassa["Imponibile movimento AVERE"], errors="coerce").fillna(0)

mask_cassa = (
    df_cassa["Codice sottoconto"].str.startswith("1034") |
    df_cassa["Codice sottoconto"].str.startswith("103500")
)
df_cassa = df_cassa[mask_cassa]

flusso_cassa = df_cassa.groupby("Mese", as_index=False).agg(
    DARE_tot=('Imponibile movimento DARE','sum'),
    AVERE_tot=('Imponibile movimento AVERE','sum')
)
flusso_cassa["Flusso_di_cassa"] = flusso_cassa["DARE_tot"] - flusso_cassa["AVERE_tot"]
flusso_cassa = flusso_cassa[["Mese","Flusso_di_cassa"]]
flusso_cassa = flusso_cassa.set_index("Mese").reindex(range(1,13), fill_value=0).reset_index()

# =========================
# 9️⃣ CALCOLO CCR con NOTE
# =========================
ccr_df = flusso_cassa.merge(ebitda[["Mese","EBITDA"]], on="Mese", how="left")

# colonna NOTE
def nota(row):
    if row["Flusso_di_cassa"] < 0 and row["EBITDA"] < 0:
        return "Rischio insolvenza"
    elif row["Flusso_di_cassa"] > 0 and row["EBITDA"] < 0:
        return "Liquidità sostenuta da fattori non core"
    elif row["Flusso_di_cassa"] < 0 and row["EBITDA"] > 0:
        return "Reddito positivo operativo non trasformato in liquidità"
    else:
        return ""

# colonna CCR
def calcola_ccr(row):
    if row["Flusso_di_cassa"] < 0 and row["EBITDA"] < 0:
        return np.nan
    elif row["EBITDA"] != 0:
        return row["Flusso_di_cassa"]/row["EBITDA"]
    else:
        return np.nan

ccr_df["CCR"] = ccr_df.apply(calcola_ccr, axis=1)
ccr_df["NOTE"] = ccr_df.apply(nota, axis=1)

# metto NOTE subito dopo CCR
ccr_df = ccr_df[["Mese","Flusso_di_cassa","EBITDA","CCR","NOTE"]]

# =========================
# 10️⃣ SCRITTURA EXCEL
# =========================
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    # Foglio Alert
    controllo.to_excel(writer, sheet_name='Alert', index=False)
    
    # Foglio EBITDA
    ebitda.to_excel(writer, sheet_name='EBITDA', index=False)

    # Foglio Flusso di cassa
    flusso_cassa.to_excel(writer, sheet_name='Flusso di cassa', index=False)

    # Foglio CCR
    ccr_df.to_excel(writer, sheet_name='CCR', index=False)

    workbook = writer.book
    ws_alert = writer.sheets['Alert']
    ws_ebitda = writer.sheets['EBITDA']
    ws_cassa = writer.sheets['Flusso di cassa']
    ws_ccr = writer.sheets['CCR']

    # =========================
    # Formati
    euro_format = workbook.add_format({'num_format': '€ #,##0.00', 'align':'center', 'valign':'vcenter'})
    percent_format = workbook.add_format({'num_format': '0.00%', 'align':'center'})
    center_format = workbook.add_format({'align':'center'})

    # centrare Mese ovunque
    for ws in [ws_alert, ws_ebitda, ws_cassa, ws_ccr]:
        ws.set_column(0,0,10,center_format)  # colonna Mese

    # Alert
    for col_num, value in enumerate(controllo.columns):
        ws_alert.set_column(col_num, col_num, 15, euro_format if 'Importo' in value or 'soglia' in value or 'media' in value else None)

    # EBITDA
    for col_num, value in enumerate(ebitda.columns):
        ws_ebitda.set_column(col_num, col_num, 15, euro_format if value in ['Ricavi','Costi','EBITDA'] else None)

    # Flusso di cassa
    for col_num, value in enumerate(flusso_cassa.columns):
        ws_cassa.set_column(col_num, col_num, 18 if value=="Flusso_di_cassa" else 12, euro_format if value=="Flusso_di_cassa" else None)

    # CCR
    for col_num, value in enumerate(ccr_df.columns):
        if value in ['Flusso_di_cassa','EBITDA']:
            ws_ccr.set_column(col_num, col_num, 18, euro_format)
        elif value == 'CCR':
            ws_ccr.set_column(col_num, col_num, 12, percent_format)
        elif value == 'NOTE':
            ws_ccr.set_column(col_num, col_num, 40)
        else:
            ws_ccr.set_column(col_num, col_num, 12)

    # =========================
    # Colori alert
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align':'center'})
    yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align':'center'})

    importo_col = controllo.columns.get_loc("Importo")
    last_row = len(controllo)

    ws_alert.conditional_format(1, importo_col, last_row, importo_col, {
        'type': 'formula',
        'criteria': f'=AND($C2="COSTO", ISNUMBER($H2), $E2>=$H2)',
        'format': red_format
    })
    ws_alert.conditional_format(1, importo_col, last_row, importo_col, {
        'type': 'formula',
        'criteria': f'=AND($C2="RICAVO", ISNUMBER($I2), $E2<=$I2)',
        'format': yellow_format
    })

    # =========================
    # Grafico EBITDA
    chart_ebitda = workbook.add_chart({'type': 'line'})
    chart_ebitda.add_series({
        'name': 'EBITDA',
        'categories': ['EBITDA', 1, 0, 12, 0],
        'values':     ['EBITDA', 1, 3, 12, 3],
        'marker': {'type':'circle','size':5},
        'line': {'color':'green'}
    })
    chart_ebitda.set_title({'name':'EBITDA Mensile'})
    chart_ebitda.set_x_axis({'name':'Mese'})
    chart_ebitda.set_y_axis({'name':'EBITDA (€)'})
    chart_ebitda.set_legend({'position':'bottom'})
    ws_ebitda.insert_chart('F2', chart_ebitda, {'x_scale':1.5, 'y_scale':1.5})

    # =========================
    # Grafico Flusso di Cassa
    chart_cassa = workbook.add_chart({'type':'column'})
    chart_cassa.add_series({
        'name': 'Flusso di cassa',
        'categories': ['Flusso di cassa', 1, 0, 12, 0],
        'values':     ['Flusso di cassa', 1, 1, 12, 1],
        'fill': {'color':'blue'}
    })
    chart_cassa.set_title({'name':'Flusso di cassa Mensile'})
    chart_cassa.set_x_axis({'name':'Mese'})
    chart_cassa.set_y_axis({'name':'€'})
    ws_cassa.insert_chart('F2', chart_cassa, {'x_scale':1.5, 'y_scale':1.5})

    # =========================
    # Grafico CCR
    chart_ccr = workbook.add_chart({'type':'line'})
    chart_ccr.add_series({
        'name': 'CCR',
        'categories': ['CCR', 1, 0, 12, 0],
        'values':     ['CCR', 1, 3, 12, 3],
        'line': {'color':'orange'},
        'marker': {'type':'circle','size':5}
    })
    chart_ccr.set_title({'name':'Cash Conversion Ratio Mensile'})
    chart_ccr.set_x_axis({'name':'Mese'})
    chart_ccr.set_y_axis({'name':'CCR'})
    ws_ccr.insert_chart('F2', chart_ccr, {'x_scale':1.5, 'y_scale':1.5})

print("Dashboard Excel creata con 4 fogli: Alert, EBITDA, Flusso di cassa, CCR")

