'''
Prosjekt oppgave. Dashbord for supportavdelingen
'''
from datetime import timedelta, datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# oppgave a. Les inn data fra excel filen og lagre kolonnene i arrays
FILENAME = "support_uke_24.xlsx"
UKEDAGER = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]

def read_xlsx(file_path:str) -> pd.DataFrame:
    '''Funksjon som leser inn en excel fil og returnerer en pandas dataframe'''
    return pd.read_excel(file_path)

support_df = read_xlsx(FILENAME)

u_dag = support_df["Ukedag"].to_numpy()
kl_slett = support_df["Klokkeslett"].to_numpy()
varighet = support_df["Varighet"].to_numpy()
score = support_df["Tilfredshet"].to_numpy()

# Fjerne NaN verdier fra score og sette de som None
score=np.where(pd.isna(score), None, score)

# oppgave b. Finne antall henvedelser per dag og visualisere ved bruk av et søylediagram

def antall_henvendelser_per_dag(dager:np) -> dict:
    '''
    Funksjon som tar inn en array med dager og
    returnerer en dictionary med antall henvendelser per dag
    '''
    return {dag:np.sum(dager==dag) for dag in UKEDAGER}

plt.bar(antall_henvendelser_per_dag(u_dag).keys(), antall_henvendelser_per_dag(u_dag).values())
plt.title("Antall henvendelser per dag")
plt.xlabel("Ukedag")
plt.ylabel("Antall henvendelser")
#plt.show()

# oppgave c. Finn minste og lengste samtaletid som er loggført for uke 2

print (f"Minste samtaletid: {np.min(varighet)}")
print (f"Lengste samtaletid: {np.max(varighet)}")

# oppgave d. Regner ut gjennomsnittlig samtaletid basert på alle henvendelser i uke 24

def regn_ut_gjennomsnittlig_samtaletid(varigheter:np) -> float:
    '''Funksjon som regner ut gjennomsnittlig samtaletid'''
    sum_varighet = timedelta(0)
    for var in varigheter:
        tid = datetime.strptime(var, "%H:%M:%S")
        sum_varighet += timedelta(hours=tid.hour, minutes=tid.minute, seconds=tid.second)
    return sum_varighet/len(varighet)

print(f"Gjennomsnittlig samtaletid: {regn_ut_gjennomsnittlig_samtaletid(varighet)}")
