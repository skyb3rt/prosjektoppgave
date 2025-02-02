'''
Prosjekt oppgave. Dashbord for supportavdelingen
'''
from datetime import timedelta, datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# Oppgave a. Les inn data fra excel filen og lagre kolonnene i arrays
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

# Oppgave b. Finne antall henvedelser per dag og visualisere ved bruk av et søylediagram

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
plt.show()

# Oppgave c. Finn minste og lengste samtaletid som er loggført for uke 2

print (f"Minste samtaletid: {np.min(varighet)}")
print (f"Lengste samtaletid: {np.max(varighet)}")

# Oppgave d. Regner ut gjennomsnittlig samtaletid basert på alle henvendelser i uke 24

def regn_ut_gjennomsnittlig_samtaletid(varigheter:np) -> float:
    '''Funksjon som regner ut gjennomsnittlig samtaletid'''
    sum_varighet = timedelta(0)
    for var in varigheter:
        tid = datetime.strptime(var, "%H:%M:%S")
        sum_varighet += timedelta(hours=tid.hour, minutes=tid.minute, seconds=tid.second)
    return sum_varighet/len(varighet)

print(f"Gjennomsnittlig samtaletid: {regn_ut_gjennomsnittlig_samtaletid(varighet)}")

# Oppgave e.  antall henvendelser tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24

def antall_henvendelser_per_bolk(klokkeslett:np)-> dict:
    '''Funksjon som returnerer antall henvendelser per tidsblokk'''
    antall= {"08-10":0, "10-12":0, "12-14":0, "14-16":0}
    for tid in klokkeslett:
        time=int(tid.split(":")[0])
        match time:
            case 8|9:
                antall["08-10"]+=1
            case 10|11:
                antall["10-12"]+=1
            case 12|13:
                antall["12-14"]+=1
            case 14|15|16:
                antall["14-16"]+=1
    return antall

antall_pr_blokk=antall_henvendelser_per_bolk(kl_slett)
plt.pie(antall_pr_blokk.values(), labels=antall_pr_blokk.keys(), autopct='%1.1f%%')
plt.title("Antall henvendelser per tidsblokk")
plt.show()

# Oppgave f. Supportavdelingens NP
def supportavdelingens_np(tilbakemeldinger:np) -> float:
    '''Funksjon som regner ut supportavdelingens NP'''
    fornoyd_score={"postiv":0, "noytral":0, "negativ":0}
    for tilbakemelding in tilbakemeldinger:
        if tilbakemelding: # Utelukker None / Nan verdiene
            match tilbakemelding:
                case _ as status if 1<=status<=6:
                    fornoyd_score["negativ"]+=1
                case _ as status if 7<=status<=8:
                    fornoyd_score["noytral"]+=1
                case _ as status if 9<=status<=10:
                    fornoyd_score["postiv"]+=1
    antall_tilbakemeldinger = sum(fornoyd_score.values())
    return ((fornoyd_score["postiv"]/antall_tilbakemeldinger)-
            (fornoyd_score["negativ"]/antall_tilbakemeldinger))*100

print(f"Supportavdelingens NP: {supportavdelingens_np(score):.2f}%")
