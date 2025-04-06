# -*- coding: utf-8 -*-

"""
Created on Sun Apr 6 19:03:41 2025

@author: Hans Petter Helsør, hph@helsor.com, +47 97052966 Github: hphelsor

Programmet leser data fra et Excel regneark, analyserer dataene og presenterer disse i tabellformat og/eller grafisk.

"""


# import pandas lib as pd
import pandas as pd
# import numpy lib as np
import numpy as np
# import matplotlib lib as plt
import matplotlib.pyplot as plt
# import datetime
from datetime import datetime

def pauselinje(): #Stegvis kjøring av programmets deler
    print("\n\n")
    next_step = input(">>>>>>>> Trykk ENTER for å gå videre >>>>>>>>")
    print("\n\n")


def konverter_til_epoch_numpy(tidsstreng_array, dato="1970-01-01"):
    """Konverterer et NumPy-array med tidsstrenger i formatet 'tt:mm:ss' til Unix-epoketid.

    Args:
        tidsstreng_array: Et NumPy-array med tidsstrenger i formatet 'tt:mm:ss'.
        dato: En streng som representerer datoen som tidsstrengene er relative til.
              Standard er '1970-01-01'.

    Returns:
        Et NumPy-array med Unix-epoketidsverdier (i sekunder).
    """
    datetimes = np.array([datetime.strptime(f"{dato} {t}", "%Y-%m-%d %H:%M:%S") for t in tidsstreng_array])
    # Konverter til NumPy datetime64[s] for å få epoketid i sekunder
    epoch_array = (datetimes.astype('datetime64[s]') - np.datetime64('1970-01-01T00:00:00')) .astype(np.int64)
    return epoch_array



# Leser data fra Excel som Pandas dataframe
df = pd.read_excel('support_uke_24.xlsx')

# Konverterer Pandas dataframe til Numpy arrayer
u_dag = df["Ukedag"].to_numpy()
kl_slett = df["Klokkeslett"].to_numpy()
varighet = df["Varighet"].to_numpy()
score = df["Tilfredshet"].to_numpy()

# Definer ønsket rekkefølge for ukedagene
ukedager_rekkefølge = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']

# Finn antall forekomster av hver verdi
verdier, antall = np.unique(u_dag, return_counts=True)

# Opprett en liste for å lagre antall i riktig rekkefølge
antall_i_riktig_rekkefølge = []
for dag in ukedager_rekkefølge:
    if dag in verdier:
        indeks = np.where(verdier == dag)[0][0]
        antall_i_riktig_rekkefølge.append(antall[indeks])
    else:
        antall_i_riktig_rekkefølge.append(0)  # Hvis en dag ikke finnes, sett antall til 0

# Lag et stolpediagram
plt.bar(ukedager_rekkefølge, antall_i_riktig_rekkefølge)

# Legg til etiketter og tittel
plt.ylabel('Antall supporttelefoner')
plt.title('Antall supporttelefoner pr. ukedag (Lukk vinduet for å fortsette)')

print("\n\nDATA FOR KUNDESUPPORT\n")
print("**** ANTALL HENVENDELSER PR. UKEDAG ****")

antall_ukedager = len(ukedager_rekkefølge)
print(f"Ukedag      Antall henvendelser")
print("--------------------------------------------------------------")
for i in range(antall_ukedager):
    print(f"{ukedager_rekkefølge[i]:10}: {antall_i_riktig_rekkefølge[i]}")

print()
utskrift = input("Vil du se en grafisk framstilling av dette? Ja/Nei ")

# Vis diagrammet hvis svaret kan tolkes som Ja
if utskrift.lower() == "ja" or utskrift.lower() == "j":
    plt.show(block=False)
    pauselinje()
    plt.close() 
else:
    print()


print("**** STATISTIKK FOR TIDSBRUK ****")

print()
plt.clf()
# Finn og skriv ut minimumsverdien i arrayet
min_verdi = np.min(varighet)
min_timer, min_minutter, min_sekunder = map(int, min_verdi.split(':'))

print(f"Den korteste henvendelsen varte", end =" ")
if min_timer > 0:
    print(f"{min_timer} timer,", end =" ")
if min_minutter > 0:
    print(f"{min_minutter} minutter og", end =" ")
print(f"{min_sekunder} sekunder")

# Finn og skriv ut maksimumsverdien i arrayet
maks_verdi = np.max(varighet)
maks_timer, maks_minutter, maks_sekunder = map(int, maks_verdi.split(':'))
print(f"Den lengste henvendelsen varte", end =" ")
if maks_timer > 0:
    print(f"{maks_timer} timer,", end =" ")
if maks_minutter > 0:
    print(f"{maks_minutter} minutter og", end =" ")
print(f"{maks_sekunder} sekunder")

# Beregn verdier for hele uka, først gjennomsnittlig samtaletid

# Lag et NumPy-array med tidsstrenger
tids_array = varighet

# Konverter tidsstrengene til epoketid
epoch_array = konverter_til_epoch_numpy(tids_array)

total_samtaletid = np.sum(epoch_array) # Sum samtaletid
antall_samtaler = np.size(epoch_array) # Antall samtaler
snitt_verdi = round(total_samtaletid/antall_samtaler) # Gjennomsnittlig samtaletid

timer = snitt_verdi // 3600
minutter = (snitt_verdi % 3600) // 60
sekunder = snitt_verdi % 60

print(f"Gjennomsnittlig henvendelsestid var ", end="")
if timer > 0:
    print(f"{timer} timer, ",end="")
if minutter > 0:
    print(f"{minutter} minutter og ",end="")  
print(f"{sekunder} sekunder")

# Beregn og vis total henvendelsestid for hele uka

timer = total_samtaletid // 3600
minutter = (total_samtaletid % 3600) // 60
sekunder = total_samtaletid % 60

print(f"\nTotal henvendelsestid var ", end="")
if timer > 0:
    print(f"{timer} timer, ",end="")
if minutter > 0:
    print(f"{minutter} minutter og ",end="")  
print(f"{sekunder} sekunder")

# Beregn og vis gjennomsnittlig henvendelsestid for hele uka

daglig_samtaletid = int(total_samtaletid/5)
timer = daglig_samtaletid // 3600
minutter = (daglig_samtaletid % 3600) // 60
sekunder = daglig_samtaletid % 60

print(f"Henvendelsestid i snitt pr dag var ", end="")
if timer > 0:
    print(f"{timer} timer, ",end="")
if minutter > 0:
    print(f"{minutter} minutter og ",end="")  
print(f"{sekunder} sekunder")



pauselinje()

# Bergen og vis data pr. økt

print("**** ANTALL HENVENDELSER FORDELT PÅ ØKTER ****")

kl_08_10 = 0
kl_10_12 = 0
kl_12_14 = 0
kl_14_16 = 0
utenfor_arbtid = 0

for i in kl_slett:
    if i[0:2] == ("08" or "09"):
        kl_08_10 += 1
    elif i[0:2] == ("10" or "11"):
        kl_10_12 += 1
    elif i[0:2] == ("12" or "13"):
        kl_12_14 += 1
    elif i[0:2] == ("14" or "15"):
        kl_14_16 += 1
    else:
        utenfor_arbtid += 1

#Legg antall henvendelser pr økt inn i et array
support_fordeling = np.array([kl_08_10,kl_10_12,kl_12_14,kl_14_16])

print()
antall_økter = np.size(support_fordeling)
økter = ["08.00-10.00" , "10.00-12.00" , "12.00-14.00" , "14.00-16.00"]
print(f"Økt             Antall henvendelser")
print("--------------------------------------------------------------")
for i in range(antall_økter):
    print(f"{økter[i]:14}: {support_fordeling[i]}")

print()

# Lag en liste med etiketter for diagrammet
etiketter = ['08.00-10.00', '10.00-12.00', '12.00-14.00', '14.00-16.00']

# Lag et kakediagram med etiketter
plt.pie(support_fordeling, labels=etiketter)

# Legg til en tittel
plt.title("Henvendelser fordelt på økter (Lukk vinduet for å fortsette)")

utskrift = input("Vil du se en grafisk framstilling av dette? Ja/Nei ")

# Vis diagrammet hvis svaret er ja eller j

if utskrift.lower() == "ja" or utskrift.lower() == "j":
    plt.show(block=False)

pauselinje()
plt.close()

print()
print("**** KUNDETILBAKEMELDINGER ****\n")

# Beregn og vis data for kundetilfredshet

resultat = []
negative = 0
nøytrale = 0
positive = 0


for i in score:
    if not np.isnan(i):
        resultat.append(i)

ant_tilbakemeldinger = np.size(resultat)

# Del karakterene inn i negativ, nøytral og positiv

for i in score:
    if i < 7:
        negative += 1
    elif (i >= 7 and i <= 8):
        nøytrale += 1
    elif i > 8:
        positive += 1
    
print(f"{ant_tilbakemeldinger} personer ga tilbakemelding på tilfredshet")
print(f"{round(negative*100/ant_tilbakemeldinger , 1)}% var negative, total {negative}")
print(f"{round(nøytrale*100/ant_tilbakemeldinger , 1)}% var nøytrale, totalt {nøytrale}")
print(f"{round(positive*100/ant_tilbakemeldinger , 1)}% var positive, totalt {positive}")

pauselinje()
print("Flott innsats! Neste uke er vi enda bedre!")
print("__________________________________________\n\n")
