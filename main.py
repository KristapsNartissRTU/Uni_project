import pandas as pd

# Excel faila nosaukums
excel_file = 'piemers.xlsx'

# Iegūst lapu nosaukumus
sheet_names = pd.ExcelFile(excel_file).sheet_names

# Pārbauda, vai ir vismaz 1 lapa
if len(sheet_names) < 1:
    raise ValueError("Failā nav nevienas lapas!")

# Nolasa pirmo lapu
df = pd.read_excel(excel_file, sheet_name=sheet_names[0])

# Pārbauda, vai ir nepieciešamās kolonnas
required_cols = {'atslēga', 'vērtība'}
if not required_cols.issubset(df.columns):
    raise ValueError("Trūkst kolonnas 'atslēga' un/vai 'vērtība'.")

# Izveido hashtabulu (vārdnīcu) bez key/value cikla
hash_table = dict(zip(df['atslēga'], df['vērtība']))

# Aprēķina rezultātus bez for-cikla – izmanto pandas funkcionalitāti
# Piemērs: pievieno 1 visām skaitliskajām vērtībām
df['rezultāts'] = pd.to_numeric(df['vērtība'], errors='coerce') + 1

# Saglabā rezultātu jaunā lapā 'Rezultāti'
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # Oriģinālā lapa (pārraksta tāpat, vai vienkārši atstāj)
    df[['atslēga', 'vērtība']].to_excel(writer, sheet_name=sheet_names[0], index=False)
    # Jaunā lapa ar rezultātiem
    df[['atslēga', 'rezultāts']].to_excel(writer, sheet_name='Rezultāti', index=False)

print("Dati apstrādāti un rezultāti saglabāti lapā 'Rezultāti'.")
