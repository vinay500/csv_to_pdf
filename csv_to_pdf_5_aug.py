import pandas as pd
from datetime import datetime,date

df = pd.DataFrame(pd.read_excel("pdf.xlsx", engine='openpyxl'))
print("df:")
print(df)
# dataframe data in list
resultsData = df.values.tolist()
print("resultsData:")
print(resultsData)
# dataframe data in list
resultDict = df.to_dict()
print("resultDict:")
print(resultDict)

# csv data
# data_in_list=[[Timestamp'('2022-05-31 00:00:00'), 'PA6052', 'STRATEGE PATRIMOINE-EUR', 'FR7615298000010037017525045EUR', 'EUR', 'nan', '31/05/2022', -47.8, 0, 'Dépositaire
# ', "22SEAV022666", "31/05/2022", "SCT - SEPA ALLER VIREMENT", 'NMSC', '-', nan, nan, -38598.62, 38598.62, nan, 69364.67, 69364.67, -30766.05], [Timestamp('
# 2022-06-03 00:00:00'), 'PA6052', 'STRATEGE PATRIMOINE-EUR', 'FR7615298000010037017525045EUR', 'EUR', nan, datetime.datetime(2022, 3, 6, 0, 0), 63418.63, 0,
#  'Dépositaire', '22ROPC020529', datetime.datetime(2022, 3, 6, 0, 0), 'SCT - SEPA ALLER VIREMENT', 'NMSC', '-', nan, nan, -38598.62, 38598.62, nan, 69364.67
# , 69364.67, -30766.05], [Timestamp('2022-06-07 00:00:00'), 'PA6052', 'STRATEGE PATRIMOINE-EUR', 'FR7615298000010037017525045EUR', 'EUR', nan, datetime.date
# time(2022, 7, 6, 0, 0), -23.41, 0, 'Dépositaire', '22AGII003938', '31/05/2022', 'RACHAT OPCVM?1?CM-AM MON PREM PARTS', 'NSEC', '-', nan, nan, -38598.62, 38
# 598.62, nan, 69364.67, 69364.67, -30766.05], [Timestamp('2022-06-13 00:00:00'), 'PA6052', 'STRATEGE PATRIMOINE-EUR', 'FR7615298000010037017525045EUR', 'EUR
# ', nan, '13/06/2022', 3005.14, 0, 'Dépositaire', '22NMSC022802', '14/06/2022', "ECHELLES D'INTERETS", 'NCOM', '-', nan, nan, -38598.62, 38598.62, nan, 6936
# 4.67, 69364.67, -30766.05], [Timestamp('2019-06-13 00:00:00'), 'PA6463', 'DPG STRATEGIES ACTIONS', 'FR361529800001AF00101085030EUR', 'EUR', 'FR24', '13/06/
# 2019', 0.0, 0, 'Dépositaire', nan, '13/06/2019', 'SOUSCRIPTION OPCVM?37?STRATEGE FIN PATRIMOINE/CAP', 'NSEC', '-', nan, nan, -38598.62, 38598.62, nan, 6936
# 4.67, 69364.67, -30766.05], [Timestamp('2019-06-13 00:00:00'), 'PA6463', 'DPG STRATEGIES ACTIONS', 'FR361529800001AF00101085030USD', 'USD', 'FR24', '13/06/
# 2019', 0.0, 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, nan, 0.0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6481', 'ALLO
# CATION PROFIL 3-EUR', 'FR361529800001AF00124885047EUR', 'EUR', nan, '13/06/2019', 0.0, 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0
# .0, nan, 0.0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6481', 'ALLOCATION PROFIL 3-GBP', 'FR361529800001AF00124885047GBP', 'GBP', nan, '13/06/2019'
# , 0.0, 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, nan, 0.0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6481', 'ALLOCATIO
# N PROFIL 3-USD', 'FR361529800001AF00124885047USD', 'USD', nan, '13/06/2019', 0.0, 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, n
# an, 0.0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6482', 'ALLOCATION PROFIL 2-EUR', 'FR361529800001AF00125825074EUR', 'EUR', nan, '13/06/2019', 0.0
# , 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, nan, 0.0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6482', 'ALLOCATION PRO
# FIL 2-GBP', 'FR361529800001AF00125825074GBP', 'GBP', nan, '13/06/2019', 0.0, 0, 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, nan, 0
# .0, 0.0, 0.0], [Timestamp('2019-06-13 00:00:00'), 'PA6482', 'ALLOCATION PROFIL 2-JPY', 'FR361529800001AF00125825074JPY', 'JPY', nan, '13/06/2019', 0.0, 0,
# 'Dépositaire', nan, '13/06/2019', nan, nan, '-', nan, nan, nan, 0.0, nan, 0.0, 0.0, 0.0]]

# data_in_dict={'Entry Date': {0: Timestamp('2022-05-31 00:00:00'), 1: Timestamp('2022-06-03 00:00:00'), 2: Timestamp('2022-06-07 00:00:00'), 3: Timestamp('2022-06-13 00:
# 00:00'), 4: Timestamp('2019-06-13 00:00:00'), 5: Timestamp('2019-06-13 00:00:00'), 6: Timestamp('2019-06-13 00:00:00'), 7: Timestamp('2019-06-13 00:00:00')
# , 8: Timestamp('2019-06-13 00:00:00'), 9: Timestamp('2019-06-13 00:00:00'), 10: Timestamp('2019-06-13 00:00:00'), 11: Timestamp('2019-06-13 00:00:00')}, 'F
# onds:': {0: 'PA6052', 1: 'PA6052', 2: 'PA6052', 3: 'PA6052', 4: 'PA6463', 5: 'PA6463', 6: 'PA6481', 7: 'PA6481', 8: 'PA6481', 9: 'PA6482', 10: 'PA6482', 11
# : 'PA6482'}, 'Account name': {0: 'STRATEGE PATRIMOINE-EUR', 1: 'STRATEGE PATRIMOINE-EUR', 2: 'STRATEGE PATRIMOINE-EUR', 3: 'STRATEGE PATRIMOINE-EUR', 4: 'D
# PG STRATEGIES ACTIONS', 5: 'DPG STRATEGIES ACTIONS', 6: 'ALLOCATION PROFIL 3-EUR', 7: 'ALLOCATION PROFIL 3-GBP', 8: 'ALLOCATION PROFIL 3-USD', 9: 'ALLOCATI
# ON PROFIL 2-EUR', 10: 'ALLOCATION PROFIL 2-GBP', 11: 'ALLOCATION PROFIL 2-JPY'}, 'Compte Dépositaire:': {0: 'FR7615298000010037017525045EUR', 1: 'FR7615298
# 000010037017525045EUR', 2: 'FR7615298000010037017525045EUR', 3: 'FR7615298000010037017525045EUR', 4: 'FR361529800001AF00101085030EUR', 5: 'FR361529800001AF
# 00101085030USD', 6: 'FR361529800001AF00124885047EUR', 7: 'FR361529800001AF00124885047GBP', 8: 'FR361529800001AF00124885047USD', 9: 'FR361529800001AF0012582
# 5074EUR', 10: 'FR361529800001AF00125825074GBP', 11: 'FR361529800001AF00125825074JPY'}, 'Currency A/C': {0: 'EUR', 1: 'EUR', 2: 'EUR', 3: 'EUR', 4: 'EUR', 5
# : 'USD', 6: 'EUR', 7: 'GBP', 8: 'USD', 9: 'EUR', 10: 'GBP', 11: 'JPY'}, 'Compte Valorisateur:': {0: nan, 1: nan, 2: nan, 3: nan, 4: 'FR24', 5: 'FR24', 6: n
# an, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Date comptable': {0: '31/05/2022', 1: datetime.datetime(2022, 3, 6, 0, 0), 2: datetime.datetime(2022, 7, 6,
#  0, 0), 3: '13/06/2022', 4: '13/06/2019', 5: '13/06/2019', 6: '13/06/2019', 7: '13/06/2019', 8: '13/06/2019', 9: '13/06/2019', 10: '13/06/2019', 11: '13/06
# /2019'}, 'Montant Dépositaire': {0: -47.8, 1: 63418.63, 2: -23.41, 3: 3005.14, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}, 'Montant
# Valorisateur': {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0}, 'Origine': {0: 'Dépositaire', 1: 'Dépositaire', 2: 'Dépositaire'
# , 3: 'Dépositaire', 4: 'Dépositaire', 5: 'Dépositaire', 6: 'Dépositaire', 7: 'Dépositaire', 8: 'Dépositaire', 9: 'Dépositaire', 10: 'Dépositaire', 11: 'Dép
# ositaire'}, 'Référence': {0: '22SEAV022666', 1: '22ROPC020529', 2: '22AGII003938', 3: '22NMSC022802', 4: nan, 5: nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: n
# an, 11: nan}, 'Date de valeur': {0: '31/05/2022', 1: datetime.datetime(2022, 3, 6, 0, 0), 2: '31/05/2022', 3: '14/06/2022', 4: '13/06/2019', 5: '13/06/2019
# ', 6: '13/06/2019', 7: '13/06/2019', 8: '13/06/2019', 9: '13/06/2019', 10: '13/06/2019', 11: '13/06/2019'}, 'Libellé': {0: 'SCT - SEPA ALLER VIREMENT', 1:
# 'SCT - SEPA ALLER VIREMENT', 2: 'RACHAT OPCVM?1?CM-AM MON PREM PARTS', 3: "ECHELLES D'INTERETS", 4: 'SOUSCRIPTION OPCVM?37?STRATEGE FIN PATRIMOINE/CAP', 5:
#  nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Trans Code': {0: 'NMSC', 1: 'NMSC', 2: 'NSEC', 3: 'NCOM', 4: 'NSEC', 5: nan, 6: nan, 7: nan, 8: n
# an, 9: nan, 10: nan, 11: nan}, 'Commentaire': {0: '-', 1: '-', 2: '-', 3: '-', 4: '-', 5: '-', 6: '-', 7: '-', 8: '-', 9: '-', 10: '-', 11: '-'}, 'Date Rég
# ul.': {0: nan, 1: nan, 2: nan, 3: nan, 4: nan, 5: nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Solde réel Dépositaire au': {0: nan, 1: nan, 2:
# nan, 3: nan, 4: nan, 5: nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Total des suspens à réguliser Dépositaire': {0: -38598.62, 1: -38598.62, 2
# : -38598.62, 3: -38598.62, 4: -38598.62, 5: nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Solde théorique Dépositaire': {0: 38598.62, 1: 38598.6
# 2, 2: 38598.62, 3: 38598.62, 4: 38598.62, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}, 'Solde réel Valorisateur au': {0: nan, 1: nan, 2: nan,
#  3: nan, 4: nan, 5: nan, 6: nan, 7: nan, 8: nan, 9: nan, 10: nan, 11: nan}, 'Total des suspens à réguliser Valorisateur': {0: 69364.67, 1: 69364.67, 2: 693
# 64.67, 3: 69364.67, 4: 69364.67, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}, 'Solde théorique Valorisateur': {0: 69364.67, 1: 69364.67, 2: 6
# 9364.67, 3: 69364.67, 4: 69364.67, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}, 'Check': {0: -30766.05, 1: -30766.05, 2: -30766.05, 3: -30766
# .05, 4: -30766.05, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0}}

# present date and time
present_date_time = datetime.now()
present_date_time_in_format = present_date_time.strftime("%d/%m/%Y %H:%M:%S")
print("present_date_time_in_format:",present_date_time_in_format)
today_date = date.today()
today_date_in_format = today_date.strftime("%d/%m/%Y")
print("today_date_in_format", today_date_in_format)

html_string = '''
<!DOCTYPE html>
<html>
  <head><title>Sample pdf</title></head>
  <link rel="stylesheet" type="text/css" href="df_style.css"/>
  <body>
     <header style="display: flex; justify-content: space-around;">
        <div>
            <img src="sample-912.jpg" style="transform:rotate(90deg); width:20px;">
        </div>
        <div>
            <h2>Rapprochement Bancaire au {0}</h2>
        </div>
        <div>
            <p>Report run on {1}</p>
        </div>
    </header>
    <br>
    <br>
     <br>
  </body>
</html>
'''.format(today_date_in_format, present_date_time_in_format)

resultKeys = ['Date comptable', 'Montant Dépositaire', 'Montant Valorisateur', 'Origine', 'Référence', 'Date de valeur',
              'Libellé', 'Trans Code', 'Commentaire', 'Date Régul.']


# unique values of fond are stored in these dictionary
duplicate_fond_values={}
for fond in resultDict['Fonds:'].values():
    print(fond)
    if fond in duplicate_fond_values:
        # print("found:",fond)
        # print(duplicate_fond_values[fond])
        duplicate_fond_values[fond]+=1
    else:
        # print("new:",fond)
        duplicate_fond_values[fond]=1

# for i in duplicate_fond_values.values():
#     print("duplicate_fond_values:",i)

# fond values list
fond_keys_list=[]
for k in duplicate_fond_values.keys():
    if k not in fond_keys_list:
        fond_keys_list.append(k)

# for k in fond_keys_list:
#     print("k:",k)

# for date in resultDict['Date comptable']:
#     print("date in Date comptable:",date)

fond_index=-1
index=-1
for value in duplicate_fond_values.values():
    fond_index+=1
    print("fond_index:",fond_keys_list[fond_index])
    html_string += """
                 <table style="min-width: 100%;">
        <thead>
            <tr>
                <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              -transform: uppercase;
              letter-spacing: 2%; padding:0.4% 4px; text-align: left; width:7%;">Fonds : 
            </th> 
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              -transform: uppercase;
              letter-spacing: 2%; padding:0.4% 4px; text-align: left; width:8%;"> {0} 
            </th>
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-align: initial;
              letter-spacing: 2%; padding: 0; text-align: left; width:20%;">
              <span>
                HELIUM OPPORTUNITIES-CAD 5
              </span>
            </th>
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              /*text-transform: uppercase;*/
              letter-spacing: 2%; text-align: left; width:9%;">Compute Depositaire :
            </th>
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-transform: uppercase;
              letter-spacing: 2%; text-align: left; width:21.5%;">
            </th>
             <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-transform: uppercase;
              letter-spacing: 2%; text-align: left; width:9%;">CAD
            </th>
            </tr>
        </thead>
    </table>
    <table style="min-width: 100%;">
        <thead>
            <tr>
                <th style="color: black; font-size: 0.875rem;
              -transform: uppercase;
              letter-spacing: 2%; padding:0.5% 0; width:34.5%;
              text-align: initial;">{1} 
            </th> 
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-transform: uppercase;
              letter-spacing: 2%; width:12.5%; ">
            </th>
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              /*text-transform: uppercase;*/
              letter-spacing: 2%; text-align: left; width:12%;">Compte Valorisateur:
            </th>
            <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-transform: uppercase;
              letter-spacing: 2%; text-align: left; width:29%;">PA768051100022CAD
            </th>
             <th style="background-color: #0A5A97;
              color: white; font-size: 0.875rem;
              text-transform: uppercase;
              letter-spacing: 2%; text-align: left;">CAD
            </th>
            </tr>
        </thead>
    </table>
    <table style="border-collapse: collapse;">
        <thead>
            <tr>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:9.5%; border: 1px solid #dddddd">Data comptable</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:10.5%; border: 1px solid #dddddd">Montant Depositaire</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:9%; border: 1px solid #dddddd">Montant Valorisateur</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:6%; border: 1px solid #dddddd">Origine</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:11.5%; border: 1px solid #dddddd">Reference</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:12%; border: 1px solid #dddddd">Date de valeur</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:15%; border: 1px solid #dddddd">Libelle</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:5%; border: 1px solid #dddddd">Trans Code</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; width:9%; border: 1px solid #dddddd">Commentaire</th>
                <th style="background-color:#0A5A97; color:white; text-align:left; border: 1px solid #dddddd">Date Regul.</th>
            </tr>
        </thead>
        <tbody>
        </tbody>
    </table>
              """.format(fond_keys_list[fond_index],
                         fond_keys_list[fond_index])
    # error in this loop
    print("value:",value)
    for i in range(0,value):
        index+=1
        print("i:",i)
        # print("index:",index)
        # print(resultDict['Date comptable'][index])
        html_string += """
                <table ">
                    <tr>
                        <td style="text-align:left; width:9.5%;">{0}</td>
                        <td style="text-align: right; text-align:left; width:content;">{1}</td>
                        <td style="text-align: right; text-align:left; width:9%;">{2}</td>
                        <td style="">{3}</td>
                        <td style="">{4}</td>
                        <td style="">{5}</td>
                        <td style="">{6}</td>
                        <td style="">{7}</td>
                        <td style="">{8}</td>
                        <td style="">{9}</td>
                    </tr>
                </table>
                    """.format(resultDict['Date comptable'][index],
                               resultDict['Montant Dépositaire'][index],
                               resultDict['Montant Valorisateur'][index],
                               resultDict['Origine'][index],
                               resultDict['Référence'][index],
                               resultDict['Date de valeur'][index],
                               resultDict['Libellé'][index],
                               resultDict['Trans Code'][index],
                               resultDict['Commentaire'][index],
                               resultDict['Date Régul.'][index]
                               )

with open('sample_5_aug.html', 'w') as f:
    f.write(html_string)




