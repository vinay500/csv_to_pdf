import pandas as pd

df = pd.DataFrame(pd.read_excel("pdf.xlsx", engine='openpyxl'))
print("df:")
print(df)
resultsData = df.values.tolist()
print("resultsData:")
print(resultsData)
resultDict = df.to_dict()
print()
print()
print("resultDict:")
print(resultDict)
html_string = '''
<html>
  <head><title>Sample pdf</title></head>
  <link rel="stylesheet" type="text/css" href="df_style.css"/>
  <body>
     <header style="display: flex; justify-content: space-around;">
        <div>
            <img src="sample-912.jpg" style="transform:rotate(90deg); width:20px;">
        </div>
        <div>
            <h2>Rapprochement Bancaire au 31/03/2020</h2>
        </div>
        <div>
            <p>Report run on 03/11/2020 18:58:23</p>
        </div>
    </header>
    <br>
    <br>
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
              letter-spacing: 2%; padding:0.4% 4px; text-align: left; width:8%;"> PA7680 
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
              text-align: initial;">PA7680 
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
     <br>
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
  </body>
</html>
'''
# print()
# print("resultsData[0]:",resultsData[0])
# print("resultsData[1]:",resultsData[1])
# print()
# print("resultsData[0][0]:",resultsData[0][0])
# print()
# print()
# print("resultDict[0]:",resultDict[0])
# print("resultDict[1]:",resultDict[1])
# print()
# print("resultDict[0][0]:",resultDict[0][0])
# print()
print()
print()
# print("Date comptable")
# print(resultDict['Date comptable'])
# print(resultDict['Date comptable'].keys())
# print(resultDict['Date comptable'].values())
# print("Date comptable values")
# for values in resultDict['Date comptable'].values():
#     print(values)
# print("Montant Dépositaire values")
# for values in resultDict['Montant Dépositaire'].values():
#     print(values)
# print("resultDict.keys():")
# resultData={}
# for key in resultDict.keys():
#     print(key,":")
#     for value in resultDict[key].values():
#         print(value)
#         resultData[key] = value
# print(resultDict.keys())
# print("resultData values:")
# print(resultDict.values())
# resultData={key:value for key,value in enumerate(resultDict)}

# for key in resultDict.keys():
#     # print(key,":")
#     if key in ['Date comptable', 'Montant Dépositaire', 'Montant Valorisateur', 'Origine', 'Référence', 'Date de valeur',
#                'Libellé', 'Trans Code', 'Commentaire', 'Date Régul.']:
#         for value in resultDict[key].values():
#             resultData[key] = value

# print("resultData['Date comptable'].values:")
# print(type(resultData['Date comptable']))
# print(type(resultDict['Date comptable']))
# for values in resultData['Date comptable'].values():
#     print(values)
# print()
# print()
# print("data")
# for data in resultDict:
#     print(data)
# print()
# print()
# for data in resultDict:
#     print(resultDict[data].keys())
#     print(resultDict[data].values())

# resultKeys=['Date comptable', 'Montant Dépositaire', 'Montant Valorisateur', 'Origine', 'Référence', 'Date de valeur',
#                'Libellé', 'Trans Code', 'Commentaire', 'Date Régul.']
resultData = {}
# for key in resultKeys:
#     for value in resultDict[key].values():
#         resultData[key]=value
#         print(key,":",resultData[key])

# for key,value in enumerate(resultData):
#     print(key,":",value)
# print(resultDict.keys())

resultKeys = ['Date comptable', 'Montant Dépositaire', 'Montant Valorisateur', 'Origine', 'Référence', 'Date de valeur',
              'Libellé', 'Trans Code', 'Commentaire', 'Date Régul.']

for index in resultDict['Date comptable'].keys():
    for key in resultKeys:
        print(resultDict[key][index])

for index in resultDict['Date comptable'].keys():
    html_string += """
            <tr style="border: 1px solid #dddddd;">
                        <td style="font-size: 0.8rem;">{0}</td>
                        <td style="text-align: right; border: 1px solid #dddddd; font-size: 0.8rem; ">{1}</td>
                        <td style="text-align: right; border: 1px solid #dddddd; font-size: 0.8rem;">{2}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{3}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{4}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{5}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{6}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{7}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{8}</td>
                        <td style="border: 1px solid #dddddd; font-size: 0.8rem;">{9}</td>
            </tr>
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

with open('sample1.html', 'w') as f:
    f.write(html_string)

below_desc_keys = ['Solde réel Dépositaire au', 'Total des suspens à réguliser Dépositaire',
                   'Solde théorique Dépositaire', 'Check', 'Solde réel Valorisateur au',
                   'Total des suspens à réguliser Valorisateur', 'Solde théorique Valorisateur']

below_desc_values = []
for key in below_desc_keys:
    below_desc_values.append(resultDict[key][0])
print("below_desc_values:")
print(below_desc_values[0])
print(below_desc_values[2])
if below_desc_values:
    # print(value)
    html_string += """
        <div id="below-description" style="display:flex; justify-content:space-around; margin-top: 10%;font-weight:600;">
            <div id="left-below-description" style="display:flex;justify-content:space-between;width:35%;">
                <div id="left-below-description-1">
                    <div>
                        <span>Solde reel  Dépositaire au</span>
                    </div>
                    <div>
                        <span>Total des suspens à réguliser Dépositaire</span>
                    </div>
                    <div>
                        <span>Solde theorique Dépositaire </span>
                    </div>
                    <div>
                        <span>check</span>
                    </div>
                </div>
                <div id="left-below-description-2">
                    <div>
                        <span>{0}</span>
                    </div>
                    <div>
                        <span>{1}</span>
                    </div>
                    <div>
                        <span>{2}</span>
                    </div>
                    <div>
                        <span>{3}</span>
                    </div>
                </div>
            </div>
            <div id="right-below-description" style="display:flex;justify-content:space-between;width:35%;">
                 <div id="right-below-description-1">
                   <div>
                     <span>Solde reel Valorisateur au</span>
                    </div>
                     <div>
                        <span>Total des suspens à réguliser Valorisateur</span>
                    </div>
                     <div>
                        <span>Solde théorique Valorisateur</span>
                    </div>
                 </div>
                 <div id="right-below-description-2">
                   <div>
                     <span>{4}</span>
                    </div>
                     <div>
                        <span>{5}</span>
                    </div>
                     <div>
                        <span>{6}</span>
                    </div>
                 </div>
            </div>
         </div>
    """.format(below_desc_values[0], below_desc_values[1], below_desc_values[2], below_desc_values[3],
               below_desc_values[4], below_desc_values[5], below_desc_values[6], )

with open('sample1.html','w') as f:
    f.write(html_string)

