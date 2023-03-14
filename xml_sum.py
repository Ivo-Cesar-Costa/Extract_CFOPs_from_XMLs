from bs4 import BeautifulSoup
import pandas as pd
import glob
# ------------------------------------- read directory ------------------------------------------
directory = sorted(glob.glob("directory path\*.xml"))
# -------------------------------- list for values ----------------------------------------------
list_nfs = []
list_values =[]
#------------------------------------ loop through directory ------------------------------------
for i in directory:
    with open(i, 'r') as f: 
        data = f.read()
        Bs_data = BeautifulSoup(data, 'xml')
# ------------------------------------ find the tags --------------------------------------------
        val_tot_not = Bs_data.find_all('ValorTotalNota')
        num_nf = Bs_data.find_all('NumeroNF')
#----------------------------------- extract text from tag NumeroNF -----------------------------
        for j in num_nf:
            var_n = int(j.text)
            list_nfs.append(var_n)
#--------------------------------- extract text from  tag ValorTotalNota ------------------------
        for k in val_tot_not:
            val = float(k.text)
            list_values.append(val)
#-------------------------------------- convert to DataFrame ------------------------------------
df = pd.DataFrame(list(zip(list_nfs, list_values)), columns = ['NUMERO NF','VALOR TOTAL NOTA'])
df = df.sort_values(by='NUMERO NF')
df.reset_index(inplace=True, drop=True)
df = df.append(df[['VALOR TOTAL NOTA']].sum().rename('SOMA').fillna('-'))
df.loc['SOMA', 'NUMERO NF'] = ' QtdeNF : {}'.format(df['NUMERO NF'].count())
#------------------------------------ save as xlsx -----------------------------------------------
df.to_excel('SOMA DOS XML.xlsx')
df