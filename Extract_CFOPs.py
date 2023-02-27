from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import glob
# -------------------------------------lê diretório -----------------------------------------
todos_xmls = sorted(glob.glob("Diretório com os XMLs"))

lista_listas_final = []
lista_base = float(0)
lista_total = float(0)
for n in todos_xmls:
    with open(n, 'r', encoding='utf-8') as f:
        nota_fiscal = []
        lista_5101 = []
        lista_5102 = []
        lista_5401 = []
        lista_5405 = []
        lista_5949 = []
        lista_6101 = []
        lista_6401 = []
        lista_5201 = []
        lista_1201 = []
        lista_1410 = [] 
        lista_icmr_5401 = []
        lista_icmr_6401 = []
        lista_icmr_1410 = []
        isolado_icmr_5401 = []
        isolado_icmr_6401 = []
        isolado_icmr_1410 = []
        data = f.read()
        Bs_data = BeautifulSoup(data, 'xml')
       
        for w in Bs_data.find_all('nNF'):
                nota_fiscal.append(w.text)
        
        for i in Bs_data.find_all('prod'):
            for j in i.CFOP:
                for k in i.vProd:
                    if i.CFOP.text == '5101':
                            lista_5101.append(float(k)) 
                    elif i.CFOP.text == '5102':
                            lista_5102.append(float(k))                            
                    elif i.CFOP.text == '5401':
                            lista_5401.append(float(k))
                            for t in Bs_data.find_all('vST'):
                                    lista_icmr_5401.append(float(t.text))
                                    isolado_icmr_5401 = max(lista_icmr_5401)                                    
                    elif i.CFOP.text == '5405':
                            lista_5405.append(float(k))                            
                    elif i.CFOP.text == '5949':
                            lista_5949.append(float(k))                            
                    elif i.CFOP.text == '6101':
                            lista_6101.append(float(k))                            
                    elif i.CFOP.text == '6401':
                            lista_6401.append(float(k))
                            for t in Bs_data.find_all('vST'):
                                lista_icmr_6401.append(float(t.text))
                                isolado_icmr_6401 = max(lista_icmr_6401)
                    elif i.CFOP.text == '5201':
                            lista_5201.append(float(k))                            
                    elif i.CFOP.text == '1201':
                            lista_1201.append(float(k))                            
                    elif i.CFOP.text == '1410':
                            lista_1410.append(float(k))
                            for t in Bs_data.find_all('vST'):
                                lista_icmr_1410.append(float(t.text))
                                isolado_icmr_1410 = max(lista_icmr_1410)
                                
                    if (isolado_icmr_5401 == []):
                        lista_icmr_5401.append(float(0))
                        isolado_icmr_5401 = lista_icmr_5401[0]
                        
                    if (isolado_icmr_6401 == []):
                        lista_icmr_6401.append(float(0))
                        isolado_icmr_6401 = lista_icmr_6401[0]
                        
                    if (isolado_icmr_1410 == []):
                        lista_icmr_1410.append(float(0))
                        isolado_icmr_1410 = lista_icmr_1410[0]          
                      

                    lista_listas_somadas = [nota_fiscal[0], sum(lista_5101), sum(lista_5102), 
                                            (sum(lista_5401) + float(isolado_icmr_5401)), float(isolado_icmr_5401), sum(lista_5401),
                                            (sum(lista_5101) + sum(lista_5102) + (sum(lista_5401) + float(isolado_icmr_5401)) + 
                                             sum(lista_5405) + sum(lista_5949) + sum(lista_6101) + 
                                             (sum(lista_6401) + float(isolado_icmr_6401)) + sum(lista_5201)), 
                                            sum(lista_5405), sum(lista_5949), sum(lista_6101), 
                                            (sum(lista_6401) + float(isolado_icmr_6401)), float(isolado_icmr_6401), sum(lista_6401),
                                            lista_total, sum(lista_5201), sum(lista_1201),
                                            (sum(lista_1410) + float(isolado_icmr_1410)),float(isolado_icmr_1410), sum(lista_1410),
                                            lista_total]
                    
            
        lista_listas_final.append(lista_listas_somadas)
df = pd.DataFrame(lista_listas_final, columns = ['NF','5101','5102','5401', 'ICMR','Base','TOTAL','5405', '5949', 
                                                        '6101','6401','ICMR','Base','TOTAL', '5201','1201', '1410',
                                                'ICMR','Base','TOTAL'] )


df.reset_index(drop=True, inplace=True)
df.loc['Total',:]= df.sum(axis=0)
df.iloc[-1,0] = 'SOMA'

df.to_excel('Nome do Arquivo.xlsx')
df