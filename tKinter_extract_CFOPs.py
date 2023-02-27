from tkinter import *
from tkinter import ttk
from bs4 import BeautifulSoup
import pandas as pd
import glob
import time

class Amalgama(object):    
    def __init__(self, i): 
        
        self.font = ('Arial', '20', 'bold')
        #fonte=('Arial', '14', 'bold')
        self.frame = Frame(i)
        self.frame1 = Frame(self.frame)
        self.frame2 = Frame(self.frame,padx=15, pady=55)
        self.frame3 = Frame(self.frame)
        self.frame4 =Frame(self.frame)
                
        self.l_p = Label(self.frame1, text = "Programa para Criação de Arquivo",
                         width=40, height=2, font = self.font)
        self.s_l = Label(self.frame2, text = 'Salvar Arquivo como: ', font=('Arial',18,'bold'))
                   
        self.form = Entry(self.frame2)
        
        self.execProg = Button(self.frame3, text = 'CRIA ARQUIVO',  
                               font = ('Arial',12,'bold'), command = self.Salva)
        self.resultado = Label(self.frame4, text ="", width=40, height=4,
                                font = ('Arial',20,'bold'))
        
        self.frame.pack()
        self.frame1.pack()
        self.frame2.pack()
        self.frame3.pack()
        self.frame4.pack()
        self.l_p.pack(side=LEFT)
        self.s_l.pack(side=LEFT)
        self.form.pack(side=LEFT)
        self.execProg.pack()
        self.resultado.pack()
        
    def Salva(self):
        #------------------------------------extrai o que foi digitado em Entry -------------------------
        self.nome_arq = self.form.get().upper()         
        #------------------------------------condição se vazio ------------------------------------------
        if (self.nome_arq == ''):
            self.resultado['text'] = 'Digite um Nome de Arquivo'
        # ----------------------------------condição preenchido -----------------------------------------
        else:
            try:
            #------------------------------------------lê diretório -----------------------------------------
                todos_xmls = sorted(glob.glob("Diretório com os XMLs"))
                if todos_xmls == []:
                    self.resultado['text'] = "Verifique se Há Arquivos XML no Diretório"
                else:
                    my_progress = ttk.Progressbar(self.frame3, orient=HORIZONTAL, length=600, mode='determinate',
                                              style='text.Horizontal.TProgressbar')
                    my_progress.pack(pady=30, expand=True)
                #-------------------------------------lista para valores ----------------------------------------
                    lista_listas_final = []
                    lista_base = float(0)
                    lista_total = float(0)
                    #----------------------------------------percorre diretório -------------------------------------
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
                            
                        my_progress['value'] += (1/len(todos_xmls))*100                                                        
                        style.configure('text.Horizontal.TProgressbar',text='{:g} %'.format(round(my_progress['value'])))                            
                        self.frame.update_idletasks()                            
                        self.resultado['text'] = "...gerando arquivo {}.xlsx".format(self.nome_arq)
                        time.sleep(0.01)
                        
                    df = pd.DataFrame(lista_listas_final, columns = ['NF','5101','5102','5401', 'ICMR','Base','TOTAL','5405', '5949', 
                                                        '6101','6401','ICMR','Base','TOTAL', '5201','1201', '1410',
                                                'ICMR','Base','TOTAL'] )
                                 
                    df.reset_index(drop=True, inplace=True)
                    df.loc['Total',:]= df.sum(axis=0)
                    df.iloc[-1,0] = 'SOMA'
                    df.to_excel('{}.xlsx'.format(self.nome_arq))
                    self.form.delete(0, END)
                    self.resultado['text'] = 'Arquivo {}.xlsx Criado com Sucesso'.format(self.nome_arq)
                    my_progress.pack_forget() 
            except:               
                self.resultado['text'] = "\tAlgo deu Errado :( "            

                
i = Tk()

e = Amalgama(i)

style = ttk.Style(i)

style.layout('text.Horizontal.TProgressbar', 
             [('Horizontal.Progressbar.trough',
                {'children': [('Horizontal.Progressbar.pbar', 
                               {'side': 'left', 'sticky': 'ns'})],
                 'sticky': 'nswe'}),
                ('Horizontal.Progressbar.label', {'sticky': ''})])
                                         
style.configure('text.Horizontal.TProgressbar', text='0 %')


i.title("EXTRACT XML")

i.geometry = ("600x600")

i.mainloop()                       