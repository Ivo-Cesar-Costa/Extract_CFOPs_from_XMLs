from tkinter import *
from tkinter import ttk
from bs4 import BeautifulSoup
import pandas as pd
import glob
import time

class Amalgama(object):    
    def __init__(self, i): 
        
        self.font = ('Arial', '20', 'bold')
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
        #------------------------------------------- extract what was typed in Entry --------------------------------
        self.nome_arq = self.form.get().upper()         
        #------------------------------------------- condition if empty ---------------------------------------------
        if (self.nome_arq == ''):
            self.resultado['text'] = 'Digite um Nome de Arquivo'
        # --------------------------------------- condition if fulfilled --------------------------------------------
        else:
            try:
            #------------------------------------------ read directory ----------------------------------------------
                directory = sorted(glob.glob("directory path\*.xml"))
                if directory == []:
                    self.resultado['text'] = "Verifique se Há Arquivos XML no Diretório"
                else:
                    my_progress = ttk.Progressbar(self.frame3, orient=HORIZONTAL, length=600, mode='determinate',
                                              style='text.Horizontal.TProgressbar')
                    my_progress.pack(pady=30, expand=True)
                #------------------------------------- list for values ----------------------------------------------
                    list_nfs = []
                    list_values =[]
                #---------------------------------------- loop through directory ------------------------------------
                    for i in directory:
                        with open(i, 'r') as f: 
                            data = f.read()
                            Bs_data = BeautifulSoup(data, 'xml')
                #----------------------------------------- find the tags --------------------------------------------
                            val_tot_not = Bs_data.find_all('ValorTotalNota')
                            num_nf = Bs_data.find_all('NumeroNF')
                #--------------------------------------- extract text from tag NumeroNF -----------------------------
                            for j in num_nf:
                                var_n = int(j.text)
                                list_nfs.append(var_n)
                #------------------------------------- extract text from  tag ValorTotalNota  -----------------------
                            for k in val_tot_not:
                                val = float(k.text)
                                list_values.append(val)
                                
                #------------------------------------------- convert to DataFrame -----------------------------------
                        my_progress['value'] += (1/len(directory))*100                                                        
                        style.configure('text.Horizontal.TProgressbar',text='{:g} %'.format(round(my_progress['value'])))                            
                        self.frame.update_idletasks()                            
                        self.resultado['text'] = "...gerando arquivo {}.xlsx".format(self.nome_arq)
                        time.sleep(0.5)

                    df = pd.DataFrame(list(zip(list_nfs,list_values)), columns = ['NUMERO NF','VALOR TOTAL NOTA'])
                    df = df.sort_values(by='NUMERO NF')
                    df.reset_index(inplace=True, drop=True)
                    df = df.append(df[['VALOR TOTAL NOTA']].sum().rename('SOMA').fillna('-'))
                    df.loc['SOMA', 'NUMERO NF'] = ' QtdeNF : {}'.format(df['NUMERO NF'].count())
                #--------------------------------------- save as xlsx -----------------------------------------------
                    df.to_excel('{}.xlsx'.format(self.nome_arq))
                    self.form.delete(0, END)
                    self.resultado['text'] = 'Arquivo {}.xlsx Criado com Sucesso'.format(self.nome_arq)
                    my_progress.pack_forget()                    
            except:               
                self.resultado['text'] = "\tAlgo deu Errado :( "
                

        
i = Tk()

# ------------------------------------------- full screen window ----------------------------------------------------
# width= i.winfo_screenwidth()  
# height= i.winfo_screenheight() 
# i.geometry("%dx%d" % (width, height))

e = Amalgama(i)

style = ttk.Style(i)

style.layout('text.Horizontal.TProgressbar', 
             [('Horizontal.Progressbar.trough',
                {'children': [('Horizontal.Progressbar.pbar', 
                               {'side': 'left', 'sticky': 'ns'})],
                 'sticky': 'nswe'}),
                ('Horizontal.Progressbar.label', {'sticky': ''})])
                                         
style.configure('text.Horizontal.TProgressbar', text='0 %')


i.title("XML SUM")

i.geometry = ("600x600")

i.mainloop()