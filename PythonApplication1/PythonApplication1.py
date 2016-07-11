import sys
#import csv #paranaue de banco de dados
import pyodbc #banco da M$
import datetime #suporte ao tipo datetime
#import TkTreectrl as treectrl #criar treeview
import tkinter as tk #GUI
import tkinter.ttk as ttk #GUI nao horrivel
import tkinter.messagebox
import tkinter.filedialog
from tkinter import * #GOTTA IMPORT THEM ALL

TITLE_FONT = ("Segoe UI Light", 18, "bold")

#Spinbox com tema
class Spinbox(ttk.Widget):
    def __init__(self, master, **kw):
        ttk.Widget.__init__(self, master, 'ttk::spinbox', kw)

#acesso ao banco
def obter(inicial, dataIni, dataFin):
    MDB = 'database.mdb'
    DRV = '{Microsoft Access Driver (*.mdb)}'
    PWD = 'pw'
    global rows
    if inicial == True:
        try:
            SQL = 'SELECT dateandtime,millitm FROM FloatTable'
        except:
            print('SQL falhou')
            messagebox.showwarning('SQL', 'FALHA')
    else:
        try:
            SQL = 'SELECT DateAndTime,millitm FROM FloatTable WHERE DateAndTime BETWEEN #' + str(dataIni) + '# AND #' + str(dataFin) + '#'
        except:
            print('SQL falhou')
    #conexao com o banco
    con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, MDB, PWD))
    cur = con.cursor()
    rows = cur.execute(SQL).fetchall() #joga resultados na lista rows
    cur.close()
    con.close()

#main
def main():
    
    print('Number of arguments:', len(sys.argv), 'arguments.')
    try:
        tk.messagebox.showinfo('OK', 'Var: ' + sys.argv[1])
    except:
        tk.messagebox.showerror('Erro', 'Erro! Sem opcao')

    root = Tk()
    root.title("Impressao")
    #root.resizable(0,0) #bloquear redimensionamento
    root.configure(bg="#ffffff")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.minsize(425, 325)

    style = ttk.Style()
    style.layout('Treeview', [('Treeview.treearea', {'sticky':'news'})])
    style.theme_use('vista')
        
    def geraData():
        dataIni = datetime.datetime(int(dataIni_entry_ano.get()), int(dataIni_entry_mes.get()), int(dataIni_entry_dia.get()), int(dataIni_entry_hor.get()), int(dataIni_entry_min.get()))
        dataFin = datetime.datetime(int(dataFin_entry_ano.get()), int(dataFin_entry_mes.get()), int(dataFin_entry_dia.get()), int(dataFin_entry_hor.get()), int(dataFin_entry_min.get()))
        if dataFin < dataIni:
            print("ERROU")
        #print(str(dataIni)+'\n'+str(dataFin))teste  #teste
        obter(False, dataIni, dataFin)
        tree_entry.delete(*tree_entry.get_children()) #esvazia treeview
        for row in rows:
            tree_entry.insert('', 0, text=row[0], values=row[1]) #repopula a treeview

    def mandarUmSalve(file):
        f = tk.filedialog.asksaveasfile(mode='w', defaultextension=".html")
        #press f to pay respects
        if f is None:
            return
        else:
            f.write(file)
            f.close()

    def tibola():
        print("todo")
        #tk.messagebox.showinfo('OK', 'Press OK to OK')
        #header
        file = (        '<html>\n'
                        '    <head>\n'
                        '        <title>Relatorio de ') + titulo + ('</title>\n'
                        '        <style type="text/css">\n'
                        '            body{ font-family:"Helvetica","Segoe UI Light",; color:#232323;}\n'
                        '            table, tr, td{ border-bottom: 1px solid #ddd; text-align:center; margin-left:auto; margin-right:auto; }\n'
                        '            .title{ border: 1px solid #3f3f3f;}\n'
                        '            tr{ height: 50px;}\n'
                        '            tr, td{ padding:10px; text-allign:left}\n'
                        '        </style>'
                        '    </head>\n'
                        '    <body>\n'
                        '        <table>\n'
                        '            <tr class=title>\n'
                        '                <td>') + row1titulo + ('</td>\n'
                        '                <td>') + row2titulo + ('</td>\n'
                        '            </tr>\n')
        #conteudo
        for row in rows:
            file=file +('            <tr>\n'
                        '                <td>') + str(row[0]) + ('</td>\n'
                        '                <td>') + str(row[1]) + ('</td>\n'
                        '            </tr>\n')
        #footer
        file = file + ( '        </table>\n'
                        '    </body>\n'
                        '</html>\n')

        #print(file) #printa o arquivo a ser exportado
        mandarUmSalve(file)

    #mainframe

    mainframe = ttk.Frame(root)
    mainframe.grid(column=0, row=0, sticky=S + N + W + E)
    mainframe.columnconfigure(0, weight=1)

    #subframe para entradas
    spinframe = ttk.Frame(mainframe)
    spinframe.grid(sticky=N + W + E, column=0, row=0, columnspan=7, ipadx=5, ipady=5)
    spinframe.rowconfigure(0, weight=1)
    spinframe.columnconfigure(0, weight=1)

    ##
    #   Variaveis e entradas de data inicial
    ##
    ttk.Label(spinframe, text="Data Inicial").grid(column=0, row=1, sticky=W)
    #dia inicial
    dataIni_entry_dia = tk.Spinbox(spinframe, width=4, from_='01', to='31')
    dataIni_entry_dia.grid(column=1, row=1, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="/").grid(column=2, row=1, sticky=W)
    #mes inicial
    dataIni_entry_mes = tk.Spinbox(spinframe, width=4, from_='01', to='12')
    dataIni_entry_mes.grid(column=3, row=1, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="/").grid(column=4, row=1, sticky=W)
    #ano inicial
    dataIni_entry_ano = tk.Spinbox(spinframe, width=6, from_='2000', to='2100')
    dataIni_entry_ano.grid(column=5, row=1, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="Hora=").grid(column=6, row=1, sticky=W)
    #hora inicial
    dataIni_entry_hor = tk.Spinbox(spinframe, width=4, from_='00', to='23')
    dataIni_entry_hor.grid(column=7, row=1, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text=":").grid(column=8, row=1, sticky=W)
    #minuto inicial
    dataIni_entry_min = tk.Spinbox(spinframe, width=4, from_=00, to=59)
    dataIni_entry_min.grid(column=9, row=1, sticky=(W, E))

    ##
    #Variaveis e entradas de data final
    ##
    ttk.Label(spinframe, text="Data Final").grid(column=0, row=2, sticky=W)
    #dia final
    dataFin_entry_dia = tk.Spinbox(spinframe, from_='01', to='31', width=4)
    dataFin_entry_dia.grid(column=1, row=2, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="/").grid(column=2, row=2, sticky=W)
    #mes final
    dataFin_entry_mes = tk.Spinbox(spinframe, from_='01', to='12', width=4)
    dataFin_entry_mes.grid(column=3, row=2, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="/").grid(column=4, row=2, sticky=W)
    #ano final
    dataFin_entry_ano = tk.Spinbox(spinframe, from_='2000', to='2100', width=6)
    dataFin_entry_ano.grid(column=5, row=2, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text="Hora=").grid(column=6, row=2, sticky=W)
    #hora final
    dataFin_entry_hor = tk.Spinbox(spinframe, from_='00', to='23', width=4)
    dataFin_entry_hor.grid(column=7, row=2, sticky=(W, E))
    #separador
    ttk.Label(spinframe, text=":").grid(column=8, row=2, sticky=W)
    #minuto final
    dataFin_entry_min = tk.Spinbox(spinframe, from_='00', to='59', width=4)
    dataFin_entry_min.grid(column=9, row=2, sticky=(W, E))

    for child in spinframe.winfo_children(): child.grid_configure(padx=5, pady=5)

    #subframe para botoes
    botaoframe = ttk.Frame(mainframe)
    botaoframe.grid(column=5, row=3, columnspan=7, ipadx=5, ipady=5, sticky=N+E)
    botaoframe.columnconfigure(0, weight=1)
    #imprimir
    titulo = 'alguma coisa'
    row1titulo = 'Data/Hora'
    row2titulo = 'Numeros'
    ttk.Button(botaoframe, text="Salvar", command=tibola).grid(column=0, row=0, sticky=E, ipadx=5)
    #Botao que faz a checagem
    ttk.Button(botaoframe, text="Pesquisar", command=geraData).grid(column=1, row=0, sticky=E, ipadx=5)

    for child in botaoframe.winfo_children(): child.grid_configure(padx=5, pady=5)


    #subframe para nao dar m* na treeview
    treeframe = ttk.Frame(mainframe)
    treeframe.grid(column=0, row=4, columnspan=10, sticky=N+E+W+S)
    treeframe.columnconfigure(0, weight=1)
    treeframe.rowconfigure(0, weight=1)

    ##treeview
    tree_entry = ttk.Treeview(treeframe, height="20", selectmode="extended")
    tree_entry.grid(column=0, row=0, sticky=N+E+W+S)
    tree_entry.heading('#0', anchor='w', text='Data/Hora')
    tree_entry["columns"]=('one', 'two')

    tree_entry.column("one", width=100)
    tree_entry.column("two", width=100)
    tree_entry.heading('one', text='PV')
    tree_entry.heading('two', text='CV')
    #scroll da treeview
    scrollbar = ttk.Scrollbar(treeframe, orient='vertical', command=tree_entry.yview)
    scrollbar.grid(column=1, row=0, sticky=N+S+W)
    tree_entry.configure(yscroll=scrollbar.set)

    for child in mainframe.winfo_children(): child.grid_configure(ipadx=5, ipady=5)

    obter(True,'ayy','lmao') #cria rows

    for row in rows:
        tree_entry.insert('', 0, text=row[0], values=row[1]) #popula inicialmente

    dataIni_entry_dia.focus()
    root.bind('<Return>', obter)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.bind('<Escape>', lambda e: root.destroy())
    root.mainloop()

main()
