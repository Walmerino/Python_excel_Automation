def csv_import(csv):
    #Abre o csv
    #Opens csv file
    from csv import reader
    open_file = open(csv)
    read = reader(open_file, dialect='excel', delimiter=';')
    data = list(read)
    return data


def agentes(data, x):
    #Cria uma lista com os nomes dos agentes
    #Creates a unique agents name list
    agente = []
    for row in data[1:-1]:
        
        agente.append(row[x])
        
    agente = list(dict.fromkeys(agente))
    
    return agente


def desempenho(data1):
    #Mostra o desempenho dos agentes em tempo falado e chamadas recebidas
    #Shows the performance of the agents, time spoken and received calls
    f.write("Desempenho dos agentes (ligações diárias)\n\n")

    for row in data1[2:]:
        
        f.write(row[1])
        f.write(";")
        f.write(row[9])
        f.write(";")
        f.write(row[15])
        f.write(";")
        f.write("\n")

    f.write("\n\n")

    
def baixa(data2):
    #Calcula os contratos novos, colchao e total.
    #Calculates the total, new contracts and subsequent payments
    f.write("Baixa de pagamento\n\n")
    f.write("Agente; Novos; Colchão; Total; AC\n")
    agente = agentes(data2, 14)
    
    for i in agente:
        
        agente = i

        total = colchao = novos = ac = 0
        
        for row in data2[1:-1]:


            parcela = int(row[6])
            valor = row[9].replace(',','.')
            valor = float(valor)
            cob = int(row[13])

            if agente == row[14]:
                ac += 1
                total += valor
                if parcela < 2:
                    novos += valor
                if parcela > 1:
                    colchao += valor
            
            
        f.write(agente)
        f.write(";")
        f.write("\t%.2f" % novos)
        f.write(";")
        f.write("\t%.2f" % colchao)
        f.write(";")
        f.write("%.2f" % total)
        f.write(";")
        f.write(str(ac))
        f.write("\n")
        

def acomp(data3):
    #Mostra os contratos fechados no dia e os valores que eles amontam
    #Shows the daily contracts and amounted values
    f.write("\n")
    f.write("Contratos diarios por agente e valor total acordado\n\n")
    f.write("Agente; Valor total dia; Contratos feitos dia; Media de acordo \n")
    agente2 = agentes(data3, 40)
    
    for i in agente2:

        agente = i
        
        total = todos = ac = 0
        for row in data3[1:-1]:

            parcela = int(row[3])
            valor = row[5].replace(',','.')
            valor = float(valor)
            todos += valor

            if agente == row[40]:
                ac += 1
                total += valor
                
        f.write(agente)
        f.write(";")
        f.write("%.2f" % total)
        f.write(";")
        f.write(str(ac))
        f.write(";")
        f.write("%.2f" % (total/ac))
        f.write(";")
        f.write("\n")

                    
    f.write("Valor total de contratos;")
    f.write("%.2f" % todos)
    f.write("\n\n")

#Main

#Cria a janela
#Creates the window
import PySimpleGUI as sg

layout = [  [sg.Text("Carregue RelDesempenho")],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text("Carregue consulta")],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text("Carregue Acomp")],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text("Digite RELATORIO_dd-mm.csv")],
            [sg.Input()],
            [sg.Button('Carregar')] ]

window = sg.Window('Syntesis', layout)

event, values = window.read()

#Importa os arquivos csv
#Imports the csv files
data1 = csv_import(values[0]) 
data2 = csv_import(values[1])
data3 = csv_import(values[2])

#Recebe o nome do arquivo que será usado para gravar
#os resultados do APP
#Gets the file output name
f = open(values[3], "w")

#Mostra a data do documento
#Shows document date
dia = str(data1[1])
f.write("Relatorios data")
f.write(dia)                            #Data do documento
f.write("\n\n")

#Funções chave do APP
#Apps key functions
desempenho(data1)
baixa(data2)
acomp(data3)

f.close()
window.close()


