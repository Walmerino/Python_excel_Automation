#O app baixadepagtos.py automatiza a criação de uma tabela contendo as 
#colunas desejadas e faz as somas necessarias sobre o relatório dado 
#pelo sistema. 

import PySimpleGUI as sg                        

# Define the window's contents
layout = [  [sg.Text("Carregue o relatorio consulta.csv")],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text("Digite o nome do arquivo de saida Baixa_dd-mm.csv")],
            [sg.Input()],
            [sg.Button('Carregar')] ]

# Create the window
# Part 3 - Window Defintion

window = sg.Window('Baixa', layout)

# Display and interact with the Window
# Part 4 - Event loop or Window.read call

event, values = window.read()

def csv_import(csv):

    from csv import reader
    open_file = open(csv)
    read = reader(open_file, dialect='excel', delimiter=';')
    data = list(read)
    return data

data = csv_import(values[0])

f = open(values[1], "w")

f.write("Baixa de Pagamento \n")


opcob = []
for row in data[1:-1]:
    
    opcob.append(row[13])
    
opcob = list(dict.fromkeys(opcob))    

f.write("Titulo; Devedor; CPF; PL; D.Pagamento; V.Pagamento; COB; Agente\n" )

for row in data[1:-1]:

    titulo = row[2]
    devedor = row[1]
    cpf = row[3]
    pl = row[6]
    dpagto = row[8]
    vpagto = row[9]
    cob = row[13]
    agente = row[14]
    parcela = int(row[6])

    f.write(titulo)
    f.write(";")
    f.write(devedor)
    f.write(";")
    f.write(cpf)
    f.write(";")
    f.write(pl)
    f.write(";")
    f.write(dpagto)
    f.write(";")
    f.write(vpagto)
    f.write(";")
    f.write(cob)
    f.write(";")
    f.write(agente)
    f.write("\n")

f.write("\n")

f.write("COB; Colchão; Novos; Total \n")
for agente in opcob:

    total = novos = colchao = ac = 0
    
    for row in data[1:-1]:

        parcela = int(row[6])
        valor = row[9].replace(',','.')
        valor = float(valor)
        
        if agente == row[13]:
            total += valor
            ac += 1
        
            if parcela < 2:
                novos += valor
            if parcela > 1:
                colchao += valor
                
    
    f.write(agente)
    f.write(";")
    f.write("%.2f" % colchao)
    f.write(";")
    f.write("%.2f" % novos)
    f.write(";")
    f.write("%.2f" % total)
    f.write("\n")
    
f.close()
window.close() 
