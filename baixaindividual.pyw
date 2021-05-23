#O app baixadepagtos.py automatiza a criação de uma tabela contendo as 
#colunas desejadas e faz as somas necessarias sobre o relatório dado 
#pelo sistema.

# Part 1 - The import
import PySimpleGUI as sg                        

# Define the window's contents
layout = [  [sg.Text("Carregue o relatorio consulta.csv")],
            [sg.Input(), sg.FileBrowse()],
            [sg.Text("Digite o nome do arquivo de saida Baixa_individual_dd-mm.csv")],
            [sg.Input()],
            [sg.Button('OK')] ]

# Create the window
# Part 3 - Window Defintion

window = sg.Window('Baixa Individual', layout)

# Display and interact with the Window
# Part 4 - Event loop or Window.read call

event, values = window.read()

def csv_import(csv):

    from csv import reader
    open_file = open(csv)
    read = reader(open_file, dialect='excel', delimiter=';')
    data = list(read)
    return data

#path = "C:\Users\DESK-30\Desktop\Relatiorios\Baixa_individual\"
data = csv_import(values[0])

f = open(values[1], "w")

opcob   = []
for row in data[1:-1]:

    opcob.append(row[14])                #creates a list of the agents names
                                         #List of unique agent names
opcob = list(dict.fromkeys(opcob))    

for cobrador in opcob:

    indf = open('%s_BAIXA.csv' % cobrador, 'w')
    f.write("Baixa de Pagamento \n")
    f.write("Confira os clientes e as datas de pagamento.\n\n")
    f.write("Titulo; Devedor; CPF; PL; D.Pagamento; V.Pagamento; COB; Agente\n" )
    indf.write("Baixa de Pagamento \n")
    indf.write("Confira os clientes e as datas de pagamento.\n\n")
    indf.write("Titulo; Devedor; CPF; PL; D.Pagamento; V.Pagamento; COB; Agente\n" )
    
    novos = colchao = total = valor = 0
    for row in data[1:-1]:
        
        titulo = row[2]
        devedor = row[1]
        cpf = row[3]
        pl = row[6]
        dpagto = row[8]
        vpagto = row[9].replace(',','.')
        valor = float(vpagto)
        cob = row[13]
        agente = row[14]
        parcela = int(row[6])

        if cobrador == agente:
        
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
            f.write(";")
            f.write("\n")
            
            indf.write(titulo)
            indf.write(";")
            indf.write(devedor)
            indf.write(";")
            indf.write(cpf)
            indf.write(";")
            indf.write(pl)
            indf.write(";")
            indf.write(dpagto)
            indf.write(";")
            indf.write(vpagto)
            indf.write(";")
            indf.write(cob)
            indf.write(";")
            indf.write(agente)
            indf.write(";")
            indf.write("\n")
 
            if parcela > 1:
                colchao += valor
            else:
                novos += valor
    
    f.write("Colchão")
    f.write(";")
    f.write("%.2f" % colchao)
    f.write("\n")
    
    indf.write("Colchão")
    indf.write(";")
    indf.write("%.2f" % colchao)
    indf.write("\n")

    f.write("Novos")
    f.write(";")
    f.write("%.2f" % novos)
    f.write("\n")
    f.write("\n\n")
    
    indf.write("Novos")
    indf.write(";")
    indf.write("%.2f" % novos)
    indf.write("\n")
    indf.write("\n\n")

f.write("\n")
indf.close()
f.close()
window.close()  
