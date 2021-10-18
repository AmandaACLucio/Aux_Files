import numpy 
import itertools
import xlwings as xw

############################################ Inputs #############################################

primeira_coluna_inicio=int(input("qual o valor inicial de pas "))
primeira_coluna_fim=int(input("qual o valor final de pas "))
primeira_coluna_variacao=int(input("qual a variação de pas "))

segunda_coluna_inicio=int(input("qual o valor inicial de rotação "))
segunda_coluna_fim=int(input("qual o valor final de rotação "))
segunda_coluna_variacao=int(input("qual a variação de rotação "))

terceira_coluna_inicio=float(input("qual o valor inicial de corda "))
terceira_coluna_fim=float(input("qual o valor final de corda "))
terceira_coluna_variacao=float(input("qual a variação de corda "))

quarta_coluna_inicio=int(input("qual o valor inicial de passo "))
quarta_coluna_fim=int(input("qual o valor inicial de passo "))
quarta_coluna_variacao=int(input("qual a variação de passo "))

quinta_coluna_inicio=int(input("qual o valor inicial de velocidade "))
quinta_coluna_fim=int(input("qual o valor inicial de velocidade "))
quinta_coluna_variacao=int(input("qual a variação de velocidade "))

#Listas  com intervalo de variações
primeira_coluna = list(range(primeira_coluna_inicio, primeira_coluna_fim+1, primeira_coluna_variacao))

segunda_coluna = list(range(segunda_coluna_inicio, segunda_coluna_fim+1, segunda_coluna_variacao))

terceiro_coluna = list(numpy.around(numpy.arange(terceira_coluna_inicio, terceira_coluna_fim+(terceira_coluna_variacao/2), terceira_coluna_variacao),2))

quarta_coluna = list(range(quarta_coluna_inicio, quarta_coluna_fim+1, quarta_coluna_variacao))

quinta_coluna = list(range(quinta_coluna_inicio, quinta_coluna_fim+1, quinta_coluna_variacao))


#print(primeira_coluna)
#print(segunda_coluna)
#print(terceiro_coluna)
#print(quarta_coluna)
#print(quinta_coluna)


listas = [primeira_coluna, segunda_coluna, terceiro_coluna, quarta_coluna, quinta_coluna]

#Writing excel
excelFile=xw.Book('AAA Exercício AeroGerador 2021-01 RevLevi.xlsx')

#nome das colunas
excelFile.sheets['Aerogerador'].range(1,10).value="Pás"
excelFile.sheets['Aerogerador'].range(1,11).value="Rotação"
excelFile.sheets['Aerogerador'].range(1,12).value="Cordas"
excelFile.sheets['Aerogerador'].range(1,13).value="Passo"
excelFile.sheets['Aerogerador'].range(1,14).value="Velocidade"
excelFile.sheets['Aerogerador'].range(1,15).value="Empuxo"
excelFile.sheets['Aerogerador'].range(1,16).value="Torque"
excelFile.sheets['Aerogerador'].range(1,17).value="Potência"
excelFile.sheets['Aerogerador'].range(1,18).value="Coef. Pot."
			

linha_number = 0

for linha in itertools.product(*listas):
    
    lista_linha=list(linha)
    
    #.range(numero da linha, numero da coluna)
    #pegando valores da primeira sheet
    excelFile.sheets['Exercício (Sem indução)'].range(10,18).value=lista_linha[3]#passo médio
    excelFile.sheets['Exercício (Sem indução)'].range(11,18).value=lista_linha[2]#corda média
    excelFile.sheets['Exercício (Sem indução)'].range(12,18).value=lista_linha[0]#número de pás
    excelFile.sheets['Exercício (Sem indução)'].range(15,18).value=lista_linha[1]#rotações por minuto
    excelFile.sheets['Exercício (Sem indução)'].range(14,18).value=lista_linha[4]#velocidade incidente

    #print(lista_linha)
    for item in range(len(linha)):
        excelFile.sheets['Aerogerador'].range(linha_number+2,10+item).value=lista_linha[item]#Linhas parâmetros de entrada (Pás, Rotação, Corda, Passo, Velo inc)

    #Linhas resultados (Empuxo, Torque, Potência, Coef. Pot.)
    excelFile.sheets['Aerogerador'].range(linha_number+2,15).value= excelFile.sheets['Exercício (Sem indução)'].range(16,8).value
    excelFile.sheets['Aerogerador'].range(linha_number+2,16).value= excelFile.sheets['Exercício (Sem indução)'].range(17,8).value
    excelFile.sheets['Aerogerador'].range(linha_number+2,17).value= excelFile.sheets['Exercício (Sem indução)'].range(18,8).value
    excelFile.sheets['Aerogerador'].range(linha_number+2,18).value= excelFile.sheets['Exercício (Sem indução)'].range(19,8).value	

    linha_number+=1

excelFile.save()
excelFile.close()

