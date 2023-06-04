#!/usr/bin/python3
import re
import matplotlib.pyplot as plot
import pandas as pd

POL = {}
entrada = "../Obras.md/Camilo-A_Morgada_de_Romariz.md"

def criagraf(xl, y):
    plot.bar(xl, y, color="red")
    plot.show()
    plot.savefig("g.png")

def carrega_sentilex():
    with open("sentilexjj.txt") as f:
        for linha in f:
            linha = re.sub(r"ANOT=.*", "", linha)
            lista = re.split(r";",linha)
            if len(lista) == 5:
                pallemapos, flex, assina, pol, lixo = lista
                pal = re.split(r",",pallemapos)[0]
                pol = int( re.sub(r"POL:N[01]=", "", pol) )
                POL[pal] = pol
            elif len(lista) == 6:
                pallemapos, flex, assina, pol1, pol2, lixo = lista
                pal = re.split(r",",pallemapos)[0]
                pol1 = int( re.sub(r"POL:N[01]=", "", pol1) )
                pol2 = int( re.sub(r"POL:N[01]=", "", pol2) )
                POL[pal] = (pol1+pol2)/2 
            else:
                print(lista)
                exit()

def sentimento(frase):
    lp = re.findall(r"\w+", frase)
    ptotalneg = 0 
    ptotalpos = 0
    qp = 0
    qn = 0

    for p in lp:
        if p in POL:
            v = POL[p]
            if  v > 0: 
                ptotalpos += v
                qp += 1

            if  v < 0: 
                ptotalneg += -v
                qn += 1

    return (ptotalpos,qp,ptotalneg,qn, len(lp)) 

carrega_sentilex()

txt = open(entrada, encoding='utf8').read()
ptotalpos, qp,ptotalneg,qn,np= sentimento(txt)
Factor = ptotalpos / ptotalneg

listacap = re.split(r"#", txt)

saida = entrada+".csv"
fo = open (saida, "wt", encoding = "utf-8")

y = []
xl = []
data = [] 
n=0 

for cap in listacap:
    ptotalpos, qp,ptotalneg,qn,np= sentimento(cap)
    
    if np != 0:
        sentcap = ((ptotalpos-(ptotalneg*Factor))/np)*1000
    else:
        sentcap = 0  
    
    y.append(sentcap)
    xl.append(f"C{n}")
    data.append([n,len(cap), ptotalpos,qp,ptotalneg,qn,np,Factor, sentcap])  # Added sentcap here
    n=n+1

# Added 'sentcap' to the DataFrame columns
df = pd.DataFrame(data, columns=["Cap", "Nº carateres", "totpos", "quanpos", "totneg", "quanneg", "palavras", "rationeg", "sentcap"])
df.to_excel("output.xlsx", index=False)

criagraf(xl[1:], y[1:])

fo.close()




