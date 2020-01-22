# -*- coding: utf-8 -*-
"""
Created on Sat Jun  8 19:51:07 2019

@author: Faster-PC
"""

import openpyxl, os
import pandas as pd
from selenium import webdriver

def Coleta(tipo):
    if tipo == 1:
        base = pd.read_csv('DadosD.csv')
    elif tipo == 2:
        base = pd.read_excel('DadosD.xlsx')
    palavras = base.iloc[:,0]
    palavras = list(palavras)
    return palavras


artigos = []
classes = []

tipo = int(input("Qual o tipo do arquivo? [1]csv , [2]xlsx : "))
palavras = Coleta(tipo)

browser = webdriver.PhantomJS()
for i in range (len(palavras)):
    browser.get("https://de.pons.com/%C3%BCbersetzung?q="+palavras[i]+"&l=dept&in=&lf=de&qnac=")
    print(palavras[i])
    try:
        artigo = browser.find_element_by_class_name('genus')
        if artigo.text == "m":
            artigos.append("Der")
        elif artigo.text == "f":
            artigos.append("Die")
        elif artigo.text == "nt":
            artigos.append("Das")
        else:
            artigos.append("ERROR")
        print("Artigo: %s" %artigo.text)
    except:
        artigos.append("")
        
    try:
        classe = browser.find_element_by_class_name('wordclass')
        classes.append(classe.text)
        print("Classe: %s\n" %classe.text)
    except:
        classes.append("ERROR")

vetorFinal = []
for i in range(len(palavras)):
    vetorFinal.append([palavras[i],artigos[i],classes[i]])
    
browser.close()

workbook = openpyxl.Workbook()

for i in range (len(vetorFinal)):
    workbook.active.append(vetorFinal[i])

os.chdir('C:\\Users\\Faster-PC\\MyPythonFiles')
workbook.save('DadosD2.xlsx')
    
    
    
    
