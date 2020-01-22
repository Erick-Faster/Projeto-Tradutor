# -*- coding: utf-8 -*-
"""
Created on Sat Jun  8  2019
Finished on Tue Jun 11 2019

###########FEATURES#############
-- Monta planilha de excel contendo
    -- Palavras inseridas em alemao
    -- Traducao das palavras
    -- Genero dos Substantivos
    -- 2 Exemplos de aplicacao
    -- Tipo (substantivo, verbo, etc...)
-- Extração de dados do Pons e Reverso Context
-- Formato de planilha para ser inserido no Memorion
-- Formatação para que os sites leiam umlauts e eszetts
-- Extração de dados em xlsx e csv
-- Busca por arquivo base, dando nome como entrada
-- Escolhe nome para arquivo de saida


@author: Faster-PC
"""

import openpyxl, os, re
import pandas as pd
from selenium import webdriver
from unidecode import unidecode

'''
###################################
Funcoes
##################################
'''

#Coleta o formato de arquivo especifico
def Coleta(nomeBase,tipo):
    if tipo == 1: #Se for um csv
        base = pd.read_csv(nomeBase+'.csv',encoding='latin-1') #Latin-1 para corrigir erro com caracteres
    elif tipo == 2: #Se for formato Excel
        base = pd.read_excel(nomeBase+'.xlsx')
    else:
        palavras = ['Tisch','Tasche','Auto']
        return palavras
    palavras = base.iloc[:,0] #Cliva a primeira coluna
    palavras = list(palavras) #Converte o DataFrame para Lista
    return palavras

#Converte caracteres estranhos
def Converte(palavras,idioma):
    regex = re.compile(r'[äöüÄÖÜß]') #Cita regras. Localiza caracteres entre []
    if idioma == '1':   
        for i in range(len(palavras)):
            Verificador = False #Criterio para manter o looping
            while Verificador == False: #Garante que todos os caracteres especiais sejam encontrados
                try:
                    mo = regex.search(palavras[i]) #Procura em 'palavras' de acordo com regra
                    aux = mo.group() #caractere especial encontrado
                    span = mo.span() #posicao do caractere especial
                    palavraAux = list(palavras[i]) #Transforma string em lista
                    
                    #Converte caractere especial em forma apropriada
                    if aux == 'Ä':
                        palavraAux[span[0]] = 'Ae'
                    elif aux == 'Ö':
                        palavraAux[span[0]] = 'Oe'
                    elif aux == 'Ü':
                        palavraAux[span[0]] = 'Ue'
                    elif aux == 'ä':
                        palavraAux[span[0]] = 'ae'
                    elif aux == 'ö':
                        palavraAux[span[0]] = 'oe'
                    elif aux == 'ü':
                        palavraAux[span[0]] = 'ue'
                    elif aux == 'ß':
                        palavraAux[span[0]] = 'ss'
                    else:
                        print('ERROR')
                        
                    palavras[i] = ''.join(palavraAux) #transforma lista em string de novo
                    print('Conversao de %s bem sucedido!'%palavras[i])
                    palavraAux.clear() #elimina lista
                except:
                    Verificador = True #Encerra busca
                    continue #Se nao encontrar, vai para o proximo caso
    else: #Para todos os outros idiomas
        for i in range(len(palavras)):
            palavras[i] = unidecode(palavras[i]) #Remove acentos e caracteres especiais
    return palavras

#Coleta Exemplos e Traducoes do Reverso Context
def Reverso(palavras,idiomaBase):
    
    if idiomaBase == 'de':
        idiomaB = 'deutsch'
    elif idiomaBase == 'fr':
        idiomaB = 'franzosisch'
    elif idiomaBase == 'en':
        idiomaB = 'englisch'
    elif idiomaBase == 'es':
        idiomaB = 'spanisch'
    
    exemplos = [] #Vetor temporario
    exemploFinal = [] #Vetor permanente
    traducoes = []
    traducaoFinal = []
    for i in range (len(palavras)):  #acao para cada palavra
        browser.get("https://context.reverso.net/%C3%BCbersetzung/"+idiomaB+"-portugiesisch/"+palavras[i]) #site no qual informacao eh extraida
        
        '''
        exemplos
        '''
        try:    
            frases = browser.find_elements_by_class_name('text') #Encontra todos os elementos de frases
            
            #Converte dados das frases de Web para String
            for j in range (len(frases)):
                exemplos.append(frases[j].text)
                
            #Elimina vazios existentes no vetor temporario
            for j in range (len(exemplos)):
                try:
                    exemplos.remove("") #Remove todos os vazios da string
                except:
                    break
            
            #Confere se nao ha Typo
            k = 0
            if exemplos[0] == 'Meinst Du:':
                k = 1
            
            #Separa frases desejadas
            exemplo = [exemplos[k],exemplos[k+1]," ~~ ",exemplos[k+2],exemplos[k+3]] #Seleciona as 2 primeiras frases
            
            #Une vetor em uma unica String
            stringExemplo = " | " #Separador entre cada elemento do vetor
            stringExemplo = stringExemplo.join(exemplo) #Transforma vetor em uma string unica
            
            #Adicionar string no vetor permanente
            exemploFinal.append(stringExemplo)
            print("Exemplo para %s processado!" %palavras[i])
            
            exemplos = [] #zera vetor temporario
        except:
            exemploFinal.append("ERROR")
        
        '''
        Traducoes
        '''
        
        try:
            traducaoWEB = browser.find_elements_by_class_name('translation')
            
            for j in range (len(traducaoWEB)):
                traducoes.append(traducaoWEB[j].text)
                
            #Elimina vazios existentes no vetor temporario
            for j in range (len(traducoes)):
                try:
                    traducoes.remove("") #Remove todos os vazios da string
                except:
                    break
            
            traducao = traducoes[0]
            traducaoFinal.append(traducao)
            print("Traducao adicionada: %s\n" %traducao)
            traducoes = []
        except:
            traducaoFinal.append("ERROR")
        
        
    return exemploFinal, traducaoFinal

#Coleta artigos classes e erros do site Pons
def Pons (palavras,idiomaBase):
    for i in range (len(palavras)): #Repete de acordo com a qtde de palavras
        browser.get("https://de.pons.com/%C3%BCbersetzung?q="+palavras[i]+"&l="+idiomaBase+"en&in=&lf=de&qnac=") #Entra no site PONS
        print(palavras[i])
        
        #Busca pelo genero
        try:
            artigo = browser.find_element_by_class_name('genus') #Busca genero
            if artigo.text == "m":
                artigos.append("Der")
            elif artigo.text == "f":
                artigos.append("Die")
            elif artigo.text == "nt":
                artigos.append("Das")
            else:
                artigos.append("ERROR")
            print("Artigo: %s" %artigo.text)
        except: #Comum quando nao eh um substantivo
            artigos.append("") #Nao retorna artigo nenhum
            
        #Busca pela classe/tipo da palavra (subst, verbo, adjetivo, etc)
        try:
            classe = browser.find_element_by_class_name('wordclass') #Busca classe
            classes.append(classe.text) #add classe
            print("Classe: %s\n" %classe.text)
        except:
            classes.append("ERROR")
        
        #Verifica a possibilidade de possiveis erros
        
        try:
            erro = browser.find_element_by_tag_name('strong') #Procura na tag <strong>
            erro = erro.text #atribui texto na variavel
            regex = re.compile(r'(Meinten Sie vielleicht:)\s(\w+)') #Cria regra para padrao
            mo = regex.search(erro) #procura padrao
            auxErro = mo.group(1) #Valor que sera except caso nao seja encontrado
            auxSugestao = mo.group(2) #Sugestao de palavra dada pelo Pons
            if auxErro == 'Meinten Sie vielleicht:': #Caso o erro seja positivo
                erros.append("WARNING -> %s"%auxSugestao) #Retorna erro com sugestao
            else:
                erros.append("") #Nao retorna nada
        except:
            erros.append("")
    return artigos, classes, erros

#Funcao que insere tudo em um vetor final e salva no Excel no formato FlashCards do Memorion
def SalvarExcel(nomeArquivo,palavrasFinais,traducoes,artigos,exemplos,classes,erros):
    vetorFinal = [] #Informacoes que irao para o Excel
    for i in range(len(palavras)):
        vetorFinal.append([palavrasFinais[i],traducoes[i],artigos[i],exemplos[i],classes[i],erros[i]]) #Add palavra, artigo, classe e exemplos
    
    workbook = openpyxl.Workbook() #Cria arquivo Excel
    
    for i in range (len(vetorFinal)): #Qtde de elementos do vetor final
        workbook.active.append(vetorFinal[i]) #Add vetor, linha por linha
    
    os.chdir('C:\\Users\\Faster-PC\\MyPythonFiles') #Seleciona Diretorio
    
    #Verifica se o arquivo ja existe
    savePoint = os.path.isfile('./'+nomeArquivo+'.xlsx')
    
    if savePoint == False: #Caso nao exista, salvara nele msm
        workbook.save(nomeArquivo+'.xlsx') #Salva Excel
        print('%s.xlsx criado com sucesso!'%nomeArquivo)
    else: #Caso ja exista
        save = 2 #Valor atribuido ao nome do arquivo
        saveStg = str(save) #Transforma int em String
         #Condicao de parada
        while savePoint == True: #Enquanto existir um arquivo igual
            savePoint = os.path.isfile('./'+nomeArquivo+saveStg+'.xlsx') #Busca arquivo com numero na frente
            if savePoint == False: #Se nao existir
                workbook.save(nomeArquivo+saveStg+'.xlsx') #Salva Excel com numero
                savePoint = False #Parou
                print('%s%s.xlsx criado com sucesso!'%(nomeArquivo,saveStg))
            else: #Se ainda existir
                save = save + 1 #Add um numero ao arquivo
                saveStg = str(save) #Transforma o numero em String

        
'''
############################################################
AQUI COMECA O MAIN
############################################################
'''

'''
Tipos de dados que serao extraidos
'''
palavrasFinais = []
artigos = []
classes = []
exemplos = []
traducoes = []
erros = []

'''
Questionario
'''

while True:
    nomeBase = input("Qual o nome do arquivo a ser formatado?: ")
    VerificaCSV = os.path.isfile('./'+nomeBase+'.csv')
    VerificaXLSX = os.path.isfile('./'+nomeBase+'.xlsx')
    if VerificaCSV == True and VerificaXLSX == False:
        tipo = 1
        break
    elif VerificaCSV == False and VerificaXLSX == True:
        tipo = 2
        break
    elif VerificaCSV == True and VerificaXLSX == True:
        tipo = int(input("Qual o formato da fonte? [1]csv , [2]xlsx : "))
        break
    else:
        VerificaTeste = int(input("Arquivo nao encontrado! Deseja atribuir teste [1]Sim , [2]Nao" ))
        if VerificaTeste == True:
            tipo = 3
            break
        else:
            continue

idiomaBase = input("Qual o idioma?\n[de]Alemao\t[fr]frances\n[en]Ingles\t[es]Espanhol\n")

nomeArquivo = input("Qual o nome do arquivo final a ser salvo?: ")
'''
Codigo de Coleta de palavras
'''

palavras = Coleta(nomeBase,tipo) #Coleta palavras de csv[1] ou excel[2]
palavrasFinais = palavras[:] #Cria nova lista de palavras nao convertidas, para ir na tabela final
#palavras = Converte(palavras,idiomaBase) #Retira umlauts e eszetts

'''
Codigo de busca no Pons e Reverso
'''

browser = webdriver.PhantomJS() #Chama Navegador fantasma

artigos, classes, erros = Pons(palavras,idiomaBase) #Elementos que usam o Pons
exemplos, traducoes = Reverso(palavras,idiomaBase) #Elementos que usam o Reverso Context

browser.close() #Fecha navegador fantasma

'''
Salvando arquivo
'''

SalvarExcel(nomeArquivo,palavrasFinais,traducoes,artigos,exemplos,classes,erros)

'''
########################################
FIM DO CODIGO
########################################
'''
   
'''Observacoes'''
