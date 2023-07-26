import xmltodict
import os
import pandas as pd

def pegar_infos(nome_arquivo):
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        info_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        chave = info_nf["@Id"].replace("NFe", "")
        numero_nota = info_nf["ide"]["nNF"]
        serie = info_nf["ide"]["serie"]
        empresa_emissora = info_nf["emit"]["xNome"]
        nome_cliente = info_nf["dest"]["xNome"]
        peso_bruto = info_nf["transp"]["vol"]["pesoB"]
        peso_liquido = info_nf["transp"]["vol"]["pesoL"]
        valor = info_nf["total"]["ICMSTot"]["vNF"]
        valores.append([chave, numero_nota, serie, empresa_emissora, nome_cliente, peso_bruto, peso_liquido, valor])

def checkFilesExists():
    if not os.path.isdir("nfs"):
        os.makedirs("nfs")

    lista_arquivos = os.listdir("nfs")

    if len(lista_arquivos) == 0:
        input("Favor colocar arquivos XML na pasta NFS")
        exit()

    return lista_arquivos

if not os.path.exists("output"):
    os.makedirs("output")

colunas = ["chave", "numero_nota", "serie", "empresa_emissora", "nome_cliente", "peso_bruto", "peso_liquido", "valor"]
valores = []
for arquivo in checkFilesExists():
    pegar_infos(arquivo)

tabela = pd.DataFrame(columns=colunas, data= valores)
tabela.to_excel("output/NotasFiscais.xlsx", index=False)