import xml.etree.ElementTree as ET
import pandas as pd
import os
from datetime import datetime

#Excel file data reading
arquivo_excel = pd.read_excel('your/path')

#Last row filled
ultima_linha = arquivo_excel.index[-1]

#Columns values extraction
colunas = arquivo_excel.columns

#[XML GENERATOR]

#[Characteristics values feeding]

#Header

#Loop for Excel file rows
for n in range(2,ultima_linha): 
    # If value = "" then print "", by default the printed values of empty rolls are zero.
    arquivo_excel = arquivo_excel.fillna(' ')
    #If not empty, it searches for the columns positions (.iloc method) in Excel file specified in row 7
    if not arquivo_excel.empty:
        Descricao_Cliente = str(arquivo_excel.iloc[n, colunas.get_loc("Descricao_Cliente")])
        Cliente = str(arquivo_excel.iloc[n, colunas.get_loc("Cliente")])
        Nota_Fiscal = str(arquivo_excel.iloc[n, colunas.get_loc("Nota_Fiscal")])
        Data_Nota_Fiscal = str(arquivo_excel.iloc[n, colunas.get_loc("Data_Nota_Fiscal")])

        #Batch Values
        Lote = str(arquivo_excel.iloc[n, colunas.get_loc("Lote")])
        Peso = str(arquivo_excel.iloc[n, colunas.get_loc("Peso")])
        Cod_SAP = str(arquivo_excel.iloc[n, colunas.get_loc("Cod_SAP")])
        Descricao_SAP = str(arquivo_excel.iloc[n, colunas.get_loc("Descricao_SAP")])
        Aco = str(arquivo_excel.iloc[n, colunas.get_loc("Aco")])
        Espessura = str(arquivo_excel.iloc[n, colunas.get_loc("Espessura")])

        #Chemical Composition Values
        Carbono = str(arquivo_excel.iloc[n, colunas.get_loc("C")])
        Silicio = str(arquivo_excel.iloc[n, colunas.get_loc("Si")])
        Manganes = str(arquivo_excel.iloc[n, colunas.get_loc("Mn")])
        Fosforo = str(arquivo_excel.iloc[n, colunas.get_loc("P")])
        Enxofre = str(arquivo_excel.iloc[n, colunas.get_loc("S")])
        Aluminio = str(arquivo_excel.iloc[n, colunas.get_loc("Al")])
        Cobre = str(arquivo_excel.iloc[n, colunas.get_loc("Cu")])
        Niobio = str(arquivo_excel.iloc[n, colunas.get_loc("Nb")])
        Vanadio = str(arquivo_excel.iloc[n, colunas.get_loc("V")])
        Titanio = str(arquivo_excel.iloc[n, colunas.get_loc("Ti")])
        Cromo = str(arquivo_excel.iloc[n, colunas.get_loc("Cr")])
        Niquel = str(arquivo_excel.iloc[n, colunas.get_loc("Ni")])
        Molibidenio = str(arquivo_excel.iloc[n, colunas.get_loc("Mo")])
        Nitrogenio = str(arquivo_excel.iloc[n, colunas.get_loc("N")])
    
#Mechanical  Values
        Alongamento = str(arquivo_excel.iloc[n, colunas.get_loc("AL2")])
        LimEscoamento = str(arquivo_excel.iloc[n, colunas.get_loc("LE")])
        LimResistencia = str(arquivo_excel.iloc[n, colunas.get_loc("LR")])

#[XML_TREE]

#Creating the root
        raiz = ET.Element("Certificado_Soufer")

#Children elements
        #Header - Child
        Header = ET.SubElement(raiz, "Header")
        cliente_num = ET.SubElement(Header, "Cliente_Num")
        cliente_num.text = Cliente
        des_cliente = ET.SubElement(Header, "Descricao_Cliente")
        des_cliente.text = Descricao_Cliente
        nf = ET.SubElement(Header, "NF")
        nf.text = Nota_Fiscal
        des_NF = ET.SubElement(Header, "Data_NF")
        des_NF.text = Data_Nota_Fiscal

        #Coil data - child
        dados_bobina = ET.SubElement(raiz, "Dados_Bobina")
        loteXML = ET.SubElement(dados_bobina, "Lote")
        loteXML.text = Lote
        acoXML = ET.SubElement(dados_bobina, "Aco")
        acoXML.text = Aco
        espessuraXML = ET.SubElement(dados_bobina, "Espessura")
        espessuraXML.text = Espessura
        pesoXML = ET.SubElement(dados_bobina, "Peso")
        pesoXML.text = Peso
        codsapXML = ET.SubElement(dados_bobina, "Codigo_SAP")
        codsapXML.text = Cod_SAP
        dessapXML = ET.SubElement(dados_bobina, "Descricao_SAP")
        dessapXML.text = Descricao_SAP

        #Adding charcaterisitics child to root
        caracteristiscas = ET.SubElement(raiz, "Caracteristiscas")

        #Chemical Composition Elements values
        prop_quim = ET.SubElement(caracteristiscas, "Propriedades_Quimicas")

        carbonoXML = ET.SubElement(prop_quim, "C")
        carbonoXML.text = str(Carbono)

        silicioXML = ET.SubElement(prop_quim, "Si")
        silicioXML.text = str(Silicio)

        manganesXML = ET.SubElement(prop_quim, "Mn")
        manganesXML.text = str(Manganes)

        fosforoXML = ET.SubElement(prop_quim, "P")
        fosforoXML.text = str(Fosforo)

        enxofreXML = ET.SubElement(prop_quim, "S")
        enxofreXML.text = str(Enxofre)

        aluminioXML = ET.SubElement(prop_quim, "Al")
        aluminioXML.text = str(Aluminio)

        cobreXML = ET.SubElement(prop_quim, "Cu")
        cobreXML.text = str(Cobre)

        niobioXML = ET.SubElement(prop_quim, "Nb")
        niobioXML.text = str(Niobio)

        vanadioXML = ET.SubElement(prop_quim, "V")
        vanadioXML.text = str(Vanadio)

        titanioXML = ET.SubElement(prop_quim, "Ti")
        titanioXML.text = str(Titanio)

        cromoXML = ET.SubElement(prop_quim, "Cr")
        cromoXML.text = str(Cromo)

        niquelXML = ET.SubElement(prop_quim, "Ni")
        niquelXML.text = str(Niquel)

        molibidenioXML = ET.SubElement(prop_quim, "Mo")
        molibidenioXML.text = str(Molibidenio)

        nitrogenioXML = ET.SubElement(prop_quim, "N")
        nitrogenioXML.text = str(Nitrogenio)

        #Mechanical Properties Elements values
        prop_mec = ET.SubElement(caracteristiscas, "Propriedades_Mecanicas")
        LeXML = ET.SubElement(prop_mec, "LE")
        LeXML.text = str(LimEscoamento)
        LrXML = ET.SubElement(prop_mec, "LR")
        LrXML.text = str(LimResistencia)
        alongXML = ET.SubElement(prop_mec, "ALONG")
        alongXML.text = str(Alongamento)

        #Creating the root and saving the files
        caminho_salvar = 'your/path'
        arvoreXML = ET.ElementTree(raiz)
        arquivo_xml = os.path.join(caminho_salvar, f"{Lote}.xml")
        arvoreXML.write(arquivo_xml, encoding="utf-8", xml_declaration=True)

        #Message "XML files created with sucess"
        print(f"Arquivo XML '{arquivo_xml}' gerado com sucesso!")
