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

#Loop for Excel file rows
for n in range(2,ultima_linha): 
    # If value = "", then print " ", by default the printed values of empty rolls are zero.
    arquivo_excel = arquivo_excel.fillna(' ')
    #If not empty, it searches for the columns positions (.iloc method) in Excel file specified in row 7
    if not arquivo_excel.empty:     
        #Header Values
        descricao_Cliente = str(arquivo_excel.iloc[n, colunas.get_loc("Descricao_Cliente")])
        cliente = str(arquivo_excel.iloc[n, colunas.get_loc("Cliente")])
        nota_Fiscal = str(arquivo_excel.iloc[n, colunas.get_loc("Nota_Fiscal")])
        data_Nota_Fiscal = str(arquivo_excel.iloc[n, colunas.get_loc("Data_Nota_Fiscal")])

        #Batch Values
        lote = str(arquivo_excel.iloc[n, colunas.get_loc("Lote")])
        peso = str(arquivo_excel.iloc[n, colunas.get_loc("Peso")])
        cod_SAP = str(arquivo_excel.iloc[n, colunas.get_loc("Cod_SAP")])
        descricao_SAP = str(arquivo_excel.iloc[n, colunas.get_loc("Descricao_SAP")])
        aco = str(arquivo_excel.iloc[n, colunas.get_loc("Aco")])
        espessura = str(arquivo_excel.iloc[n, colunas.get_loc("Espessura")])

        #Chemical Composition Values
        carbono = str(arquivo_excel.iloc[n, colunas.get_loc("C")])
        silicio = str(arquivo_excel.iloc[n, colunas.get_loc("Si")])
        manganes = str(arquivo_excel.iloc[n, colunas.get_loc("Mn")])
        fosforo = str(arquivo_excel.iloc[n, colunas.get_loc("P")])
        enxofre = str(arquivo_excel.iloc[n, colunas.get_loc("S")])
        aluminio = str(arquivo_excel.iloc[n, colunas.get_loc("Al")])
        cobre = str(arquivo_excel.iloc[n, colunas.get_loc("Cu")])
        niobio = str(arquivo_excel.iloc[n, colunas.get_loc("Nb")])
        vanadio = str(arquivo_excel.iloc[n, colunas.get_loc("V")])
        titanio = str(arquivo_excel.iloc[n, colunas.get_loc("Ti")])
        cromo = str(arquivo_excel.iloc[n, colunas.get_loc("Cr")])
        niquel = str(arquivo_excel.iloc[n, colunas.get_loc("Ni")])
        molibidenio = str(arquivo_excel.iloc[n, colunas.get_loc("Mo")])
        nitrogenio = str(arquivo_excel.iloc[n, colunas.get_loc("N")])
    
        #Mechanical Properties Values
        alongamento = str(arquivo_excel.iloc[n, colunas.get_loc("AL2")])
        limEscoamento = str(arquivo_excel.iloc[n, colunas.get_loc("LE")])
        limResistencia = str(arquivo_excel.iloc[n, colunas.get_loc("LR")])

#[XML_TREE]

        #Creating the root
        raiz = ET.Element("Certificado_Soufer")

        #Header Parent
        header = ET.SubElement(raiz, "Header")
        
        #Children
        cliente_num = ET.SubElement(header, "Cliente_Num")
        cliente_num.text = cliente
        des_cliente = ET.SubElement(header, "Descricao_Cliente")
        des_cliente.text = descricao_Cliente
        nf = ET.SubElement(header, "NF")
        nf.text = nota_Fiscal
        des_NF = ET.SubElement(header, "Data_NF")
        des_NF.text = data_Nota_Fiscal

        #Coil data Parent
        dados_bobina = ET.SubElement(raiz, "Dados_Bobina")
        
        #Coil data Children
        loteXML = ET.SubElement(dados_bobina, "Lote")
        loteXML.text = lote
        acoXML = ET.SubElement(dados_bobina, "Aco")
        acoXML.text = aco
        espessuraXML = ET.SubElement(dados_bobina, "Espessura")
        espessuraXML.text = espessura
        pesoXML = ET.SubElement(dados_bobina, "Peso")
        pesoXML.text = peso
        codsapXML = ET.SubElement(dados_bobina, "Codigo_SAP")
        codsapXML.text = cod_SAP
        dessapXML = ET.SubElement(dados_bobina, "Descricao_SAP")
        dessapXML.text = descricao_SAP

        #charcaterisitics Parent
        caracteristiscas = ET.SubElement(raiz, "Caracteristiscas")

        #prop_quim Child
        prop_quim = ET.SubElement(caracteristiscas, "Propriedades_Quimicas")

        #Chemical Composition Grandchildren
        carbonoXML = ET.SubElement(prop_quim, "C")
        carbonoXML.text = str(carbono)

        silicioXML = ET.SubElement(prop_quim, "Si")
        silicioXML.text = str(silicio)

        manganesXML = ET.SubElement(prop_quim, "Mn")
        manganesXML.text = str(manganes)

        fosforoXML = ET.SubElement(prop_quim, "P")
        fosforoXML.text = str(fosforo)

        enxofreXML = ET.SubElement(prop_quim, "S")
        enxofreXML.text = str(enxofre)

        aluminioXML = ET.SubElement(prop_quim, "Al")
        aluminioXML.text = str(aluminio)

        cobreXML = ET.SubElement(prop_quim, "Cu")
        cobreXML.text = str(cobre)

        niobioXML = ET.SubElement(prop_quim, "Nb")
        niobioXML.text = str(niobio)

        vanadioXML = ET.SubElement(prop_quim, "V")
        vanadioXML.text = str(vanadio)

        titanioXML = ET.SubElement(prop_quim, "Ti")
        titanioXML.text = str(titanio)

        cromoXML = ET.SubElement(prop_quim, "Cr")
        cromoXML.text = str(cromo)

        niquelXML = ET.SubElement(prop_quim, "Ni")
        niquelXML.text = str(niquel)

        molibidenioXML = ET.SubElement(prop_quim, "Mo")
        molibidenioXML.text = str(molibidenio)

        nitrogenioXML = ET.SubElement(prop_quim, "N")
        nitrogenioXML.text = str(nitrogenio)

        #prop_mec Child
        prop_mec = ET.SubElement(caracteristiscas, "Propriedades_Mecanicas")
        
        #Mechanical Properties Grandchildren
        leXML = ET.SubElement(prop_mec, "LE")
        leXML.text = str(limEscoamento)
        lrXML = ET.SubElement(prop_mec, "LR")
        lrXML.text = str(limResistencia)
        alongXML = ET.SubElement(prop_mec, "ALONG")
        alongXML.text = str(alongamento)

        #Creating the root and saving the files
        caminho_salvar = 'your/path'
        arvoreXML = ET.ElementTree(raiz)
        arquivo_xml = os.path.join(caminho_salvar, f"{lote}.xml")
        arvoreXML.write(arquivo_xml, encoding="utf-8", xml_declaration=True)

        #Message "XML files created with sucess"
        print(f"Arquivo XML '{arquivo_xml}' gerado com sucesso!")
