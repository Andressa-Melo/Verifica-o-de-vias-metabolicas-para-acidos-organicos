# Os 'nomes K' das enzimas são salvos em uma lista de lista, onde cada etapa da via é uma lista
# Se na via metabólica tem enzimas ortologas, ou um conjunto de enzima no mesmo grupo, é posto na mesma lista.
# Na rota da via metabólica é pego apenas as enzimas que começam com K
# A função 'acidos_organicos' verifica se a enzima da lista esta presente no arquivo de entrada (organismo)
# A função 'verificar_via' para verificar se a via está completa ou incompleta e verifica se todas as etapas estao completas ou não. Adiciona os status da via na 2 coluna abaixo da via
# A função 'adicionar_status' adicioa os status da via na aba resultados e nas demias abas
# A funçao 'combinar_dataframes_com_numeracao' combina as listas (dfs) em um único DataFrame, adicionando uma numeração e um prefixo às colunas de cada DataFrame e inserindo colunas em branco entre os DataFrames para separação visual.


import pandas as pd   

# entrada
arquivo = '/home/jessica/Área de Trabalho/acido_organico/pararodar(txt)/result_ko.GO024741.txt'   

#saida 
arquivo_saida = '/home/jessica/Área de Trabalho/acido_organico/resutado_final/GO024741.xlsx'   

### Lista de enzimas de interesse FIXO  ####
#metabolismo das purinas(00230) parti da adenina 
enzimasoxalate_viaadenina1 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K00839'], ['K22601'], ['K22602']]
enzimasoxalate_viaadenina2 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K00839'], ['K22601'], ['K22602']]
enzimasoxalate_viaadenina3 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K14977'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaadenina4 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K14977'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaadenina5 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K01477'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaadenina6 = [['K01486', 'K21053'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K01477'], ['K00073'], ['K22601'], ['K22602']]

#metabolismo das purinas(00230) partir da guanina 
enzimasoxalate_viaguanina1 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K00839'], ['K22601'], ['K22602']]
enzimasoxalate_viaguanina2 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K00839'], ['K22601'], ['K22602']]
enzimasoxalate_viaguanina3 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K14977'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaguanina4 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K02083', 'K24871'], ['K14977'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaguanina5 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K13485', 'K13484', 'K16838', 'K16840'], ['K01466', 'K16842'], ['K01477'], ['K00073'], ['K22601'], ['K22602']]
enzimasoxalate_viaguanina6 = [['K01487'], ['K00106', 'K00087', 'K13479', 'K13480', 'K13481', 'K13482', 'K11177', 'K11178', 'K13483'], ['K16839', 'K22879', 'K00365', 'K16838'], ['K07127', 'K13484', 'K19964'], ['K16841'], ['K01466', 'K16842'], ['K01477'], ['K00073'], ['K22601'], ['K22602']]
#metabolismo piruvato (00620) # partir da glicolise
enzimaslactato_mp = [['K00873', 'K12406'], ['K00016']]

enzimasmalato_mp1 = [['K00873', 'K12406'], ['K01958', 'K01959', 'K01960'], ['K00025', 'K00026', 'K00024', 'K00051', 'K00116']]
enzimasmalato_mp2 = [['K00873', 'K12406'], ['K00028', 'K00027', 'K00029']]
enzimasmalato_mp3 = [['K00873', 'K12406'], ['K00163', 'K00161', 'K00162'], ['K00627'], ['K01638']]
enzimasmalato_mp4 = [['K00873', 'K12406'], ['K00174', 'K00175'], ['K01638']]
enzimasmalato_mp5 = [['K00873', 'K12406'], ['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737', 'K03737'], ['K01638']]

enzimasacetato_mp1 = [['K00873', 'K12406'], ['K00016'], ['K00467']]
enzimasacetato_mp2 = [['K00873', 'K12406'], ['K00156']]
enzimasacetato_mp3 = [['K00873', 'K12406'], ['K00158'], ['K00925']]
enzimasacetato_mp4 = [['K00873', 'K12406'], ['K00158'], ['K01512']]
enzimasacetato_mp5 = [['K00873', 'K12406'],['K00163', 'K00161', 'K00162'], ['K00627'], ['K01067']]
enzimasacetato_mp6 = [['K00873', 'K12406'],['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K00627'], ['K01067']]
enzimasacetato_mp7 = [['K00873', 'K12406'], ['K00174', 'K00175'], ['K00627'], ['K01067']]
enzimasacetato_mp8 = [['K00873', 'K12406'], ['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K01067']]
enzimasacetato_mp9 = [['K00873', 'K12406'],['K00163', 'K00161', 'K00162'], ['K00627'], ['K01026', 'K18118', 'K01905', 'K22224', 'K24012']]
enzimasacetato_mp10 = [['K00873', 'K12406'],['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K01026', 'K18118', 'K01905', 'K22224', 'K24012']]
enzimasacetato_mp11 = [['K00873', 'K12406'], ['K00174', 'K00175'], ['K01026', 'K18118', 'K01905', 'K22224', 'K24012']]
enzimasacetato_mp12 = [['K00873', 'K12406'], ['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K01026', 'K18118', 'K01905', 'K22224', 'K24012']]
enzimasacetato_mp13 = [['K00873', 'K12406'], ['K00174', 'K00175'], ['K01895', 'K01913']]
enzimasacetato_mp14 = [['K00873', 'K12406'], ['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K01895', 'K01913']]
enzimasacetato_mp15 = [['K00873', 'K12406'], ['K00163', 'K00161', 'K00162'], ['K01895', 'K01913']]
enzimasacetato_mp16 = [['K00873', 'K12406'], ['K00163', 'K00161', 'K00162'], ['K00132', 'K04072', 'K04073', 'K18366', 'K04021'], ['K00128', 'K14085', 'K00149', 'K00138']]
enzimasacetato_mp17 = [['K00873', 'K12406'], ['K00174', 'K00175'], ['K00132', 'K04072', 'K04073', 'K18366', 'K04021'], ['K00128', 'K14085', 'K00149', 'K00138']]
enzimasacetato_mp18 = [['K00873', 'K12406'], ['K00169', 'K00170', 'K00172', 'K00189', 'K00171', 'K03737'], ['K00132', 'K04072', 'K04073', 'K18366', 'K04021'], ['K00128', 'K14085', 'K00149', 'K00138']]
#metabolismo metano (00680) partir do acetil CoA
enzimasmalato_mm = [['K00169', 'K00170', 'K00172', 'K00189', 'K00171'], ['K01007'], ['K01595'], ['K00024']]
#ciclo do citrato (00020) partir da gliclises/gluconeogenesis
enzimasmalato_cc= [['K01958', 'K01959', 'K01960'], ['K00025', 'K00026', 'K00024', ]]
#metabolismo de glioxilato e dicarboxilato (00630) partir do oxalacetato
enzimasaoxalato1 = [['K01647', 'K01659'], ['K01681', 'K27802', 'K01682'], ['K01681', 'K27802', 'K01682'], ['K01637']]
enzimasaoxalato2 = [['K00025', 'K00026', 'K00024'], ['K01638']]

#metabolismo fosfato e fosfanato (00440) partir da glicolises
enzimasacetato = [['K01841', 'K23999'], ['K09459'], ['K00206'], ['K19670']]
#via das pentoses fosfato (00030) partir da glicolise
enzimasgluconato1 = [['K21840'], ['K01053']]
enzimasgluconato2 = [['K00034', 'K22969'], ['K01053']]
enzimasgluconato3 = [['K18124', 'K18125', 'K00115', 'K19813', 'K00117'], ['K01053']]
#via da frutose e manose (00051) partir da fucose
enzimaslactato_fm1 = [['K18333'], ['K07046'], ['K18334'], ['K18335'], ['K18336']]
enzimaslactato_fm2 = [['K00064'], ['K07046'], ['K18334'], ['K18335'], ['K18336']]


### FUNÇÃO QUE VERIFICA AS ETAPAS DAS ENZIMAS ####
def acidos_organicos(valores, conteudo):
    resultados = []
    for i, valores in enumerate(valores, start=1):
        etapa = i
        enzima_encontrada = False
        for valor in valores:
            encontrado = any(valor in linha for linha in conteudo)
            if encontrado:
                resultados.append([valor, 'Existe', f'Etapa {etapa}'])
                enzima_encontrada = True   
            else:
                resultados.append([valor, 'N/E', f'Etapa {etapa}'])
        if enzima_encontrada:
            resultados.append(['', '', f'Etapa {etapa} Completa'])
        else:
            resultados.append(['', '', f'Etapa {etapa} Incompleta'])
    return resultados

### Ler o arquivo e obter o conteúdo ####
with open(arquivo, 'r') as file:
    conteudo = file.readlines()  # Lê todas as linhas do arquivo (readlines) (NÃO USEI O SEPARADOR SPLIT ('\t') e ele conseguiu fazer a busca, se surgir algum bug nesse sentido, adicionar linha comentada)
    #arquivo = [linha.strip().split('\t') for linha in conteudo]

### Chamar a função para verificar enzimas na via ###
resultados_oxalate_viaadenina1 = acidos_organicos(enzimasoxalate_viaadenina1, conteudo)
resultados_oxalate_viaadenina2 = acidos_organicos(enzimasoxalate_viaadenina2, conteudo)
resultados_oxalate_viaadenina3 = acidos_organicos(enzimasoxalate_viaadenina3, conteudo)
resultados_oxalate_viaadenina4 = acidos_organicos(enzimasoxalate_viaadenina4, conteudo)
resultados_oxalate_viaadenina5 = acidos_organicos(enzimasoxalate_viaadenina5, conteudo)
resultados_oxalate_viaadenina6 = acidos_organicos(enzimasoxalate_viaadenina6, conteudo)

resultados_oxalate_viaguanina1 = acidos_organicos(enzimasoxalate_viaguanina1, conteudo)
resultados_oxalate_viaguanina2 = acidos_organicos(enzimasoxalate_viaguanina2, conteudo)
resultados_oxalate_viaguanina3 = acidos_organicos(enzimasoxalate_viaguanina3, conteudo)
resultados_oxalate_viaguanina4 = acidos_organicos(enzimasoxalate_viaguanina4, conteudo)
resultados_oxalate_viaguanina5 = acidos_organicos(enzimasoxalate_viaguanina5, conteudo)
resultados_oxalate_viaguanina6 = acidos_organicos(enzimasoxalate_viaguanina6, conteudo)

resultados_lactato_via_piruvato = acidos_organicos(enzimaslactato_mp, conteudo)

resultados_malato_via_piruvato1 = acidos_organicos(enzimasmalato_mp1, conteudo)
resultados_malato_via_piruvato2 = acidos_organicos(enzimasmalato_mp2, conteudo)
resultados_malato_via_piruvato3 = acidos_organicos(enzimasmalato_mp3, conteudo)
resultados_malato_via_piruvato4 = acidos_organicos(enzimasmalato_mp4, conteudo)
resultados_malato_via_piruvato5 = acidos_organicos(enzimasmalato_mp5, conteudo)

resultados_acetato_via_piruvato1 = acidos_organicos(enzimasacetato_mp1, conteudo)
resultados_acetato_via_piruvato2 = acidos_organicos(enzimasacetato_mp2, conteudo)
resultados_acetato_via_piruvato3 = acidos_organicos(enzimasacetato_mp3, conteudo)
resultados_acetato_via_piruvato4 = acidos_organicos(enzimasacetato_mp4, conteudo)
resultados_acetato_via_piruvato5 = acidos_organicos(enzimasacetato_mp5, conteudo)
resultados_acetato_via_piruvato6 = acidos_organicos(enzimasacetato_mp6, conteudo)
resultados_acetato_via_piruvato7 = acidos_organicos(enzimasacetato_mp7, conteudo)
resultados_acetato_via_piruvato8 = acidos_organicos(enzimasacetato_mp8, conteudo)
resultados_acetato_via_piruvato9 = acidos_organicos(enzimasacetato_mp9, conteudo)
resultados_acetato_via_piruvato10 = acidos_organicos(enzimasacetato_mp10, conteudo)
resultados_acetato_via_piruvato11 = acidos_organicos(enzimasacetato_mp11, conteudo)
resultados_acetato_via_piruvato12 = acidos_organicos(enzimasacetato_mp12, conteudo)
resultados_acetato_via_piruvato13 = acidos_organicos(enzimasacetato_mp13, conteudo)
resultados_acetato_via_piruvato14 = acidos_organicos(enzimasacetato_mp14, conteudo)
resultados_acetato_via_piruvato15 = acidos_organicos(enzimasacetato_mp15, conteudo)
resultados_acetato_via_piruvato16 = acidos_organicos(enzimasacetato_mp16, conteudo)
resultados_acetato_via_piruvato17 = acidos_organicos(enzimasacetato_mp17, conteudo)
resultados_acetato_via_piruvato18 = acidos_organicos(enzimasacetato_mp18, conteudo)

resultados_malato_meta_metano = acidos_organicos(enzimasmalato_mm, conteudo)

resultados_malato_ciclo_citrato = acidos_organicos(enzimasmalato_cc, conteudo)

resultados_oxalato_via_glio1 = acidos_organicos(enzimasaoxalato1, conteudo)
resultados_oxalato_via_glio2 = acidos_organicos(enzimasaoxalato2, conteudo)

resultados_acetato_meta_fosfa = acidos_organicos(enzimasacetato, conteudo)

resultados_gluconato_via_glicolise1 = acidos_organicos(enzimasgluconato1, conteudo)
resultados_gluconato_via_glicolise2 = acidos_organicos(enzimasgluconato2, conteudo)
resultados_gluconato_via_glicolise3 = acidos_organicos(enzimasgluconato3, conteudo)

resultados_enzimaslactato_fm1 = acidos_organicos(enzimaslactato_fm1, conteudo)
resultados_enzimaslactato_fm2 = acidos_organicos(enzimaslactato_fm2, conteudo)


###  Criar DataFrames para cada conjunto de resultados
df_oxalate_viaadenina1 = pd.DataFrame(resultados_oxalate_viaadenina1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaadenina2 = pd.DataFrame(resultados_oxalate_viaadenina2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaadenina3 = pd.DataFrame(resultados_oxalate_viaadenina3, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaadenina4 = pd.DataFrame(resultados_oxalate_viaadenina4, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaadenina5 = pd.DataFrame(resultados_oxalate_viaadenina5, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaadenina6 = pd.DataFrame(resultados_oxalate_viaadenina6, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_oxalate_viaguanina1 = pd.DataFrame(resultados_oxalate_viaguanina1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaguanina2 = pd.DataFrame(resultados_oxalate_viaguanina2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaguanina3 = pd.DataFrame(resultados_oxalate_viaguanina3, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaguanina4 = pd.DataFrame(resultados_oxalate_viaguanina4, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaguanina5 = pd.DataFrame(resultados_oxalate_viaguanina5, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalate_viaguanina6 = pd.DataFrame(resultados_oxalate_viaguanina6, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_lactato_vp = pd.DataFrame(resultados_lactato_via_piruvato, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_malato_vp1 = pd.DataFrame(resultados_malato_via_piruvato1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_malato_vp2 = pd.DataFrame(resultados_malato_via_piruvato2 , columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_malato_vp3 = pd.DataFrame(resultados_malato_via_piruvato3 , columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_malato_vp4 = pd.DataFrame(resultados_malato_via_piruvato4 , columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_malato_vp5 = pd.DataFrame(resultados_malato_via_piruvato5 , columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_acetato_vp1 = pd.DataFrame(resultados_acetato_via_piruvato1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp2 = pd.DataFrame(resultados_acetato_via_piruvato2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp3 = pd.DataFrame(resultados_acetato_via_piruvato3, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp4 = pd.DataFrame(resultados_acetato_via_piruvato4, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp5 = pd.DataFrame(resultados_acetato_via_piruvato5, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp6 = pd.DataFrame(resultados_acetato_via_piruvato6, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp7 = pd.DataFrame(resultados_acetato_via_piruvato7, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp8 = pd.DataFrame(resultados_acetato_via_piruvato8, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp9 = pd.DataFrame(resultados_acetato_via_piruvato9, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp10 = pd.DataFrame(resultados_acetato_via_piruvato10, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp11 = pd.DataFrame(resultados_acetato_via_piruvato11, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp12 = pd.DataFrame(resultados_acetato_via_piruvato12, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp13 = pd.DataFrame(resultados_acetato_via_piruvato13, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp14 = pd.DataFrame(resultados_acetato_via_piruvato14, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp15 = pd.DataFrame(resultados_acetato_via_piruvato15, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp16 = pd.DataFrame(resultados_acetato_via_piruvato16, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp17 = pd.DataFrame(resultados_acetato_via_piruvato17, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_acetato_vp18 = pd.DataFrame(resultados_acetato_via_piruvato18, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_malato_mm = pd.DataFrame(resultados_malato_meta_metano, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_malato_cc = pd.DataFrame(resultados_malato_ciclo_citrato, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_oxalato_vg1 = pd.DataFrame(resultados_oxalato_via_glio1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_oxalato_vg2 = pd.DataFrame(resultados_oxalato_via_glio2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_acetato_mf = pd.DataFrame(resultados_acetato_meta_fosfa, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_gluconato_vg1 = pd.DataFrame(resultados_gluconato_via_glicolise1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_gluconato_vg2 = pd.DataFrame(resultados_gluconato_via_glicolise2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_gluconato_vg3 = pd.DataFrame(resultados_gluconato_via_glicolise3, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_enzimaslactato_fm1 = pd.DataFrame(resultados_enzimaslactato_fm1, columns=['Enzima', 'Status', 'Etapas da via metabólica'])
df_enzimaslactato_fm2 = pd.DataFrame(resultados_enzimaslactato_fm2, columns=['Enzima', 'Status', 'Etapas da via metabólica'])

df_resultado_total = pd.DataFrame(columns=['Resultado Geral'])  # DataFrame geral para resultados

### Função para verificar se a via metabólica está completa ###
def verificar_via(df, total_etapas):
    etapas_completas = set()  
    for etapa in df[df['Status'] == 'Existe']['Etapas da via metabólica']:
        etapa_numero = int(etapa.split(' ')[1])  # Extrai o número 0 da etapa
        etapas_completas.add(etapa_numero)
    # Verifica se todas as etapas esperadas foram completadas
    return len(etapas_completas) == total_etapas

### Contar total de etapas para cada grupo ###
total_etapas_oxalate_adenina1 = len(enzimasoxalate_viaadenina1) 
total_etapas_oxalate_adenina2 = len(enzimasoxalate_viaadenina2) 
total_etapas_oxalate_adenina3 = len(enzimasoxalate_viaadenina3) 
total_etapas_oxalate_adenina4 = len(enzimasoxalate_viaadenina4) 
total_etapas_oxalate_adenina5 = len(enzimasoxalate_viaadenina5) 
total_etapas_oxalate_adenina6 = len(enzimasoxalate_viaadenina6) 

total_etapas_oxalate_guanina1 = len(enzimasoxalate_viaguanina1)
total_etapas_oxalate_guanina2 = len(enzimasoxalate_viaguanina2)
total_etapas_oxalate_guanina3 = len(enzimasoxalate_viaguanina3)
total_etapas_oxalate_guanina4 = len(enzimasoxalate_viaguanina4)
total_etapas_oxalate_guanina5 = len(enzimasoxalate_viaguanina5)
total_etapas_oxalate_guanina6 = len(enzimasoxalate_viaguanina6)

total_etapas_lactato = len(enzimaslactato_mp)

total_etapas_malatovp1 = len(enzimasmalato_mp1)
total_etapas_malatovp2 = len(enzimasmalato_mp2)
total_etapas_malatovp3 = len(enzimasmalato_mp3)
total_etapas_malatovp4 = len(enzimasmalato_mp4)
total_etapas_malatovp5 = len(enzimasmalato_mp5)

total_etapas_acetatomp1 = len(enzimasacetato_mp1)
total_etapas_acetatomp2 = len(enzimasacetato_mp2)
total_etapas_acetatomp3 = len(enzimasacetato_mp3)
total_etapas_acetatomp4 = len(enzimasacetato_mp4)
total_etapas_acetatomp5 = len(enzimasacetato_mp5)
total_etapas_acetatomp6 = len(enzimasacetato_mp6)
total_etapas_acetatomp7 = len(enzimasacetato_mp7)
total_etapas_acetatomp8 = len(enzimasacetato_mp8)
total_etapas_acetatomp9 = len(enzimasacetato_mp9)
total_etapas_acetatomp10 = len(enzimasacetato_mp10)
total_etapas_acetatomp11 = len(enzimasacetato_mp11)
total_etapas_acetatomp12 = len(enzimasacetato_mp12)
total_etapas_acetatomp13 = len(enzimasacetato_mp13)
total_etapas_acetatomp14 = len(enzimasacetato_mp14)
total_etapas_acetatomp15 = len(enzimasacetato_mp15)
total_etapas_acetatomp16 = len(enzimasacetato_mp16)
total_etapas_acetatomp17 = len(enzimasacetato_mp17)
total_etapas_acetatomp18 = len(enzimasacetato_mp18)

total_etapas_malatomm = len(enzimasmalato_mm)

total_etapas_malatocc = len(enzimasmalato_cc)

total_etapas_oxalato1 = len(enzimasaoxalato1)
total_etapas_oxalato2 = len(enzimasaoxalato2)


total_etapas_acetato = len(enzimasacetato)

total_etapas_gluconato1 = len(enzimasgluconato1)
total_etapas_gluconato2 = len(enzimasgluconato2)
total_etapas_gluconato3 = len(enzimasgluconato3)

total_etapas_enzimaslactato_fm1 = len(enzimaslactato_fm1)
total_etapas_enzimaslactato_fm2 = len(enzimaslactato_fm2)

### Verificar se a via é completa ###
via_completa_oxalate_adenina1 = verificar_via(df_oxalate_viaadenina1, total_etapas_oxalate_adenina1)
via_completa_oxalate_adenina2 = verificar_via(df_oxalate_viaadenina2, total_etapas_oxalate_adenina2)
via_completa_oxalate_adenina3 = verificar_via(df_oxalate_viaadenina3, total_etapas_oxalate_adenina3)
via_completa_oxalate_adenina4 = verificar_via(df_oxalate_viaadenina4, total_etapas_oxalate_adenina4)
via_completa_oxalate_adenina5 = verificar_via(df_oxalate_viaadenina5, total_etapas_oxalate_adenina5)
via_completa_oxalate_adenina6 = verificar_via(df_oxalate_viaadenina6, total_etapas_oxalate_adenina6)

via_completa_oxalate_guanina1 = verificar_via(df_oxalate_viaguanina1, total_etapas_oxalate_guanina1)
via_completa_oxalate_guanina2 = verificar_via(df_oxalate_viaguanina2, total_etapas_oxalate_guanina2)
via_completa_oxalate_guanina3 = verificar_via(df_oxalate_viaguanina3, total_etapas_oxalate_guanina3)
via_completa_oxalate_guanina4 = verificar_via(df_oxalate_viaguanina4, total_etapas_oxalate_guanina4)
via_completa_oxalate_guanina5 = verificar_via(df_oxalate_viaguanina5, total_etapas_oxalate_guanina5)
via_completa_oxalate_guanina6 = verificar_via(df_oxalate_viaguanina6, total_etapas_oxalate_guanina6)

via_completa_lactato = verificar_via(df_lactato_vp, total_etapas_lactato)

via_piruvato_malatovp1 = verificar_via(df_malato_vp1, total_etapas_malatovp1)
via_piruvato_malatovp2 = verificar_via(df_malato_vp2, total_etapas_malatovp2)
via_piruvato_malatovp3 = verificar_via(df_malato_vp3, total_etapas_malatovp3)
via_piruvato_malatovp4 = verificar_via(df_malato_vp4, total_etapas_malatovp4)
via_piruvato_malatovp5 = verificar_via(df_malato_vp5, total_etapas_malatovp5)

via_completa_acetatomp1 = verificar_via(df_acetato_vp1, total_etapas_acetatomp1)
via_completa_acetatomp2 = verificar_via(df_acetato_vp2, total_etapas_acetatomp2)
via_completa_acetatomp3 = verificar_via(df_acetato_vp3, total_etapas_acetatomp3)
via_completa_acetatomp4 = verificar_via(df_acetato_vp4, total_etapas_acetatomp4)
via_completa_acetatomp5 = verificar_via(df_acetato_vp5, total_etapas_acetatomp5)
via_completa_acetatomp6 = verificar_via(df_acetato_vp6, total_etapas_acetatomp6)
via_completa_acetatomp7 = verificar_via(df_acetato_vp7, total_etapas_acetatomp7)
via_completa_acetatomp8 = verificar_via(df_acetato_vp8, total_etapas_acetatomp8)
via_completa_acetatomp9 = verificar_via(df_acetato_vp9, total_etapas_acetatomp9)
via_completa_acetatomp10 = verificar_via(df_acetato_vp10, total_etapas_acetatomp10)
via_completa_acetatomp11 = verificar_via(df_acetato_vp11, total_etapas_acetatomp11)
via_completa_acetatomp12 = verificar_via(df_acetato_vp12, total_etapas_acetatomp12)
via_completa_acetatomp13 = verificar_via(df_acetato_vp13, total_etapas_acetatomp13)
via_completa_acetatomp14 = verificar_via(df_acetato_vp14, total_etapas_acetatomp14)
via_completa_acetatomp15 = verificar_via(df_acetato_vp15, total_etapas_acetatomp15)
via_completa_acetatomp16 = verificar_via(df_acetato_vp16, total_etapas_acetatomp16)
via_completa_acetatomp17 = verificar_via(df_acetato_vp17, total_etapas_acetatomp17)
via_completa_acetatomp18 = verificar_via(df_acetato_vp18, total_etapas_acetatomp18)

via_completa_malatomm = verificar_via(df_malato_mm, total_etapas_malatomm)

via_completa_malatocc = verificar_via(df_malato_cc, total_etapas_malatocc)

via_completa_oxalato1 = verificar_via(df_oxalato_vg1, total_etapas_oxalato1)
via_completa_oxalato2 = verificar_via(df_oxalato_vg2, total_etapas_oxalato2)

via_completa_acetato = verificar_via(df_acetato_mf, total_etapas_acetato)

via_completa_gluconato1 = verificar_via(df_gluconato_vg1, total_etapas_gluconato1)
via_completa_gluconato2 = verificar_via(df_gluconato_vg2, total_etapas_gluconato2)
via_completa_gluconato3 = verificar_via(df_gluconato_vg3, total_etapas_gluconato3)

via_completa_enzimaslactato_fm1 = verificar_via(df_enzimaslactato_fm1, total_etapas_enzimaslactato_fm1)
via_completa_enzimaslactato_fm2 = verificar_via(df_enzimaslactato_fm2, total_etapas_enzimaslactato_fm2)


### função para adicionar o resultado geral a partir dos resultados dos status da via
def adicionar_status(df_resultado, nome_via, status_via):
    nova_linha = pd.DataFrame({'Resultado Geral': [f'{nome_via}: {status_via}']})
    df_resultado = pd.concat([df_resultado, nova_linha], ignore_index=True)
    return df_resultado

status_oxalatea1 = 'Via Completa' if via_completa_oxalate_adenina1 else 'Via Incompleta'
status_oxalatea2 = 'Via Completa' if via_completa_oxalate_adenina2 else 'Via Incompleta'
status_oxalatea3 = 'Via Completa' if via_completa_oxalate_adenina3 else 'Via Incompleta'
status_oxalatea4 = 'Via Completa' if via_completa_oxalate_adenina4 else 'Via Incompleta'
status_oxalatea5 = 'Via Completa' if via_completa_oxalate_adenina5 else 'Via Incompleta'
status_oxalatea6 = 'Via Completa' if via_completa_oxalate_adenina6 else 'Via Incompleta'

status_oxalateg1 = 'Via Completa' if via_completa_oxalate_guanina1 else 'Via Incompleta'
status_oxalateg2 = 'Via Completa' if via_completa_oxalate_guanina2 else 'Via Incompleta'
status_oxalateg3 = 'Via Completa' if via_completa_oxalate_guanina3 else 'Via Incompleta'
status_oxalateg4 = 'Via Completa' if via_completa_oxalate_guanina4 else 'Via Incompleta'
status_oxalateg5 = 'Via Completa' if via_completa_oxalate_guanina5 else 'Via Incompleta'
status_oxalateg6 = 'Via Completa' if via_completa_oxalate_guanina6 else 'Via Incompleta'

status_lactato = 'Via Completa' if via_completa_lactato else 'Via Incompleta'

status_malatovp1 = 'Via Completa' if via_piruvato_malatovp1 else 'Via Incompleta'
status_malatovp2 = 'Via Completa' if via_piruvato_malatovp2 else 'Via Incompleta'
status_malatovp3 = 'Via Completa' if via_piruvato_malatovp3 else 'Via Incompleta'
status_malatovp4 = 'Via Completa' if via_piruvato_malatovp4 else 'Via Incompleta'
status_malatovp5 = 'Via Completa' if via_piruvato_malatovp5 else 'Via Incompleta'

status_acetatomp1 = 'Via Completa' if via_completa_acetatomp1 else 'Via Incompleta'
status_acetatomp2 = 'Via Completa' if via_completa_acetatomp2 else 'Via Incompleta'
status_acetatomp3 = 'Via Completa' if via_completa_acetatomp3 else 'Via Incompleta'
status_acetatomp4 = 'Via Completa' if via_completa_acetatomp4 else 'Via Incompleta'
status_acetatomp5 = 'Via Completa' if via_completa_acetatomp5 else 'Via Incompleta'
status_acetatomp6 = 'Via Completa' if via_completa_acetatomp6 else 'Via Incompleta'
status_acetatomp7 = 'Via Completa' if via_completa_acetatomp7 else 'Via Incompleta'
status_acetatomp8 = 'Via Completa' if via_completa_acetatomp8 else 'Via Incompleta'
status_acetatomp9 = 'Via Completa' if via_completa_acetatomp9 else 'Via Incompleta'
status_acetatomp10 = 'Via Completa' if via_completa_acetatomp10 else 'Via Incompleta'
status_acetatomp11 = 'Via Completa' if via_completa_acetatomp11 else 'Via Incompleta'
status_acetatomp12 = 'Via Completa' if via_completa_acetatomp12 else 'Via Incompleta'
status_acetatomp13 = 'Via Completa' if via_completa_acetatomp13 else 'Via Incompleta'
status_acetatomp14 = 'Via Completa' if via_completa_acetatomp14 else 'Via Incompleta'
status_acetatomp15 = 'Via Completa' if via_completa_acetatomp15 else 'Via Incompleta'
status_acetatomp16 = 'Via Completa' if via_completa_acetatomp16 else 'Via Incompleta'
status_acetatomp17 = 'Via Completa' if via_completa_acetatomp17 else 'Via Incompleta'
status_acetatomp18 = 'Via Completa' if via_completa_acetatomp18 else 'Via Incompleta'

status_malatomm = 'Via Completa' if via_completa_malatomm else 'Via Incompleta'

status_malatocc = 'Via Completa' if via_completa_malatocc else 'Via Incompleta'

status_oxalato_vg1 = 'Via Completa' if via_completa_oxalato1 else 'Via Incompleta'
status_oxalato_vg2 = 'Via Completa' if via_completa_oxalato2 else 'Via Incompleta'

status_acetato = 'Via Completa' if via_completa_acetato else 'Via Incompleta'

status_glaconato1 = 'Via Completa' if via_completa_gluconato1 else 'Via Incompleta'
status_glaconato2 = 'Via Completa' if via_completa_gluconato2 else 'Via Incompleta'
status_glaconato3 = 'Via Completa' if via_completa_gluconato3 else 'Via Incompleta'

status_enzimaslactato_fm1 = 'Via Completa' if via_completa_enzimaslactato_fm1 else 'Via Incompleta'
status_enzimaslactato_fm2 = 'Via Completa' if via_completa_enzimaslactato_fm2 else 'Via Incompleta'


# Adicionar o status das vias ao DataFrame 'Resultado Geral'
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 1 (purinas)', status_oxalatea1)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 2 (purinas)', status_oxalatea2)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 3 (purinas)', status_oxalatea3)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 4 (purinas)', status_oxalatea4)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 5 (purinas)', status_oxalatea5)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via adenina 6 (purinas)', status_oxalatea6)

df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 1 (purinas)', status_oxalateg1)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 2 (purinas)', status_oxalateg2)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 3 (purinas)', status_oxalateg3)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 4 (purinas)', status_oxalateg4)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 5 (purinas)', status_oxalateg5)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato via guanina 6 (purinas)', status_oxalateg6)

df_resultado_total = adicionar_status(df_resultado_total, 'Lactato (piruvato)', status_lactato)

df_resultado_total = adicionar_status(df_resultado_total, 'Malato 1 (piruvato)', status_malatovp1)
df_resultado_total = adicionar_status(df_resultado_total, 'Malato 2 (piruvato)', status_malatovp2)
df_resultado_total = adicionar_status(df_resultado_total, 'Malato 3 (piruvato)', status_malatovp3)
df_resultado_total = adicionar_status(df_resultado_total, 'Malato 4 (piruvato)', status_malatovp4)
df_resultado_total = adicionar_status(df_resultado_total, 'Malato 5 (piruvato)', status_malatovp5)

df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 1 (piruvato)', status_acetatomp1)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 2 (piruvato)', status_acetatomp2)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 3 (piruvato)', status_acetatomp3)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 4 (piruvato)', status_acetatomp4)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 5 (piruvato)', status_acetatomp5)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 6 (piruvato)', status_acetatomp6)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 7 (piruvato)', status_acetatomp7)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 8 (piruvato)', status_acetatomp8)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 9 (piruvato)', status_acetatomp9)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 10 (piruvato)', status_acetatomp10)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 11 (piruvato)', status_acetatomp11)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 12 (piruvato)', status_acetatomp12)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 13 (piruvato)', status_acetatomp13)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 14 (piruvato)', status_acetatomp14)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 15 (piruvato)', status_acetatomp15)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 16 (piruvato)', status_acetatomp16)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 17 (piruvato)', status_acetatomp17)
df_resultado_total = adicionar_status(df_resultado_total, 'Acetato 18 (piruvato)', status_acetatomp18)

df_resultado_total = adicionar_status(df_resultado_total, 'Gluconato 1 (pentose fosfato)', status_glaconato1)
df_resultado_total = adicionar_status(df_resultado_total, 'Gluconato 2 (pentose fosfato)', status_glaconato2)
df_resultado_total = adicionar_status(df_resultado_total, 'Gluconato 3 (pentose fosfato)', status_glaconato3)

df_resultado_total = adicionar_status(df_resultado_total, 'Malato (ciclo citrato)', status_malatocc)

df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato 1 (glioxilato)', status_oxalato_vg1)
df_resultado_total = adicionar_status(df_resultado_total, 'Oxalato 2 (glioxilato)', status_oxalato_vg2)

df_resultado_total = adicionar_status(df_resultado_total, 'Acetato (fosfato)', status_acetato)

df_resultado_total = adicionar_status(df_resultado_total, 'Malato (metano)', status_malatomm)


df_resultado_total = adicionar_status(df_resultado_total, 'Lactato 1 (frutose e manose)', status_enzimaslactato_fm1)
df_resultado_total = adicionar_status(df_resultado_total, 'Lactato 2 (frutose e manose)', status_enzimaslactato_fm2)



# Função para combinar DataFrames com cabeçalhos numerados
def combinar_dataframes_com_numeracao(dfs, prefixo_nome):
    df_combined = pd.DataFrame()
    
    for i, df in enumerate(dfs, start=1):
        # Adiciona o prefixo e a numeração ao cabeçalho das colunas
        df_temp = df.copy()
        df_temp.columns = [f'{col} {i}' for col in df.columns]
        
        # Adiciona os DataFrames ao combinado
        df_combined = pd.concat([df_combined, df_temp], axis=1)
        
        # Adiciona uma coluna em branco para separar
        if i < len(dfs):
            df_blank = pd.DataFrame('', index=df_combined.index, columns=[''] * len(df.columns))
            df_combined = pd.concat([df_combined, df_blank], axis=1)
    return df_combined

# Listas de DataFrames para cada categoria
dataframeadenina = [
    df_oxalate_viaadenina1, df_oxalate_viaadenina2, df_oxalate_viaadenina3,
    df_oxalate_viaadenina4, df_oxalate_viaadenina5, df_oxalate_viaadenina6
]

dataframeguanina = [
    df_oxalate_viaguanina1, df_oxalate_viaguanina2, df_oxalate_viaguanina3,
    df_oxalate_viaguanina4, df_oxalate_viaguanina5, df_oxalate_viaguanina6
]

dataframemalato_vp1 = [
    df_malato_vp1, df_malato_vp2, df_malato_vp3,
    df_malato_vp4, df_malato_vp5
]

dataframeacetato_vp = [
    df_acetato_vp1, df_acetato_vp2, df_acetato_vp3, df_acetato_vp4, df_acetato_vp5, df_acetato_vp6,
    df_acetato_vp7, df_acetato_vp8, df_acetato_vp9, df_acetato_vp10, df_acetato_vp11, df_acetato_vp12,
    df_acetato_vp13, df_acetato_vp14, df_acetato_vp15, df_acetato_vp16, df_acetato_vp17, df_acetato_vp18
]

dataframgluconato = [
    df_gluconato_vg1, df_gluconato_vg2, df_gluconato_vg3
]

dataframlactato_fm = [
    df_enzimaslactato_fm1, df_enzimaslactato_fm2
]

dataframoxalato_vg = [
    df_oxalato_vg1, df_oxalato_vg2
]


# Crie um arquivo Excel com múltiplas abas
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    # Combine e salve cada grupo de DataFrames em abas diferentes
    df_adenina = combinar_dataframes_com_numeracao(dataframeadenina, 'Oxalato Via Adenina (Purinas)')
    df_guanina = combinar_dataframes_com_numeracao(dataframeguanina, 'Oxalato Via Guanina (Purinas)')
    df_malato_vp = combinar_dataframes_com_numeracao(dataframemalato_vp1, 'Malato')
    df_acetato_vp = combinar_dataframes_com_numeracao(dataframeacetato_vp, 'Acetato')
    df_gluconato = combinar_dataframes_com_numeracao(dataframgluconato, 'Gluconato')
    df_lactato = combinar_dataframes_com_numeracao(dataframlactato_fm, 'Lactato')
    df_oxalato = combinar_dataframes_com_numeracao(dataframoxalato_vg, 'Oxalato')


    # Salve cada DataFrame em uma aba separada
    df_adenina.to_excel(writer, sheet_name='Oxalato Via Adenina (Purinas)', index=False)
    df_guanina.to_excel(writer, sheet_name='Oxalato Via Guanina (Purinas)', index=False)
    df_malato_vp.to_excel(writer, sheet_name='Malato (Piruvato)', index=False)
    df_gluconato.to_excel(writer, sheet_name='Gluconato (Pentoses fosfato )', index=False)
    df_lactato.to_excel(writer, sheet_name='Lactato (Frutose e Manose)', index=False)
    df_acetato_vp.to_excel(writer, sheet_name='Acetato (Piruvato)', index=False)

    df_lactato_vp.to_excel(writer, sheet_name='Lactato (Piruvato)', index=False)

    df_malato_mm.to_excel(writer, sheet_name='Malato (Metano)', index=False)

    df_malato_cc.to_excel(writer, sheet_name='Malato (Ciclo citrato)', index=False)

    df_oxalato.to_excel(writer, sheet_name='Oxalato (Glioxilato)', index=False)

    df_acetato_mf.to_excel(writer, sheet_name='Acetato (Fosfato)', index=False)

    df_resultado_total.to_excel(writer, sheet_name='Resultado Geral', index=False)
    
print(f'Resultados foram gravados em {arquivo_saida}')


