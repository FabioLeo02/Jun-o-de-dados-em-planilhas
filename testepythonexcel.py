import pandas as pd
#mostra a tabela no terminal do python
#OBS: é necessario colocar o local onde o arquivo está para assim o codigo o localizar
tabela = pd.read_excel("C:/Users/Leandro/Documents/python/Excel com Python/Produtos.xlsx")
print(tabela) #mostra a tablea no prompt 


#atualizar o multiplicador da planilha
# (nome da variavel).loc - localiza a coluna ou linha na planilha
# a coluna ou linha pode ser localizada pela sua letra, numero ou apartir de seu nome, caso o possui
# OBS: caso a linha ou coluna possua nome, ele deve ser escrito no codigo do mesmo jeito que está na planilha
tabela.loc[tabela["Tipo"]=='Serviço', "Multiplicador imposto"] = 1.5

#fazer a conta da coluna "Preço Base Reais"
tabela['Preço Base Reais'] = tabela['Multiplicador Imposto'] * tabela['Preço Base Reais']

#cria a nova versão do arquivo alterado caso sejá escrito um novo nome, do contrario 
#será criado um novo arquivo, cujo o mesmo tbm irá substituir o outro caso sejá atualizado
#com o mesmo nome 

tabela.to_excel('ProdutosPandas.xlsx',index=False) 
# (nome da varivavel).to_excel('nome da planilha') - cria uma nova planilha
# index=False - tira a numeração das linhas ao cirar uma nova planilh