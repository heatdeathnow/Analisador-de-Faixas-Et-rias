### Analisador de vidas
1. Aceita ambas planilhas (.xlsx e .xls)
2. Lê as colunas de sexo, titularidade e idade.
3. Cria uma planilha de análise separando isso por faixas-etárias.

##### Graus de redundância:
1. Coluna de sexo não é necessária, o programa busca o nome numa tabela lookup que os associa aos sexos.
2. Se o programa não consegue encontrar o nome nessa tabela de associação, ele usa a última letra do nome para decidir o
sexo da pessoa.
3. A coluna de idade não é necessária. Se o programa não a encontra, ele usa a coluna de idade de nascimento para
calcular a idade da pessoa e classificá-la.
4. A coluna de titularidade não é necessária. Se o programa não a encontra, ele simplesmente desconsidera a 
titularidade.
