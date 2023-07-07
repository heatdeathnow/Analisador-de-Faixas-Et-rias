from pandas import read_csv

# Valores que o programa vai procurar nas colunas da planilha passada para descobrir onde ficam as informações úteis.
lookup_names = ('Beneficiário', 'Colaborador', 'Titular', 'Nome', )
lookup_birth = ('Nascimento', 'Data de nasc', 'nasc', )
lookup_age = ('Idade', )
lookup_sex = ('M/F', 'M / F', 'M/ F', 'M /F', 'Sexo', 'Gênero', )
lookup_title = ('Tipo', 'Titularidade', 'Dependência', )
banned_words = ('Mensalidade', )  # Isso é necessário para casos como "idade" existindo dentro de "mensalidade".

lookup_table = read_csv('dados\\sex.csv', sep=';')  # Arquivos .csv brasileiros tem ";" como separador e não ",".
