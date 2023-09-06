from openpyxl.worksheet.worksheet import Worksheet
from dateutil.relativedelta import relativedelta
from var import lookup_table, banned_words
from typing import Sequence, Literal
from unidecode import unidecode
from pandas import DataFrame
from datetime import date


def tostr(x: str) -> str:
    """
    Recebe uma string e retorna uma string em caixa alta, sem acentos e sem espaços antes e depois.
    """

    return unidecode(str(x).upper().strip())

def fill_column(ws: Worksheet, ranges: Sequence, sex: str, title: str | None) -> None:
    """
    Preenche uma coluna específica da planilha modelo passada dependendo do sexo e da titularidade passados como
    argumentos. Função void, ela não retorna nada, apenas modifica a planilha passada.
    """

    match (tostr(sex), tostr(title)):
        case ('F', 'T'):
            col = 'B'

        case ('M', 'T'):
            col = 'C'
        
        case ('F', 'D'):
            col = 'D'
        
        case ('M', 'D'):
            col = 'E'
        
        case ('F', None):
            col = 'F'
        
        case ('M', None):
            col = 'G'

        case _ :
            raise ValueError(f'Os únicos valores que essa função aceita são "M" e "F" para sexo, e "T" e "D" para titularidade.\n \
                             Porém fora passado "{sex}" para sexo e "{title}" para titularidade.')

    for i in range(3, 13):
        ws[f'{col}{i}'] = ranges[i - 3]

def get_sex_from_name(df: DataFrame, index_name: int, row: int, lookup_table: DataFrame) -> str:
    name = tostr(df.loc[row][index_name].split(' ')[0])  # Retira acentos, coloca tudo em maiúsculo e pega o primeiro nome.

    try:
        index = lookup_table.nome[lookup_table.nome == name].index[0]  # Acha a posição desse nome na tabela de lookup.
        return lookup_table.sexo[index]  # Retorna o sexo associado ao nome nessa posição.
    
    except IndexError:  # Caso o nome não esteja na tabela de lookup, use esse algoritmo.
        if name[-1] in 'AEIYZS':  # Se o nome termina com essas letras, assumir que é feminino.
            return 'F'
        
        else:  # Senão, assumir que é masculino.
            return 'M'

def get_index(df: DataFrame, names: Sequence) -> int:
    """
    Procura os nomes passados como parâmetro nos títulos das colunas do Dataframe e retorna a sua posição.
    """

    if not isinstance(names, Sequence):
        raise TypeError

    i = 0
    for column in df.columns:
        for name in names:
            if tostr(name) in tostr(column) and tostr(column) not in [tostr(word) for word in banned_words]:
                return i
            
        i += 1

    raise IndexError(f'Não foi possível encontrar {name} nas colunas da planilha passada...')

def get_ranges(df: DataFrame,
               index_name: int | None,
               index_sex: int | None,
               sex: str,
               index_title: int | None,
               title: str,
               index_age: int | None,
               index_birth: int | None) -> list[int]:
    """
    Retorna uma lista com 10 inteiros, cada um representando a quantidade de pessoas do sexo e titularidade
    especificadas nos parâmetros da função naquela faixa etária específica encontrados no Dataframe. Se não há índice de
    sexo, é chamada a função get_sex_from_name para tentar determiná-lo.Se não há titularidade, ou seja,
    index_tile=None, a função a desconsidera.
    """

    sex = tostr(sex)
    if sex != 'M' and sex != 'F':
        raise ValueError(f'Apenas as strings "M" e "F" são aceitas por essa função, porém o valor "{sex}" fora passado.')

    title = tostr(title)
    if title != 'T' and title != 'D':
        raise ValueError(f'Apenas as string "T" e "D" são aceitas por essa função, porém o valor "{title}" fora passado.')

    if index_age is None and index_birth is None:
        raise ValueError('Essa função precisa que pelo menos um dos dois parâmetros tenham valores não-nulos: "index_age" ou "index_birth".')
    
    elif index_age is not None:
        def age(index: int) -> int: return df.loc[index][index_age]

    else:
        def age(index: int) -> int: return relativedelta(date.today(), df.loc[index][index_birth].date()).years

    if index_sex is None and index_name is None:
        raise ValueError('Essa função precisa que pelo menos um dos dois parâmetros tenham valores não-nulos: "index_sex" ou "index_name".')
    
    elif index_sex is None:
        def sex_(index: int) -> str: return get_sex_from_name(df, index_name, index, lookup_table)

    else:
        def sex_(index: int) -> str: return df.loc[index][index_sex]

    if index_title is None:  # Se não houver titularidade...
        def stewardness(_) -> Literal['T', 'D']: return title  # A última parte dos if statements será sempre verdadeira (não interferirá).
        
    else:
        def stewardness(index: int) -> Literal['T', 'D']: return 'T' if tostr(df.loc[index][index_title]) in ('TITULAR', 'T') else 'D'

    ranges = [0] * 10
    for i in range(len(df.index)):
        match (age(i), sex_(i), stewardness(i)):
            case x, y, z if x <= 18 and y == sex and z == title:
                ranges[0] += 1

            case x, y, z if x <= 23 and y == sex and z == title:
                ranges[1] += 1

            case x, y, z if x <= 28 and y == sex and z == title:
                ranges[2] += 1

            case x, y, z if x <= 33 and y == sex and z == title:
                ranges[3] += 1
            
            case x, y, z if x <= 38 and y == sex and z == title:
                ranges[4] += 1
            
            case x, y, z if x <= 43 and y == sex and z == title:
                ranges[5] += 1
            
            case x, y, z if x <= 48 and y == sex and z == title:
                ranges[6] += 1
            
            case x, y, z if x <= 53 and y == sex and z == title:
                ranges[7] += 1
            
            case x, y, z if x <= 58 and y == sex and z == title:
                ranges[8] += 1
            
            case x, y, z if x >= 59 and y == sex and z == title:
                ranges[9] += 1

    return ranges
