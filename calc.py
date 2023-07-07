from openpyxl.worksheet.worksheet import Worksheet
from dateutil.relativedelta import relativedelta
from var import lookup_table, banned_words
from unidecode import unidecode
from pandas import DataFrame
from typing import Sequence
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

    if tostr(sex) == 'F' and tostr(title) == 'T':
        col = 'B'
    elif tostr(sex) == 'M' and tostr(title) == 'T':
        col = 'C'
    elif tostr(sex) == 'F' and tostr(title) == 'D':
        col = 'D'
    elif tostr(sex) == 'M' and tostr(title) == 'D':
        col = 'E'
    elif tostr(sex) == 'F' and title is None:
        col = 'F'
    elif tostr(sex) == 'M' and title is None:
        col = 'G'
    else:
        raise ValueError('Os únicos que essa função aceita para sexo são "M" e "F". E os únicos valores que essa função'
                         'aceita para titularidade as strings "T" e "D" e o valor None.')

    for i in range(3, 13):
        ws[f'{col}{i}'] = ranges[i - 3]


def get_sex_from_name(df: DataFrame, index_name: int, row: int, lookup_table: DataFrame) -> str:
    name = tostr(df.loc[row][index_name].split(' ')[0])  # Retira acentos, coloca tudo em maiúsculo e pega o primeiro nome.

    try:
        index = lookup_table.nome[lookup_table.nome == name].index[0]  # Acha a posição desse nome na tabela de lookup.
        return lookup_table.sexo[index]  # Retorna o sexo associado ao nome nessa posição.
    except IndexError:  # Caso o nome não esteja na tabela de lookup, use esse algoritmo.
        if name[-1] in 'AEIYZ':  # Se o nome termina com essas letras, assumir que é feminino.
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

    raise IndexError


def get_ranges(df: DataFrame,
               index_name: int | None,
               index_sex: int | None,
               sex: str,
               index_title: int | None,
               title: str,
               index_age: int | None,
               index_birth: int | None) -> list:
    """
    Retorna uma lista com 10 inteiros, cada um representando a quantidade de pessoas do sexo e titularidade
    especificadas nos parâmetros da função naquela faixa etária específica encontrados no Dataframe. Se não há índice de
    sexo, é chamada a função get_sex_from_name para tentar determiná-lo.Se não há titularidade, ou seja,
    index_tile=None, a função a desconsidera.
    """

    sex = tostr(sex)
    if sex != 'M' and sex != 'F':
        raise ValueError('Apenas as strings "M" e "F" são aceitas por essa função.')

    title = tostr(title)
    if title != 'T' and title != 'D':
        raise ValueError('Apenas as string "T" e "D" são aceitas por essa função.')

    if index_age is None and index_birth is None:
        raise ValueError('Essa função precisa que pelo menos um desses dois parâmetros tenham valores não-nulos: "index_age" ou "index_birth".')
    elif index_age is not None:
        def age(j): return df.loc[j][index_age]
    else:
        def age(j): return relativedelta(date.today(), df.loc[j][index_birth].date()).years

    if index_sex is None and index_name is None:
        raise ValueError('Essa função precisa que pelo menos um desses dois parâmetros tenham valores não-nulos: "index_sex" ou "index_name".')
    elif index_sex is None:
        def sex_(j): return get_sex_from_name(df, index_name, j, lookup_table)
    else:
        def sex_(j): return df.loc[j][index_sex]

    if index_title is None:  # Se não houver titularidade...
        def stewardness(_): return title  # A última parte dos if statements será sempre verdadeira (não interferirá).
    else:
        def stewardness(j): return 'T' if tostr(df.loc[j][index_title]) in ('TITULAR', 'T') else 'D'

    ranges = [0] * 10
    for i in range(len(df.index)):
        if age(i) <= 18 and sex_(i) == sex and stewardness(i) == title:
            ranges[0] += 1
        elif age(i) <= 23 and sex_(i) == sex and stewardness(i) == title:
            ranges[1] += 1
        elif age(i) <= 28 and sex_(i) == sex and stewardness(i) == title:
            ranges[2] += 1
        elif age(i) <= 33 and sex_(i) == sex and stewardness(i) == title:
            ranges[3] += 1
        elif age(i) <= 38 and sex_(i) == sex and stewardness(i) == title:
            ranges[4] += 1
        elif age(i) <= 43 and sex_(i) == sex and stewardness(i) == title:
            ranges[5] += 1
        elif age(i) <= 48 and sex_(i) == sex and stewardness(i) == title:
            ranges[6] += 1
        elif age(i) <= 53 and sex_(i) == sex and stewardness(i) == title:
            ranges[7] += 1
        elif age(i) <= 58 and sex_(i) == sex and stewardness(i) == title:
            ranges[8] += 1
        elif age(i) >= 59 and sex_(i) == sex and stewardness(i) == title:
            ranges[9] += 1
    return ranges
