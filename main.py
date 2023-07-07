from openpyxl import load_workbook
import pandas as pd
from calc import *
from var import *
import argparse


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('dir')
    args = parser.parse_args()

    lookup_sex_dir = 'dados\\sex.csv'  # Diretório para tabela lookup que associa os nomes aos sexos.
    model_dir = 'dados\\model.xlsx'  # Diretório para o arquivo modelo formatado.
    excel_type = '.xlsx'

    if args.dir[-5:] != '.xlsx' and args.dir[-4:] != '.xls':
        raise ValueError('As únicas extensões aceitas são .xlsx e .xls')
    elif args.dir[-4:] == '.xls':
        input_file = args.dir
        output_file = input_file.replace('.xls', '') + ' - análise.xls'
    else:
        input_file = args.dir
        output_file = input_file.replace('.xlsx', '') + ' - análise.xlsx'
    df = pd.read_excel(input_file)

    try:
        index_age = get_index(df, lookup_age)
        index_birth = None
    except IndexError:
        try:
            index_birth = get_index(df, lookup_birth)
            index_age = None
        except IndexError:
            raise KeyError('Nem idade nem data de nascimento foram encontrados na planilha.')

    try:
        index_sex = get_index(df, lookup_sex)
        index_name = None
    except IndexError:
        try:
            index_name = get_index(df, lookup_names)
            index_sex = None
        except IndexError:
            raise KeyError('Nem sexo nem nome foram encontrados na planilha.')

    try:
        index_title = get_index(df, lookup_title)
    except IndexError:
        index_title = None

    model = load_workbook(model_dir)
    ws = model.worksheets[0]
    if index_title is not None:
        female_stewards_range = get_ranges(df, index_name, index_sex, 'F', index_title, 'T', index_age, index_birth)
        male_stewards_range   = get_ranges(df, index_name, index_sex, 'M', index_title, 'T', index_age, index_birth)
        female_proteges_range = get_ranges(df, index_name, index_sex, 'F', index_title, 'D', index_age, index_birth)
        male_proteges_range   = get_ranges(df, index_name, index_sex, 'M', index_title, 'D', index_age, index_birth)

        fill_column(ws, female_stewards_range, 'F', 'T')
        fill_column(ws, male_stewards_range, 'M', 'T')
        fill_column(ws, female_proteges_range, 'F', 'D')
        fill_column(ws, male_proteges_range, 'M', 'D')

    else:
        female_range = get_ranges(df, index_name, index_sex, 'F', index_title, '.', index_age, index_birth)
        male_range   = get_ranges(df, index_name, index_sex, 'M', index_title, '.', index_age, index_birth)

        fill_column(ws, female_range, 'F', None)
        fill_column(ws, male_range, 'F', None)

    model.save(output_file)
