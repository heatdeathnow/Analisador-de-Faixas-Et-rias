from openpyxl import load_workbook
from colorama import Fore, Style
from time import perf_counter
import pandas as pd
from calc import *
from var import *
import argparse


if __name__ == '__main__':
    runtime = perf_counter()
    parser = argparse.ArgumentParser()
    parser.add_argument('dir')
    args = parser.parse_args()

    lookup_sex_dir = 'dados\\sex.csv'  # Diretório para tabela lookup que associa os nomes aos sexos.
    model_dir = 'dados\\model.xlsx'  # Diretório para o arquivo modelo formatado.

    if args.dir[-5:] != '.xlsx' and args.dir[-4:] != '.xls':
        raise ValueError(f'As únicas extensões aceitas são .xlsx e .xls, porém fora passado {args.dir[-5:]}.')

    else:
        input_file = args.dir
        output_file = input_file.replace('.xlsx', '').replace('.xls', '') + ' - análise.xlsx'

    df = pd.read_excel(input_file)

    try:
        index_age = get_index(df, lookup_age)
        index_birth = None

    except IndexError:
        try:
            index_birth = get_index(df, lookup_birth)
            index_age = None

        except IndexError:
            raise KeyError('Incapaz de identificar uma coluna de idade ou de data de nascimento na planilha dada.')

    try:
        index_sex = get_index(df, lookup_sex)
        index_name = None

    except IndexError:
        try:
            index_name = get_index(df, lookup_names)
            index_sex = None

        except IndexError:
            raise KeyError('Incapaz de identificar uma coluna de sexo ou de nome na planilha dada.')

    try:
        index_title = get_index(df, lookup_title)

    except IndexError:
        index_title = None
        print(f'{Fore.RED}Incapaz de identificar uma coluna de titularidade na planilha dada. A execução continuará sem a titularidade.{Style.RESET_ALL}')

    model = load_workbook(model_dir)
    ws = model.worksheets[0]

    if index_title is not None:
        start_time = perf_counter()
        print('Calculando a faixa etária das mulheres titulares... ', end = '')
        female_stewards_range = get_ranges(df, index_name, index_sex, 'F', index_title, 'T', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        start_time = perf_counter()
        print('Calculando a faixa etária dos homens titulares... ', end = '')
        male_stewards_range   = get_ranges(df, index_name, index_sex, 'M', index_title, 'T', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        start_time = perf_counter()
        print('Calculando a faixa etária das mulheres dependentes... ', end = '')
        female_proteges_range = get_ranges(df, index_name, index_sex, 'F', index_title, 'D', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        start_time = perf_counter()
        print('Calculando a faixa etária dos homens dependentes... ', end = '')
        male_proteges_range   = get_ranges(df, index_name, index_sex, 'M', index_title, 'D', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        fill_column(ws, female_stewards_range, 'F', 'T')
        fill_column(ws, male_stewards_range, 'M', 'T')
        fill_column(ws, female_proteges_range, 'F', 'D')
        fill_column(ws, male_proteges_range, 'M', 'D')

    else:
        start_time = perf_counter()
        print('Calculando a faixa etária das mulheres... ', end = '')
        female_range = get_ranges(df, index_name, index_sex, 'F', index_title, '.', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        start_time = perf_counter()
        print('Calculando a faixa etária dos homens... ', end = '')
        male_range   = get_ranges(df, index_name, index_sex, 'M', index_title, '.', index_age, index_birth)
        print(f'cálculo concluído em {perf_counter() - start_time:.2f} segundos.')

        fill_column(ws, female_range, 'F', None)
        fill_column(ws, male_range, 'F', None)

    model.save(output_file)
    print(f'A execução foi bem-sucedida e terminada em {perf_counter() - runtime:.2f} segundos.')
