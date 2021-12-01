import numpy as np
import pandas as pd
from glob import glob
import os

look_for = 'MSP/LABEO (₴/MWh)'
zone = ['CA_UA_IPS', 'CA_UA_BEI']
direction = ['Down', 'Up']
zone_direction = [f'{y} {x}' for y in set(zone) for x in set(direction)]


def create_df(path):
    df = pd.read_excel(path, sheet_name='Details')
    df['index'] = df.apply(lambda x: str(x['Date'])[:10] + ' ' + x['Settlement Period'], axis=1)
    df.index = df['index']
    df = df.drop(columns=['Date', 'Settlement Period', 'index'])
    df = df.sort_values(['index'])
    return df


def manipulation(df, time_index, col):
    mod_col_list = col.split()
    return df.loc[(df['Market Balance Area'] == mod_col_list[0]) & (df['Direction'] == mod_col_list[1])
                  & (df.index == time_index)][look_for][0]


def select(df):
    df_new = pd.DataFrame(columns=zone_direction, index=df.index)
    for item in df_new.index:
        for col in df_new.columns:
            df_new.loc[item, col] = manipulation(df, item, col)
    df_new = df_new.drop_duplicates()
    return df_new


def main():
    cwd = os.path.dirname(os.path.abspath(__file__))
    num = 0
    target = os.path.join(cwd, 'input', '*.xlsx')
    dir_list = glob(target)
    if not os.path.isdir("output"):
        os.mkdir("output")
    for item in dir_list:
        df = create_df(item)
        df = select(df)
        num += 1
        df.to_excel(f'{cwd}/output/result{num}.xlsx')


if __name__ == '__main__':
    main()
