import numpy as np
import pandas as pd
from glob import glob
import os
import time

look_for = 'MSP/LABEO (₴/MWh)'
zone = ['CA_UA_IPS', 'CA_UA_BEI']
direction = ['Up', 'Down']
zone_direction = [f'{y} {x}' for y in set(zone) for x in set(direction)]


def create_df(path):
    df = pd.read_excel(path, sheet_name='Details')
    df['index'] = df.apply(lambda x: str(x['Date'])[:10] + ' ' + x['Settlement Period'], axis=1)
    df = df.drop_duplicates(subset=['index', 'Market Balance Area', 'Direction'])
    df.index = df['index']
    df = df.drop(columns=['Date', 'Settlement Period', 'index'])
    df = df.sort_values(['index'])
    return df


def manipulation(df, time_index, col):
    mod_col_list = col.split()
    results = df.loc[(df['Market Balance Area'] == mod_col_list[0]) & (df['Direction'] == mod_col_list[1]) & (
                time_index == df.index)][look_for]
    if not results.empty:
        return results[0]
    else:
        return '-'


def select(df):
    df_new = pd.DataFrame(columns=zone_direction, index=df.index.unique())
    for item in df_new.index:
        for col in df_new.columns:
            df_new.loc[item, col] = manipulation(df, item, col)
    return df_new


def main():
    cwd = os.path.dirname(os.path.abspath(__file__))
    num = 0
    target = os.path.join(cwd, 'input', 'Вхід*.xlsx')
    dir_list = glob(target)
    if not os.path.isdir("output"):
        os.mkdir("output")
    for item in dir_list:
        df = create_df(item)
        num += 1
        df.to_excel(f'{cwd}/input/check{num}.xlsx')
        df = select(df)
        df.to_excel(f'{cwd}/output/result{num}.xlsx')


if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))