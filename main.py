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


def manipulation_fast(df, col):
    mod_col_list = col.split()
    result = df[(df['Market Balance Area'] == mod_col_list[0]) & (df['Direction'] == mod_col_list[1])]
    return result[look_for]


def select(df):
    df_new = pd.DataFrame(index=df.index.unique())
    for col in zone_direction:
        df_new = pd.merge(df_new, manipulation_fast(df, col), left_index=True, right_index=True)
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
        df.columns = zone_direction
        df.to_excel(f'{cwd}/output/result{num}.xlsx')


if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))