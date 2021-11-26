import numpy as np
import pandas as pd
from glob import glob
import os

columns_res = ['UA_IPS_UP', 'UA_IPS_DOWN', 'UA_BEI_UP', 'UA_BEI_DOWN']
look_for = 'MSP/LABEO (₴/MWh)'


def create_df(path):
    df = pd.read_excel(path, sheet_name='Details')
    df['index'] = df.apply(lambda x: str(x['Date'])[:10] + ' ' + x['Settlement Period'], axis=1)
    df.index = df['index']
    df = df.drop(columns=['Date', 'Settlement Period', 'index'])
    df = df.sort_values(['index'])
    # print(df.to_string())
    return df


def select(df):
    df_new = pd.DataFrame(columns=columns_res, index=df.index)
    for item in df_new.index:
        df_new.loc[[item], columns_res[0]] = df.loc[
            (df['Market Balance Area'] == 'CA_UA_IPS') & (df['Direction'] == 'Up') & (df.index == item)][look_for][0]
        df_new.loc[[item], columns_res[1]] = df.loc[
            (df['Market Balance Area'] == 'CA_UA_IPS') & (df['Direction'] == 'Down') & (df.index == item)][look_for][0]
        df_new.loc[[item], columns_res[2]] = df.loc[
            (df['Market Balance Area'] == 'CA_UA_BEI') & (df['Direction'] == 'Up') & (df.index == item)][look_for][0]
        df_new.loc[[item], columns_res[3]] = df.loc[
            (df['Market Balance Area'] == 'CA_UA_BEI') & (df['Direction'] == 'Down') & (df.index == item)][look_for][0]
    df_new.replace(0.01, '-', regex=True, inplace=True)
    print(df_new.to_string())
    return df_new


def main():
    num = 0
    cwd = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(cwd, 'input', '*.xlsx')
    dir_list = glob(target)
    if not os.path.isdir("output"):
        os.mkdir("output")
    for item in dir_list:
        df = create_df(item)
        df = select(df)
        num += 1
        df.to_excel(f'./output/result{num}.xlsx', index=False)


if __name__ == '__main__':
    main()
