#%%
import pandas as pd
from pathlib import Path
#%%

# for test in VSCode
excelpath = Path(r"C:\Users\gingi\Downloads\target_scival_publication.xlsx")
df = pd.read_excel(excelpath,dtype=str,usecols=['key','target'])

tmp_dict = {}
for record in df.to_dict(orient='records'):

    # 区切り文字の処理
    target_list = [_.strip() for _ in record['target'].split('|')]
    tmp_dict.update({record['key']:target_list})
    # break

# unstack
df_from_dict = pd.DataFrame.from_dict(tmp_dict, orient='index').T

df_unstack = df_from_dict.unstack() \
                .dropna() \
                .reset_index()
                
# 列名の定義・整理
df_unstack.columns = ['key','drop_dummy','target']
df_key_target = df_unstack.drop(columns='drop_dummy')
df_key_target.to_excel('../output/sample_output_unstack.xlsx',index=False)

#%%
# for test in VSCode
# excelpath = Path('../input/sample_target_data.xlsx')
excelpath = Path(r"C:\Users\gingi\Downloads\target_scival_publication.xlsx")
df = pd.read_excel(excelpath,dtype=str,usecols=['key','target'])

# 格納先リストの定義
datalist = []

# dfのレコードに対する繰り返し処理：区切り文字を取ってdatalistに格納する
for record in df.to_dict(orient='records'):

    # 区切り文字の処理
    target_list = [_.strip() for _ in record['target'].split('|')]

    # DataFarmeの元データを作成
    for target in target_list:
        datalist.append({'key':record['key'],'target':target})


pd.DataFrame(datalist).to_excel('../output/sample_output_target_list.xlsx',index=False)
# df_key_target.to_excel('../output/sample_output_target_list.xlsx',index=False)



#%%
# applyの使用確認
def tmp_func(x):
    print(x['key'])
    return x + ":tmp_func"


# df.head(3).apply(tmp_func,axis=1)

#%% 
# unstack ver.
# for test in VSCode

def format_dict_data(row):
    target_list = [_.strip() for _ in row['target'].split('|')]
    tmp_dict.update({row['key']:target_list})

def preprocess_df(df_):
    df_ = df_.unstack() \
        .dropna() \
        .reset_index()
    df_.drop(columns = df_.columns[1],inplace=True)
    df_.columns=['key','target']
    return df_
    

excelpath = Path(r"C:\Users\gingi\Downloads\target_scival_publication.xlsx")
df = pd.read_excel(excelpath,dtype=str,usecols=['key','target'])


df.apply(format_dict_data,axis=1)
df_from_dict = pd.DataFrame.from_dict(tmp_dict, orient='index').T

df_key_target = df_from_dict.pipe(preprocess_df)


df_key_target.to_excel('../output/sample_output_unstack.xlsx',index=False)


# %%
excelpath = Path(r"C:\Users\gingi\Downloads\target_scival_publication.xlsx")
df = pd.read_excel(excelpath,dtype=str,usecols=['key','target'])

# 格納先リストの定義
datalist = []

# # dfのレコードに対する繰り返し処理：区切り文字を取ってdatalistに格納する
# for record in df.to_dict(orient='records'):

#     # 区切り文字の処理
#     target_list = [_.strip() for _ in record['target'].split('|')]

#     # DataFarmeの元データを作成
#     for target in target_list:
#         datalist.append({'key':record['key'],'target':target})

def testfunc_2(row):
    target_list = [_.strip() for _ in row['target'].split('|')]

    for target in target_list:
        datalist.append({'key':row['key'],'target':target})

df.apply(testfunc_2,axis=1)

pd.DataFrame(datalist).to_excel('../output/sample_output_target_list.xlsx',index=False)
# %%

# ファイルテーブルの拡張子を問わず処理したい
path = Path('../input/sample_target_data_csv.csv')
# %%
