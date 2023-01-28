# パッケージのインポート
import pandas as pd
from pathlib import Path
import PySimpleGUI as sg



# 区切り文字排除関数の定義
def rm_delimiter(delimiter,input_file,output_file):
    '''
    区切り文字排除関数です。
        1. テーブルデータの読み込み
        2. 区切り文字の排除し辞書データを作成(function:formmat_dict_data)
        3. 得られた辞書データをDataFrameとして再定義
        4. DataFrameに対する整形処理(function:preprocess_df)
        5. 出力

    Parameters
    ----------
    delimiter : str
        対象の区切り文字。
    input_file: str
        対象のファイル
    output_file: str
        出力先のファイル

    Returns
    -------
    True : bool
        正常終了
    False : bool
        異常終了（ファイルの読み込み失敗）
    
    
    See Also
    --------
    formmat_dict_data : 区切り文字の排除し辞書データを作成する。
    preprocess_df : DataFrameに対する整形処理。unstackして必要な情報以外削除する。
    '''

    # データの読み込み
    try:
        input_filepath = Path(input_file)

        if input_filepath.suffix == '.csv':
            df = pd.read_csv(input_filepath,dtype=str,usecols=['key','target'])
        elif input_filepath.suffix == '.xlsx' or input_filepath.suffix == '.xls':
            df = pd.read_excel(input_filepath,dtype=str,usecols=['key','target'])
        else:
            raise ValueError
    except ValueError:
        print(f'読み込み時にエラー！{input_filepath}が正しくありません。サンプルデータを参考にして修正してください')
        return False
    print('データを読み込みました。ただいま処理中です...')

    tmp_dict = {}
    # 区切り文字を排除したリストを辞書に格納する関数
    def format_dict_data(row):
        target_list = [_.strip() for _ in row['target'].split(delimiter)]
        tmp_dict.update({row['key']:target_list})

    # 関数の適用
    df.apply(format_dict_data,axis=1)

    # 得られた辞書をDataFrameとして再定義
    df_from_dict = pd.DataFrame.from_dict(tmp_dict, orient='index').T

    # DataFrameを取り回す関数
    def preprocess_df(df_):
        df_ = df_.unstack() \
            .dropna() \
            .reset_index()
        df_.drop(columns = df_.columns[1],inplace=True)
        df_.columns=['key','target']
        return df_

    # パイプの適用
    df_key_target = df_from_dict.pipe(preprocess_df)

    print('処理が完了しました。ファイルを出力中です...')

    # 出力先ディレクトリの作成
    output_filepath = Path(output_file)
    if not output_filepath.parent.exists():
        output_filepath.parent.mkdir(parents=True)

    # 出力

    if output_filepath.suffix == '.csv':
        df_key_target.to_csv(output_filepath,index=False,encoding='cp932')
    elif output_filepath.suffix == '.xlsx' or input_filepath.suffix == '.xls':
        print('Excel形式で出力中です...CSVならもっと早いよ!')
        df_key_target.to_excel(output_filepath,index=False)
    
    print('出力が完了しました。')
    return True

# GUIの操作
def main():

    init_delimiter = '|'
    init_INPUT = './input/sample_target_data.xlsx'
    init_OUTPUT = './output/sample_output_data.csv'
    sg.theme('GreenMono')

    # Windowレイアウト定義
    layout = [
            [sg.Text('いらっしゃい！ExcelのTarget列の区切り文字を分解して縦持ちにするプログラムだよー')],
            
            # 区切り文字
            [sg.Text('まずは区切り文字を指定してね！初期値は "|" だよ')],
            [sg.Input(init_delimiter,key='-DELIMITER-',size=(2,1))],

            # Inputファイル
            [sg.Text('用意したExcel or CSVデータを選択してね ※列名を指定しています：列名=[key,target]')],
            [sg.Input(init_INPUT,key='-INPUT-'),sg.FileBrowse()],

            # Outputファイル&ディレクトリ
            [sg.Text('出力するExecel or CSVデータ名と保存先を教えてね')],
            [sg.Input(init_OUTPUT,key='-OUTPUT-'),sg.FileSaveAs()],
            [sg.Button('実行'),sg.Exit('終了')]]

    window = sg.Window('区切り文字排除くん', layout)


    # Window操作
    while True:
        event, values = window.read()


        if event == '実行':
            success = rm_delimiter(
                delimiter = values['-DELIMITER-'],
                input_file = values['-INPUT-'],
                output_file = values['-OUTPUT-'])
            
            if not success:
                sg.popup_ok('処理が失敗したよ \n 黒い画面を確認しておくれ')
            else:
                sg.popup_ok('処理が正常に終了したよ \n 出力したデータを確認しておくれ')

        if event == sg.WIN_CLOSED or event == '終了':
            break

    window.close()