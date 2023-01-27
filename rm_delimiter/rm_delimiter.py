
# パッケージのインポート
import pandas as pd
from pathlib import Path
import PySimpleGUI as sg


# 区切り文字排除関数の定義
def rm_delimiter(delimiter,inputfile,outputfile):

    try:
        # データの読み込み
        excelpath = Path(inputfile)
        df = pd.read_excel(excelpath,dtype=str,usecols=['key','target'])
    except ValueError:
        print('読み込み時にエラー！処理データの列名が正しくありません サンプルデータを参考にして列名を決めてください')
        return False

    # 格納先リストの定義
    datalist = []

    # dfのレコードに対する繰り返し処理：区切り文字を取ってdatalistに格納する
    for record in df.to_dict(orient='records'):

        # 区切り文字の処理
        target_list = [_.strip() for _ in record['target'].split(delimiter)]

        # DataFarmeの元データを作成
        for target in target_list:
            datalist.append({'key':record['key'],'target':target})


    # 出力先ディレクトリの作成
    output_filename = Path(outputfile)
    if not output_filename.parent.exists():
        output_filename.parent.mkdir(parents=True)

    # 出力
    pd.DataFrame(datalist).to_excel(output_filename,index=False)

    return True

def main():

    init_delimiter = '|'
    init_INPUT = './input/sample_target_data.xlsx'
    init_OUTPUT = './output/sample_output_data.xlsx'
    sg.theme('GreenMono')

    # Windowレイアウト定義
    layout = [
            [sg.Text('いらっしゃい！ExcelのTarget列の区切り文字を分解して縦持ちにするプログラムだよー')],
            
            # 区切り文字
            [sg.Text('まずは区切り文字を指定してね！初期値は "|" だよ')],
            [sg.Input(init_delimiter,key='-DELIMITER-',size=(2,1))],

            # Inputファイル
            [sg.Text('用意したExcelデータを選択してね ※形式を指定しています：列名=[key,target]')],
            [sg.Input(init_INPUT,key='-INPUT-'),sg.FileBrowse()],

            # Outputファイル&ディレクトリ
            [sg.Text('出力するExecelデータ名と保存先を教えてね')],
            [sg.Input(init_OUTPUT,key='-OUTPUT-'),sg.FileBrowse()],
            [sg.Button('実行'),sg.Exit('終了')]]

    window = sg.Window('区切り文字排除くん', layout)


    # Window操作


    while True:
        event, values = window.read()


        if event == '実行':
            success = rm_delimiter(
                delimiter = values['-DELIMITER-'],
                inputfile = values['-INPUT-'],
                outputfile = values['-OUTPUT-'])
            
            if not success:
                sg.popup_ok('処理が失敗したよ \n 黒い画面を確認しておくれ')
            else:
                sg.popup_ok('処理が正常に終了したよ \n 出力したデータを確認しておくれ')

        if event == sg.WIN_CLOSED or event == '終了':
            break

    window.close()