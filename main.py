"""エクセルのパスをつなげるためのインストール"""
import os
import sys
import webbrowser
import PySimpleGUI as sg
import pandas
import pyperclip
import sqlite3

db = sqlite3.connect(":memory:")
sg.theme("DarkBrown3")

#レイアウト
layout = [[sg.T('下にキャラ名を入力してください')],
          [sg.I(k='in1', size=20, key='text',),sg.B('入力',key = 'btn'),sg.B('クリア',key= 'clear')],
          [sg.Combo(key='search_result', size=19, enable_events=True, values=[])],
          [sg.I(key='text1', size=20, readonly=True,default_text=""),sg.B('サイト遷移',key='btn1')],
          [sg.I(key='copyright_key',size=(20),readonly=True),
           sg.B('コピー',key='btn2')]
          ]

window = sg.Window(title='合成音声規約検索',
                   layout=layout,
                   font=('BIZ UDPゴシック',16),
                   finalize=True,
                   resizable=True,)

#エクセルにまとめてあるurlを表示
def excel(value):
    """pathlibを使いエクセルのパスをつなげる"""
    excel_pass = os.path.join(os.path.dirname(sys.argv[0]),'./excel/index.xlsx')
    work_sheet = pandas.read_excel(excel_pass)
    work_sheet.columns = work_sheet.columns.str.strip()
    #A列から検索
    row = work_sheet.loc[work_sheet["A"] == value]
    if not row.empty:
        voice = row["B"].values[0]
        voice_copyright = row["C"].values[0]
        #そのままリンクを貼るとエラーを吐くので文字列に変換する
        text = f"{voice}"
        window["text1"].update(text)
        if not voice_copyright == "なし":
            #そのまま貼るとエラー
            copyright_text = f"{voice_copyright}"
            window["copyright_key"].update(copyright_text)
        else:
            window["copyright_key"].update("")
    else:
        return None

#jump()に渡す準備
def jump():
    """そのまま入れるとエラーが起きるのでtext1の値をweb_jumpに引き継ぐ"""
    web_jump = values['text1']
    if web_jump == "":
        pass
    else:
        webbrowser.get().open(web_jump,new=0)

def copy():
    """クレジットをコピペするためのイベント"""
    pyperclip.copy(values["copyright_key"])

def setup_db():
    global db
    excel_pass = os.path.join(os.path.dirname(sys.argv[0]),'./excel/index.xlsx')
    work_sheet = pandas.read_excel(excel_pass)
    db.execute("CREATE TABLE credit( character text, URL text, credit text)")
    for row in work_sheet.values:
        db.execute(f"INSERT INTO credit (character,URL,credit) values ('{row[0]}', '{row[1]}' , '{row[2]}')")
    db.commit()

def display_credit(name):
    global db
    cursor = db.cursor()
    result = cursor.execute(f"select * from credit where character like '%{name}%'").fetchall()
    if len(result) == 0:
        return
    voice = result[0][1]
    voice_copyright = result[0][2]
    # そのままリンクを貼るとエラーを吐くので文字列に変換する
    text = f"{voice}"
    window["text"].update(result[0][0])
    window["text1"].update(text)
    if not voice_copyright == "なし":
        # そのまま貼るとエラー
        copyright_text = f"{voice_copyright}"
        window["copyright_key"].update(copyright_text)
    else:
        window["copyright_key"].update("")

def get_value(name):
    global db
    cursor = db.cursor()
    result = cursor.execute(f"select * from credit where character like '%{name}%'").fetchall()
    result_list = []
    for item in result:
        result_list.append(item[0])
    window["search_result"].update(values=result_list)
    if len(result) != 0:
        display_credit(result[0][0])


setup_db()

while True:
    event,values = window.read()
    if event == 'btn1':
        jump()
    if event == 'clear':
        window["text"].update("")
        window["text1"].update("")
        window["copyright_key"].update("")
    if event == 'btn' and not values["text"] == "":
        get_value(values["text"])
    if event =='btn2':
        if not values["copyright_key"] == "":
            copy()
        else:
            pyperclip.copy("")
    if event == "search_result":
        display_credit(values['search_result'])
    if event == sg.WIN_CLOSED:
        break

window.close()
