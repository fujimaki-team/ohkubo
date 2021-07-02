import openpyxl as excel
import datetime
import os
import json
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill


def set_sheet_name(jsn, sheet1):
    '''シート名を設定する'''
    key_name = "sheet_name"
    try:
        sheet1.title = jsn[key_name]
    except ValueError:
        print("sheet_name_Error")
    pass


def set_column(cols, sheet1):
    '''列の幅を設定する'''
    key_cols = "column"
    try:
        sheet1.column_dimensions[cols[key_cols]].width = cols["width"]
    except ValueError:
        print("cols_Error")
    pass


def set_line(rows, sheet1):
    '''行の幅を設定する'''
    key_rows = "line"
    try:
        sheet1.row_dimensions[rows[key_rows]].height = rows["height"]
    except ValueError:
        print("rows_Error")
    pass


def set_merges(merges, sheet1):
    '''セルの結合箇所を設定する'''
    key_merg = "merge"
    try:
        sheet1.merge_cells(merges[key_merg])
    except ValueError:
        print("merge_Error")
    pass


def set_borders(borders, sheet1):
    '''枠線の場所を設定する'''
    key_border = "cells"
    border = borders["border"]
    top = border["top"]
    bottom = border["bottom"]
    left = border["left"]
    right = border["right"]
    t = Side(style=top["style"], color=top["color"])
    b = Side(style=bottom["style"], color=bottom["color"])
    le = Side(style=left["style"], color=left["color"])
    r = Side(style=right["style"], color=right["color"])
    try:
        for rows in sheet1[borders[key_border]]:
            for cell in rows:
                cell.border = Border(
                    top=t, bottom=b, left=le, right=r)
    except ValueError:
        print("border_Error")
    pass


def set_sysdate(sysdate, sheet1, key):
    '''日付を設定する'''
    if "todey" == sysdate["value"]:
        try:
            sheet1[sysdate[key]] = datetime.date.today()
        except ValueError:
            print("sysdate_Error")
    else:
        print("JSONファイルの日付書式が間違えています")
    pass


def set_title(cell1, sheet1, key):
    '''出力文字を設定する'''
    try:
        sheet1[cell1[key]] = cell1["value"]
    except ValueError:
        print("Title_Error")
    pass


def set_font(cell1, sheet1, key):
    ''' フォントを設定する '''
    try:
        fonts = cell1["font"]
        sheet1[cell1[key]].font = Font(
            name=fonts["name"], bold=fonts["bold"], size=fonts["size"])
    except ValueError:
        print("Title_Font_Error")
    pass


def set_alignment(cell1, sheet1, key):
    ''' アラインメントを設定する '''
    try:
        alig = cell1["alignment"]
        sheet1[cell1[key]].alignment = Alignment(
            horizontal=alig["horizontal"], vertical=alig["vertical"])
    except ValueError:
        print("Title_alignment_Error")
    pass


def set_fill(cell1, sheet1, key):
    ''' セル内の色を設定する '''
    try:
        fills = cell1["fill"]
        sheet1[cell1[key]].fill = PatternFill(
            patternType=fills["patternType"], fgColor=fills["fgColor"])
    except ValueError:
        print("Title_fill_Error")
    pass


def main(sheet1):

    with open(os.path.join('.', '表紙test.json'), encoding='utf-8') as f:
        jsn = json.load(f)

    # シート名設定
    set_sheet_name(jsn, sheet1)

    # セルの幅
    for cols in jsn["cols"]:
        set_column(cols, sheet1)

    # セルの高さ
    for rows in jsn["rows"]:
        set_line(rows, sheet1)

    # セルの結合設定
    for merges in jsn["merges"]:
        set_merges(merges, sheet1)

    # ----枠線の場所設定
    for borders in jsn["borders"]:
        set_borders(borders, sheet1)

    '''
    # 枠線の種類と色設定
    # side_b = Side(style='thick', color='000000')
    # side_s = Side(style='thin', color='000000')

    # タイトル
    for rows in sheet1["A4:E6"]:
        for cell in rows:
            cell.border = Border(
                top=side_b, bottom=side_b, left=side_b, right=side_b)

    # サブタイトル
    for rows in sheet1["B13:D17"]:
        for cell in rows:
            if cell == (sheet1["B13"] or sheet1["C13"] or sheet1["D13"]):
                cell.border = Border(
                    top=side_b, bottom=side_s, left=side_b, right=side_b)
            elif cell == (sheet1["B17"] or sheet1["C17"] or sheet1["D17"]):
                cell.border = Border(
                    top=side_s, bottom=side_b, left=side_b, right=side_b)
            else:
                cell.border = Border(
                    top=side_s, bottom=side_s, left=side_b, right=side_b)

    # バージョン、作成日、作成者欄
    for rows in sheet1["B20:D21"]:
        for cell in rows:
            if cell == (sheet1["B20"] or sheet1["C20"] or sheet1["D20"]):
                cell.border = Border(
                    top=side_b, bottom=side_s, left=side_b, right=side_b)
            else:
                cell.border = Border(
                    top=side_s, bottom=side_b, left=side_b, right=side_b)
    '''

    # 文字表示設定
    # 日付
    key = 'coordinate'
    set_sysdate(jsn["sysdate"], sheet1, key)

    for cell1 in jsn["cells_title"]:

        # 出力文字設定
        set_title(cell1, sheet1, key)

        # フォント設定
        set_font(cell1, sheet1, key)

        # 文字上下左右そろえ
        set_alignment(cell1, sheet1, key)

        set_fill(cell1, sheet1, key)


if __name__ == "__main__":

    # ブック作成
    book = excel.Workbook()
    sheet = book.active

    # メイン処理
    main(sheet)

    # 保存
    book.save(os.path.join('.', 'test.xlsx'))
