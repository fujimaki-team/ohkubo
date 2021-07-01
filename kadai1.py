import openpyxl as excel
import datetime
import os
import json
from openpyxl.styles.borders import Border,Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill


def set_font(font):
    ''' フォントを設定する '''
    pass


def set_alignment(alignment):
    ''' アラインメントを設定する '''
    pass


def main(sheet1):

    with open(os.path.join('.','表紙test.json'), encoding='utf-8') as f:
        jsn = json.load(f)
    
    #----シート名設定--------------------------------------------------------------------
    sheet1.title = "表紙"

    #セルの幅
    key_cols = "cell"
    for cols in jsn["cols"]:
        try:
            sheet1.column_dimensions[cols[key_cols]].width = cols["width"]
        except:
            print("cols_Error")    
    
    #セルの高さ
    key_rows = "cell_r"
    for rows in jsn["rows"]:
        try:
            sheet1.row_dimensions[rows[key_rows]].height = rows["height"]
        except:
            print("rows_Error")

    #セルの結合設定
    key_merg = "merge"
    for merges in jsn["merges"]:
        try:
            sheet1.merge_cells(merges[key_merg])
        except:
            print("merge_Error")
    

    #----セル内の色設定-----------------------------------------------------------------------
    fill = PatternFill(patternType = 'solid', fgColor = 'd3d3d3')

    #サブタイトル
    for rows in sheet1["B13:B17"]:
       for cell in rows:
            cell.fill = fill

    #バージョン、作成日、作成者欄
    for rows in sheet1["B20:D20"]:
        for cell in rows:
            cell.fill = fill     



    #----枠線の種類と色設定-------------------------------------------------------------------
    side_b = Side(style='thick', color='000000')
    side_s = Side(style='thin', color='000000')



    #----枠線の場所設定-----------------------------------------------------------------------

    #タイトル
    for rows in sheet1["A4:E6"]:
       for cell in rows:
           cell.border = Border(top=side_b ,bottom=side_b ,left=side_b ,right=side_b)

    #サブタイトル
    for rows in sheet1["B13:D17"]:
        for cell in rows:
       
            if cell == sheet1["B13"] or cell == sheet1["C13"] or cell == sheet1["D13"]:
                cell.border = Border(top=side_b ,bottom=side_s ,left=side_b ,right=side_b)
    
            elif cell == sheet1["B17"] or cell == sheet1["C17"] or cell == sheet1["D17"]:
                cell.border = Border(top=side_s ,bottom=side_b ,left=side_b ,right=side_b)
        
            else :
                cell.border = Border(top=side_s ,bottom=side_s ,left=side_b ,right=side_b)

    #バージョン、作成日、作成者欄
    for rows in sheet1["B20:D21"]:
        for cell in rows:
       
            if cell == sheet1["B20"] or cell == sheet1["C20"] or cell == sheet1["D20"]:
                cell.border = Border(top=side_b ,bottom=side_s ,left=side_b ,right=side_b)
        
            else :
                cell.border = Border(top=side_s ,bottom=side_b ,left=side_b ,right=side_b)



    #日付
    sheet1["E2"] = datetime.date.today()


    #文字表示設定
    key = 'coordinate'
    for cell1 in jsn["cells_title"]:
        try:
            #出力文字設定
            sheet1[cell1[key]] = cell1['value']
        except:
            print("Title_Error")

        try:
            #フォント設定
            sheet1[cell1[key]].font = Font(name = cell1['font_name'], bold = cell1['bold'], size = cell1['size'])
        except:
            print("Title_Font_Error")

        try:
            #文字上下左右そろえ
            sheet1[cell1[key]].alignment = Alignment(horizontal = cell1["horizontal"], vertical = cell1["vertical"])
        except:
            print("Title_alignment_Error")
    

if __name__ == "__main__":

    #ブック作成
    book = excel.Workbook()
    sheet = book.active
 
    #メイン処理
    main(sheet)

    #保存
    book.save(os.path.join('.','test.xlsx'))