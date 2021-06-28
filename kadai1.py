import openpyxl as excel
import datetime
import os.path
from openpyxl.styles.borders import Border,Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill



def main():
    #----ブック作成----------------------------------------------------------------------
    book = excel.Workbook()
    sheet1 = book.active
    

    #----シート名設定--------------------------------------------------------------------
    sheet1.title = "表紙"

    #----セルの幅------------------------------------------------------------------------
    sheet1.column_dimensions['A'].width = 21.38
    sheet1.column_dimensions['B'].width = 21.38
    sheet1.column_dimensions['C'].width = 21.38
    sheet1.column_dimensions['D'].width = 21.38
    sheet1.column_dimensions['E'].width = 21.38



    #----セルの高さ----------------------------------------------------------------------
    sheet1.row_dimensions[13].height = 19.50
    sheet1.row_dimensions[14].height = 19.50
    sheet1.row_dimensions[15].height = 19.50
    sheet1.row_dimensions[16].height = 19.50
    sheet1.row_dimensions[17].height = 19.50

    sheet1.row_dimensions[20].height = 19.50
    sheet1.row_dimensions[21].height = 19.50



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


        
    #----セルの結合設定-------------------------------------------------------------------------
    sheet1.merge_cells("A4:E6")
    sheet1.merge_cells("C13:D13")
    sheet1.merge_cells("C14:D14")
    sheet1.merge_cells("C15:D15")
    sheet1.merge_cells("C16:D16")
    sheet1.merge_cells("C17:D17")



    #----文字列そろえ設定-----------------------------------------------------------------------
    for rows in sheet1:
        for cell in rows:
        
            if cell == sheet1["A4"] :
                #上下左右センターそろえ
                cell.alignment = Alignment(horizontal = "center" ,vertical = "center")
        
            elif cell == sheet1["B20"] or cell == sheet1["C20"] or cell == sheet1["D20"]:
                #上下左右センターそろえ
                cell.alignment = Alignment(horizontal = "center" ,vertical = "center")
        
            else :
                #上下センターそろえ
                cell.alignment = Alignment(vertical = "center")



    #----フォント変更設定-----------------------------------------------------------------------
    fontname = "ＭＳ Ｐゴシック"

    #タイトル
    sheet1["A4"].font = Font(name = fontname ,bold = "true" ,size = 36 )

    #サブタイトル
    sheet1["B13"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["B14"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["B15"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["B16"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["B17"].font = Font(name = fontname ,bold = "true" ,size = 11 )

    #バージョン、作成日、作成者欄
    sheet1["B20"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["C20"].font = Font(name = fontname ,bold = "true" ,size = 11 )
    sheet1["D20"].font = Font(name = fontname ,bold = "true" ,size = 11 )



    #----出力文字列設定-------------------------------------------------------------------------

    #タイトル
    sheet1["A4"] = "設計書"

    #サブタイトル
    sheet1["B13"] = "モジュールID"
    sheet1["B14"] = "サブメニュー名"
    sheet1["B15"] = "コマンドID"
    sheet1["B16"] = "コマンド名"
    sheet1["B17"] = "プログラムID"

    #バージョン、作成日、作成者欄
    sheet1["B20"] = "バージョン"
    sheet1["C20"] = "作成日"
    sheet1["D20"] = "作成者"

    #日付
    sheet1["E2"] = datetime.date.today()


    #----保存------------------------------------------------------------------------------
    book.save('test.xlsx')

if __name__ == "__main__":

    
    main()