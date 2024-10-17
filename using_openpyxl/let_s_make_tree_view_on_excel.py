#
# cd using_openpyxl
# python let_s_make_tree_view_on_excel.py
#
# エクセルで樹形図を描こう
#

import traceback
import datetime
import pandas as pd
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side

from xl_tree.database import TreeNode, TreeRecord, TreeTable
from xl_tree.models import TreeModel
from xl_tree.workbooks import TreeDrawer, TreeEraser

CSV_FILE_PATH = '../data/tree_shiritori.csv'
WB_FILE_PATH = '../temp/tree.xlsx'


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        # ワークブックを生成
        wb = xl.Workbook()

        # Tree シートを作成
        wb.create_sheet('Tree')

        # 既存の Sheet シートを削除
        wb.remove(wb['Sheet'])

        # CSV読込
        tree_table = TreeTable.from_csv(file_path=CSV_FILE_PATH)

#         # CSV確認
#         print(f"""\
# tree_table.df:
# {tree_table.df}""")

        tree_drawer = TreeDrawer(df=tree_table.df, wb=wb)

        # GTWB の Sheet シートへのヘッダー書出し
        tree_drawer.on_header()

        # GTWB の Sheet シートへの各行書出し
        tree_table.for_each(on_each=tree_drawer.on_each_record)

        # 最終行の実行
        tree_drawer.on_each_record(next_row_number=len(tree_table.df), next_record=TreeRecord.new_empty())


        # 要らない罫線を消す
        # DEBUG_TIPS: このコードを不活性にして、必要な線は全部描かれていることを確認してください
        if True:
            tree_eraser = TreeEraser(wb=wb)
            tree_eraser.execute()
        else:
            print(f"消しゴム　使用中止中")


        # ワークブックの保存
        wb.save(WB_FILE_PATH)


    except Exception as err:
        print(f"[unexpected error] {err=}  {type(err)=}")

        # スタックトレース表示
        print(traceback.format_exc())
