#
# cd using_openpyxl
# python let_s_make_tree_view_on_excel.py
#
# エクセルで樹形図を描こう
#

import traceback
import datetime
import openpyxl as xl

from xl_tree.database import TreeRecord, TreeTable
from xl_tree.workbooks import TreeDrawer, TreeEraser

CSV_FILE_PATH = '../data/tree_shiritori.csv'
WB_FILE_PATH = '../temp/tree.xlsx'
SHEET_NAME = 'Tree'


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        # ワークブックを生成
        wb = xl.Workbook()

        # シートを作成
        wb.create_sheet(SHEET_NAME)

        # 既存の Sheet シートを削除
        wb.remove(wb['Sheet'])

        # CSV読込
        tree_table = TreeTable.from_csv(file_path=CSV_FILE_PATH)

        # ツリードロワーを用意、描画（都合上、要らない罫線が付いています）
        tree_drawer = TreeDrawer(tree_table=tree_table, ws=wb[SHEET_NAME])
        tree_drawer.render()


        # 要らない罫線を消す
        # DEBUG_TIPS: このコードを不活性にして、必要な線は全部描かれていることを確認してください
        if True:
            tree_eraser = TreeEraser(ws=wb[SHEET_NAME])
            tree_eraser.render()
        else:
            print(f"消しゴム　使用中止中")


        # ワークブックの保存
        wb.save(WB_FILE_PATH)


    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
