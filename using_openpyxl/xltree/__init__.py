import openpyxl as xl

from xltree.database import TreeTable
from xltree.workbooks import TreeDrawer, TreeEraser


class XlTree():


    def __init__(self):
        pass


    def generate(self, csv_file_path, wb_file_path, sheet_name):
        # ワークブックを生成
        wb = xl.Workbook()

        # シートを作成
        wb.create_sheet(sheet_name)

        # 既存の Sheet シートを削除
        wb.remove(wb['Sheet'])

        # CSV読込
        tree_table = TreeTable.from_csv(file_path=csv_file_path)

        # ツリードロワーを用意、描画（都合上、要らない罫線が付いています）
        tree_drawer = TreeDrawer(tree_table=tree_table, ws=wb[sheet_name])
        tree_drawer.render()


        # 要らない罫線を消す
        # DEBUG_TIPS: このコードを不活性にして、必要な線は全部描かれていることを確認してください
        if True:
            tree_eraser = TreeEraser(ws=wb[sheet_name])
            tree_eraser.render()
        else:
            print(f"消しゴム　使用中止中")


        # ワークブックの保存
        wb.save(wb_file_path)
