#
# cd using_openpyxl
# python let_s_make_tree_view_on_excel.py
#
# エクセルで樹形図を描こう
#

import traceback

from xltree import Config, Renderer

SHEET_NAME = 'Tree'


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        csv_file_path = input(f"""\

CSVファイルへのパスを入力してください
Enter the path to the CSV file

    Example: ../data/tree_shiritori.csv

> """)

        wb_file_path = input(f"""\

エクセルのワークブック・ファイルへのパスを入力してください
Enter the path to the Excel workbook(.xlsx) file

    Example: ../temp/tree.xlsx

> """)

        # 構成
        config = Config()

        # レンダラー生成
        renderer = Renderer(config=config)
        renderer.render(
                csv_file_path=csv_file_path,
                wb_file_path=wb_file_path,
                sheet_name=SHEET_NAME)

    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
