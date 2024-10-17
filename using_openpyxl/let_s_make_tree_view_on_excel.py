#
# cd using_openpyxl
# python let_s_make_tree_view_on_excel.py
#
# エクセルで樹形図を描こう
#

import traceback
import datetime

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
        config = Config(
                # 省略可能
                dictionary = {
                    # 列の幅
                    'no_width':                         4,      # A列の幅。no列
                    'row_header_separator_width':       3,      # B列の幅。空列
                    'node_width':                       20,     # 例：C, F, I ...列の幅。ノードの箱の幅
                    'parent_side_edge_width':           2,      # 例：D, G, J ...列の幅。エッジの水平線のうち、親ノードの方
                    'child_side_edge_width':            4,      # 例：E, H, K ...列の幅。エッジの水平線のうち、子ノードの方

                    # 行の高さ
                    'header_height':                    13,     # 第１行。ヘッダー
                    'column_header_separator_height':   13,     # 第２行。空行
                })

        # レンダラー生成
        renderer = Renderer(config=config)
        renderer.render(
                csv_file_path=csv_file_path,
                wb_file_path=wb_file_path,
                sheet_name=SHEET_NAME)

        print(f"[{datetime.datetime.now()}] {wb_file_path} ファイルを確認してください")


    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
