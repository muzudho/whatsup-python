#
# cd using_openpyxl
# python let_s_make_tree_view_on_excel.py
#
# エクセルで樹形図を描こう
#

import traceback

from xltree import Config, Renderer

CSV_FILE_PATH = '../data/tree_shiritori.csv'
WB_FILE_PATH = '../temp/tree.xlsx'
SHEET_NAME = 'Tree'


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        # 構成
        config = Config()

        # レンダラー生成
        renderer = Renderer(config=config)
        renderer.render(
                csv_file_path=CSV_FILE_PATH,
                wb_file_path=WB_FILE_PATH,
                sheet_name=SHEET_NAME)

    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
