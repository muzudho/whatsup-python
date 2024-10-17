import openpyxl as xl

from xltree.database import TreeTable
from xltree.workbooks import TreeDrawer, TreeEraser


class Config():
    """構成"""


    def __init__(
        self,
        dictionary=None):
        """初期化
        
        Parameters
        ----------
        dictionary : dict
            設定

            列の幅設定。width はだいたい 'ＭＳ Ｐゴシック' サイズ11 の半角英文字の個数
            * `no_width` - A列の幅。no列
            * `row_header_separator_width` - B列の幅。空列
            * `node_width` - 例：C, F, I ...列の幅。ノードの箱の幅
            * `parent_side_edge_width` - 例：D, G, J ...列の幅。エッジの水平線のうち、親ノードの方
            * `child_side_edge_width` - 例：E, H, K ...列の幅。エッジの水平線のうち、子ノードの方

            行の高さ設定。height の単位はポイント。既定値 8。昔のアメリカ人が椅子に座ってディスプレイを見たとき 1/72 インチに見える大きさが 1ポイント らしいが、そんなんワカラン。目視確認してほしい
            * `header_height` - 第１行。ヘッダー
            * `column_header_separator_height` - 第２行。空行
        """

        # 既定のディクショナリー
        self._dictionary = {
            # 列の幅
            'no_width':                         4,
            'row_header_separator_width':       3,
            'node_width':                       20,
            'parent_side_edge_width':           2,
            'child_side_edge_width':            4,

            # 行の高さ
            'header_height':                    13,
            'column_header_separator_height':   13,
        }

        # 上書き
        if dictionary is not None:
            for key, value in dictionary.items():
                self._dictionary[key] = value


    @property
    def dictionary(self):
        return self._dictionary


class Renderer():
    """描画"""


    def __init__(self, config=Config()):
        """初期化"""
        self._config = config


    def render(self, csv_file_path, wb_file_path, sheet_name):
        """描画"""

        # ワークブックを生成
        wb = xl.Workbook()

        # シートを作成
        wb.create_sheet(sheet_name)

        # 既存の Sheet シートを削除
        wb.remove(wb['Sheet'])

        # CSV読込
        tree_table = TreeTable.from_csv(file_path=csv_file_path)

        # ツリードロワーを用意、描画（都合上、要らない罫線が付いています）
        tree_drawer = TreeDrawer(tree_table=tree_table, ws=wb[sheet_name], config=self._config)
        tree_drawer.render()


        # 要らない罫線を消す
        # DEBUG_TIPS: このコードを不活性にして、必要な線は全部描かれていることを確認してください
        if True:
            tree_eraser = TreeEraser(tree_table=tree_table, ws=wb[sheet_name])
            tree_eraser.render()
        else:
            print(f"Eraser disabled")


        # ワークブックの保存
        wb.save(wb_file_path)
