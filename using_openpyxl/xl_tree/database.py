import pandas as pd
from xl_tree import INDENT


############
# MARK: Node
############
class TreeNode():
    """ノード（節）
    ノードの形に合わせて改造してください"""

    def __init__(self, edge_text, text):
        """初期化
        
        Parameters
        ----------
        edge_text : str
            辺のテキスト
        text : str
            節のテキスト
        """
        self._edge_text = edge_text
        self._text = text


    @property
    def edge_text(self):
        return self._edge_text


    @property
    def text(self):
        return self._text


    def stringify_dump(self, indent):
        succ_indent = indent + INDENT
        return f"""\
{indent}TreeNode
{indent}--------
{succ_indent}{self._edge_text=}
{succ_indent}{self._text=}
"""


##############
# MARK: Record
##############
class TreeRecord():


    def __init__(self, no, node_list):
        """初期化
        
        Parameters
        ----------
        no : int
            1から始まる連番。数詞は葉
        node_list : list<TreeNode>
            固定長ノード０～４。
            第０層は根。
            TODO 第５層以降欲しい場合は改造してください
        """
        self._no = no
        self._node_list = node_list


    @staticmethod
    def new_empty():
        return TreeRecord(
                no=None,
                node_list=[None, None, None, None, None])


    @property
    def no(self):
        return self._no


    @property
    def len_node_list(self):
        return len(self._node_list)


    def node_at(self, depth_th):
        """
        Parameters
        ----------
        round_th : int
            th は forth や fifth の th。
            例：根なら０を指定してください。
            例：第１層なら 1 を指定してください
        """

        # NOTE -1 を指定すると最後尾の要素になるが、固定長配列の最後尾の要素が、思っているような最後尾の要素とは限らない。うまくいかない
        if depth_th < 0:
            raise ValueError(f'depth_th に負数を設定しないでください。意図した動作はしません {depth_th=}')

        return self._node_list[depth_th]


    def update(self, no=None, node_list=None):
        """no inplace
        何も更新しなければシャロー・コピーを返します"""

        def new_or_default(new, default):
            if new is None:
                return default
            return new

        return TreeRecord(
                no=new_or_default(no, self._no),
                node_list=new_or_default(node_list, self._node_list))


    def stringify_dump(self, indent):
        succ_indent = indent + INDENT

        blocks = []
        for node in self._node_list:
            blocks.append(node.stringify_dump(succ_indent))

        return f"""\
{indent}TreeRecord
{indent}----------
{succ_indent}{self._no=}
{'\n'.join(blocks)}
"""


    def get_th_of_leaf_node(self):
        """葉要素の層番号を取得。
        th は forth や fifth の th。
        葉要素は、次の層がない要素"""

        for depth_th in range(0, len(self._node_list)):
            nd = self._node_list[depth_th]
            if nd is None or nd.text is None:
                return depth_th

        return len(self._node_list)


##############
# MARK: Record
##############
class TreeTable():
    """樹形図データのテーブル"""


    _dtype = {
        # no はインデックス

        'node0':'object',   # string 型は無いので object 型にする
        'edge1':'object',
        'node1':'object',
        'edge2':'object',
        'node2':'object',
        'edge3':'object',
        'node3':'object',
        'edge4':'object',
        'node4':'object'}


    def __init__(self, df):
        self._df = df


    @classmethod
    def new_empty_table(clazz):
        df = pd.DataFrame(
                columns=[
                    # 'no' は後でインデックスに変換
                    'no',

                    'node0',
                    'edge1',
                    'node1',
                    'edge2',
                    'node2',
                    'edge3',
                    'node3',
                    'edge4',
                    'node4'])
        clazz.setup_data_frame(df=df, shall_set_index=True)
        return TreeTable(df=df)


    @classmethod
    def from_csv(clazz, file_path):
        """ファイル読込

        Parameters
        ----------
        file_path : str
            CSVファイルパス
        
        Returns
        -------
        table : TreeTable
            テーブル、またはナン
        file_read_result : FileReadResult
            ファイル読込結果
        """
        df = pd.read_csv(file_path, encoding="utf8", index_col=['no'])

        # テーブルに追加の設定
        clazz.setup_data_frame(df=df, shall_set_index=False)

        return TreeTable(df=df)


    @property
    def df(self):
        return self._df


    @classmethod
    def setup_data_frame(clazz, df, shall_set_index):
        """データフレームの設定"""

        if shall_set_index:
            # インデックスの設定
            df.set_index('no',
                    inplace=True)   # NOTE インデックスを指定したデータフレームを戻り値として返すのではなく、このインスタンス自身を更新します

        # データ型の設定
        df.astype(clazz._dtype)


    def upsert_record(self, welcome_record):
        """該当レコードが無ければ新規作成、あれば更新

        Parameters
        ----------
        welcome_record : GameTreeRecord
            レコード

        Returns
        -------
        shall_record_change : bool
            レコードの新規追加、または更新があれば真。変更が無ければ偽
        """

        # インデックス
        # -----------
        # index : any
        #   インデックス。整数なら numpy.int64 だったり、複数インデックスなら tuple だったり、型は変わる。
        #   <class 'numpy.int64'> は int型ではないが、pandas では int型と同じように使えるようだ
        index = welcome_record.no

        # データ変更判定
        # -------------
        is_new_index = index not in self._df.index

        # インデックスが既存でないなら
        if is_new_index:
            shall_record_change = True

        else:
            # 更新の有無判定
            # no はインデックス
            shall_record_change =\
                self._df['node0'][index] != welcome_record.node_at(0).text or\
                \
                self._df['edge1'][index] != welcome_record.node_at(1).edge_text or\
                self._df['node1'][index] != welcome_record.node_at(1).node or\
                \
                self._df['edge2'][index] != welcome_record.node_at(2).edge_text or\
                self._df['node2'][index] != welcome_record.node_at(2).node or\
                \
                self._df['edge3'][index] != welcome_record.node_at(3).edge_text or\
                self._df['node3'][index] != welcome_record.node_at(3).node or\
                \
                self._df['edge4'][index] != welcome_record.node_at(4).edge_text or\
                self._df['node4'][index] != welcome_record.node_at(4).node


        # 行の挿入または更新
        if shall_record_change:
            self._df.loc[index] = {
                # no はインデックス
                'node0': welcome_record.node_at(0).text,

                'edge1': welcome_record.node_at(1).edge_text,
                'node1': welcome_record.node_at(1).text,

                'edge2': welcome_record.node_at(2).edge_text,
                'node2': welcome_record.node_at(2).text,

                'edge3': welcome_record.node_at(3).edge_text,
                'node3': welcome_record.node_at(3).text,

                'edge4': welcome_record.node_at(4).edge_text,
                'node4': welcome_record.node_at(4).text}

        if is_new_index:
            # NOTE ソートをしておかないと、インデックスのパフォーマンスが機能しない
            self._df.sort_index(
                    inplace=True)   # NOTE ソートを指定したデータフレームを戻り値として返すのではなく、このインスタンス自身をソートします


        return shall_record_change


    def to_csv(self, file_path):
        """ファイル書き出し
        
        Parameters
        ----------
        file_path : str
            CSVファイルパス
        """

        self._df.to_csv(
                csv_file_path,
                # no はインデックス
                columns=[
                    'node0',
                    'edge1', 'node1',
                    'edge2', 'node2',
                    'edge3', 'node3',
                    'edge4', 'node4'])


    def for_each(self, on_each):
        """
        Parameters
        ----------
        on_each : func
            TreeRecord 引数を受け取る関数
        """

        df = self._df

        for row_number,(
                node0,
                edge1, node1,
                edge2, node2,
                edge3, node3,
                edge4, node4) in\
                enumerate(zip(
                    df['node0'],
                    df['edge1'], df['node1'],
                    df['edge2'], df['node2'],
                    df['edge3'], df['node3'],
                    df['edge4'], df['node4'])):

            # no はインデックス
            no = df.index[row_number]

            # レコード作成
            record = TreeRecord(
                    no=no,
                    # TODO 今のところ固定長サイズ
                    node_list=[
                        TreeNode(edge_text=None, text=node0),
                        TreeNode(edge_text=edge1, text=node1),
                        TreeNode(edge_text=edge2, text=node2),
                        TreeNode(edge_text=edge3, text=node3),
                        TreeNode(edge_text=edge4, text=node4)])

            on_each(row_number, record)
