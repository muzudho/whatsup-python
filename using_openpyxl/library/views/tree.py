import datetime


class TreeView():


    @staticmethod
    def can_connect_to_parent(curr_record, prev_record, node_th):
        """前ラウンドのノードに接続できるか？"""

        prenode_th = node_th - 1

        # 先頭行は、ラウンド０も含め、全部親ノードに接続できる
        if curr_record.no == 1:
            return True

        # 先頭行以外の第１ラウンドは、親ノードに接続できない
        elif node_th == 1:
            return False

        try:
            # 前ラウンドは、前行とノードテキストが異なるか？
            return curr_record.node_at(node_th=prenode_th).text != prev_record.node_at(node_th=prenode_th).text

        except AttributeError as e:
            raise AttributeError(f"{node_th=}  {prenode_th=}  {prev_record.node_at(node_th=prenode_th)=}  {curr_record.node_at(node_th=prenode_th)=}") from e

        # IndexError: node_th=9  prenode_th=7  prev_record.len_node_list=6  curr_record.len_node_list=6
        except IndexError as e:
            raise IndexError(f"{node_th=}  {prenode_th=}  {prev_record.len_node_list=}  {curr_record.len_node_list=}") from e


    @staticmethod
    def is_same_as_avobe(curr_record, prev_record, node_th):
        # 先頭行に兄は無い
        if curr_record.no == 1:
            return False

        # 現業と前行は、現ラウンドについて、テキストが等しい
        a = curr_record.node_at(node_th=node_th).text
        b = prev_record.node_at(node_th=node_th).text
        print(f"[{datetime.datetime.now()}] {curr_record.no}件目 {node_th=}  is_same_as_avobe  {a=}  {b=}")
        return a == b


    @staticmethod
    def prev_row_is_elder_sibling(curr_record, prev_record, node_th):
        """前行は兄か？"""

        # 先頭行に兄は無い
        if curr_record.no == 1:
            return False

        # 第0節は根なので、兄弟はいないものとみなす
        if node_th == 0:
            return False

        prenode_th = node_th - 1

        # 前ラウンドは、現業と前行で、テキストが等しいか？
        return curr_record.node_at(node_th=prenode_th).text == prev_record.node_at(node_th=prenode_th).text


    @staticmethod
    def next_row_is_younger_sibling(curr_record, next_record, node_th):
        """次行は（自分または）弟か？

        TODO 下方に弟ノードがあるかどうかは、数行読み進めないと分からない
        TODO 自分がラスト・シブリングかどうかの情報がほしい。プリフェッチするか？
        """

        # 次行が無ければ弟は無い
        if next_record.no is None:
            return False

        # 第0節は根なので、兄弟はいないものとみなす
        if node_th == 0:
            return False

        prenode_th = node_th - 1

        # 前節は、現業と次行で、ノードテキストが等しいか？
        return curr_record.node_at(node_th=prenode_th).text == next_record.node_at(node_th=prenode_th).text


    @staticmethod
    def get_kind_connect_to_child(prev_record, curr_record, next_record, node_th):
        """
        子ノードへの接続は４種類の線がある
        
        (1) Horizontal
          .    under_border
        ...__  
          .    None
        
        (2) Down
          .    under_border
        ..+__  
          |    leftside_border
        
        (3) TLetter
          |    l_letter_border
        ..+__  
          |    leftside_border
        
        (4) Up
          |    l_letter_border
        ..+__  
          .    None
        """

        # 前行は兄か？
        if TreeView.prev_row_is_elder_sibling(curr_record=curr_record, prev_record=prev_record, node_th=node_th):

            # 次行は（自分または）弟か？
            if TreeView.next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, node_th=node_th):
                return 'TLetter'

            else:
                return 'Up'

        # 次行は（自分または）弟か？
        elif TreeView.next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, node_th=node_th):
            return 'Down'


        prenode_th = node_th - 1
        if prenode_th < 0:
            raise ValueError(f"node_th は負数であってはいけません {prenode_th=}")


        node = curr_record.node_at(node_th=node_th)
        prenode = curr_record.node_at(node_th=prenode_th)
        print(f"""[{datetime.datetime.now()}] 水平線 {node_th}節：{node.text=}  {prenode_th}節：{prenode.text=}""")
#         print(f"""\
# prenode_thde:
# {prenode_thde.stringify_dump('')}

# curr_record:
# {curr_record.stringify_dump('')}

# next_record:
# {next_record.stringify_dump('')}
# """)

        return 'Horizontal'
