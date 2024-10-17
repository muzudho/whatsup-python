import datetime


class TreeView():


    @staticmethod
    def is_same_between_ancestor_and_myself_as_avobe(curr_record, prev_record, depth_th):
        """ノード番号で指定したノードについて、前件と同じ祖先を持つか？"""

        # 前件が無い、または未設定なら偽
        if prev_record is None or prev_record.no is None:
            return False

        if curr_record.len_node_list < depth_th:
            raise ValueError(f"現件のノード数 {curr_record.len_node_list} が、 {depth_th=} に足りていません")

        if prev_record.len_node_list < depth_th:
            return False

        for cur_depth_th in range(0, depth_th + 1):
            if curr_record.node_at(depth_th=cur_depth_th).text != prev_record.node_at(depth_th=cur_depth_th).text:
                return False

        return True


    @staticmethod
    def can_connect_to_parent(curr_record, prev_record, depth_th):
        """前ラウンドのノードに接続できるか？"""

        predepth_th = depth_th - 1

        # 先頭行は、ラウンド０も含め、全部親ノードに接続できる
        if curr_record.no == 1:
            return True

        # 先頭行以外の第１ラウンドは、親ノードに接続できない
        elif depth_th == 1:
            return False

        try:
            # 前ラウンドは、前行とノードテキストが異なるか？
            return curr_record.node_at(depth_th=predepth_th).text != prev_record.node_at(depth_th=predepth_th).text

        except AttributeError as e:
            raise AttributeError(f"{depth_th=}  {predepth_th=}  {prev_record.node_at(depth_th=predepth_th)=}  {curr_record.node_at(depth_th=predepth_th)=}") from e

        # IndexError: depth_th=9  predepth_th=7  prev_record.len_node_list=6  curr_record.len_node_list=6
        except IndexError as e:
            raise IndexError(f"{depth_th=}  {predepth_th=}  {prev_record.len_node_list=}  {curr_record.len_node_list=}") from e


    @staticmethod
    def prev_row_is_elder_sibling(curr_record, prev_record, depth_th):
        """前行は兄か？"""

        # 先頭行に兄は無い
        if curr_record.no == 1:
            return False

        # 第0層は根なので、兄弟はいないものとみなす
        if depth_th == 0:
            return False

        predepth_th = depth_th - 1

        # 前ラウンドは、現業と前行で、テキストが等しいか？
        return curr_record.node_at(depth_th=predepth_th).text == prev_record.node_at(depth_th=predepth_th).text


    @staticmethod
    def next_row_is_younger_sibling(curr_record, next_record, depth_th):
        """次行は（自分または）弟か？

        TODO 下方に弟ノードがあるかどうかは、数行読み進めないと分からない
        TODO 自分がラスト・シブリングかどうかの情報がほしい。プリフェッチするか？
        """

        # 次行が無ければ弟は無い
        if next_record.no is None:
            return False

        # 第0層は根なので、兄弟はいないものとみなす
        if depth_th == 0:
            return False

        predepth_th = depth_th - 1

        # 前層は、現業と次行で、ノードテキストが等しいか？
        return curr_record.node_at(depth_th=predepth_th).text == next_record.node_at(depth_th=predepth_th).text


    @staticmethod
    def get_kind_connect_to_child(prev_record, curr_record, next_record, depth_th):
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
        if TreeView.prev_row_is_elder_sibling(curr_record=curr_record, prev_record=prev_record, depth_th=depth_th):

            # 次行は（自分または）弟か？
            if TreeView.next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, depth_th=depth_th):
                return 'TLetter'

            else:
                return 'Up'

        # 次行は（自分または）弟か？
        elif TreeView.next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, depth_th=depth_th):
            return 'Down'


        predepth_th = depth_th - 1
        if predepth_th < 0:
            raise ValueError(f"depth_th は負数であってはいけません {predepth_th=}")


        node = curr_record.node_at(depth_th=depth_th)
        prenode = curr_record.node_at(depth_th=predepth_th)
        print(f"""[{datetime.datetime.now()}] 水平線 第{depth_th}層：{node.text=}  第{predepth_th}層：{prenode.text=}""")
#         print(f"""\
# predepth_thde:
# {predepth_thde.stringify_dump('')}

# curr_record:
# {curr_record.stringify_dump('')}

# next_record:
# {next_record.stringify_dump('')}
# """)

        return 'Horizontal'
