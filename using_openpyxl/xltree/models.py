class TreeModel():


    @staticmethod
    def is_same_path_as_avobe(curr_record, prev_record, depth_th):
        """自件と前件を比較して、根から自ノードまで、ノードテキストが等しいか？"""

        # 前件が無い、または未設定なら偽
        if prev_record is None or prev_record.no is None:
            return False

        if curr_record.len_node_list < depth_th:
            raise ValueError(f"引数の指定ミスです。現件のノード数 {curr_record.len_node_list} が、 {depth_th=} に足りていません")

        if prev_record.len_node_list < depth_th:
            return False

        for cur_depth_th in range(0, depth_th + 1):
            # 同じかどうかはノードテキストで決める
            if curr_record.node_at(depth_th=cur_depth_th).text != prev_record.node_at(depth_th=cur_depth_th).text:
                return False

        return True


    @staticmethod
    def get_kind_of_edge(prev_record, curr_record, next_record, depth_th):
        """
        子ノードへの接続は４種類の線がある
        
        (1) ─字
        (2) ┬字
        (3) ├字
        (4) └字
        """

        # 前行は兄か？
        if TreeModel._prev_row_is_elder_sibling(curr_record=curr_record, prev_record=prev_record, depth_th=depth_th):

            # 次行は（自分または）弟か？ 自分が複数行に跨っていることはある
            if TreeModel._next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, depth_th=depth_th):
                return '├字'

            else:
                return '└字'

        # 次行は（自分または）弟か？ 自分が複数行に跨っていることはある
        elif TreeModel._next_row_is_younger_sibling(curr_record=curr_record, next_record=next_record, depth_th=depth_th):
            return '┬字'


        return '─字'


    @staticmethod
    def _prev_row_is_elder_sibling(curr_record, prev_record, depth_th):
        """前件は兄か？"""

        # 先頭行に兄は無い
        if curr_record.no == 1:
            return False

        # 第0層は根なので、兄弟はいないものとみなす
        if depth_th == 0:
            return False

        predepth_th = depth_th - 1

        # 自件と前件を比較して、根から自ノードまで、ノードテキストが等しいか？
        return TreeModel.is_same_path_as_avobe(
                curr_record=curr_record,
                prev_record=prev_record,
                depth_th=predepth_th)


    @staticmethod
    def _next_row_is_younger_sibling(curr_record, next_record, depth_th):
        """次件は（自分または）弟か？"""

        # 次行が無ければ弟は無い
        if next_record.no is None:
            return False

        # 第0層は根なので、兄弟はいないものとみなす
        if depth_th == 0:
            return False

        predepth_th = depth_th - 1

        # 自件と前件を比較して、根から自ノードまで、ノードテキストが等しいか？
        return TreeModel.is_same_path_as_avobe(
                curr_record=next_record,
                prev_record=curr_record,
                depth_th=predepth_th)
