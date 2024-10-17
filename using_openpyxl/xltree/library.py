# 循環参照を防ぐために、何もインポートしません

# ダンプで使う
INDENT = '    '


def nth(n):
    """序数の接尾辞
    📖 [Gareth on codegolf](https://codegolf.stackexchange.com/questions/4707/outputting-ordinal-numbers-1st-2nd-3rd#answer-4712)
    📖 [Ordinal numbers replacement](https://stackoverflow.com/questions/9647202/ordinal-numbers-replacement)
    """
    return "%d%s" % (n,"tsnrhtdd"[(n//10%10!=1)*(n%10<4)*n%10::4])
