#
# python let_s_replace_the_names_of_all_files_in_the_directory_at_once.py
#
# ディレクトリ内のすべてのファイル名を一度に置換しよう！
#
import traceback
import enum
import os
import glob
import re


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    try:

        while True:
            # ディレクトリーを指定させます
            prompt = """\

| 入力例：
| Example: C:\\Users\\Muzudho\\Documents\\Hello
|
| ディレクトリーを指定してください。
| Specify the directory path.
> """
            path = input(prompt)

            os.chdir(path)

            # いったん、指定されたディレクトリーの中にあるファイル名を一覧します
            #
            #   NOTE ディレクトリーがここで合っているかの確認です
            #
            print(f"""Current directory: {os.getcwd()}

Files
-----""")

            files = glob.glob("./*")

            # とりあえず一覧します
            for file in files:
                # `file` - Example: `.\20210815shogi67.png`
                basename = os.path.basename(file)
                print(basename)

            prompt = """\

| このディレクトリーで間違いないですか？
| Are you sure this is the right directory?
(y/n)> """
            answer = input(prompt)

            if answer == "y":
                break

            print("Canceld")

        # 正規表現のパターンを入力させます
        while True:
            prompt = r"""
| 入力例：
| Example: ^example-([\d\w]+)-([\d\w]+).txt$
|
| 正規表現パターンを入力してください。
| Please enter a regular expression pattern.
> """
            patternText = input(prompt)
            pattern = re.compile(patternText)

            # とりあえず一覧します
            for i, file in enumerate(files):
                basename = os.path.basename(file)
                result = pattern.match(basename)
                if result:
                    # Matched
                    # グループ数
                    groupCount = len(result.groups())
                    buf = f"({i+1}) {basename}"
                    for j in range(0, groupCount):
                        buf += f" \\{j+1}=[{result.group(j+1)}]"

                    print(buf)
                else:
                    # Unmatched
                    print(f"( ) {basename}")

            prompt = """\

| これで合っていますか？
| Is this what you want?
(y/n)> """
            answer = input(prompt)

            if answer == "y":
                break
            else:
                print("Canceld")

        # 置換のシミュレーション
        while True:
            prompt = r"""
| 入力例：
| Example: example-\2-\1.txt
|
| 置換後のパターンを入力してください。
| Enter the pattern after the conversion.
> """
            replacement = input(prompt)

            print("""
Simulation
----------""")
            for i, file in enumerate(files):
                basename = os.path.basename(file)
                result = pattern.match(basename)
                if result:
                    # Matched
                    converted = re.sub(patternText, replacement, basename)
                    print(f"({i+1}) {basename} --> {converted}")

            prompt = """\

| 実行しますか？
| Do you want to run it?
(y/n)> """
            answer = input(prompt)

            if answer == "y":
                break

            print("Canceld")

        # 置換実行
        for i, file in enumerate(files):
            basename = os.path.basename(file)
            result = pattern.match(basename)
            if result:
                # Matched
                converted = re.sub(patternText, replacement, basename)
                oldPath = os.path.join(os.getcwd(), basename)
                newPath = os.path.join(os.getcwd(), converted)
                print(f"({i})Rename {oldPath} --> {newPath}")
                os.rename(oldPath, newPath)


        print("でーきたっ！")


    except Exception as err:
        print(f"""\
おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
