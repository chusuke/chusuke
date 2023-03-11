import tkinter as tk


class Application(tk.Frame):
    # 関数に紐づく、イベントの詳細情報を取得する関数
    # event : イベント詳細情報
    def getEventDetail(self, event):
        print(event.keysym)

    # 追加bind用の関数。実行される場合に、hello worldを出力する。
    def getAddFunc(self, event):
        print('hello world')

    def __init__(self, master=None):
        # Windowの初期設定を行う。
        super().__init__(master)

        # Windowの画面サイズを設定する。
        # geometryについて : https://kuroro.blog/python/rozH3S2CYE0a0nB3s2QL/
        self.master.geometry("300x200")

        # Windowを親要素として、label Widgetを作成する。
        # text : テキスト情報
        # bg : 背景色を設定
        # 色について : https://kuroro.blog/python/YcZ6Yh4PswqUzaQXwnG2/
        # Labelについて : https://kuroro.blog/python/Pj4Z7JBNRvcHZvtFqiKD/
        label = tk.Label(self.master, text="test", bg='blue')

        # イベントと関数の実行を紐づける。
        # 第一引数 : イベント内容。label Widgetがクリックされた場合
        # 第二引数 : 第一引数が実行された場合に、呼び出す関数。self.getEventDetailとする。
        # 戻り値 : bind関数に紐づく独自ID
        label.bind("<ButtonPress>", self.getEventDetail)
        # イベントと関数の実行を紐づける。
        # 第一引数 : イベント内容。label Widgetがクリックされた場合
        # 第二引数 : 第一引数が実行された場合に、呼び出す関数。self.getAddFuncとする。
        # 第三引数 : ''か'+'を指定する。(デフォルト : '')
        # ''の場合、一つ前のbind関数で宣言された関数を実行しないで、今回bindする関数を実行する。今回の場合、self.getEventDetailを実行しないで、self.getAddFuncを実行する。
        # '+'の場合、一つ前のbind関数で宣言された関数を実行して、今回bindする関数を実行する。今回の場合、self.getEventDetailを実行して、self.getAddFuncを実行する。
        # 戻り値 : bind関数に紐づく独自ID
        label.bind("<ButtonPress>", self.getAddFunc, '+')

        # Windowを親要素とした場合に、label Widgetをどのように配置するのか?
        # packについて : https://kuroro.blog/python/UuvLfIBIEaw98BzBZ3FJ/
        label.pack()


# Tkinter初学者参考 : https://docs.python.org/ja/3/library/tkinter.html#a-simple-hello-world-program
if __name__ == "__main__":
    # Windowを生成する。
    # Windowについて : https://kuroro.blog/python/116yLvTkzH2AUJj8FHLx/
    root = tk.Tk()
    app = Application(master=root)

    # Windowをループさせて、継続的にWindow表示させる。
    # mainloopについて : https://kuroro.blog/python/DmJdUb50oAhmBteRa4fi/
    app.mainloop()