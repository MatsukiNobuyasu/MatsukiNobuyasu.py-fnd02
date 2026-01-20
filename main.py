import tkinter as tk
from tkinter import ttk
root = tk.Tk()
root.title("SQLサーバー接続アプリ")
root.geometry('1200x800')

# rootのグリッド設定
root.columnconfigure(0, weight=0,minsize=300)  # 左カラム（テーブル）
root.columnconfigure(1, weight=5)  # 右カラム（抽出データ）
root.rowconfigure(0, weight=0)     # 上段(Server設定) 高さ固定
root.rowconfigure(1, weight=1,minsize=300)     # 下段（テーブル・抽出データ）高さ伸縮

# frame1: Server設定（上段、列2つ分に跨る）
frame1 = tk.LabelFrame(root, text='Server設定', bd=2, relief=tk.GROOVE)
frame1.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky=tk.EW)

# frame1の列幅設定
for i in range(9):
    frame1.columnconfigure(i, weight=1)

# frame1のウィジェット配置
Label1 = tk.Label(frame1, text='サーバー')
Label1_Entry = tk.Entry(frame1)
Label1.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
Label1_Entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
Label2 = tk.Label(frame1, text='データベース')
Label2.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
Label2_Entry = tk.Entry(frame1)
Label2_Entry.grid(row=0, column=3, padx=5, pady=5, sticky=tk.EW)
Label3 = tk.Label(frame1, text='ユーザーID')
Label3.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
Label3_Entry = tk.Entry(frame1)
Label3_Entry.grid(row=0, column=5, padx=5, pady=5, sticky=tk.EW)
Label4 = tk.Label(frame1, text='pass')
Label4.grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
Label4_Entry = tk.Entry(frame1, show="*")
Label4_Entry.grid(row=0, column=7, padx=5, pady=5, sticky=tk.EW)

from AR_traceability import *
Button1_1 = tk.Button(frame1, text='接続', command=lambda : connect_server(Label1_Entry.get(), Label2_Entry.get(), Label3_Entry.get(), Label4_Entry.get(),table_list))
Button1_1.grid(row=0, column=8, padx=5, pady=5, sticky=tk.NSEW)



# frame2: テーブル（左下）
frame2 = tk.LabelFrame(root, text='テーブル', bd=2, relief=tk.GROOVE)
frame2.grid(row=1, column=0, padx=10, pady=5, sticky=tk.NSEW)
# frame2のグリッド設定
frame2.columnconfigure(0, weight=1)
frame2.rowconfigure(0, weight=1)
# リストボックスの設定
table_list = tk.Listbox(frame2, selectmode=tk.SINGLE)
table_list.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
# リストボックスの選択時に選択テーブルの列情報を取得
selected_tablename_global=None
def list_selected(event):
    # イベント発生元（ウィジェット）を取得
    select_index = table_list.curselection()  # タプルで返る

    #リストから選択が離れたときにエラーが出るので回避処理
    if not select_index:        
        return

    index = select_index[0]  # タプルの最初の要素（選択インデックス）
    slect_table = table_list.get(index)
    global selected_tablename_global
    selected_tablename_global = slect_table
    
    load_table_columns(Label1_Entry.get(), Label2_Entry.get(), Label3_Entry.get(), Label4_Entry.get(),slect_table, filter_combobox)
table_list.bind('<<ListboxSelect>>', list_selected)



# frame3: 抽出データ（右下）
frame3 = tk.LabelFrame(root, text='抽出データ', bd=2, relief=tk.GROOVE)
frame3.grid(row=1, column=1, padx=10, pady=5, sticky=tk.NSEW)
frame3.columnconfigure(0, weight=1)
frame3.rowconfigure(0, weight=1)

#frame3にフィルターや抽出ボタンなどの配置用のサブフレームを設定
frame3_subframe = tk.LabelFrame(frame3, text='フィルター条件', bd=2, relief=tk.GROOVE)
frame3_subframe.grid(row=2, column=0, padx=10, pady=5, sticky=tk.NSEW)

# フィルター用Comboboxの作成
filter_combobox = ttk.Combobox(frame3_subframe,values=filter_columns_global,state="readonly")
filter_combobox.grid(row=2, column=2, padx=5, pady=5, sticky=tk.NSEW)
# Comboboxの選択時の処理
filter_selected_column =None #抽出ボタン押した時にSQLコードに使用
def combobox_selected(event):
    # イベント発生元（ウィジェット）を取得
    widget = event.widget
    # 選択された値を取得
    value = widget.get()
    print(f"Selected: {value}")
    global filter_selected_column
    filter_selected_column = value
    print(filter_selected_column)
    
filter_combobox.bind('<<ComboboxSelected>>',combobox_selected)
#抽出ボタン設定
btn_load_tabledata = tk.Button(
    frame3_subframe, text='抽出', 
    command=lambda : load_tabledata(Label1_Entry.get(), Label2_Entry.get(), Label3_Entry.get(), Label4_Entry.get(),selected_tablename_global,filter_selected_column,tree)
    )
btn_load_tabledata.grid(row=2, column=10, padx=5, pady=5, sticky=tk.NSEW)

#保存ボタン設定
import datetime
def save_csv():
    # 現在の日時を取得
    now = datetime.datetime.now()    
    filename = f"{selected_tablename_global}_{now.strftime("%Y%m%d_%H%M%S")}.csv"
    with open(filename,"w",encoding="utf-8") as f:
        columns = tree["columns"]
        writer = csv.writer(f)
        #ヘッダーの書込み
        writer.writerow(columns)
        
        #すべてのアイテムを取得
        for item_id in tree.get_children():
            #各カラムの値を取得
            row = [tree.set(item_id,col) for col in columns]
            writer.writerow(row)
    print(f"{filename}を保存しました。")
        
btn_save_tabledata = tk.Button(frame3_subframe, text='保存', command=save_csv)
btn_save_tabledata.grid(row=2, column=20, padx=5, pady=5, sticky=tk.E)
frame3_subframe.columnconfigure(20, weight=1)
# frame3のグリッド設定
frame3.columnconfigure(0, weight=1)
frame3.rowconfigure(0, weight=1)
frame3.rowconfigure(1, weight=0)  # スクロールバー用行


# Treeview作成(テーブルデータを表示する表)
tree = ttk.Treeview(frame3, columns=columns_global, show='headings')
tree.grid(row=0, column=0, sticky='nsew')

# 縦スクロールバー
vsb = tk.Scrollbar(frame3, orient='vertical', command=tree.yview)
vsb.grid(row=0, column=1, sticky='ns')
tree.configure(yscrollcommand=vsb.set)
# 横スクロールバー
hsb = tk.Scrollbar(frame3, orient='horizontal', command=tree.xview)
hsb.grid(row=1, column=0, sticky='ew')
tree.configure(xscrollcommand=hsb.set)

# サンプルデータ挿入
for i in range(50):
    tree.insert('', 'end', values=(i, f'Name {i}', i * 10))
root.mainloop()
