
import pyodbc
import csv
import pandas as pd
import tkinter as tk
from tkinter import ttk


df_table_global = None  # グローバル変数として用意
columns_global = []

def connect_info(servre_address, database, username, password):
    connect_info_str = (
        f'DRIVER={{SQL Server}};'
        f'SERVER={servre_address};'
        f'DATABASE={database};'
        f'UID={username};'
        f'PWD={password}'
    )
    return connect_info_str

def connect_server(servre_address, database, username, password, listbox):
    # 接続文字列
    conn_str = connect_info(servre_address, database, username, password)
    # 接続
    try:
        conn = pyodbc.connect(conn_str)
        print("接続に成功しました。")
        cursor = conn.cursor()
    # クエリ実行
        cursor.execute("SELECT name, create_date, modify_date FROM sys.tables;")
        # cursor.execute("SELECT TOP 5 * FROM T_Casting_Hist ORDER BY ProductTime DESC;")
    # カラム名を取得
        columns = [column[0] for column in cursor.description]
        # 全行取得
        rows = cursor.fetchall().copy()
        # DataFrameに変換
        global df_table_global
        df_table_global = pd.DataFrame.from_records(rows, columns=columns)
        print(df_table_global)
        print(type(df_table_global))        
        # Listboxの内容をクリア
        listbox.delete(0, 'end')
        # name列の値をListboxに追加
        for name in df_table_global['name']:
            listbox.insert('end', name)
            
    except pyodbc.Error as e:
        print("接続エラー:", e)
    # 接続終了
    finally:
        cursor.close()
        conn.close()

def load_tabledata(servre_address, database, username, password, table_name, filter_column, table_tree):
    # 接続文字列
    conn_str = connect_info(servre_address, database, username, password)
    # 接続
    try:
        conn = pyodbc.connect(conn_str)
        print("接続に成功しました。")
        cursor = conn.cursor()
    # クエリ実行
        if filter_column == None:
            cursor.execute(f"SELECT TOP 50 * FROM {table_name};")
        else:
            cursor.execute(f"SELECT TOP 50 * FROM {table_name} ORDER BY {filter_column} DESC;")
    # カラム名を取得        
        columns = [column[0] for column in cursor.description]
        # 全行取得
        rows = cursor.fetchall().copy()
        # DataFrameに変換        
        df_table_global = pd.DataFrame.from_records(rows, columns=columns)
        print(columns)
        print(df_table_global)
        print(type(df_table_global))
        def update_treeview(table_tree, update_columns, data):
            # 列の更新
            table_tree["columns"] = update_columns
            for col in update_columns:
                table_tree.heading(col, text=col)
                table_tree.column(col, minwidth=100, anchor='center')
            # データのクリア
            for item in table_tree.get_children():
                table_tree.delete(item)
            # データ挿入
            for row in data:
                table_tree.insert('', 'end', values=row)                
        def update_treeview_from_df(table_tree, df_table_global):
            update_columns = list(df_table_global.columns)
            data = df_table_global.values.tolist()
            update_treeview(table_tree, update_columns, data)
        
        update_treeview_from_df(table_tree, df_table_global)
    except pyodbc.Error as e:
        print("接続エラー:", e)
    # 接続終了
    finally:
        cursor.close()
        conn.close()


filter_columns_global=[]        
def load_table_columns(servre_address, database, username, password, table_name, combobox):
    # 接続文字列
    conn_str = connect_info(servre_address, database, username, password)
    # 接続
    try:
        conn = pyodbc.connect(conn_str)
        print("接続に成功しました。")
        cursor = conn.cursor()
    # クエリ実行
        cursor.execute(f"SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table_name}';")
    # カラム名を取得        
        columns = [column[0] for column in cursor.description]
        print(cursor.description)
        # 全行取得
        rows = cursor.fetchall().copy()
        # DataFrameに変換        
        df_table_calumns = pd.DataFrame.from_records(rows, columns=columns)
        # print(rows)
        # print(columns)
        # print(df_table_calumns["COLUMN_NAME"].to_list())
        # print(type(df_table_global))
        global filter_columns_global        
        filter_columns_global = df_table_calumns["COLUMN_NAME"].to_list() #カラム一覧をリストにしてグローバル変数に格納
        print(filter_columns_global)
        combobox["values"] = filter_columns_global
    except pyodbc.Error as e:
        print("接続エラー:", e)
    # 接続終了
    finally:
        cursor.close()
        conn.close()


