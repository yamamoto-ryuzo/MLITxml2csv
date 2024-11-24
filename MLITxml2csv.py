# -*- coding: utf-8 -*-

# インストール
# python本体
# https://www.python.org/downloads/
# インストール時【Add python.exe to PATH】にチェックを入れる
# pip install pyinstaller
# アップデートがあるとき
# python.exe -m pip install --upgrade pip

# EXE作成
#　ディレクトリは適宜変更
# cd C:\github\MLITxml2csv
# pyinstaller MLITxml2csv.py --onefile --noconsole --distpath ./ --clean
#　完成したら　-----.exeとかメッセージが出て完成

import csv
import os
import re
from lxml import etree
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import xml.etree.ElementTree as ET
import codecs

def safe_find_text(element, path, default=''):
    if element is None:
        return default
    found = element.find(path)
    return found.text if found is not None and found.text is not None else default

def convert_coordinates(coord_str):
    if not coord_str or len(coord_str) < 5:
        return ''
    try:
        if len(coord_str) == 6:  # 緯度の場合
            degrees = int(coord_str[:2])
            minutes = int(coord_str[2:4])
            seconds = int(coord_str[4:6])
        elif len(coord_str) == 7:  # 経度の場合
            degrees = int(coord_str[:3])
            minutes = int(coord_str[3:5])
            seconds = int(coord_str[5:7])
        else:
            return coord_str

        decimal = degrees + (minutes / 60.0) + (seconds / 3600.0)
        return f"{decimal:.6f}"
    except ValueError:
        return coord_str

def calculate_average_coordinates(west, east, north, south):
    try:
        avg_longitude = (float(west) + float(east)) / 2
        avg_latitude = (float(north) + float(south)) / 2
        # エラーは皇居に設定
        if avg_longitude > 180:
            avg_longitude = 139.7528144  
        if avg_latitude > 180:
            avg_latitude = 35.6852211
        return f"{avg_longitude:.6f}", f"{avg_latitude:.6f}"
    except ValueError:
        return '', ''

def parse_xml(xml_file_path):
    try:
        parser = etree.XMLParser(encoding='shift_jis')
        tree = etree.parse(xml_file_path, parser=parser)
        root = tree.getroot()
    except Exception as e:
        print(f"警告: ファイル '{xml_file_path}' の解析中にエラーが発生しました: {str(e)}")
        return None

    basic_info = root.find('基礎情報')
    project_info = root.find('業務件名等')
    location_info = root.find('場所情報')
    facility_info = root.findall('施設情報')
    client_info = root.find('発注者情報')
    contractor_info = root.find('受注者情報')
    work_info = root.find('業務情報')

    boundary_info = location_info.find('境界座標情報') if location_info is not None else None

    west = convert_coordinates(safe_find_text(boundary_info, '西側境界座標経度'))
    east = convert_coordinates(safe_find_text(boundary_info, '東側境界座標経度'))
    north = convert_coordinates(safe_find_text(boundary_info, '北側境界座標緯度'))
    south = convert_coordinates(safe_find_text(boundary_info, '南側境界座標緯度'))

    avg_longitude, avg_latitude = calculate_average_coordinates(west, east, north, south)

    # report_XMLの変換
    report_xml_to_csv(xml_file_path,safe_find_text(project_info, '業務名称'))

    # 属性の保存
    return [
        avg_longitude,
        avg_latitude,
        safe_find_text(basic_info, '適用要領基準'),
        safe_find_text(project_info, '業務名称'),
        safe_find_text(project_info, '履行期間-着手'),
        safe_find_text(project_info, '履行期間-完了'),
        safe_find_text(location_info, '測地系'),
        west, east, north, south,
        safe_find_text(client_info, '発注者機関事務所名'),
        safe_find_text(contractor_info, '受注者名'),
        safe_find_text(work_info, '業務概要'),
        safe_find_text(work_info, 'BIMCIM対象'),
        ','.join([safe_find_text(keyword, '.') for keyword in work_info.findall('業務キーワード')]) if work_info is not None else '',
        ','.join([safe_find_text(facility, '施設名称') for facility in facility_info]),
        xml_file_path
    ]

def find_index_d_xml(folder):
    for file in os.listdir(folder):
        if file.lower() == 'index_d.xml':
            return os.path.join(folder, file)
    return None

def process_index_d_xml(folder, writer, report_csvwriter):
    index_d_path = find_index_d_xml(folder)
    
    if index_d_path:
        try:
            data = parse_xml(index_d_path)
            if data:
                writer.writerow(data)
                print(f"ファイル '{os.path.basename(index_d_path)}' の処理が完了しました。")
            else:
                print(f"警告: '{os.path.basename(index_d_path)}' の処理中にエラーが発生しました。")
            
            # report.XML　の取得
            #  data = parse_xml(index_d_path)の返り値 4列目が業務名　
            process_xml_to_csv(folder, data[3],report_csvwriter)
            
        except Exception as e:
            print(f"エラー: ファイル '{os.path.basename(index_d_path)}' の処理中に問題が発生しました: {str(e)}")

def process_folders(input_folder, writer,report_csvwriter):
    for root, dirs, files in os.walk(input_folder):
        process_index_d_xml(root, writer,report_csvwriter)

import os
import tkinter as tk
from tkinter import filedialog

def select_input_folder():
    global input_folder
    input_folder = filedialog.askdirectory(title="入力フォルダを選択してください")
    folder_label.config(text=f"選択されたフォルダ: {input_folder}")

def start_processing():
    root.destroy()  # GUIを閉じる
    
# ==========================
# === ジオパッケージの作成 ===
# ==========================
def csv_to_geopackage(csv_file, output_gpkg, lon_col='平均境界経度', lat_col='平均境界緯度', crs="EPSG:4326"):
    """
    CSVファイルをジオパッケージに変換する関数

    Parameters:
    csv_file (str): 入力CSVファイルのパス
    output_gpkg (str): 出力ジオパッケージファイルのパス
    lon_col (str): 経度のカラム名（デフォルト: '平均境界経度'）
    lat_col (str): 緯度のカラム名（デフォルト: '平均境界緯度'）
    crs (str): 座標参照系（デフォルト: "EPSG:4326"）

    Returns:
    None
    """
    try:
        # CSVファイルを読み込む
        df = pd.read_csv(csv_file)

        # Pointジオメトリを作成
        geometry = [Point(xy) for xy in zip(df[lon_col], df[lat_col])]

        # GeoDataFrameを作成
        gdf = gpd.GeoDataFrame(df, geometry=geometry, crs=crs)

        # ジオパッケージとして保存
        gdf.to_file(output_gpkg, driver="GPKG")

        print(f"ジオパッケージが作成されました: {output_gpkg}")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

# ===========================
# === report.XMLの読み込み ===
# ===========================
def open_csv_file(csv_file_path):
    csvfile = open(csv_file_path, 'w', newline='', encoding='utf-8')
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(['業務名', '報告書名', '報告書副題', '報告書ファイル名', '報告書ファイル日本語名', '情報取得ファイル'])
    return csvfile, csvwriter

def close_csv_file(csvfile):
    csvfile.close()

def process_xml_to_csv(folder, project_name, csvwriter):
    report_xml_file_path = os.path.join(folder, 'report', 'report.xml')
    print(f'report_XMLを確定: {report_xml_file_path}')
    
    if not os.path.exists(report_xml_file_path):
        print(f'XMLファイルが見つかりません: {report_xml_file_path}')
        return
    
    # XMLファイルを解析
    parser = etree.XMLParser(encoding='shift_jis')
    report_tree = etree.parse(report_xml_file_path, parser=parser)
    report_root = report_tree.getroot()

    # XMLから必要なデータを抽出してCSVに書き込む
    for element in report_root.findall('.//報告書ファイル情報'):
        report_name = element.findtext('報告書名', default='')
        report_subtitle = element.findtext('報告書副題', default='')
        report_filename = element.findtext('報告書ファイル名', default='')
        report_japanese_filename = element.findtext('報告書ファイル日本語名', default='')
        csvwriter.writerow([project_name, report_name, report_subtitle, report_filename, report_japanese_filename, report_xml_file_path])


def report_xml_to_csv(report_file_path,project_name):
    report_folder_path = os.path.dirname(report_file_path)
    report_xml_file_path = os.path.join(report_folder_path, 'report', 'report.xml')
    print(f'report_XMLを確定: {report_xml_file_path}')
    
    if not os.path.exists(report_xml_file_path):
        print(f'XMLファイルが見つかりません: {report_xml_file_path}')
        return
    
    report_csv_file_path = os.path.splitext(report_xml_file_path)[0] + '.csv'
    
    # XMLファイルを解析
    parser = etree.XMLParser(encoding='shift_jis')
    report_tree = etree.parse(report_xml_file_path,parser=parser)
    report_root = report_tree.getroot()

    # CSVファイルを開く
    with open(report_csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        
        # ヘッダーを書き込む（必要に応じて）
        csvwriter.writerow(['業務名','報告書名','報告書副題', '報告書ファイル名', '報告書ファイル日本語名','情報取得ファイル'])
        
        # XMLから必要なデータを抽出してCSVに書き込む
        for element in report_root.findall('.//報告書ファイル情報'):
            report_name = element.findtext('報告書名', default='')
            report_subtitle = element.findtext('報告書副題', default='')
            report_filename = element.findtext('報告書ファイル名', default='')
            report_japanese_filename = element.findtext('報告書ファイル日本語名', default='')
            csvwriter.writerow([report_name, report_subtitle, report_filename, report_japanese_filename,report_xml_file_path])
    print(f"reportファイルを保存しました: {report_csv_file_path}")

            
# ===========================
# ====== プログラム本体 ======
# ===========================       
if __name__ == "__main__":
    current_dir = os.getcwd()
    # ＝＝＝＝＝＝＝＝＝＝＝
    # ＝＝＝GUIの設定＝＝＝
    # ＝＝＝＝＝＝＝＝＝＝＝
    root = tk.Tk()
    root.title("業務委託電子納品　概要書の集約")
    root.geometry("400x200")

    select_button = tk.Button(root, text="入力フォルダを選択", command=select_input_folder)
    select_button.pack(pady=20)

    folder_label = tk.Label(root, text="フォルダが選択されていません", wraplength=380)
    folder_label.pack(pady=10)

    start_button = tk.Button(root, text="INDEX_D.XMLをMLITxml.csvに集約\n\nREPORT.XMLをMLITreportxml.csvに集約", command=start_processing)
    start_button.pack(pady=20)

    root.mainloop()

    # ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
    # ＝＝＝ GUIが閉じられた後の処理 ＝＝＝
    # ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
    
    # ＝＝＝ index_D.XMLを変換したCSVファイルの作成 ＝＝＝
    output_csv = os.path.join(current_dir, "MLITxml.csv")
    
    print(f"選択された入力フォルダ: {input_folder}")
    print(f"出力CSVファイル: {output_csv}")
    
    if not os.path.isdir(input_folder):
        print(f"エラー: 'input' フォルダが見つかりません。スクリプトと同じディレクトリに 'input' フォルダを作成し、index_D.xmlファイルを配置してください。")
        exit(1)

    # index_D.XMLを変換したCSVファイル属性情報
    headers = [
        '平均境界経度', '平均境界緯度',
        '適用要領基準', '業務名称', '履行期間-着手', '履行期間-完了', '測地系',
        '西側境界座標経度', '東側境界座標経度', '北側境界座標緯度', '南側境界座標緯度',
        '発注者機関事務所名', '受注者名', '業務概要', 'BIMCIM対象', '業務キーワード', '施設名称' , '情報取得ファイル'
    ]

    # report.XMLを変換したCSVファイルの保存先を作成
    report_csv_file_path = os.path.join(current_dir, "MLITreportxml.csv")
    report_csvfile, report_csvwriter = open_csv_file(report_csv_file_path)

    # index_D.XMLを変換したCSVファイルの書き込み
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        process_folders(input_folder, writer, report_csvwriter)
         
    # report.XMLを変換したCSVファイルの保存先を保存
    close_csv_file(report_csvfile)

    # === ジオパッケージの作成 ===
    input_csv = output_csv
    output_gpkg = os.path.join(current_dir, "MLITxml.gpkg")
    csv_to_geopackage(input_csv, output_gpkg)

    print(f"処理が完了しました。結果は\n '{output_csv}' \n '{report_csv_file_path}' \nに保存されました。")
    
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示にする
    messagebox.showerror("報告", f"結果は\n '{output_csv}' \n '{report_csv_file_path}' \nに保存されました。")
    root.destroy()  # Tkinterのインスタンスを破棄

