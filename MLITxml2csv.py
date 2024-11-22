import csv
import os
from lxml import etree

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
            return coord_str  # 不正な形式の場合はそのまま返す

        decimal = degrees + (minutes / 60.0) + (seconds / 3600.0)
        return f"{decimal:.6f}"
    except ValueError:
        return coord_str  # 数値変換に失敗した場合はそのまま返す

def parse_xml(xml_file_path):
    try:
        parser = etree.XMLParser(encoding='shift_jis')
        tree = etree.parse(xml_file_path, parser=parser)
        root = tree.getroot()
    except etree.XMLSyntaxError as e:
        print(f"警告: ファイル '{xml_file_path}' の解析中にXMLエラーが発生しました: {str(e)}")
        return None
    except UnicodeDecodeError:
        print(f"警告: ファイル '{xml_file_path}' のエンコーディングがShift-JISではない可能性があります。")
        return None
    except Exception as e:
        print(f"警告: ファイル '{xml_file_path}' の読み込み中にエラーが発生しました: {str(e)}")
        return None

    basic_info = root.find('基礎情報')
    project_info = root.find('業務件名等')
    location_info = root.find('場所情報')
    facility_info = root.findall('施設情報')
    client_info = root.find('発注者情報')
    contractor_info = root.find('受注者情報')
    work_info = root.find('業務情報')

    boundary_info = location_info.find('境界座標情報') if location_info is not None else None

    return [
        safe_find_text(basic_info, '適用要領基準'),
        safe_find_text(project_info, '業務名称'),
        safe_find_text(project_info, '履行期間-着手'),
        safe_find_text(project_info, '履行期間-完了'),
        safe_find_text(location_info, '測地系'),
        convert_coordinates(safe_find_text(boundary_info, '西側境界座標経度')),
        convert_coordinates(safe_find_text(boundary_info, '東側境界座標経度')),
        convert_coordinates(safe_find_text(boundary_info, '北側境界座標緯度')),
        convert_coordinates(safe_find_text(boundary_info, '南側境界座標緯度')),
        safe_find_text(client_info, '発注者機関事務所名'),
        safe_find_text(contractor_info, '受注者名'),
        safe_find_text(work_info, '業務概要'),
        safe_find_text(work_info, 'BIMCIM対象'),
        ','.join([safe_find_text(keyword, '.') for keyword in work_info.findall('業務キーワード')]) if work_info is not None else '',
        ','.join([safe_find_text(facility, '施設名称') for facility in facility_info])
    ]

def find_xml_files(folder):
    xml_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith('.xml'):
                xml_files.append(os.path.join(root, file))
    return xml_files

def process_xml_folder(input_folder, output_csv):
    headers = [
        '適用要領基準', '業務名称', '履行期間-着手', '履行期間-完了', '測地系',
        '西側境界座標経度', '東側境界座標経度', '北側境界座標緯度', '南側境界座標緯度',
        '発注者機関事務所名', '受注者名', '業務概要', 'BIMCIM対象', '業務キーワード', '施設名称'
    ]

    xml_files = find_xml_files(input_folder)
    
    if not xml_files:
        print(f"警告: フォルダ '{input_folder}' とそのサブフォルダにXMLファイルが見つかりません。")
        return

    processed_files = 0
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)

        for xml_file_path in xml_files:
            try:
                data = parse_xml(xml_file_path)
                if data:
                    writer.writerow(data)
                    processed_files += 1
                    print(f"ファイル '{xml_file_path}' の処理が完了しました。")
            except Exception as e:
                print(f"エラー: ファイル '{xml_file_path}' の処理中に問題が発生しました: {str(e)}")

    print(f"全ての処理が完了しました。{processed_files}個のファイルが正常に処理され、'{output_csv}' に保存されました。")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # フォルダinputを実際に業務委託データを格納しているフォルダに変更して利用
    input_folder = os.path.join(current_dir, "input")
    # 出力は常に、この階層
    output_csv = os.path.join(current_dir, "output.csv")

    if not os.path.isdir(input_folder):
        print(f"エラー: 'input' フォルダが見つかりません。スクリプトと同じディレクトリに 'input' フォルダを作成し、XMLファイルを配置してください。")
        exit(1)

    process_xml_folder(input_folder, output_csv)