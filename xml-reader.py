import csv
import os
from lxml import etree

def safe_find_text(element, path, default=''):
    found = element.find(path) if element is not None else None
    return found.text if found is not None else default

def convert_coordinates(coord_str):
    if not coord_str or len(coord_str) < 5:
        return ''
    if len(coord_str) == 6:  # 緯度の場合
        degrees = int(coord_str[:2])
        minutes = int(coord_str[3:4])
        seconds = int(coord_str[5:6])
    elif len(coord_str) == 7:  # 経度の場合
        degrees = int(coord_str[:3])
        minutes = int(coord_str[4:5])
        seconds = int(coord_str[6:7])
    else:
        return coord_str  # 不正な形式の場合はそのまま返す

    decimal = degrees + (minutes / 60.0) + (seconds / 3600.0)
    return f"{decimal:.6f}"

def parse_xml(xml_file_path):
    try:
        with open(xml_file_path, 'r', encoding='shift_jis') as file:
            parser = etree.XMLParser(encoding='shift_jis')
            root = etree.fromstring(file.read().encode('shift_jis'), parser=parser)
    except etree.XMLSyntaxError as e:
        print(f"警告: ファイル '{xml_file_path}' の解析中にXMLエラーが発生しました: {str(e)}")
        return None
    except UnicodeDecodeError:
        print(f"警告: ファイル '{xml_file_path}' のエンコーディングがShift-JISではない可能性があります。")
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
        ','.join([keyword.text for keyword in work_info.findall('業務キーワード')]) if work_info is not None else '',
        ','.join([safe_find_text(facility, '施設名称') for facility in facility_info])
    ]

def process_xml_folder(input_folder, output_csv):
    headers = [
        '適用要領基準', '業務名称', '履行期間-着手', '履行期間-完了', '測地系',
        '西側境界座標経度', '東側境界座標経度', '北側境界座標緯度', '南側境界座標緯度',
        '発注者機関事務所名', '受注者名', '業務概要', 'BIMCIM対象', '業務キーワード', '施設名称'
    ]

    xml_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xml')]
    
    if not xml_files:
        print(f"警告: フォルダ '{input_folder}' にXMLファイルが見つかりません。")
        return

    processed_files = 0
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)

        for filename in xml_files:
            xml_file_path = os.path.join(input_folder, filename)
            try:
                data = parse_xml(xml_file_path)
                if data:
                    writer.writerow(data)
                    processed_files += 1
                    print(f"ファイル '{filename}' の処理が完了しました。")
            except Exception as e:
                print(f"エラー: ファイル '{filename}' の処理中に問題が発生しました: {str(e)}")

    print(f"全ての処理が完了しました。{processed_files}個のファイルが正常に処理され、'{output_csv}' に保存されました。")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(current_dir, "input")
    output_csv = os.path.join(current_dir, "output.csv")

    if not os.path.isdir(input_folder):
        print(f"エラー: 'input' フォルダが見つかりません。スクリプトと同じディレクトリに 'input' フォルダを作成し、XMLファイルを配置してください。")
        exit(1)

    process_xml_folder(input_folder, output_csv)