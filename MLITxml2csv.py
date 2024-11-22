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
            return coord_str

        decimal = degrees + (minutes / 60.0) + (seconds / 3600.0)
        return f"{decimal:.6f}"
    except ValueError:
        return coord_str

def calculate_average_coordinates(west, east, north, south):
    try:
        avg_longitude = (float(west) + float(east)) / 2
        avg_latitude = (float(north) + float(south)) / 2
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
        ','.join([safe_find_text(facility, '施設名称') for facility in facility_info])
    ]

def find_index_d_xml(folder):
    for file in os.listdir(folder):
        if file.lower() == 'index_d.xml':
            return os.path.join(folder, file)
    return None

def process_index_d_xml(folder, writer):
    index_d_path = find_index_d_xml(folder)
    
    if index_d_path:
        try:
            data = parse_xml(index_d_path)
            if data:
                writer.writerow(data)
                print(f"ファイル '{os.path.basename(index_d_path)}' の処理が完了しました。")
            else:
                print(f"警告: '{os.path.basename(index_d_path)}' の処理中にエラーが発生しました。")
        except Exception as e:
            print(f"エラー: ファイル '{os.path.basename(index_d_path)}' の処理中に問題が発生しました: {str(e)}")

def process_folders(input_folder, writer):
    for root, dirs, files in os.walk(input_folder):
        process_index_d_xml(root, writer)

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(current_dir, "input")
    output_csv = os.path.join(current_dir, "output.csv")

    if not os.path.isdir(input_folder):
        print(f"エラー: 'input' フォルダが見つかりません。スクリプトと同じディレクトリに 'input' フォルダを作成し、index_D.xmlファイルを配置してください。")
        exit(1)

    headers = [
        '平均境界経度', '平均境界緯度',
        '適用要領基準', '業務名称', '履行期間-着手', '履行期間-完了', '測地系',
        '西側境界座標経度', '東側境界座標経度', '北側境界座標緯度', '南側境界座標緯度',
        '発注者機関事務所名', '受注者名', '業務概要', 'BIMCIM対象', '業務キーワード', '施設名称'
    ]

    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        process_folders(input_folder, writer)

    print(f"処理が完了しました。結果は '{output_csv}' に保存されました。")