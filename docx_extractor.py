import os
import zipfile
import shutil
from datetime import datetime

def extract_docx_structure(docx_file_path, output_dir="word/structure"):
    """
    .docxファイルの構成要素を指定ディレクトリに展開する
    
    Args:
        docx_file_path (str): .docxファイルのパス
        output_dir (str): 出力ディレクトリのパス
    """
    try:
        if not os.path.exists(docx_file_path):
            print(f"ファイルが見つかりません: {docx_file_path}")
            return
        
        print("=" * 80)
        print(f"DOCX構成要素展開: {docx_file_path}")
        print("=" * 80)
        
        # 出力ディレクトリを作成（既存の場合は削除して再作成）
        if os.path.exists(output_dir):
            print(f"既存のディレクトリを削除中: {output_dir}")
            shutil.rmtree(output_dir)
        
        os.makedirs(output_dir, exist_ok=True)
        print(f"出力ディレクトリを作成: {output_dir}")
        print()
        
        # ファイル基本情報
        file_size = os.path.getsize(docx_file_path)
        print(f"元ファイルサイズ: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        extracted_files = 0
        total_extracted_size = 0
        
        with zipfile.ZipFile(docx_file_path, 'r') as zip_file:
            file_list = zip_file.filelist
            print(f"展開対象ファイル数: {len([f for f in file_list if not f.is_dir()])}")
            print()
            
            # 進捗表示用
            print("展開中...")
            print(f"{'No':<4} {'ファイル名':<60} {'サイズ':<12} {'種別'}")
            print("-" * 85)
            
            for idx, file_info in enumerate(file_list, 1):
                if file_info.is_dir():
                    # ディレクトリの場合は作成のみ
                    dir_path = os.path.join(output_dir, file_info.filename)
                    os.makedirs(dir_path, exist_ok=True)
                    continue
                
                # ファイルを展開
                file_data = zip_file.read(file_info.filename)
                output_file_path = os.path.join(output_dir, file_info.filename)
                
                # 必要に応じてディレクトリを作成
                os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
                
                # ファイルを書き込み
                with open(output_file_path, 'wb') as output_file:
                    output_file.write(file_data)
                
                # ファイル種別を判定
                file_type = get_file_type(file_info.filename)
                
                print(f"{idx:<4} {file_info.filename:<60} {file_info.file_size:<12,} {file_type}")
                
                extracted_files += 1
                total_extracted_size += file_info.file_size
        
        print("-" * 85)
        print(f"展開完了: {extracted_files} ファイル")
        print(f"総展開サイズ: {total_extracted_size:,} bytes ({total_extracted_size / 1024 / 1024:.2f} MB)")
        print()
        
        # 展開されたディレクトリ構造を表示
        print("=" * 80)
        print("展開されたディレクトリ構造:")
        print("=" * 80)
        show_directory_tree(output_dir)
        
        # 画像ファイルの詳細
        show_media_files_info(output_dir)
        
        print("=" * 80)
        print(f"展開完了: {output_dir}")
        print("=" * 80)
        
    except zipfile.BadZipFile:
        print(f"無効なZIPファイルです: {docx_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

def get_file_type(filename):
    """
    ファイル名から種別を判定する
    
    Args:
        filename (str): ファイル名
        
    Returns:
        str: ファイル種別
    """
    if filename.startswith('word/media/'):
        return "画像"
    elif filename.endswith('.xml'):
        if 'document.xml' in filename:
            return "文書本体"
        elif 'styles.xml' in filename:
            return "スタイル"
        elif 'settings.xml' in filename:
            return "設定"
        elif 'fontTable.xml' in filename:
            return "フォント"
        elif 'numbering.xml' in filename:
            return "番号付け"
        elif 'theme' in filename:
            return "テーマ"
        elif '_rels' in filename:
            return "関係定義"
        elif 'customXml' in filename:
            return "カスタムXML"
        else:
            return "XML"
    elif filename == '[Content_Types].xml':
        return "コンテンツ型定義"
    else:
        return "その他"

def show_directory_tree(root_dir, prefix="", max_depth=3, current_depth=0):
    """
    ディレクトリツリーを表示する
    
    Args:
        root_dir (str): ルートディレクトリ
        prefix (str): 表示用プレフィックス
        max_depth (int): 最大表示深度
        current_depth (int): 現在の深度
    """
    if current_depth >= max_depth:
        return
    
    try:
        items = os.listdir(root_dir)
        items.sort()
        
        for i, item in enumerate(items):
            item_path = os.path.join(root_dir, item)
            is_last = i == len(items) - 1
            
            if os.path.isdir(item_path):
                print(f"{prefix}{'└── ' if is_last else '├── '}{item}/")
                extension = "    " if is_last else "│   "
                show_directory_tree(item_path, prefix + extension, max_depth, current_depth + 1)
            else:
                file_size = os.path.getsize(item_path)
                print(f"{prefix}{'└── ' if is_last else '├── '}{item} ({file_size:,} bytes)")
    except PermissionError:
        print(f"{prefix}[アクセス権限エラー]")

def show_media_files_info(root_dir):
    """
    展開された画像ファイルの詳細情報を表示する
    
    Args:
        root_dir (str): ルートディレクトリ
    """
    media_dir = os.path.join(root_dir, "word", "media")
    
    if not os.path.exists(media_dir):
        print("\n画像ファイルは見つかりませんでした。")
        return
    
    print()
    print("=" * 80)
    print("展開された画像ファイル詳細:")
    print("=" * 80)
    
    media_files = []
    for filename in os.listdir(media_dir):
        file_path = os.path.join(media_dir, filename)
        if os.path.isfile(file_path):
            file_size = os.path.getsize(file_path)
            file_ext = os.path.splitext(filename)[1].lower()
            modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            media_files.append((filename, file_size, file_ext, modified_time, file_path))
    
    if not media_files:
        print("画像ファイルが見つかりませんでした。")
        return
    
    # ファイル名順にソート
    media_files.sort()
    
    print(f"{'No':<4} {'ファイル名':<25} {'サイズ':<12} {'拡張子':<8} {'更新日時':<20} {'パス'}")
    print("-" * 100)
    
    total_size = 0
    for idx, (filename, file_size, file_ext, modified_time, file_path) in enumerate(media_files, 1):
        rel_path = os.path.relpath(file_path, root_dir)
        print(f"{idx:<4} {filename:<25} {file_size:<12,} {file_ext:<8} {modified_time.strftime('%Y-%m-%d %H:%M:%S'):<20} {rel_path}")
        total_size += file_size
    
    print("-" * 100)
    print(f"総画像ファイル数: {len(media_files)}")
    print(f"総画像サイズ: {total_size:,} bytes ({total_size / 1024 / 1024:.2f} MB)")

def main():
    # 差し替え後のファイルを展開
    docx_file = "replace_data/replace_result/img_food_word.docx"
    output_dir = "word/structure"
    
    if not os.path.exists(docx_file):
        print(f"ファイルが見つかりません: {docx_file}")
        return
    
    extract_docx_structure(docx_file, output_dir)

if __name__ == "__main__":
    main() 