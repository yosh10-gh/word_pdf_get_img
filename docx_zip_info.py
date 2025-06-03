import os
import zipfile
from datetime import datetime

def show_docx_zip_info(docx_file_path):
    """
    .docxファイルのzipファイル構造を詳細表示する
    
    Args:
        docx_file_path (str): .docxファイルのパス
    """
    try:
        if not os.path.exists(docx_file_path):
            print(f"ファイルが見つかりません: {docx_file_path}")
            return
        
        print("=" * 80)
        print(f"DOCXファイル ZIP構造情報: {docx_file_path}")
        print("=" * 80)
        
        # ファイル基本情報
        file_size = os.path.getsize(docx_file_path)
        modified_time = datetime.fromtimestamp(os.path.getmtime(docx_file_path))
        print(f"ファイルサイズ: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        print(f"更新日時: {modified_time}")
        print()
        
        with zipfile.ZipFile(docx_file_path, 'r') as zip_file:
            # ZIP全体の情報
            file_list = zip_file.filelist
            total_files = len(file_list)
            total_compressed_size = sum(info.compress_size for info in file_list)
            total_uncompressed_size = sum(info.file_size for info in file_list)
            
            print(f"総ファイル数: {total_files}")
            print(f"圧縮後サイズ: {total_compressed_size:,} bytes")
            print(f"展開後サイズ: {total_uncompressed_size:,} bytes")
            print(f"圧縮率: {(1 - total_compressed_size / total_uncompressed_size) * 100:.1f}%")
            print()
            
            # ディレクトリ別ファイル数
            directories = {}
            for info in file_list:
                if info.is_dir():
                    continue
                
                dir_name = os.path.dirname(info.filename)
                if dir_name == '':
                    dir_name = '(ルート)'
                
                if dir_name not in directories:
                    directories[dir_name] = []
                directories[dir_name].append(info)
            
            print("ディレクトリ別ファイル数:")
            for dir_name, files in sorted(directories.items()):
                print(f"  {dir_name}: {len(files)} ファイル")
            print()
            
            # 詳細ファイル一覧
            print("=" * 80)
            print("詳細ファイル一覧:")
            print("=" * 80)
            print(f"{'No':<4} {'ファイル名':<50} {'サイズ':<12} {'圧縮後':<12} {'圧縮率':<8} {'更新日時'}")
            print("-" * 95)
            
            for idx, info in enumerate(file_list, 1):
                if info.is_dir():
                    continue
                
                # 圧縮率計算
                if info.file_size > 0:
                    compression_ratio = (1 - info.compress_size / info.file_size) * 100
                else:
                    compression_ratio = 0
                
                # 更新日時
                file_datetime = datetime(*info.date_time)
                
                print(f"{idx:<4} {info.filename:<50} {info.file_size:<12,} {info.compress_size:<12,} {compression_ratio:<7.1f}% {file_datetime}")
            
            # 画像ファイル詳細
            media_files = [info for info in file_list if info.filename.startswith('word/media/') and not info.is_dir()]
            if media_files:
                print()
                print("=" * 80)
                print("画像ファイル詳細:")
                print("=" * 80)
                print(f"{'No':<4} {'画像ファイル名':<30} {'サイズ':<12} {'拡張子':<8} {'更新日時'}")
                print("-" * 70)
                
                for idx, info in enumerate(media_files, 1):
                    file_ext = os.path.splitext(info.filename)[1].lower()
                    file_datetime = datetime(*info.date_time)
                    basename = os.path.basename(info.filename)
                    
                    print(f"{idx:<4} {basename:<30} {info.file_size:<12,} {file_ext:<8} {file_datetime}")
                
                print()
                print(f"総画像ファイル数: {len(media_files)}")
                total_image_size = sum(info.file_size for info in media_files)
                print(f"総画像サイズ: {total_image_size:,} bytes ({total_image_size / 1024 / 1024:.2f} MB)")
        
        print("=" * 80)
        print("ZIP構造表示完了")
        print("=" * 80)
        
    except zipfile.BadZipFile:
        print(f"無効なZIPファイルです: {docx_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

def main():
    # 元ファイルと差し替え後ファイルの両方を表示
    original_file = "./test_directry/data/img_food_word.docx"
    replaced_file = "replace_data/replace_result/img_food_word.docx"
    
    if os.path.exists(original_file):
        print("【元ファイル】")
        show_docx_zip_info(original_file)
        print("\n\n")
    
    if os.path.exists(replaced_file):
        print("【差し替え後ファイル】")
        show_docx_zip_info(replaced_file)

if __name__ == "__main__":
    main() 