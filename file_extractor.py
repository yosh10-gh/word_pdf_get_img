import os
import glob
import pandas as pd
import zipfile
from PIL import Image
import io
import tempfile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows

def extract_docx_pdf_files(directory_path):
    """
    指定されたディレクトリを再帰的にクロールして、
    .docx と .pdf ファイルのパスを抽出する
    
    Args:
        directory_path (str): 検索対象のディレクトリパス
        
    Returns:
        list: 見つかった .docx と .pdf ファイルのパスのリスト
    """
    found_files = []
    
    try:
        # ディレクトリの存在確認
        if not os.path.exists(directory_path):
            print(f"警告: ディレクトリが存在しません: {directory_path}")
            return found_files
        
        # .docx ファイルを検索
        docx_pattern = os.path.join(directory_path, "**", "*.docx")
        docx_files = glob.glob(docx_pattern, recursive=True)
        # 有効なファイルのみを追加
        for file_path in docx_files:
            if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
                found_files.append(file_path)
        
        # .pdf ファイルを検索
        pdf_pattern = os.path.join(directory_path, "**", "*.pdf")
        pdf_files = glob.glob(pdf_pattern, recursive=True)
        # 有効なファイルのみを追加
        for file_path in pdf_files:
            if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
                found_files.append(file_path)
    
    except Exception as e:
        print(f"ファイル検索エラー: {str(e)}")
    
    return found_files

def has_images_in_docx(file_path):
    """
    .docxファイルに画像が含まれているかチェックする
    
    Args:
        file_path (str): .docxファイルのパス
        
    Returns:
        bool: 画像が含まれている場合True
    """
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return False
            
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # word/media/ フォルダ内のファイルをチェック
            for file_info in zip_file.filelist:
                if file_info.filename.startswith('word/media/'):
                    return True
        return False
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"  → docxファイル読み込みエラー: {str(e)}")
        return False
    except Exception:
        return False

def extract_images_from_docx(file_path):
    """
    .docxファイルから画像を抽出する
    
    Args:
        file_path (str): .docxファイルのパス
        
    Returns:
        list: 抽出された画像のPILImageオブジェクトのリスト
    """
    images = []
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return images
            
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # word/media/ フォルダ内のファイルを検索
            for file_info in zip_file.filelist:
                if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                    try:
                        # 画像データを読み込み
                        image_data = zip_file.read(file_info.filename)
                        if len(image_data) == 0:  # 空のファイルをスキップ
                            continue
                            
                        # PILImageオブジェクトに変換
                        image = Image.open(io.BytesIO(image_data))
                        # 画像形式を確認してRGBに変換
                        if image.mode in ('RGBA', 'LA', 'P'):
                            image = image.convert('RGB')
                        images.append(image)
                    except (Image.UnidentifiedImageError, OSError) as e:
                        print(f"    → 画像読み込みエラー ({file_info.filename}): {str(e)}")
                        continue
                    except Exception:
                        continue
    except (zipfile.BadZipFile, PermissionError, OSError) as e:
        print(f"  → docxファイル処理エラー: {str(e)}")
    except Exception:
        pass
    return images

def has_images_in_pdf(file_path):
    """
    .pdfファイルに画像が含まれているかチェックする
    
    Args:
        file_path (str): .pdfファイルのパス
        
    Returns:
        bool: 画像が含まれている場合True
    """
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return False
            
        import PyPDF2
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            # 暗号化されたPDFの場合
            if pdf_reader.is_encrypted:
                try:
                    pdf_reader.decrypt('')  # 空のパスワードで試行
                except:
                    print(f"  → 暗号化されたPDFです: {file_path}")
                    return False
            
            for page_num, page in enumerate(pdf_reader.pages):
                try:
                    if '/Resources' in page and '/XObject' in page['/Resources']:
                        xObject = page['/Resources']['/XObject'].get_object()
                        for obj in xObject:
                            if xObject[obj]['/Subtype'] == '/Image':
                                return True
                except Exception:
                    continue
        return False
    except (PyPDF2.errors.PdfReadError, PermissionError, OSError) as e:
        print(f"  → PDFファイル読み込みエラー: {str(e)}")
        return False
    except Exception:
        return False

def extract_images_from_pdf(file_path):
    """
    .pdfファイルから画像を抽出する
    
    Args:
        file_path (str): .pdfファイルのパス
        
    Returns:
        list: 抽出された画像のPILImageオブジェクトのリスト
    """
    images = []
    try:
        # ファイルの存在とアクセス権限をチェック
        if not os.path.exists(file_path) or not os.access(file_path, os.R_OK):
            return images
            
        import PyPDF2
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            # 暗号化されたPDFの場合
            if pdf_reader.is_encrypted:
                try:
                    pdf_reader.decrypt('')  # 空のパスワードで試行
                except:
                    print(f"  → 暗号化されたPDFをスキップ: {file_path}")
                    return images
            
            for page_num, page in enumerate(pdf_reader.pages):
                try:
                    if '/Resources' in page and '/XObject' in page['/Resources']:
                        xObject = page['/Resources']['/XObject'].get_object()
                        for obj in xObject:
                            if xObject[obj]['/Subtype'] == '/Image':
                                try:
                                    # 画像データを抽出
                                    img_obj = xObject[obj]
                                    if '/Filter' in img_obj:
                                        if img_obj['/Filter'] == '/DCTDecode':
                                            # JPEG画像
                                            img_data = img_obj._data
                                            if len(img_data) > 0:  # 空のデータをスキップ
                                                image = Image.open(io.BytesIO(img_data))
                                                if image.mode in ('RGBA', 'LA', 'P'):
                                                    image = image.convert('RGB')
                                                images.append(image)
                                        elif img_obj['/Filter'] == '/FlateDecode':
                                            # PNG/その他の圧縮画像
                                            try:
                                                width = int(img_obj['/Width'])
                                                height = int(img_obj['/Height'])
                                                img_data = img_obj._data
                                                if len(img_data) > 0 and '/ColorSpace' in img_obj:
                                                    if img_obj['/ColorSpace'] == '/DeviceRGB':
                                                        expected_size = width * height * 3
                                                        if len(img_data) >= expected_size:
                                                            image = Image.frombytes('RGB', (width, height), img_data[:expected_size])
                                                            images.append(image)
                                            except (ValueError, TypeError):
                                                continue
                                except (Image.UnidentifiedImageError, OSError, ValueError) as e:
                                    print(f"    → PDF画像抽出エラー (page {page_num + 1}): {str(e)}")
                                    continue
                                except Exception:
                                    continue
                except Exception:
                    continue
    except (PyPDF2.errors.PdfReadError, PermissionError, OSError) as e:
        print(f"  → PDFファイル処理エラー: {str(e)}")
    except Exception:
        pass
    return images

def filter_files_with_images(file_list):
    """
    ファイルリストから画像が含まれているファイルのみを抽出する
    
    Args:
        file_list (list): ファイルパスのリスト
        
    Returns:
        list: 画像が含まれているファイルのリスト
    """
    files_with_images = []
    
    for file_path in file_list:
        try:
            print(f"チェック中: {file_path}")
            
            if file_path.endswith('.docx'):
                if has_images_in_docx(file_path):
                    files_with_images.append(file_path)
                    print(f"  → 画像あり")
                else:
                    print(f"  → 画像なし")
                    
            elif file_path.endswith('.pdf'):
                if has_images_in_pdf(file_path):
                    files_with_images.append(file_path)
                    print(f"  → 画像あり")
                else:
                    print(f"  → 画像なし")
        except Exception as e:
            print(f"  → ファイルチェックエラー: {str(e)}")
            continue
    
    return files_with_images

def resize_image_to_100px(image):
    """
    画像を100px×100pxにリサイズする（アスペクト比を保持）
    
    Args:
        image: PILImageオブジェクト
        
    Returns:
        PILImageオブジェクト: リサイズされた画像
    """
    try:
        # 画像のサイズをチェック
        if image.size[0] == 0 or image.size[1] == 0:
            raise ValueError("無効な画像サイズ")
        
        # アスペクト比を保持して100px以内にリサイズ
        image.thumbnail((100, 100), Image.Resampling.LANCZOS)
        return image
    except Exception as e:
        print(f"    → 画像リサイズエラー: {str(e)}")
        # フォールバック: 最小サイズの白い画像を作成
        fallback_image = Image.new('RGB', (50, 50), 'white')
        return fallback_image

def save_to_excel_with_images(file_list, output_filename="検索結果.xlsx"):
    """
    ファイルリストと画像をExcelファイルに保存する
    
    Args:
        file_list (list): ファイルパスのリスト
        output_filename (str): 出力するExcelファイル名
    """
    if not file_list:
        print("保存するファイルがありません。")
        return
    
    # 新しいワークブックを作成
    wb = Workbook()
    ws = wb.active
    
    # ヘッダーを設定
    ws['A1'] = 'ファイルパス'
    ws['B1'] = 'image1'
    ws['C1'] = 'image2' 
    ws['D1'] = 'image3'
    
    # 行の高さを設定（100px用）
    for row in range(2, len(file_list) + 2):
        ws.row_dimensions[row].height = 75  # Excelのポイント単位
    
    # 列の幅を設定
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    
    temp_files = []  # 一時ファイルのリストを保持
    
    try:
        for idx, file_path in enumerate(file_list, start=2):
            try:
                print(f"画像抽出中: {file_path}")
                
                # A列にファイルパスを設定
                ws[f'A{idx}'] = file_path
                
                # 画像を抽出
                images = []
                if file_path.endswith('.docx'):
                    images = extract_images_from_docx(file_path)
                elif file_path.endswith('.pdf'):
                    images = extract_images_from_pdf(file_path)
                
                # 最大3つの画像をB、C、D列に配置
                columns = ['B', 'C', 'D']
                for img_idx, image in enumerate(images[:3]):  # 最大3つまで
                    try:
                        # 画像を100px×100pxにリサイズ
                        resized_image = resize_image_to_100px(image.copy())
                        
                        # 一時ファイルを作成（手動で削除）
                        temp_fd, temp_path = tempfile.mkstemp(suffix='.png')
                        temp_files.append(temp_path)  # 削除用にリストに追加
                        
                        try:
                            # ファイルディスクリプタを閉じる
                            os.close(temp_fd)
                            
                            # 画像を保存
                            resized_image.save(temp_path, 'PNG', optimize=True)
                            
                            # ファイルサイズをチェック（異常に大きい場合はスキップ）
                            if os.path.getsize(temp_path) > 10 * 1024 * 1024:  # 10MB制限
                                print(f"    → image{img_idx + 1}: ファイルサイズが大きすぎます")
                                continue
                            
                            # Excelに画像を挿入
                            img = OpenpyxlImage(temp_path)
                            img.width = 100
                            img.height = 100
                            
                            cell_location = f'{columns[img_idx]}{idx}'
                            ws.add_image(img, cell_location)
                            
                            print(f"  → image{img_idx + 1}: 配置完了")
                        except (OSError, PermissionError) as e:
                            print(f"  → image{img_idx + 1}: ファイル操作エラー ({str(e)})")
                        except Exception as e:
                            print(f"  → image{img_idx + 1}: エラー ({str(e)})")
                    except Exception as e:
                        print(f"  → image{img_idx + 1}: 画像処理エラー ({str(e)})")
                
                print(f"  → 合計 {len(images)} 個の画像を処理")
            except Exception as e:
                print(f"  → ファイル処理エラー: {str(e)}")
                continue
        
        # Excelファイルを保存
        try:
            wb.save(output_filename)
            print(f"Excelファイルを保存しました: {output_filename}")
        except (PermissionError, OSError) as e:
            print(f"Excelファイル保存エラー: {str(e)}")
            # 代替ファイル名で保存を試行
            import time
            alt_filename = f"検索結果_{int(time.time())}.xlsx"
            wb.save(alt_filename)
            print(f"代替ファイル名で保存しました: {alt_filename}")
    
    finally:
        # 一時ファイルをクリーンアップ
        for temp_path in temp_files:
            try:
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
            except Exception:
                pass  # 削除できなくても続行

def main():
    # 検索対象のディレクトリを指定（現在のディレクトリから開始）
    search_directory = "."
    
    print(f"ディレクトリをクロール中: {os.path.abspath(search_directory)}")
    print("-" * 50)
    
    # ファイルを抽出
    files = extract_docx_pdf_files(search_directory)
    
    # 結果を表示
    if files:
        print(f"見つかったファイル数: {len(files)}")
        
        # 画像が含まれているファイルのみを抽出
        print("\n" + "-" * 50)
        print("画像が含まれているファイルをチェック中...")
        print("-" * 50)
        
        files_with_images = filter_files_with_images(files)
        
        print("\n" + "-" * 50)
        print(f"画像が含まれているファイル数: {len(files_with_images)}")
        print("\n画像が含まれているファイル:")
        for i, file_path in enumerate(files_with_images, 1):
            print(f"{i:3d}. {file_path}")
        
        # Excelファイルに保存（画像含有ファイルのみ + 埋め込み画像表示）
        if files_with_images:
            print("\n" + "-" * 50)
            print("画像を埋め込んだExcelファイルを作成中...")
            print("-" * 50)
            save_to_excel_with_images(files_with_images, "画像含有ファイル検索結果.xlsx")
        else:
            print("\n画像が含まれているファイルがありませんでした。")
    else:
        print("該当するファイルが見つかりませんでした。")
    
    print("-" * 50)
    print("抽出完了")

if __name__ == "__main__":
    main() 