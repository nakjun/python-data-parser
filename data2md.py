import os
import tempfile
import base64
import urllib.parse
import shutil
from main import ppt_to_pdf
from pdf2md import convert_pdf_to_markdown

def decode_filename(filename):
    decoding_attempts = [
        # 1. Base64 디코딩
        lambda name: base64.b64decode(name).decode('utf-8'),
        # 2. URI 디코딩
        lambda name: urllib.parse.unquote(name),
        # 3. 이중 URI 디코딩
        lambda name: urllib.parse.unquote(urllib.parse.unquote(name)),
        # 4. UTF-8 디코딩
        lambda name: name.encode('iso-8859-1').decode('utf-8'),
        # 5. ISO-8859-1로 해석 후 UTF-8로 변환
        lambda name: name.encode('iso-8859-1').decode('utf-8')
    ]

    for attempt in decoding_attempts:
        try:
            decoded = attempt(filename)
            if any('\u3131' <= char <= '\uD79D' for char in decoded):
                return decoded
        except:
            continue

    # 모든 시도가 실패하면 원본 반환
    return filename

def change_filename(input_folder):    
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            input_path = os.path.join(root, file)
            filename_without_ext, ext = os.path.splitext(file)
            
            decoded_filename = decode_filename(filename_without_ext)
            print(f"원본 파일명: {file}")
            print(f"디코딩된 파일명: {decoded_filename}")
            
            # 디코딩된 파일명으로 파일 이름 변경
            new_filename = f"{decoded_filename}{ext}"
            new_input_path = os.path.join(root, new_filename)
            os.rename(input_path, new_input_path)
            print(f"파일명 변경: {file} -> {new_filename}")

def process_files(input_folder, output_folder):
    print(f"입력 폴더: {input_folder}")
    
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith((".ppt", ".pptx")):
                input_path = os.path.abspath(os.path.join(root, file))
                filename_without_ext, ext = os.path.splitext(file)
                
                relative_path = os.path.relpath(root, input_folder)
                pdf_output_path = os.path.abspath(os.path.join(output_folder, relative_path, f"{filename_without_ext}.pdf"))
                md_output_path = os.path.abspath(os.path.join(output_folder, relative_path, f"{filename_without_ext}.md"))
                
                os.makedirs(os.path.dirname(pdf_output_path), exist_ok=True)
                
                print(f"PPT 파일 처리 중: {file}")
                print(f"입력 경로: {input_path}")
                print(f"PDF 출력 경로: {pdf_output_path}")
                
                if not os.path.exists(input_path):
                    print(f"오류: 입력 파일을 찾을 수 없습니다: {input_path}")
                    continue
                
                try:
                    ppt_to_pdf(input_path, pdf_output_path)
                    print(f"{file}을(를) PDF로 변환했습니다.")
                    
                    # md_content = convert_pdf_to_markdown(pdf_output_path)
                    # with open(md_output_path, 'w', encoding='utf-8') as f:
                    #     f.write(md_content)
                    # print(f"PDF를 Markdown으로 변환 완료: {md_output_path}")
                except Exception as e:
                    print(f"오류 발생: {file} 처리 중 문제가 발생했습니다.")
                    print(f"오류 메시지: {str(e)}")
            else:
                print(f"PPT 파일이 아님, 처리하지 않음: {file}")

        # 하위 디렉토리 처리
        for dir in dirs:
            sub_input_folder = os.path.join(root, dir)
            sub_output_folder = os.path.join(output_folder, os.path.relpath(sub_input_folder, input_folder))
            process_files(sub_input_folder, sub_output_folder)

input_folder = os.path.abspath("data")
output_folder = os.path.abspath("/result/data")
#change_filename(input_folder)
process_files(input_folder, output_folder)