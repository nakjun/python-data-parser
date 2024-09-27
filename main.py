# pip install comtypes pdf2image Pillow transformers torch pytesseract easyocr reportlab PyMuPDF streamlit
import streamlit as st
import os
import tempfile
from comtypes import client
from pdf2image import convert_from_path
from PIL import Image
from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError
from transformers import AutoTokenizer, AutoModel
import torch
import easyocr
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import fitz

def ppt_to_pdf(ppt_path, pdf_path):
    """
    PPT 파일을 PDF로 변환하는 함수
    
    :param ppt_path: PPT 파일 경로
    :param pdf_path: 저장할 PDF 파일 경로
    :return: 성공 시 True, 실패 시 False
    """
    try:
        powerpoint = client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(ppt_path)
        presentation.SaveAs(pdf_path, 32)  # 32는 PDF 형식을 나타냅니다
        presentation.Close()
        powerpoint.Quit()
        
        print(f"PPT 파일이 {pdf_path}로 변환되었습니다.")
        return True
    except Exception as e:
        print(f"PPT를 PDF로 변환하는 중 오류 발생: {str(e)}")
        return False

def pdf_to_images(pdf_path, output_folder):
    """
    PDF 파일을 이미지로 변환하는 함수
    
    :param pdf_path: PDF 파일 경로
    :param output_folder: 이미지를 저장할 폴더 경로
    :return: 저장된 이미지 파일 경로 리스트, 실패 시 빈 리스트
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    poppler_bin_path = r"D:\Library\poppler-24.07.0\Library\bin"
    if not os.path.exists(poppler_bin_path):
        print("Poppler 경로가 잘못되었습니다.")
        return []
    
    try:
        images = convert_from_path(pdf_path, poppler_path=poppler_bin_path)
        image_paths = []
        
        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f"page_{i+1}.png")
            image.save(image_path, "PNG")
            image_paths.append(image_path)
        
        print(f"{len(image_paths)}개의 이미지가 {output_folder}에 저장되었습니다.")
        return image_paths
    except Exception as e:
        print(f"PDF를 이미지로 변환하는 중 오류 발생: {str(e)}")
        return []

def ppt_to_pdf_files(test_folder, output_folder):
    for filename in os.listdir(test_folder):
        if filename.endswith(".ppt") or filename.endswith(".pptx"):
            ppt_path = os.path.join(test_folder, filename)
            pdf_filename = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(output_folder, pdf_filename)
            
            ppt_to_pdf(ppt_path, pdf_path)
            print(f"{filename}을(를) {pdf_filename}로 변환했습니다.")

def load_ocr_model():
    """
    OCR 모델과 토크나이저를 로드하는 함수
    """
    tokenizer = AutoTokenizer.from_pretrained('ucaslcl/GOT-OCR2_0', trust_remote_code=True)
    model = AutoModel.from_pretrained('ucaslcl/GOT-OCR2_0', trust_remote_code=True, low_cpu_mem_usage=True, device_map='cuda', use_safetensors=True, pad_token_id=tokenizer.eos_token_id)
    return tokenizer, model.eval().cuda()

def ppt_to_image_ocr(ppt_path, pdf_path, output_folder):
    """
    PPT를 PDF로 변환하고, 이미지로 변환한 후 OCR을 수행하는 함수
    
    :param ppt_path: PPT 파일 경로
    :param pdf_path: 저장할 PDF 파일 경로
    :param output_folder: 이미지와 OCR 결과를 저장할 폴더 경로
    :return: OCR 결과 텍스트 리스트, 실패 시 빈 리스트
    """
    try:
        # PPT를 PDF로 변환
        if not ppt_to_pdf(ppt_path, pdf_path):
            raise Exception("PPT를 PDF로 변환하는데 실패했습니다.")
        
        # PDF를 이미지로 변환
        image_paths = pdf_to_images(pdf_path, output_folder)
        if not image_paths:
            raise Exception("PDF를 이미지로 변환하는데 실패했습니다.")
        
        # OCR 수행
        reader = easyocr.Reader(['en', 'ko'])
        ocr_results = []
        
        for i, image_path in enumerate(image_paths):
            image = Image.open(image_path)
            res = reader.readtext(image, detail=0, paragraph=True)
            
            # OCR 결과를 텍스트 파일로 저장
            txt_path = os.path.join(output_folder, f"ocr_result_{i+1}.txt")
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(res))
            
            ocr_results.append('\n'.join(res))
            print(f"이미지 {i+1}의 OCR 결과가 {txt_path}에 저장되었습니다.")
        
        return ocr_results
    except Exception as e:
        print(f"PPT를 이미지로 변환하고 OCR을 수행하는 중 오류 발생: {str(e)}")
        return []

def pdf_to_image_ocr(pdf_path, output_folder):
    """
    PDF를 이미지로 변환하고 OCR을 수행하는 함수
    
    :param pdf_path: PDF 파일 경로
    :param output_folder: 이미지와 OCR 결과를 저장할 폴더 경로
    :return: OCR 결과 텍스트 리스트, 실패 시 빈 리스트
    """
    try:
        # PDF를 이미지로 변환
        image_paths = pdf_to_images(pdf_path, output_folder)
        if not image_paths:
            raise Exception("PDF를 이미지로 변환하는데 실패했습니다.")
        
        # OCR 수행
        reader = easyocr.Reader(['en', 'ko'])
        ocr_results = []
        
        for i, image_path in enumerate(image_paths):
            res = reader.readtext(image_path, detail=0, paragraph=True)
            
            # OCR 결과를 텍스트 파일로 저장
            txt_path = os.path.join(output_folder, f"ocr_result_{i+1}.txt")
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(res))
            
            ocr_results.append('\n'.join(res))
            print(f"이미지 {i+1}의 OCR 결과가 {txt_path}에 저장되었습니다.")
        
        return ocr_results
    except Exception as e:
        print(f"PDF를 이미지로 변환하고 OCR을 수행하는 중 오류 발생: {str(e)}")
        return []

def txt_to_pdf_convert(txt_files_path, pdf_file_name):
    """
    특정 경로의 모든 txt 파일을 하나의 PDF로 변환합니다.
    각 txt 파일은 PDF에서 1페이지를 차지합니다.

    :param txt_files_path: txt 파일들이 있는 디렉토리 경로
    :param pdf_file_name: 생성할 PDF 파일의 이름
    :return: 성공 시 True, 실패 시 False
    """
    try:
        # 한글 폰트 등록
        font_path = 'NanumGothic.ttf'  # 폰트 파일 경로를 적절히 수정해주세요
        pdfmetrics.registerFont(TTFont('NanumGothic', font_path))

        # PDF 생성
        c = canvas.Canvas(pdf_file_name, pagesize=letter)
        width, height = letter

        # txt 파일 목록 가져오기
        txt_files = sorted([f for f in os.listdir(txt_files_path) if f.endswith('.txt')])

        for txt_file in txt_files:
            file_path = os.path.join(txt_files_path, txt_file)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # 텍스트를 PDF 페이지에 추가
            c.setFont('NanumGothic', 12)
            textobject = c.beginText(40, height - 40)
            for line in content.split('\n'):
                textobject.textLine(line)
            c.drawText(textobject)

            c.showPage()  # 새 페이지 시작

        c.save()  # PDF 저장

        print(f"PDF 파일 '{pdf_file_name}'이 생성되었습니다.")
        return True
    except Exception as e:
        print(f"TXT를 PDF로 변환하는 중 오류 발생: {str(e)}")
        return False

def pdf_to_html(pdf_path, html_path):
    """
    PDF 파일을 HTML로 변환하는 함수
    
    :param pdf_path: PDF 파일 경로
    :param html_path: 저장할 HTML 파일 경로
    :return: 성공 시 True, 실패 시 False
    """
    try:
        doc = fitz.open(pdf_path)
        html_content = "<html><body>"

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("html")  # 페이지의 내용을 HTML로 추출
            html_content += text

        html_content += "</body></html>"

        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        print(f"PDF가 성공적으로 HTML로 변환되었습니다: {html_path}")
        return True
    except Exception as e:
        print(f"PDF를 HTML로 변환하는 중 오류 발생: {str(e)}")
        return False

def extract_image_from_pdf(pdf_path, output_folder):
    """
    PDF 파일에서 이미지를 추출하고 OCR을 수행하는 함수
    
    :param pdf_path: PDF 파일 경로
    :param output_folder: 이미지와 OCR 결과를 저장할 폴더 경로
    :return: 추출된 이미지 수, 실패 시 0
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    try:
        doc = fitz.open(pdf_path)
        image_count = 0
        reader = easyocr.Reader(['en', 'ko'])

        for page_num in range(len(doc)):
            page = doc[page_num]
            images = page.get_images()
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_path = os.path.join(output_folder, f"image_{page_num+1}_{img_index+1}.{image_ext}")
                
                with open(image_path, "wb") as image_file:
                    image_file.write(image_bytes)
                
                # OCR 수행
                res = reader.readtext(image_path, detail=0, paragraph=True)
                
                # OCR 결과를 텍스트 파일로 저장
                text_path = os.path.join(output_folder, f"ocr_text_{page_num+1}_{img_index+1}.txt")
                with open(text_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(res))
                
                image_count += 1
        
        print(f"{image_count}개의 이미지가 추출되고 OCR이 수행되었습니다.")
        return image_count
    except Exception as e:
        print(f"PDF에서 이미지 추출 및 OCR 수행 중 오류 발생: {str(e)}")
        return 0

def main():
    st.title("파일 파서 애플리케이션")

    menu = ["PPT to PDF", "PDF to Images", "Image Analysis", "OCR", "TXT to PDF", "PDF to HTML", "Extract Images from PDF"]
    choice = st.sidebar.selectbox("기능 선택", menu)

    if choice == "PPT to PDF":
        st.subheader("PPT를 PDF로 변환")
        uploaded_file = st.file_uploader("PPT 파일을 업로드하세요", type=["ppt", "pptx"])
        if uploaded_file:
            with st.spinner("PPT를 PDF로 변환 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                ppt_to_pdf(temp_file_path, output_pdf)
                st.success(f"PDF로 변환되었습니다: {output_pdf}")
                
                with open(output_pdf, "rb") as file:
                    st.download_button(
                        label="변환된 PDF 다운로드",
                        data=file,
                        file_name="converted.pdf",
                        mime="application/pdf"
                    )

    elif choice == "PDF to Images":
        st.subheader("PDF를 이미지로 변환")
        uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
        if uploaded_file:
            with st.spinner("PDF를 이미지로 변환 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_folder = tempfile.mkdtemp()
                image_paths = pdf_to_images(temp_file_path, output_folder)
                
                for image_path in image_paths:
                    st.image(Image.open(image_path), caption=os.path.basename(image_path))

    elif choice == "OCR":
        st.subheader("OCR 수행")
        uploaded_file = st.file_uploader("PDF 또는 이미지 파일을 업로드하세요", type=["pdf", "png", "jpg", "jpeg"])
        if uploaded_file:
            # 업로드된 파일 표시
            if uploaded_file.type.startswith('image'):
                st.image(uploaded_file, caption="업로드된 이미지", use_column_width=True)
            elif uploaded_file.type == "application/pdf":
                st.write("PDF 파일이 업로드되었습니다.")
                st.write(f"파일명: {uploaded_file.name}")
            else:
                st.write(f"업로드된 파일: {uploaded_file.name}")
            with st.spinner("OCR 수행 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_folder = tempfile.mkdtemp()
                if uploaded_file.type == "application/pdf":
                    pdf_to_image_ocr(temp_file_path, output_folder)
                else:
                    reader = easyocr.Reader(['en', 'ko'])
                    res = reader.readtext(temp_file_path, detail=0, paragraph=True)
                    txt_path = os.path.join(output_folder, "ocr_result.txt")
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(res))
                
                for file in os.listdir(output_folder):
                    if file.endswith('.txt'):
                        with open(os.path.join(output_folder, file), 'r', encoding='utf-8') as f:
                            st.text_area(f"OCR 결과 - {file}", f.read(), height=200)

    elif choice == "TXT to PDF":
        st.subheader("TXT를 PDF로 변환")
        uploaded_file = st.file_uploader("TXT 파일을 업로드하세요", type=["txt"])
        if uploaded_file:
            with st.spinner("TXT를 PDF로 변환 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                txt_to_pdf_convert(os.path.dirname(temp_file_path), output_pdf)
                st.success(f"PDF로 변환되었습니다: {output_pdf}")
                
                with open(output_pdf, "rb") as file:
                    st.download_button(
                        label="변환된 PDF 다운로드",
                        data=file,
                        file_name="converted.pdf",
                        mime="application/pdf"
                    )

    elif choice == "PDF to HTML":
        st.subheader("PDF를 HTML로 변환")
        uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
        if uploaded_file:
            with st.spinner("PDF를 HTML로 변환 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_html = tempfile.NamedTemporaryFile(delete=False, suffix=".html").name
                pdf_to_html(temp_file_path, output_html)
                st.success(f"HTML로 변환되었습니다: {output_html}")
                
                with open(output_html, "r", encoding="utf-8") as file:
                    st.download_button(
                        label="변환된 HTML 다운로드",
                        data=file.read(),
                        file_name="converted.html",
                        mime="text/html"
                    )

    elif choice == "Extract Images from PDF":
        st.subheader("PDF에서 이미지 추출")
        uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
        if uploaded_file:
            with st.spinner("PDF에서 이미지 추출 중..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_file_path = temp_file.name
                
                output_folder = tempfile.mkdtemp()
                extract_image_from_pdf(temp_file_path, output_folder)
                
                for file in os.listdir(output_folder):
                    if file.endswith(('.png', '.jpg', '.jpeg')):
                        st.image(Image.open(os.path.join(output_folder, file)), caption=file)
                    elif file.endswith('.txt'):
                        with open(os.path.join(output_folder, file), 'r', encoding='utf-8') as f:
                            st.text_area(f"OCR 결과 - {file}", f.read(), height=200)

if __name__ == "__main__":
    main()