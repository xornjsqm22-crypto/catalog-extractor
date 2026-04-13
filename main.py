import ollama
import openpyxl
import base64
import fitz  # pymupdf
import os
from pathlib import Path

def image_to_base64(image_path):
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')

def pdf_to_images(pdf_path, output_folder):
    doc = fitz.open(pdf_path)
    image_paths = []
    os.makedirs(output_folder, exist_ok=True)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        img_path = f"{output_folder}/page_{i+1}.jpg"
        pix.save(img_path)
        image_paths.append(img_path)
    return image_paths

def extract_spec(image_path):
    image_data = image_to_base64(image_path)
    prompt = """
    이 카탈로그 이미지에서 창호 제품 스펙을 추출해줘.
    스펙이 없는 페이지면 "스펙없음" 이라고만 답해.
    스펙이 있으면 아래 형식으로만 답해:
    
    제조사: 
    시리즈명: 
    개폐방식: 
    프로파일깊이(mm): 
    유리종류: 
    유리두께(T): 
    Uf(W/m²K): 
    Uw(W/m²K): 
    기밀등급: 
    방범성능: 
    차음성능(dB): 
    최대폭(mm): 
    최대높이(mm): 
    최소폭(mm): 
    최소높이(mm): 
    색상옵션: 
    방화여부: 
    비고: 
    """
    response = ollama.chat(
        model='gemma3',
        messages=[{
            'role': 'user',
            'content': prompt,
            'images': [image_data]
        }]
    )
    return response['message']['content']

def save_to_excel(spec_text, output_path='창호DB.xlsx'):
    if '스펙없음' in spec_text:
        print('  → 스펙 없는 페이지, 건너뜀')
        return

    try:
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
    except:
        wb = openpyxl.Workbook()
        ws = wb.active
        headers = [
            '제조사', '시리즈명', '개폐방식', '프로파일깊이(mm)',
            '유리종류', '유리두께(T)', 'Uf(W/m²K)', 'Uw(W/m²K)',
            '기밀등급', '방범성능', '차음성능(dB)',
            '최대폭(mm)', '최대높이(mm)', '최소폭(mm)', '최소높이(mm)',
            '색상옵션', '방화여부', '비고'
        ]
        ws.append(headers)

    data = {}
    for line in spec_text.strip().split('\n'):
        if ':' in line:
            key, _, value = line.partition(':')
            data[key.strip()] = value.strip()

    row = [
        data.get('제조사', ''),
        data.get('시리즈명', ''),
        data.get('개폐방식', ''),
        data.get('프로파일깊이(mm)', ''),
        data.get('유리종류', ''),
        data.get('유리두께(T)', ''),
        data.get('Uf(W/m²K)', ''),
        data.get('Uw(W/m²K)', ''),
        data.get('기밀등급', ''),
        data.get('방범성능', ''),
        data.get('차음성능(dB)', ''),
        data.get('최대폭(mm)', ''),
        data.get('최대높이(mm)', ''),
        data.get('최소폭(mm)', ''),
        data.get('최소높이(mm)', ''),
        data.get('색상옵션', ''),
        data.get('방화여부', ''),
        data.get('비고', '')
    ]
    ws.append(row)
    wb.save(output_path)
    print(f'  ✅ 저장완료')

def process_folder(folder_path):
    folder = Path(folder_path)
    files = list(folder.glob('*.pdf')) + \
            list(folder.glob('*.jpg')) + \
            list(folder.glob('*.jpeg')) + \
            list(folder.glob('*.png'))

    print(f'\n총 {len(files)}개 파일 처리 시작\n')

    for file in files:
        print(f'처리중: {file.name}')
        if file.suffix.lower() == '.pdf':
            temp_folder = f"temp_{file.stem}"
            images = pdf_to_images(str(file), temp_folder)
            for i, img in enumerate(images):
                print(f'  페이지 {i+1}/{len(images)} 분석중...')
                spec = extract_spec(img)
                save_to_excel(spec)
        else:
            spec = extract_spec(str(file))
            save_to_excel(spec)

    print('\n🎉 전체 처리 완료!')

if __name__ == '__main__':
    print('=== 창호 카탈로그 자동 추출 ===')
    print('1. 단일 파일 (PDF/이미지)')
    print('2. 폴더 전체 처리')
    choice = input('\n선택 (1 or 2): ')

    if choice == '1':
        path = input('파일 경로 입력: ')
        if path.endswith('.pdf'):
            images = pdf_to_images(path, 'temp_pdf')
            for i, img in enumerate(images):
                print(f'페이지 {i+1} 분석중...')
                spec = extract_spec(img)
                save_to_excel(spec)
        else:
            spec = extract_spec(path)
            save_to_excel(spec)
    elif choice == '2':
        folder = input('폴더 경로 입력: ')
        process_folder(folder)

    print('\n창호DB.xlsx 저장 완료!')