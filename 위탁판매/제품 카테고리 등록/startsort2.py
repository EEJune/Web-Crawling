import pandas as pd
from openpyxl import load_workbook

def read_category_info_from_sheet(file_name, sheet_name, category_numbers, row_numbers):
    try:
        # 엑셀 파일의 특정 시트를 읽어옵니다.
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        
        # 각 카테고리 번호에 대해 해당하는 행의 정보를 저장할 딕셔너리를 초기화합니다.
        category_info = {}
        
        # 카테고리 번호가 일치하는 행의 정보를 딕셔너리에 저장합니다.
        for category_number, row_number in zip(category_numbers, row_numbers):
            selected_row = df[df['카테고리번호'] == category_number]
            if not selected_row.empty:
                category_info[row_number] = {
                    '카테고리번호': selected_row['카테고리번호'].iloc[0],
                    '대카테고리명': selected_row['대카테고리명'].iloc[0],
                    '중카테고리명': selected_row['중카테고리명'].iloc[0],
                    '소카테고리명': selected_row['소카테고리명'].iloc[0],
                    '세카테고리명': selected_row['세카테고리명'].iloc[0] if '세카테고리명' in df.columns else None,
                    '행번호': row_number
                }
        
        return category_info
    except Exception as e:
        print(f"{sheet_name} 시트를 읽어오는 도중 오류가 발생했습니다:", e)
        return None

def read_category_numbers_and_product_names_from_excel(file_name, sheet_name):
    try:
        # 엑셀 파일의 특정 시트를 읽어옵니다.
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        
        # 카테고리 번호* 열과 상품명* 열의 값을 읽어 배열에 저장합니다.
        category_numbers = df.iloc[1:]['카테고리 번호*'].tolist()  # 3행부터 시작
        product_names = df.iloc[1:]['상품명*'].tolist()  # 3행부터 시작
        
        # 행 번호도 함께 저장합니다.
        row_numbers = df.index[1:].tolist()  # 3행부터 시작
        
        return category_numbers, product_names, row_numbers
    except Exception as e:
        print(f"{sheet_name} 시트를 읽어오는 도중 오류가 발생했습니다:", e)
        return None, None, None

def modify_category_name(category_info, product_name):
    try:
        # 입력으로 받은 category_info가 올바른 형식인지 확인
        if isinstance(category_info, dict):
            for category_number, info in category_info.items():
                # 소카테고리명을 가져옵니다.
                category_name = info.get('소카테고리명')
                if category_name:
                    # 소카테고리명을 단어로 분할하여 리스트로 변환합니다.
                    category_name_list = category_name.split('/')
                    # 상품명을 단어로 분할하여 리스트로 변환합니다.
                    product_name_list = product_name.split()
                    # 상품명에 존재하는 소카테고리명의 단어를 찾아서 리스트에 저장합니다.
                    matching_words = [word for word in category_name_list if word in product_name_list]
                    # 일치하는 단어들을 다시 문자열로 결합하여 소카테고리명으로 설정합니다.
                    modified_category_name = ' '.join(matching_words)
                    if modified_category_name:
                        # 수정된 소카테고리명을 category_info에 반영합니다.
                        category_info[category_number]['소카테고리명'] = modified_category_name
        else:
            print("오류 발생: 입력으로 받은 category_info가 올바른 형식이 아닙니다.")
        
        return category_info
    except Exception as e:
        print("오류 발생:", e)
        return None

def print_category_info(category_info, product_names):
    if category_info is not None:
        print("카테고리 정보:")
        i=0
        for category_number, info in category_info.items(): 
            print("제품명: ",product_names[i])
            print(f"행번호: {info["행번호"]}")
            print(f"카테고리번호: {info['카테고리번호']}")
            print(f"대카테고리명: {info['대카테고리명']}")
            print(f"중카테고리명: {info['중카테고리명']}")
            print(f"소카테고리명: {info['소카테고리명']}")
            print(f"세카테고리명: {info['세카테고리명']}")
            print()
            i+=1
    else:
        print("카테고리 정보가 없습니다.")


def cutCategory(category_info):
    for category_number, info in category_info.items():
        for key, value in info.items():
            if isinstance(value, str) and "/" in value:
                # 문자열인 경우 "/"로 분할하여 딕셔너리 형태로 다시 저장합니다.
                category_info[category_number][key] = value.split("/")
            else:
                # 그 외의 경우에는 리스트로 변경하여 저장합니다.
                category_info[category_number][key] = [value] if value else []
    return category_info

def main():
    file_name = "원본상품양식.xlsx"
    basic_info_sheet_name = "기본정보"
    standard_category_sheet_name = "이셀러스표준카테고리"
    
    # 카테고리 번호* 값을 읽어옵니다.
    category_numbers, product_names, row_numbers = read_category_numbers_and_product_names_from_excel(file_name, basic_info_sheet_name)
    
    # "이셀러스표준카테고리" 시트로 넘어가서 작업을 계속합니다.
    category_info = read_category_info_from_sheet(file_name, standard_category_sheet_name, category_numbers, row_numbers)
    if category_info is not None:
        #print(f"카테고리 번호 {category_numbers}에 대한 정보:")
        # 상품명*과 소카테고리명을 비교하여 일치하는 단어를 찾고 수정된 소카테고리명을 반환합니다.
        for product_name, category_number in zip(product_names, category_numbers):
            modify_category_name(category_info, product_name)
        cutCategory(category_info)
        print_category_info(category_info,product_names)
    return category_info
        
if __name__ == "__main__":
    main()
#오류시 엑셀파일 내 카테고리번호 숫자로 변환