import startsort2
import pandas as pd


def searchcategory(df, categoryLevel, category_name, startIdx=0, endIdx=0):
    oriStartIdx = startIdx
    oriEndIdx = endIdx
    if startIdx == 0 and endIdx == 0:
        search_range = df.index  # 전체 행을 검사
    else:
        search_range = range(startIdx, endIdx + 1)  # 지정된 범위의 행을 검사
    for idx in search_range:
        exact_match = df.loc[idx, categoryLevel]
        if exact_match == category_name:
            if startIdx == oriStartIdx:
                startIdx = idx  # 첫 번째 일치하는 행
            else:
                endIdx = idx  # 두 번째 이후 일치하는 행의 경우 마지막 행으로 업데이트
    if startIdx == oriStartIdx and endIdx == oriEndIdx:
        # 완전 일치하는 행이 없을 경우 단어가 포함된 행을 찾음
        for idx in search_range:
            if str(category_name) in str(df.loc[idx, categoryLevel]):
                if startIdx == oriStartIdx:
                    startIdx = idx  # 첫 번째 단어 포함된 행
                else:
                    endIdx = idx  # 두 번째 이후 단어 포함된 행의 경우 마지막 행으로 업데이트

    if startIdx == oriStartIdx and endIdx == oriEndIdx:
        print(f"입력된 카테고리명 '{category_name}'에 해당하는 정보를 찾을 수 없습니다.")
        return oriStartIdx, oriEndIdx
    if startIdx != oriStartIdx and endIdx == oriEndIdx:
        endIdx = startIdx
    return startIdx, endIdx

def IsValid(startIdx, endIdx):
    if startIdx != 0 and endIdx != 0:
        if startIdx == endIdx:
            return True
        else:
            return False
    else:
        return False

def finCategory(df, startIdx):
    try:
        # startIdx 행에 해당하는 '최종 카테고리' 열의 값을 가져옵니다.
        coupang_value = df.loc[startIdx, '최종 카테고리']
        return coupang_value
    except Exception as e:
        print(f"행 인덱스 '{startIdx}'의 '최종 카테고리' 값을 가져오는 도중 오류가 발생했습니다:", e)
        return None  # 오류 발생 시 None을 반환합니다.

def compareCategory(category_info):
    file_name = "초대형 마켓카테고리 완성본(내가 바꾼거).xlsx"
    df = pd.read_excel(file_name)
    results = []
    status = -1
    
    for category_number in category_info.keys():
        start_index = 0
        end_index = 0
        status = -1
        
        for 대카테고리명 in category_info[category_number]['대카테고리명']:
            if status != -1:
                break
            categoryLevel = "1단계 카테고리명"
            start_index, end_index = searchcategory(df, categoryLevel, 대카테고리명, start_index, end_index)
            
            for 중카테고리명 in category_info[category_number]['중카테고리명']:
                if status != -1:
                    break
                categoryLevel = "2단계 카테고리명"
                start_index, end_index = searchcategory(df, categoryLevel, 중카테고리명, start_index, end_index)
                twoLevelStartIdx = start_index
                twoLevelEndIdx = end_index
                
                for 소카테고리명 in category_info[category_number]['소카테고리명']:
                    categoryLevel = "3단계 카테고리명"
                    start_index, end_index = searchcategory(df, categoryLevel, 소카테고리명, start_index, end_index)
                    
                    if IsValid(start_index, end_index):
                        status = 1
                        results.append(finCategory(df, start_index))
                        break
                    
                    categoryLevel = "4단계 카테고리명"
                    start_index, end_index = searchcategory(df, categoryLevel, 소카테고리명, start_index, end_index)
                    
                    if IsValid(start_index, end_index):
                        status = 1
                        results.append(finCategory(df, start_index))
                        break
                    
                    start_index = 0
                    end_index = 0
                    start_index, end_index = searchcategory(df, categoryLevel, 소카테고리명, start_index, end_index)
                    
                    if IsValid(start_index, end_index):
                        status = 1
                        results.append(finCategory(df, start_index))
                        break
                    
                    start_index = twoLevelStartIdx
                    end_index = twoLevelEndIdx
        if status == -1:
            results.append("None")
    return results

def write_coupang_column2(values):
    try:
        output_file_name = "카테고리번호정리완료.xlsx"
        
        # 새로운 DataFrame을 생성하여 카테고리 값을 추가합니다.
        new_df = pd.DataFrame({'최종 카테고리': values}, index=range(1, len(values) + 1))
        
        # DataFrame을 엑셀 파일로 저장합니다.
        new_df.to_excel(output_file_name)
        
        print("카테고리 열이 성공적으로 추가되었습니다.")
    except Exception as e:
        print(f"파일 '{output_file_name}'을 읽어오는 도중 오류가 발생했습니다:", e)

def main():
    category_info = startsort2.main()
    value = compareCategory(category_info)
    write_coupang_column2(value)
    return 0
    
if __name__ == "__main__":
    main()

