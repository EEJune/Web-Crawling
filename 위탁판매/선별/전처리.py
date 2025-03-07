import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time

# Selenium을 이용한 상품 가격 검색 함수
def search_product_price_by_code(management_code):
    search_url = f"https://ownerclan.com/V2/product/search.php?topSearchKeywordInfo=&topSearchKeyword={management_code}&topSearchType=all"
    
    try:
        # Chrome 옵션 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # 브라우저를 표시하지 않음
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        # WebDriver 경로 설정
        driver_service = Service(r'c:\chromedriver-win64\chromedriver-win64\chromedriver.exe')  # chromedriver 경로
        driver = webdriver.Chrome(service=driver_service, options=chrome_options)
        
        # URL로 이동
        driver.get(search_url)
        time.sleep(2)  # 페이지 로딩 대기
        
        # 가격 요소 찾기
        try:
            price_element = driver.find_element(By.CSS_SELECTOR, "p.price2 > span.won_color1")
            price_text = price_element.text.strip().replace(',', '')  # 쉼표 제거
            return int(price_text)
        except Exception as e:
            print(f"[DEBUG] 가격 요소를 찾지 못했습니다: {e}")
            return None
        finally:
            driver.quit()
    
    except Exception as e:
        print(f"[ERROR] Selenium 검색 중 오류 발생: {e}")
        return None

# 메인 코드
def main():
    file_path = input("원본 엑셀 파일명을 입력하세요 (예: 원본_파일.xlsx): ").strip()

    try:
        # 엑셀 파일 읽기 (3행부터 데이터, B열과 E열만 추출)
        df = pd.read_excel(file_path, usecols=[1, 4], skiprows=1)  # B열(1), E열(4)을 읽음
        
        # 열 이름 설정
        df.columns = ["판매자 관리코드", "상품명*"]
        
        # 결과 저장을 위한 리스트
        processed_data = []

        for _, row in df.iterrows():
            management_code = row["판매자 관리코드"]
            product_name = row["상품명*"]
            
            try:
                original_price = search_product_price_by_code(management_code)
                if original_price is not None:
                    adjusted_price = original_price * 1.8  # 원가의 1.7배
                    search_price = original_price * 1.2
                    processed_data.append([management_code, product_name, adjusted_price, search_price])
                else:
                    processed_data.append([management_code, product_name, "X", "X"])  # 결과 없음
            except Exception as e:
                print(f"[ERROR] {management_code} 처리 중 오류: {e}")
                processed_data.append([management_code, product_name, "X", "X"])  # 에러 시도 "X" 처리

        # 새로운 엑셀 파일로 저장
        output_df = pd.DataFrame(processed_data, columns=["판매자 관리코드", "상품명*", "판매가(1.8배)", "검색가(최저가 1.2배)"])
        output_file_path = "전처리_완료.xlsx"
        output_df.to_excel(output_file_path, index=False)

        print(f"결과가 {output_file_path}에 저장되었습니다!")

    except FileNotFoundError:
        print(f"[ERROR] 파일 {file_path}을(를) 찾을 수 없습니다.")
    except Exception as e:
        print(f"[ERROR] 처리 중 알 수 없는 오류 발생: {e}")

# 실행
if __name__ == "__main__":
    main()
