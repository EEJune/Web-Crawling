import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import random
import sys  # 프로그램 종료를 위해 추가
import urllib.parse  # URL 인코딩을 위한 라이브러리


# Selenium을 이용한 쿠팡 검색 함수
def search_coupang(product_name, search_price):
    # URL 인코딩 처리
    encoded_product_name = urllib.parse.quote(product_name)  # URL 인코딩 적용
    search_url = f"https://www.coupang.com/np/search?component=&q={encoded_product_name}&channel=user"

    try:
        # Chrome 옵션 설정
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # 브라우저 비표시 모드
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        )

        # WebDriver 설정
        driver_service = Service(r"c:\chromedriver-win64\chromedriver-win64\chromedriver.exe")  # Chromedriver 경로
        driver = webdriver.Chrome(service=driver_service, options=chrome_options)

        # 검색 페이지 열기
        driver.get(search_url)
        time.sleep(5)  # 페이지 로딩 대기

        # 결과 탐색
        products = driver.find_elements(By.CSS_SELECTOR, "li.search-product")
        print(f"[DEBUG] 검색 URL: {search_url}")
        print(f"[DEBUG] 검색된 상품 수: {len(products)}")

        # 상품이 없는 경우
        if len(products) == 0:
            driver.quit()
            return "zero"  # 상품이 없는 경우 "zero" 반환

        prices = []
        for product in products:
            try:
                # 제품명과 가격 추출
                product_name_element = product.find_element(By.CSS_SELECTOR, "div.name")
                product_name_text = product_name_element.text.strip()

                # 제품명이 포함된 경우에만 처리
                if product_name in product_name_text:
                    price_element = product.find_element(By.CSS_SELECTOR, "strong.price-value")
                    price_text = price_element.text.strip().replace(",", "")  # 쉼표 제거
                    price = int(price_text)
                    prices.append(price)

                    # 디버깅용 출력
                    print(f"[DEBUG] 상품명: {product_name_text}, 가격: {price}")
            except Exception as e:
                # 요소 탐색 중 문제가 발생해도 다음으로 진행
                print(f"[ERROR] 상품 정보 처리 중 문제 발생: {e}")
                continue

        driver.quit()

        # 가격 비교
        for price in prices:
            if price <= search_price:
                return "x"  # 검색가보다 낮은 가격이 있으면 "x"
        return None  # 모든 가격이 검색가보다 높으면 결과 없음

    except Exception as e:
        print(f"[ERROR] Selenium 검색 중 오류 발생: {e}")
        return "error"  # 오류 발생 시 "error" 반환


# 메인 코드
def main():
    input_file = input("원본 엑셀 파일명을 입력하세요 (예: 원본_파일.xlsx): ").strip()

    try:
        # 엑셀 데이터 읽기
        df = pd.read_excel(input_file, usecols=[1, 2, 3], skiprows=1, header=None)
        df.columns = ["상품명*", "판매가(1.8배)", "검색가(1.2배)"]

        result_data = []

        for index, row in df.iterrows():
            product_name = row["상품명*"]
            max_price = int(row["판매가(1.8배)"])  # 참고용
            search_price = int(row["검색가(1.2배)"])  # 비교 기준 가격

            print(f"[DEBUG] 검색 제품명: {product_name}, 검색가(1.2배): {search_price}")
            result = search_coupang(product_name, search_price)

            # 검색된 상품 수가 0인 경우
            if result == "zero":
                print(f"[INFO] '{product_name}'에 대해 검색된 상품이 없습니다. 프로그램을 종료합니다.")

                # zero가 발생한 행까지 기록
                result_data.append([product_name, max_price, search_price, "zero"])
                output_df = pd.DataFrame(result_data, columns=["상품명*", "판매가(1.8배)", "검색가(1.2배)", "결과"])
                output_file = "쿠팡_검색_결과.xlsx"
                output_df.to_excel(output_file, index=False)

                print(f"결과가 {output_file}에 저장되었습니다!")
                sys.exit(0)  # 프로그램 종료

            # "x"나 "error"만 기록
            if result in ["x", "error"]:
                result_data.append([product_name, max_price, search_price, result])

            # 요청 간 대기 시간 추가
            time.sleep(random.uniform(23, 29))  # 매크로 방지용 지연 시간

        # 모든 작업 완료 후 결과 저장
        output_df = pd.DataFrame(result_data, columns=["상품명*", "판매가(1.8배)", "검색가(1.2배)", "결과"])
        output_file = "쿠팡_검색_결과.xlsx"
        output_df.to_excel(output_file, index=False)

        print(f"결과가 {output_file}에 저장되었습니다!")

    except FileNotFoundError: 
        print(f"[ERROR] 파일 {input_file}을(를) 찾을 수 없습니다.")
    except Exception as e:
        print(f"[ERROR] 처리 중 알 수 없는 오류 발생: {e}")


# 실행
if __name__ == "__main__":
    main()
