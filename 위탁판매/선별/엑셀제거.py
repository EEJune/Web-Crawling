import pandas as pd

def remove_matching_rows(a_file, b_file, output_file):
    try:
        # A엑셀 읽기
        a_df = pd.read_excel(a_file, usecols=["상품명*"])
        a_product_names = a_df["상품명*"].dropna().tolist()  # "상품명*" 열에서 NaN 제거 후 리스트화

        # B엑셀 읽기
        b_df = pd.read_excel(b_file)

        # B엑셀에서 A엑셀의 상품명을 제외한 행만 유지
        filtered_b_df = b_df[~b_df["상품명*"].isin(a_product_names)]

        # 결과 저장
        filtered_b_df.to_excel(output_file, index=False)
        print(f"일치하는 행을 제거한 결과가 '{output_file}'에 저장되었습니다!")

    except FileNotFoundError as e:
        print(f"[ERROR] 파일을 찾을 수 없습니다: {e}")
    except KeyError as e:
        print(f"[ERROR] 엑셀에 '상품명*' 열이 없습니다: {e}")
    except Exception as e:
        print(f"[ERROR] 처리 중 오류 발생: {e}")

# 실행
if __name__ == "__main__":
    a_file = input("비교대상 엑셀 파일명을 입력하세요 (예: A.xlsx): ").strip()
    b_file = input("제거할 원본 엑셀 파일명을 입력하세요 (예: B.xlsx): ").strip()
    output_file = input("결과를 저장할 파일명을 입력하세요 (예: 결과.xlsx): ").strip()

    remove_matching_rows(a_file, b_file, output_file)
