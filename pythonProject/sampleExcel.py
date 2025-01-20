import pandas as pd

# 엑셀 파일 열기
file_path = "/pythonProject/sampledata.xlsx"  # 파일 이름
df = pd.read_excel(file_path, engine="openpyxl")

# 품목별로 묶어서 각 품목의 개수와 총 금액 계산
grouped_items = df.groupby("구매내역").agg(
    개수=("구매내역", "count"),
    총금액=("금액", "sum")
).reset_index()

print("품목별 개수와 총금액:")
print(grouped_items)

# 중복된 사업자등록번호 파악하기
duplicated_business_numbers = df["사업자등록번호"].duplicated(keep=False)
duplicated_data = df[duplicated_business_numbers].groupby("사업자등록번호").size()

# 중복된 데이터 테이블로 정리
duplicated_data_table = duplicated_data.reset_index(name="중복 횟수")
print("\n중복된 사업자 번호와 중복 횟수 (정리된 테이블):")
print(duplicated_data_table)

# 결과 저장
grouped_items.to_excel("grouped_items_result.xlsx", index=False)
duplicated_data_table.to_excel("duplicated_business_numbers_result.xlsx", index=False)
print("\n결과가 'grouped_items_result.xlsx'와 'duplicated_business_numbers_result.xlsx'에 저장되었습니다!")
