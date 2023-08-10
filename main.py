import requests
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import Workbook

while True:  # 무한 루프
    try:
        response_text = ""

        user_input = input("입력: ")
        user_input_enc = requests.utils.quote(user_input, encoding='EUC-KR')

        today = datetime.now().strftime("%Y/%m/%d").replace("/", "%2F")

        url = f"https://www.g2b.go.kr:8101/ep/tbid/tbidList.do?area=&bidNm={user_input_enc}&bidSearchType=1&fromBidDt={today}&fromOpenBidDt=&instNm=&maxPageViewNoByWshan=2&radOrgan=1&regYn=Y&searchDtType=1&searchType=1&taskClCds=&toBidDt={today}&toOpenBidDt=&currentPageNo=1"

        print(url)

        response = requests.get(url)
        document = BeautifulSoup(response.content, 'html.parser')

        rows = document.select("table tr")

        # 엑셀 파일 생성 및 시트 추가
        wb = Workbook()
        ws = wb.active
        ws.append(["공고 제목", "기관", "작성일", "세부링크", "사업금액"])

        data_added = False  # 데이터가 추가되었는지를 나타내는 플래그 변수

        for row in rows:
            columns = row.select("td div")

            if len(columns) >= 8:
                content4 = columns[3].text
                content4url = columns[3].select_one("a")["href"]
                content5 = columns[4].text
                content8 = columns[7].text

                detail_response = requests.get(content4url)
                document_detail = BeautifulSoup(detail_response.content, 'html.parser')

                detail_table = document_detail.find("table", summary="예정가격 결정 및 입찰금액 정보")

                if detail_table:
                    detail_rows = detail_table.find_all("tr")[1]
                    cash_div = detail_rows.find("td").find("div")
                    cash = cash_div.text.strip() if cash_div else ""

                    try:
                        cash_value = int(''.join(filter(str.isdigit, cash)))
                    except ValueError:
                        cash_value = 0

                    if cash_value > 149999999:
                        # 엑셀 시트에 데이터 추가
                        ws.append([content4, content5, content8, content4url, cash])
                        print(f"추가된 데이터: {content4} - {content5} - {content8} - {content4url} - {cash}")
                        data_added = True  # 데이터가 추가되었음을 표시

        if data_added:
            # 엑셀 파일 저장
            current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

            excel_filename = f"g2b_{user_input}_{current_time}.xlsx"
            wb.save(excel_filename)
            print(f"데이터가 {excel_filename} 파일에 저장되었습니다.")
        else:
            print("조회된 데이터가 없습니다.")

        # again = input("프로그램을 다시 실행하시겠습니까? (Y/N): ")
        # if again.lower() != "y":
        #     break  # 루프 종료

    except Exception as e:
        raise RuntimeError(str(e))
