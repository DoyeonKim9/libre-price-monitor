import urllib.request
import json
import pandas as pd
import re
import time
import requests
import asyncio
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from datetime import datetime

from no import Product, render_card_to_png

# ==========================================
# 1. API 키 설정
# ==========================================
CLIENT_ID = "NXeGJZXkxK8ZyE4l4bsR"
CLIENT_SECRET = "9c5ZGASXBK"

# ==========================================
# 2. 분석 함수 (수량 및 단가 계산) - [수정됨]
# ==========================================
def analyze_product(title, total_price):
    clean_title = title
    
    # 함정 단어 제거
    black_list = [
        r"아메리카노\s*\d+개", r"커피\s*\d+잔", r"커피\s*\d+개",
        r"패치\s*\d+매", r"패치\s*\d+개", r"알콜솜\s*\d+매",
        r"방수필름\s*\d+매", r"멤버십\s*\d+일", r"유효기간\s*\d+일",
        r"\d+일\s*체험", r"\d+일\s*멤버십"
    ]
    for pattern in black_list:
        clean_title = re.sub(pattern, " ", clean_title)

    # 수량 추출
    qty_candidates = []
    matches = re.findall(r"[\sxX](\d+)\s*(개|세트|팩|박스|ea|set)", clean_title, re.IGNORECASE)
    for m in matches: qty_candidates.append(int(m[0]))
    matches_mul = re.findall(r"[xX*]\s*(\d+)", clean_title)
    for m in matches_mul: qty_candidates.append(int(m))

    extracted_qty = qty_candidates[-1] if qty_candidates else 1

    # 가격 검증 (리브레2 기준가 6.5만 ~ 13만)
    MIN_PRICE, MAX_PRICE = 65000, 130000
    calc_unit_price = total_price // extracted_qty
    
    # [수정 포인트] return 값 뒤에 clean_title을 추가해서 4개를 맞췄습니다.
    if MIN_PRICE <= calc_unit_price <= MAX_PRICE:
        return extracted_qty, calc_unit_price, "텍스트분석", clean_title
    else:
        estimated_qty = round(total_price / 90000)
        if estimated_qty == 0: estimated_qty = 1
        recalc_price = total_price // estimated_qty
        
        if MIN_PRICE <= recalc_price <= MAX_PRICE:
            return estimated_qty, recalc_price, "가격역산(보정)", clean_title
        else:
            return extracted_qty, calc_unit_price, "확인필요", clean_title

# ==========================================
# 3. 데이터 수집 함수 (이미지 포함)
# ==========================================
def get_naver_shopping_data_with_image(query, max_items=50):
    print(f"🔍 '{query}' 데이터 수집 중 (목표: {max_items}개)...")
    
    encText = urllib.parse.quote(query)
    all_results = []
    start = 1
    display = 100 
    
    # [중요] API 정렬은 'sim'(정확도순)으로 둡니다.
    # 'asc'(가격순)으로 하면 500원짜리 케이스만 잔뜩 가져오기 때문입니다.
    # 진짜 정렬은 아래에서 파이썬으로 직접 합니다.
    
    while start < max_items:
        if start > 1000: break
        
        url = f"https://openapi.naver.com/v1/search/shop.json?query={encText}&display={display}&start={start}&sort=sim"
        request = urllib.request.Request(url)
        request.add_header("X-Naver-Client-Id", CLIENT_ID)
        request.add_header("X-Naver-Client-Secret", CLIENT_SECRET)

        try:
            response = urllib.request.urlopen(request)
            if response.getcode() == 200:
                data = json.loads(response.read().decode('utf-8'))
                items = data['items']
                if not items: break

                for item in items:
                    raw_title = item['title'].replace("<b>", "").replace("</b>", "")
                    total_price = int(item['lprice'])
                    image_url = item['image'] 
                    
                    qty, unit_price, method, _ = analyze_product(raw_title, total_price)
                    
                    all_results.append({
                        "이미지": "", 
                        "이미지URL": image_url, 
                        "상품명": raw_title,
                        "총 가격": total_price,
                        "수량": qty,
                        "개당 단가": unit_price,
                        "계산방식": method,
                        "판매처": item['mallName'],
                        "링크": item['link']
                    })
                start += display
                time.sleep(0.1)
            else: break
        except Exception as e:
            print(f"API 에러: {e}")
            break
            
    return all_results

# ==========================================
# 4. 이미지 삽입 함수 (엑셀 후처리)
# ==========================================
def insert_images_to_excel(filename, df):
    print("🖼️ 엑셀에 이미지를 넣는 중... (잠시만 기다려주세요)")
    wb = load_workbook(filename)
    ws = wb.active
    ws.column_dimensions['A'].width = 15 
    
    for index, row in df.iterrows():
        excel_row = index + 2
        img_url = row['이미지URL']
        
        if img_url:
            try:
                res = requests.get(img_url)
                img_data = BytesIO(res.content)
                img = ExcelImage(img_data)
                img.width, img.height = 100, 100
                ws.add_image(img, f"A{excel_row}")
                ws.row_dimensions[excel_row].height = 80
            except: continue

    wb.save(filename)
    print("✨ 이미지 삽입 완료!")


# ==========================================
# 5. 카드 이미지 생성 (no.py 연동)
# ==========================================
async def generate_card_images_for_df(df, out_dir="product_cards"):
    """
    no.py의 Product / render_card_to_png를 활용해
    DataFrame의 각 상품에 대한 카드 이미지를 생성하고 경로를 반환합니다.
    """
    results = []

    for idx, row in df.iterrows():
        try:
            product = Product(
                platform="naver",
                name=str(row["상품명"]),
                price=int(row["개당 단가"]),
                url=str(row["링크"]),
                image_url=str(row["이미지URL"]) if pd.notna(row["이미지URL"]) else None,
            )
            card_path = await render_card_to_png(product, out_dir)
            results.append((idx, card_path))
            print(f"[CARD OK] {row['상품명'][:25]}... -> {card_path}")
        except Exception as e:
            print(f"[CARD ERROR] {row.get('상품명', '')}: {e}")

    return results

# ==========================================
# 6. 실행 및 정렬 설정 + 카드 이미지/CSV 저장
# ==========================================
if __name__ == "__main__":
    keyword = "프리스타일 리브레2"
    
    # 50개만 수집 (테스트용, 원하면 늘리세요)
    data = get_naver_shopping_data_with_image(keyword, max_items=50)

    if data:
        df = pd.read_json(json.dumps(data))
        
        # ---------------------------------------------------------
        # [핵심] 1. 액세서리 필터링 (너무 싼 건 가짜다!)
        # 개당 단가가 50,000원보다 싼 건 리스트에서 지워버립니다.
        # ---------------------------------------------------------
        print(f"🧹 필터링 전: {len(df)}개 -> 액세서리(5만원 이하) 제거 중...")
        df = df[df['개당 단가'] >= 50000]
        print(f"✨ 필터링 후: {len(df)}개 남음 (진짜 센서만)")

        # ---------------------------------------------------------
        # [핵심] 2. 낮은 가격 순으로 정렬
        # ascending=True가 '오름차순(낮은 게 위로)' 입니다.
        # ---------------------------------------------------------
        df = df.sort_values(by='개당 단가', ascending=True)
        
        # ---------------------------------------------------------
        # 3. 카드 이미지 생성 (no.py 연동)
        # ---------------------------------------------------------
        print("\n🖼 상품 카드 이미지 생성 중... (Playwright)")
        card_results = asyncio.run(generate_card_images_for_df(df, out_dir="product_cards"))

        # 카드 이미지 경로 컬럼 추가
        df["카드이미지경로"] = ""
        for idx, path in card_results:
            df.at[idx, "카드이미지경로"] = path

        # 엑셀/CSV 저장
        filename_base = f"리브레_최저가순_{datetime.now().strftime('%H%M')}"
        xlsx_filename = f"{filename_base}.xlsx"
        csv_filename = f"{filename_base}.csv"

        # 컬럼 순서 (이미지, 단가, 수량 순으로 보기 좋게)
        cols = ['이미지', '상품명', '개당 단가', '수량', '총 가격', '판매처', '계산방식', '링크', '이미지URL', '카드이미경로']
        # 오타 수정: '카드이미경로' -> '카드이미지경로' 가 컬럼명과 일치해야 하므로 조정
        cols = ['이미지', '상품명', '개당 단가', '수량', '총 가격', '판매처', '계산방식', '링크', '이미지URL', '카드이미지경로']
        df = df[cols]

        # 엑셀 저장 + 원본 이미지 삽입
        df.to_excel(xlsx_filename, index=False)
        insert_images_to_excel(xlsx_filename, df)

        # CSV 저장 (이미지 삽입 없이 경로/데이터만)
        df.to_csv(csv_filename, index=False, encoding="utf-8-sig")

        print(f"\n💾 저장 완료: {xlsx_filename} / {csv_filename}")
        print("\n🏆 [가장 싼 상품 TOP 5 미리보기]")
        print("-" * 60)
        # 화면에 미리보기 출력
        for i, row in df.head(5).iterrows():
            print(f"{i+1}등: {row['개당 단가']:,}원 | {row['상품명'][:30]}...")
