import pandas as pd
import pandas_datareader.data as web
from datetime import datetime
import os

# ==========================================
# 1. 데이터 수집 및 연산 엔진 (Data Engine)
# ==========================================
def fetch_macro_data():
    """FRED에서 데이터를 수집하고 퀀트 지표를 계산하여 반환합니다."""
    start_date = '2000-01-01' 
    end_date = datetime.today().strftime('%Y-%m-%d')
    
    print(f"[{end_date}] FRED에서 매크로 데이터를 수집 중입니다...\n")
    
    tickers = ['DFF', 'DGS10', 'CPIAUCSL', 'GDPC1', 'GDPPOT']
    raw_data = web.DataReader(tickers, 'fred', start_date, end_date)
    
    # 지표 연산 (월간/분기 데이터를 일간 기준으로 변환)
    cpi = raw_data[['CPIAUCSL']].dropna()
    inflation = (cpi.pct_change(12) * 100).rename(columns={'CPIAUCSL': 'Inflation_YoY(%)'})
    
    gdp = raw_data[['GDPC1', 'GDPPOT']].dropna()
    gdp_growth = (gdp['GDPC1'].pct_change(4) * 100).rename('GDP_Growth_YoY(%)')
    gdp_growth_5y_avg = gdp_growth.rolling(20).mean().rename('GDP_Growth_5Y_Avg(%)')
    gdp_gap = (((gdp['GDPC1'] - gdp['GDPPOT']) / gdp['GDPPOT']) * 100).rename('GDP_Gap(%)')
    
    # 일간 금리 데이터
    daily_market = raw_data[['DFF', 'DGS10']].rename(columns={
        'DFF': 'Fed_Fund_Rate(%)', 
        'DGS10': 'US_Treasury_10Y(%)'
    })
    
    # 병합 및 Forward Fill (새 발표 전까지 과거 값 유지)
    df = daily_market.join([inflation, gdp_growth, gdp_growth_5y_avg, gdp_gap], how='outer')
    df = df.ffill().dropna()
    
    # 퀀트 파생 지표 계산
    df['Real_Rate_10Y(%)'] = df['US_Treasury_10Y(%)'] - df['Inflation_YoY(%)']
    df['Financial_Stance(r-g)'] = df['Real_Rate_10Y(%)'] - df['GDP_Growth_5Y_Avg(%)']
    
    target_inflation = 2.0
    df['Taylor_Target_Rate(%)'] = (
        df['GDP_Growth_5Y_Avg(%)'] + 
        df['Inflation_YoY(%)'] + 
        0.5 * (df['Inflation_YoY(%)'] - target_inflation) + 
        0.5 * df['GDP_Gap(%)']
    )
    df['Taylor_Gap(%)'] = df['Fed_Fund_Rate(%)'] - df['Taylor_Target_Rate(%)']
    
    # 날짜 인덱스 이름 정리 및 최신 날짜가 위로 오도록 내림차순 정렬
    df.index.name = 'Date'
    df = df.sort_index(ascending=False) 
    
    return df

# ==========================================
# 2. 콘솔 출력 엔진 (Dashboard Output)
# ==========================================
def print_dashboard(df):
    """가장 최신(오늘) 데이터를 바탕으로 텍스트 대시보드를 출력합니다."""
    # 내림차순 정렬되어 있으므로 0번째 행이 가장 최신 데이터
    latest = df.iloc[0]
    latest_date = df.index[0].strftime('%Y-%m-%d')
    
    print("="*60)
    print(f"📊 [퀀트 매크로 대시보드] 데이터 업데이트 기준일: {latest_date}")
    print("="*60)
    print(f"• [일간] 미 국채 10년물 금리 : {latest['US_Treasury_10Y(%)']:.2f}%")
    print(f"• [일간] 실제 유효 기준금리 : {latest['Fed_Fund_Rate(%)']:.2f}%")
    print(f"• [월간 반영] 현재 인플레이션 : {latest['Inflation_YoY(%)']:.2f}%")
    print(f"• [분기 반영] 5개년 평균 성장률(g): {latest['GDP_Growth_5Y_Avg(%)']:.2f}%")
    print(f"• [분기 반영] 현재 GDP 갭 추정   : {latest['GDP_Gap(%)']:.2f}%")
    print("-" * 60)
    
    # 1. 실질금리 vs 성장률
    print("💡 분석 1: 시장 실질금리 기반 환경 판단 (r - g)")
    print(f"  - 현재 시장 실질금리(r) : {latest['Real_Rate_10Y(%)']:.2f}%")
    print(f"  - 실질금리 - 성장률 격차: {latest['Financial_Stance(r-g)']:.2f}%p")
    if latest['Financial_Stance(r-g)'] > 0:
        print("  📢 결론: [긴축 국면] 자금 조달 비용이 경제 성장 탄력보다 높습니다.")
    else:
        print("  📢 결론: [완화 국면] 금융 환경이 완화적이며 투자가 장려되는 구간입니다.")
        
    print("-" * 60)
    
    # 2. 테일러 준칙 
    print("💡 분석 2: 테일러 준칙(Taylor's Rule) 기반 정책 판단")
    print(f"  - 모델 제시 적정 기준금리: {latest['Taylor_Target_Rate(%)']:.2f}%")
    print(f"  - 실제 기준금리와의 격차 : {latest['Taylor_Gap(%)']:.2f}%p")
    if latest['Taylor_Gap(%)'] > 0:
        print("  📢 결론: [준칙 대비 긴축] 연준이 경제 체력 대비 금리를 높게 유지하고 있습니다.")
    else:
        print("  📢 결론: [준칙 대비 완화] 연준이 경제 지표 대비 통화 정책을 느슨하게 펴고 있습니다.")
    print("="*60)

# ==========================================
# 3. 엑셀 출력 및 포맷팅 엔진 (Export Engine)
# ==========================================
def update_excel_db(df, file_name="Macro_Daily_DB.xlsx"):
    """데이터프레임을 엑셀로 저장하고 포맷팅합니다."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    file_path = os.path.join(output_dir, file_name)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Daily_Macro_Data')
        
        worksheet = writer.sheets['Daily_Macro_Data']
        worksheet.freeze_panes = "B2"
        
        for idx, col in enumerate(df.columns):
            excel_col = chr(66 + idx)
            column_length = max(len(col), 12)
            worksheet.column_dimensions[excel_col].width = column_length + 2
            
            for row in range(2, len(df) + 2):
                cell = worksheet[f"{excel_col}{row}"]
                cell.number_format = '0.00'

    print(f"\n✅ 엑셀 파일 업데이트 완료: {file_path}")

# ==========================================
# 4. 메인 실행 블록
# ==========================================
if __name__ == "__main__":
    try:
        # 1. 데이터 수집 및 계산
        macro_df = fetch_macro_data()
        
        # 2. 대시보드 요약 텍스트 출력
        print_dashboard(macro_df)
        
        # 3. 엑셀 파일 백업
        update_excel_db(macro_df)
        
    except Exception as e:
        print(f"\n❌ 오류가 발생했습니다: {e}")