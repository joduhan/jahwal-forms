# 제주광역자활센터 업무 자동화 시스템

## 개요
2025년도 중앙자활자금 개발원·광역 지원사업 서식 자동 생성 시스템

## 실행 방법
1. 패키지 설치: pip install -r requirements.txt
2. 웹 UI 실행: streamlit run app.py
3. CLI 실행: python main.py

## 폴더 설명
- templates/: 서식 DOCX 템플릿
- outputs/: 생성된 문서 저장
- data/: 입력 Excel/CSV 및 HWP 원본
- modules/: 기능별 Python 모듈

## 생성 가능한 서식
### 공통 서식
- 표준협약서
- 개인정보 수집·이용 동의서
- 청렴이행각서

### 개별 서식 (창업지원)
- 사업신청서
- 중간집행정산보고
