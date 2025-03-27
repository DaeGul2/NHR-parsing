# 📊 Excel 기반 이력서/자기소개서 정제 및 세로화 도구

이 프로젝트는 엑셀로 받은 이력서/자기소개서 데이터를 손쉽게 **분할**하고 **세로화**된 형태로 **전처리**해주는 도구입니다.  
Step1, Step2 로 나눠 사용하며, 사용자가 기준 컬럼을 선택하고 시트를 설정하여 자동으로 가공된 엑셀을 다운로드 받을 수 있습니다.

---



🪜 사용법
✅ Step1: 시트 분리 및 기준 컬럼 설정
엑셀 파일 업로드 (.xlsx)

기준 컬럼 선택 (예: 지원자번호, 지원직무, 이름)

분리할 시트 그룹 선택

Step1 엑셀 생성 버튼 클릭

📌 기준 컬럼은 병합된 셀을 풀고 난 뒤, 상단 1행과 2행을 결합하여 표시됩니다.
예: 기본정보 : 지원자번호, 인적정보 : 이름 형태

<!-- 이미지: step1 화면 업로드, 그룹 선택, 기준 컬럼 선택 -->
✅ Step2: 시트 세로화 및 정렬 설정
Step1 결과 엑셀이 자동으로 메모리에 로드됨

세로화할 시트 선택

자동 그룹 감지 (자동으로 실행됨, 필요 시 수동 가능)

컬럼 순서 드래그로 조정

정렬 기준 컬럼 및 방식 설정 (선택 사항)

Step2 엑셀 생성 버튼 클릭

📌 자동 감지는 동일 접두사 + 숫자 패턴 (자격증명1, 자격증명2, ...) 을 그룹으로 인식합니다.
📌 같은 지원자 내 여러 항목이 있을 경우 자동으로 연번이 부여됩니다.

<!-- 이미지: step2 화면 - 시트 선택, 자동 그룹 감지, 컬럼 순서 설정 -->
📝 결과 파일 예시
step1.xlsx: 기준 컬럼 + 시트별 분리된 데이터 포함

step2.xlsx: 세로화된 형식의 완성된 엑셀 파일



📌 유의사항
업로드할 엑셀 파일은 .xlsx 형식이어야 합니다.

기준 컬럼의 이름은 파일마다 다를 수 있으므로 직접 선택해 주세요.

병합된 셀은 자동으로 풀어지며, 그룹명과 컬럼명은 :로 연결됩니다.
