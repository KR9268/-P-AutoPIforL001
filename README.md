# AutoPIforL001

## 개요
- 아웃룩으로 특정 조건의 메일에서 첨부를 열어, pdf파일을 Regex로 확인 후 시스템에 자동으로 입력하는 코드
- 작업완료 후 조회한 pdf파일을, 시스템에 등록한 등록번호를 파일이름으로 해 저장

## 필수사항
- SAPGUI : SAP Scripting 활용
- SAP권한 : ZSDP10200_A
- 아웃룩설정 : 메일을 수신할 수 있는 상태로 세팅
- NERP_PI_LC, pkb_selenium 파일을 폴더에 함께 두기

## 사용법
(1) 기본값 입력
```python
# SAP 준비  *시스템 접속용 아이디, 비밀번호 입력*
sessions = NERP_PI_LC.check_and_open_sap('SEP', '실제아이디', '실제비번',3)
session = sessions[1]

# 주소 마스터 엑셀 준비    *주소 조회용 마스터 엑셀 파일 경로 입력*
wb = xw.Book('C:\\python_source\\' + 'pi_address.xlsx')

# 기준값 세팅         *파일 저장할 경로 입력*
file_path = "C:\\Users\\USER\\Downloads\\"
copy_path_all = ['D:\\()PI보관\\']
```
(2) 마스터 엑셀 파일 작성 (pi_address.xlsx)
- PI형식 : 파일 내의 번호형식과 대조되는 부분으로 중요
- 거래선코드, POL1, POL2, POD1, POD2, NEGO ORG : 시스템에 입력될 값으로 중요
  APPLICANT, SELLER, NOTIFY, CONSIGNEE, SHIPPING MARK, ADDNOTI, ADDCNEE : 입력될 주소값으로 중요

(3) 박스 순서대로 실행

## 특이사항
- 작업완료 후 메일로 통보를 하고자 하였으나, 아웃룩 발송기능이 막혀있고 잦은 Chrome버전 업데이트로 기능 제외

## 기능
- handle_regex_error : crawl_pdf_pi 함수에서 사용 (오류발생시 코드를 멈추며, skip_exit에 None외의 값 할당시 멈추지 않고 넘김)
- crawl_pdf_pi : 스캔이 아닌 pdf를 읽어들여 사용, pdf형식 바뀌면 관련 로직도 변경해야함
- replyFinishedByMail(미사용) : 작업완료 후 전체회신으로 메일회신. 관리상 불편한 점이 있어 미사용
