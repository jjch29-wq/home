import openpyxl
import os
import sys

data_file_path = r"c:\Users\jjch2\Desktop\PMI\home\data\위험성평가_기본데이터.xlsx"

if not os.path.exists(data_file_path):
    print("기본 데이터 엑셀 파일이 아직 생성되지 않았습니다.")
    sys.exit(0)

try:
    wb = openpyxl.load_workbook(data_file_path)
    if "RT" in wb.sheetnames:
        ws = wb["RT"]
        
        new_items = [
            ["비파괴검사", "가이드튜브 꺾임, 장비 결함 등으로 동위원소 선원 미회수 시 고선량 피폭 위험", "", 1, 3, "III", "사용 전 장비 점검 / 비상용 차폐복 및 집게(Handling tong) 비치 / 비상훈련"],
            ["비파괴검사", "야간 검사 시 조명 부족 및 시야 미확보로 인한 전도, 추락, 장비 충돌 위험", "", 2, 2, "IV", "작업구간 투광기 설치 / 안전조끼(반사띠) 착용 / 개인용 랜턴 지급"],
            ["사전준비", "배관 내부 및 깊은 트렌치 내부 출입 시 산소 결핍 또는 유해가스 체류로 인한 질식", "", 2, 3, "V", "출입 전 산소/유해가스 농도 측정 / 가스마스크 등 개인보호구 지급"],
            ["비파괴검사", "야간 조명등 및 검사장비 케이블 피복 손상, 누전, 우천 시 감전 위험", "", 2, 2, "IV", "누전차단기 부착 분전반 사용 / 전선 바닥 방치 금지(거치) / 우천 시 작업 통제"],
            ["이동", "야간 암실 차량 주·정차 및 이동 시 시야 확보 불량으로 근로자 및 타 장비 충돌", "", 2, 3, "V", "동승자 하차 후 신호봉 유도 / 차량 경광등 작동 / 작업자 반사조끼 착용"],
            ["작업 후 정리", "차량 이동 중 동위원소 저장용기 전도/낙하로 인한 방사능 유출 및 장비 파손", "", 1, 3, "III", "전용 운반차량 사용(경고표지 부착) / 저장소 시건장치 및 고정 상태 확인"]
        ]
        
        for item in new_items:
            ws.append(item)
            
        wb.save(data_file_path)
        print("성공적으로 업데이트 되었습니다.")
    else:
        print("RT 시트가 없습니다.")
except Exception as e:
    print(f"오류: {e}")
    sys.exit(1)
