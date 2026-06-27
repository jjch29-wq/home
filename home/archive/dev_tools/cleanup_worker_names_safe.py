"""
Safe Database Cleanup Script - Remove Shift Information from Worker Names
이 스크립트는 데이터베이스의 작업자 이름에서 shift 정보를 제거합니다.
예: "(주간) 김진환" → "김진환"

안전 기능:
- 자동 백업 생성
- 변경사항 미리보기
- 파일 잠금 감지
- 안전한 데이터 처리
"""
import pandas as pd
import re
import os
import shutil
from datetime import datetime

# 설정
db_path = 'Material_Inventory.xlsx'
backup_suffix = datetime.now().strftime('_backup_%Y%m%d_%H%M%S')

def clean_worker_name(name):
    """작업자 이름에서 shift 마커 제거"""
    if pd.isna(name) or not str(name).strip():
        return ''
    
    name_str = str(name).strip()
    
    # "(주간/야간/휴일) 이름" 패턴 체크
    match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", name_str)
    if match:
        actual_name = match.group(2).strip()
        return actual_name if actual_name else ''
    
    # Shift만 있고 이름 없는 경우
    if re.match(r"^\((주간|야간|휴일)\)$", name_str):
        return ''
    
    # 이미 정리된 이름
    return name_str

def main():
    print("="*70)
    print("작업자 이름 정리 스크립트")
    print("="*70)
    
    # 1. 파일 존재 확인
    if not os.path.exists(db_path):
        print(f"\n❌ 오류: '{db_path}' 파일을 찾을 수 없습니다!")
        input("\nPress Enter to exit...")
        return
    
    # 2. 파일 크기 확인
    file_size = os.path.getsize(db_path)
    print(f"\n파일 정보:")
    print(f"  - 파일명: {db_path}")
    print(f"  - 크기: {file_size:,} bytes")
    
    if file_size < 10000:
        print("\n⚠️  경고: 파일 크기가 비정상적으로 작습니다!")
        response = input("계속하시겠습니까? (y/N): ")
        if response.lower() != 'y':
            return
    
    # 3. 데이터 로드
    print("\n📂 데이터베이스를 읽는 중...")
    try:
        daily_usage_df = pd.read_excel(db_path, sheet_name='DailyUsage', engine='openpyxl')
    except PermissionError:
        print("\n❌ 오류: 파일이 다른 프로그램에서 열려있습니다!")
        print("MaterialManager를 종료하고 다시 시도해주세요.")
        input("\nPress Enter to exit...")
        return
    except Exception as e:
        print(f"\n❌ 오류: 파일을 읽을 수 없습니다: {e}")
        input("\nPress Enter to exit...")
        return
    
    print(f"✓ 총 {len(daily_usage_df)}개의 기록 로드됨")
    
    # 4. 작업자 이름 분석
    user_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
    
    print("\n🔍 작업자 이름 분석 중...")
    changes = []
    
    for col in user_cols:
        if col not in daily_usage_df.columns:
            continue
            
        for idx in daily_usage_df.index:
            original = daily_usage_df.loc[idx, col]
            if pd.isna(original):
                continue
                
            original_str = str(original).strip()
            if not original_str:
                continue
            
            cleaned = clean_worker_name(original)
            
            if original_str != cleaned:
                changes.append({
                    'row': idx + 2,  # Excel row number (1-indexed + header)
                    'column': col,
                    'before': original_str,
                    'after': cleaned if cleaned else '(삭제됨)'
                })
    
    # 5. 변경사항 미리보기
    if not changes:
        print("\n✓ 정리가 필요한 이름이 없습니다. 데이터베이스가 이미 깨끗합니다!")
        input("\nPress Enter to exit...")
        return
    
    print(f"\n📋 총 {len(changes)}개의 변경사항 발견:")
    print("\n변경 미리보기 (최대 20개):")
    print("-" * 70)
    for i, change in enumerate(changes[:20]):
        print(f"{i+1}. Row {change['row']}, {change['column']}: '{change['before']}' → '{change['after']}'")
    
    if len(changes) > 20:
        print(f"... 외 {len(changes) - 20}개 더")
    
    # 6. 사용자 확인
    print("\n" + "="*70)
    response = input(f"\n이 {len(changes)}개의 변경사항을 적용하시겠습니까? (y/N): ")
    
    if response.lower() != 'y':
        print("\n취소되었습니다.")
        input("\nPress Enter to exit...")
        return
    
    # 7. 백업 생성
    backup_path = db_path.replace('.xlsx', f'{backup_suffix}.xlsx')
    print(f"\n💾 백업 생성 중: {backup_path}")
    try:
        shutil.copy2(db_path, backup_path)
        print(f"✓ 백업 생성 완료")
    except Exception as e:
        print(f"❌ 백업 실패: {e}")
        input("\nPress Enter to exit...")
        return
    
    # 8. 데이터 정리
    print("\n🔧 데이터 정리 중...")
    for col in user_cols:
        if col in daily_usage_df.columns:
            daily_usage_df[col] = daily_usage_df[col].apply(clean_worker_name)
    
    # 9. 데이터 저장
    print("💾 변경사항 저장 중...")
    try:
        # 다른 시트들도 함께 로드
        materials_df = pd.read_excel(db_path, sheet_name='Materials', engine='openpyxl')
        transactions_df = pd.read_excel(db_path, sheet_name='Transactions', engine='openpyxl')
        monthly_usage_df = pd.read_excel(db_path, sheet_name='MonthlyUsage', engine='openpyxl')
        
        # 모든 시트를 다시 저장
        with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
            materials_df.to_excel(writer, sheet_name='Materials', index=False)
            transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
            monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
            daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
        
        print("✓ 저장 완료!")
        
    except Exception as e:
        print(f"\n❌ 저장 실패: {e}")
        print(f"\n백업 파일로 복원하세요: {backup_path}")
        input("\nPress Enter to exit...")
        return
    
    # 10. 결과 확인
    print("\n" + "="*70)
    print("✅ 정리 완료!")
    print("="*70)
    
    # 정리 후 작업자 목록 표시
    all_users = set()
    for col in user_cols:
        if col in daily_usage_df.columns:
            users = daily_usage_df[col].dropna().unique()
            all_users.update([u for u in users if u and str(u).strip()])
    
    print(f"\n정리된 작업자 목록 ({len(all_users)}명):")
    for user in sorted(all_users):
        print(f"  - {user}")
    
    print(f"\n백업 파일: {backup_path}")
    print("\n이제 MaterialManager를 실행하세요!")
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n사용자에 의해 취소되었습니다.")
    except Exception as e:
        print(f"\n\n예상치 못한 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
