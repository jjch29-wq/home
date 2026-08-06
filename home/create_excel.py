import pandas as pd

data = {
    "항목 (Section)": [
        "1.1 Scope of Work", 
        "1.2 Items and Materials",
        "3.2 Personnel Qualification",
        "4.2 Equipment and software",
        "4.2.1 Instrument Performance Verification"
    ],
    "영문 (English)": [
        "This procedure covers the minimum requirements for Phased Array Ultrasonic Testing (PAUT) of weld joints.",
        'This procedure shall be applied to the Venturi Tube (Line Size: 52", 56", 60", 68") manufactured or installed for the project. The applicable material is A516 Gr.70 and the thickness range is from a minimum of 15.88 mm to a maximum of 28.58 mm.',
        "All testing shall be carried out by qualified NDT service provider and certified personnel. All personnel shall be experienced, certified and have knowledge regarding PAUT equipment and technique for a minimum of 3 years with 3 different PAUT projects relevant to the project work scope.",
        "The following equipment shall be used, and all equipment and software shall comply with ASTM E2700-20 §7.",
        "PAUT instrument performance shall be verified by the manufacturer, the owner, or a laboratory, at 12-month intervals of the ultrasonic phased array instrument during its lifetime."
    ],
    "국문 (Korean)": [
        "본 절차서는 용접부의 위상배열 초음파탐상검사(PAUT)에 대한 최소한의 요건을 다룬다.",
        '본 절차서는 프로젝트를 위해 제작 및 설치되는 벤투리 튜브(Venturi Tube) (관경: 52", 56", 60", 68")에 적용된다. 적용되는 재질은 A516 Gr.70이며, 적용 가능한 재료의 두께 범위는 최소 15.88 mm에서 최대 28.58 mm이다.',
        "모든 검사는 자격을 갖춘 비파괴검사(NDT) 전문 업체 및 자격이 인증된 인원에 의해 수행되어야 한다. 모든 검사원은 프로젝트 작업 범위와 관련된 3건의 타 PAUT 프로젝트에서 최소 3년 이상의 PAUT 장비 및 기술에 대한 경험, 자격 인증 및 지식을 보유하여야 한다.",
        "다음의 장비가 사용되어야 하며, 모든 장비 및 소프트웨어는 ASTM E2700-20 §7 요건을 준수하여야 한다.",
        "PAUT 장비의 성능은 장비의 수명 기간 동안 12개월 주기로 제조업체, 소유주 또는 공인 시험 기관에 의해 검증되어야 한다."
    ]
}

df = pd.DataFrame(data)
writer = pd.ExcelWriter(r"c:\Users\-\PMI\home\Procedure_Scope_Update_v6.xlsx", engine='openpyxl')
df.to_excel(writer, index=False, sheet_name='Procedure')

worksheet = writer.sheets['Procedure']
worksheet.column_dimensions['A'].width = 40
worksheet.column_dimensions['B'].width = 80
worksheet.column_dimensions['C'].width = 80

writer.close()
print("Excel file v6 created successfully.")
