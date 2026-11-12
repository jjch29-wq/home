/**
 * 경인사무소 24시간 모바일 현장 입력 시스템 (Google Apps Script)
 * 컴퓨터 전원이 꺼져 있어도 스마트폰에서 언제든 저장 가능합니다.
 * 구글 드라이브 스프레드시트와 실시간 연동됩니다.
 */

// 1. 기본 설정 및 마스터 데이터
var APP_CONFIG = {
  companies: [
    "경인사무소",
    "D&K켐텍", "K&M", "SG엔지니어링", "SK에코플랜트",
    "경성기공", "글로벌 크레인", "나라삼양감속기", "대성기계", "대성테크놀로지",
    "대화유공업", "대화콘베어체인", "도원이앤씨", "동남ENG", "동아기업DAC",
    "동해펌프", "디에이치건설산업(주)", "롯데건설", "리니트코리아", "모나스펌프",
    "부강풀리", "삼남엔지니어링(주)", "삼성파워텍", "삼승종합건설", "삼화테크",
    "서울발전본부", "서진인스텍", "성헌", "세진공조기계", "수성웹툰(주)",
    "스페이스 체인지", "신성플랜트", "아이조음", "아이피에스", "양영",
    "에버랜드", "에스엠테크", "에이스냉동공조", "영흥발전", "올스톤산업",
    "위드엔지니어링", "유성씨앤에프", "의성기어", "이노텍코리아", "자인건축",
    "전진이앤씨", "제이에스티", "제일푸리", "청원기연", "캐논코리아",
    "캐리비안베이", "케이엔엠", "코아베큠", "코오롱글로벌", "태성기전",
    "태창기계공업", "포스코이앤씨", "포스코인터내셔널", "피테크", "한국산업인력공단",
    "한양스틸", "함감속기제작소", "현대건설", "현대산업", "현성펌프",
    "현진산업"
  ],
  users: [
    "이사 곽재운", "이사 박홍규",
    "부장 주진철", "부장 박광복", "부장 이병국", "부장 이용주", "부장 주동기", "부장 김규년", "부장 곽상돈", "부장 정효철", "부장 노현태",
    "차장 김춘호", "차장 권인경", "차장 박원준",
    "과장 주영광", "과장 유상훈",
    "대리 김상현", "대리 윤석현", "대리 우명광", "대리 박만영",
    "계장 박상태", "계장 이준우", "계장 장승대", "계장 이은수", "계장 손주언", "계장 이경재", "계장 박재욱", "계장 한지원",
    "주임 김진환", "주임 황조현", "주임 김태현"
  ],
  vehicles: [
    "84저1259 (스타렉스)",
    "95가0200 (스타렉스)",
    "91주8839 (스타렉스)",
    "90서8616 (스타렉스)",
    "76마3422 (스타렉스 12인승)",
    "71고4405 (스타렉스 12인승)",
    "89보4028 (탑차)",
    "81두1580 (탑차)",
    "81루5100 (스타렉스)",
    "95가0175 (스타렉스)",
    "81도5958 (스타렉스)",
    "90너4889 (탑차)",
    "자가용/대중교통"
  ],
  equipments: [
    "PAUT MX2(OMNI2-103085)",
    "PAUT MX2(OMNI2-100055)",
    "PAUT MX2(OMNI2-103918)",
    "PAUT X3(QC-0093084)",
    "Sitescan 150(I000760)",
    "Sitescan 150(I000950)",
    "Sitescan 140(1403264TC)",
    "Masterscan 380(I002561)",
    "Masterscan 380(I003201)",
    "Masterscan 380(I005587)",
    "Masterscan 380(I005593)",
    "Epoch 650(150128312)",
    "USM-35(8814A)",
    "USM-GO(I001008)",
    "두께측정기 37DL-PLUS(051301807)",
    "Se-75(880D D10692)",
    "Se-75(880D D10626)",
    "Se-75(880D D4507)",
    "Se-75(880D D7748)",
    "Ir-192(880D D10700)",
    "Ir-192(880D D3086)",
    "Ir-192(880D D4567)",
    "Ir-192(880D D10674)",
    "Ir-192(880S S1684)",
    "방사선발생장치 300EGM2(FR0020-2)",
    "방사선발생장치 RF300EGM2(GR0002-3)",
    "Yoke AC(MY-2 NW0453)",
    "Yoke AC(MP-A3 MP31646)",
    "Yoke AC(MY-2 NW6999)",
    "Yoke AC(MP-A3 KI-01)",
    "Yoke AC(MP-A3 NW1510097)",
    "Yoke AC(MY-2 NW4424)",
    "Yoke AC(MP-A3 KI-02)",
    "Yoke AC(MY-2 NW7328)",
    "Yoke AC(MY-2 NW8064)",
    "Yoke AC(MY-2 NW5338)",
    "Yoke AC(MY-2 NW4413)",
    "Yoke AC(MY-2 NW1301154)",
    "Yoke AC(MY-2 NW4700)",
    "Yoke DC(MP-A2D MP3403)",
    "Yoke DC(MP-A2D MP1336)",
    "Yoke DC(MP-A2D MP9693)",
    "Yoke DC(MP-A2D MP2458)",
    "AC POWER SUPPLY DA-PB(NW0589)",
    "AC POWER SUPPLY DA-PB(NW0604)",
    "AC POWER SUPPLY DA-PB(NW0639)",
    "AC POWER SUPPLY DA-PB(NW0718)",
    "AC POWER SUPPLY DA-PB(NW0582)",
    "AC POWER SUPPLY DA-PB(NW0230)",
    "PT 약품",
    "합금분석기(X-ray) XL2 100G(200983)",
    "합금분석기(RI) XLi818(24673)",
    "경도계 MIC10(1-2129)",
    "페이라트함량측정기 FMP30(160002524)",
    "내시경장비 ER400(E30969R)",
    "해당없음"
  ],
  aux_equipments: [
    "PAUT SCANNER (MANUAL)",
    "MICROBE 스캐너 (Body-137331, Encoder-136861)",
    "COBRA 스캐너 (139310)",
    "PAUT 스캐너 (133490)",
    "PAUT PROBE(A2) 2.25L64-0.75*12-D62-U-110-2 (GBA008)",
    "PAUT PROBE(A2) 2.25L64-0.75*12-D62-U-110-2 (GBA009)",
    "PAUT PROBE(A2) 2.25L64-A2 (P3378)",
    "PAUT PROBE(A2) 2.25L64-A2 (R0080)",
    "PAUT PROBE(A2) 5L64-06*10-D62-U-110-2.0-P1 (GBA011)",
    "PAUT PROBE(A2) 5L64-A2 (S1599)",
    "PAUT PROBE(A2) 5L64-A2 (S1694)",
    "PAUT PROBE(A2) 5L64-A2 (S5225)",
    "PAUT PROBE(A15) 7.5CCEV35-16-A15 (V4063)",
    "PAUT PROBE(A15) 7.5CCEV35 (V4109)",
    "PAUT PROBE(A25) XAAB-10882 (W2479)",
    "PAUT PROBE(A25) 4DL16-12X5-A25 (U3188)",
    "PAUT PROBE(A31) 5L32-A31 (P4301)",
    "PAUT PROBE(A25) 5DL16-12X5-A25-P-2.5-OM (Y1193)",
    "PAUT PROBE(A10) 5L16-A10 (S1873)",
    "PAUT PROBE(A32) 5L64-32X10-A32 (U1462)",
    "PAUT PROBE(A32) 5L64-32X10-A32 (U14651)",
    "엔코더(PAUT) COBRA-SP-ENC (25A2505)",
    "엔코더(PAUT) (소형 2EA, 대형 2EA)",
    "엔코더(PAUT) CKS007 (135275)",
    "엔코더(PAUT) CKS007 (136682)",
    "케이블(PAUT) (140828-03)",
    "채널분배기(PAUT) (QC-008086)",
    "채널분배기(PAUT) OMNI-A-ADP05 (974106-03)",
    "채널분배기(PAUT) (1011294-01)",
    "모듈(PAUT) Omni-M-PA32 (OMNI-600336)",
    "TomoView Lock Key (1930372742)",
    "Index Control Bar (N/A(Bar 300mm))",
    "엔코더(PAUT) (145552)",
    "채널분배기(PAUT) (QC003541)",
    "충전기 (70180-0004709)",
    "충전기 (70180-0024164)",
    "빔툴키 PRO H (MN720433)",
    "OMNISCAN PC TOOLKIT / OMNIPC-A (134846330)",
    "WELDSIGHT-A (1522192766)",
    "채널분배기(PAUT) Omni-A2-SPLIT128 4UT (QC-080522)",
    "엔코더(PAUT) (132801)",
    "스캐너 (133490)",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC",
    "PAUT PROBE-WEDGE(SA2) SA2-OL-IHC",
    "PAUT PROBE-WEDGE(SA10) SA10-N55S-IHC",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IHC",
    "PAUT PROBE-WEDGE(SA25) SA25-DL70L 평웻지",
    "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD1.05",
    "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD1.315",
    "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD2.875",
    "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD3.5",
    "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD8",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD10",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD12",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD14",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD16",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD18",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD20",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD24",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD30",
    "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD22 DUAL",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD0.84",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.05",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.315",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.66",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.9",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD2.375",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD2.875",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD3.25",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD3.5",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD4",
    "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD4.5",
    "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC-AOD8.625",
    "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC-AOD12.75",
    "PAUT PROBE-WEDGE(SD62) SD62-N55S-IHC-DUAL",
    "PAUT PROBE-WEDGE(M15) M10-N55S-IHC(PS)",
    "PAUT PROBE-WEDGE(M15) M15-N60S-IH",
    "PAUT PROBE-WEDGE(M15) M15-N60S-IH-AOD1.66",
    "PAUT PROBE-WEDGE(M15) M15-N60S-IH-AOD2.8",
    "PAUT PROBE-WEDGE(SA31) N55S-I HC-PLATE",
    "표준시험편 IIW TYPEⅠ(C/S) (V1-CS-25-D46)",
    "표준시험편 IIW TYPEⅠ(STS304) (16-3060)",
    "표준시험편 IIW TYPEⅠ(STS304) (A19374)",
    "표준시험편 IIW (BS-2704)",
    "표준시험편 ISO19675 PAUT CAL BLOCK",
    "표준시험편 ISO19675 PAUT CAL BLOCK (56340)",
    "표준시험편 STB-A1 (2)",
    "표준시험편 STB-A1 (09148)",
    "표준시험편 STB-A1 (1830)",
    "표준시험편 STB-A1 (356)",
    "표준시험편 STB-A1 (No.09159)",
    "표준시험편 STB-A1 (244)",
    "표준시험편 STB-A1 (257)",
    "사각 PROBE(70˚) MWB70-2 (56923-5617)",
    "사각 PROBE(70˚) MWB70-2 (56923-5124)",
    "사각 PROBE(70˚) A5013 (572702)",
    "사각 PROBE(70˚) 5C10X10A70 (AFPC017)",
    "사각 PROBE(70˚) 5C10X10A70 (AFPC018)",
    "사각 PROBE(70˚) 2C14X14A70 (ACUC002)",
    "사각 PROBE(70˚) 2C16X16 A70 (ACUC020)",
    "사각 PROBE(70˚) 5C10X10A70 (AFPB11)",
    "사각 PROBE(70˚) D791-RM (556223)",
    "사각 PROBE(70˚) A609-S (514619)",
    "사각 PROBE(70˚) A609-S (16306)",
    "사각 PROBE(70˚) D7908 (1271887)",
    "사각 PROBE(70˚) MWB70-4 (56929-92244)",
    "사각 PROBE(70˚) MWB70-4 (56929-92222)",
    "사각 PROBE(70˚) MWB70-4 (56929-92258)",
    "사각 PROBE(70˚) MWB70-4 (56929-40641)",
    "사각 PROBE(70˚) SWB70-2 (58416-01771)",
    "사각 PROBE(70˚) U-PRO70-4 (H5817)",
    "사각 PROBE(70˚) MWB70-4 (56929-84005)",
    "사각 PROBE(70˚) MWB70-4 (56929-71089)",
    "사각 PROBE(70˚) TKS70-2 (ACUC19)",
    "사각 PROBE(70˚) TKS70-4 (H0006)",
    "사각 PROBE(70˚) KA4-70 (8×9) (H6323)",
    "사각 PROBE(70˚) 5C10×10A70 (AFPC023)",
    "사각 PROBE(70˚) MWB70-4 (8×9) (56929-59715)",
    "사각 PROBE(70˚) WB70-2 (20×22) (57001-51556)",
    "사각 PROBE(60˚) KL5-60 (H7793)",
    "사각 PROBE(60˚) KL5-60 (H7794)",
    "사각 PROBE(60˚) KL5-60 (H7798)",
    "사각 PROBE(60˚) MWB60-2 (56922-05004)",
    "사각 PROBE(60˚) MWB60-2 (57000-51382)",
    "사각 PROBE(60˚) MWB60-4 (56928-17156)",
    "사각 PROBE(60˚) MWB60-4 (56928-17283)",
    "사각 PROBE(60˚) MWB60-4 (56928-35552)",
    "사각 PROBE(60˚) MWB60-4 (56929-52302)",
    "사각 PROBE(60˚) MWB60-4 (56928-17312)",
    "사각 PROBE(60˚) TKS60 (ACUB001)",
    "사각 PROBE(60˚) MWB60-4 (56928-57116)",
    "사각 PROBE(60˚) MWB60-4 (56928-52153)",
    "사각 PROBE(60˚) TKS60 (AFPB54)",
    "사각 PROBE(60˚) TKS60 (ACUB17)",
    "사각 PROBE(60˚) TKS60-2 (16×16) (ACUB016)",
    "사각 PROBE(60˚) TKS5C5×5A60 (AFKA98)",
    "사각 PROBE(45˚) 2Z10X10R (JR1399)",
    "사각 PROBE(45˚) KL5-45 (H6005)",
    "사각 PROBE(45˚) KL5-45 (H6006)",
    "사각 PROBE(45˚) HS-7427-2/45 (H9163)",
    "사각 PROBE(45˚) MWB45-2 (56921-6976)",
    "사각 PROBE(45˚) MWB45-4 (56927-56190)",
    "사각 PROBE(45˚) MWB45-4 (56927-58799)",
    "사각 PROBE(45˚) MWB45-4 (56927-26651)",
    "사각 PROBE(45˚) 2C16X16A45 (ACUA017)",
    "사각 PROBE(45˚) 2C14X14A45 (ACUA001)",
    "사각 PROBE(45˚) 2C10X10A45 (ACPA32)",
    "사각 PROBE(45˚) 2C10X10A45 (ACPA33)",
    "사각 PROBE(45˚) MWB45 (56927-35905)",
    "사각 PROBE(45˚) U-PRO45 (D4065)",
    "사각 PROBE(45˚) TKS45 (ACUC31)",
    "사각 PROBE(45˚) TKS45 (ACUA009)",
    "사각 PROBE(45˚) MWB45 (56927-58233)",
    "사각 PROBE(45˚) MWB45 (56927-52290)",
    "사각 PROBE(45˚) TKS5C5×5A45 (AFKA87)",
    "사각 PROBE(30˚) TKS5C5×5A30 (AFKA79)",
    "사각 PROBE(35˚) TKS5C5×5A35 (AFKA81)",
    "사각 PROBE(40˚) TKS5C5×5A40 (AFKA84)",
    "사각 PROBE(50˚) TKS5C5×5A50 (AFKA92)",
    "사각 PROBE(55˚) TKS5C5×5A55 (AFKA93)",
    "수직 PROBE(Φ5) 1C24N (NAGA67)",
    "수직 PROBE(Φ5) A109-S (57744-54161)",
    "수직 PROBE(Φ5) A109-S (57745-76316)",
    "수직 PROBE(Φ5) 5C5N (57745-76327)",
    "수직 PROBE(Φ10) MB2S (57749-52552)",
    "수직 PROBE(Φ10) MB2S (57749-53039)",
    "수직 PROBE(Φ10) MB2S (57749-75353)",
    "수직 PROBE(Φ10) MB2S (57749-75360)",
    "수직 PROBE(Φ10) MB4S (57749-75378)",
    "수직 PROBE(Φ10) MB4S (502788)",
    "수직 PROBE(Φ10) MB4S (503983)",
    "수직 PROBE(Φ10) MB4S (NFBA 56)",
    "수직 PROBE(Φ10) MB4S (NFGIA 15)",
    "수직 PROBE(Φ20) 5C20N (57746-51530)",
    "수직 PROBE(Φ24) B1S (I 130)",
    "수직 PROBE(Φ24) B2S (I 037)",
    "수직 PROBE(Φ24) B2S (I 138)",
    "수직 PROBE(Φ24) B2S (57748-4571)",
    "수직 PROBE(Φ24) SW2-24 (57748-52567)",
    "수직 PROBE(Φ24) SW2-24 (57748-56177)",
    "수직 PROBE(Φ24) SW2-24 (57748-56178)",
    "수직 PROBE(Φ24) B4S (T029092)",
    "수직 PROBE(Φ24) B4S (57746-51052)",
    "수직 PROBE(Φ25) SLF4-25 (57745-07614)",
    "수직 PROBE(Φ10) MB4S (57749-55037)",
    "수직 PROBE(Φ10) MB4S (57749-74386)",
    "수직 PROBE(Φ10) MB4S (57749-71405)",
    "수직 PROBE(Φ10) MB2S (57748-4757)",
    "수직 PROBE(Φ20) TKS5C20N (NFGIA20)",
    "수직 PROBE(Φ24) B4S (57746-51384)",
    "듀얼-사각 PROBE(70˚) 2C8ⅹ14LAD70 (DLANC006)",
    "듀얼-사각 PROBE(기타) 2C6ⅹ18CRD (DCCRJA032)",
    "듀얼-사각 PROBE(기타) 2C8ⅹ14CRD (DCCNA05)",
    "듀얼-사각 PROBE(기타) TRCr2-AUST (12-1944)",
    "듀얼-수직 PROBE(Φ5) MTD705 (869315)",
    "듀얼-수직 PROBE(Φ10) MSEB4 (57462-78164)",
    "듀얼-수직 PROBE(Φ25) SEB4 (57469-50297)",
    "표준시험편 1mm~30mm 10단",
    "표준시험편 ASME C - BCB",
    "표준시험편 ASME RB FBH (Ø3) (SIS-GI-001~006)",
    "표준시험편 ASME 감도교정",
    "표준시험편 ASME 직진성 (SIS-CAL-2201)",
    "표준시험편 AWS DS (35161)",
    "표준시험편 AWS RC",
    "표준시험편 BEARING 시험편 (4EA)",
    "표준시험편 SS0902 스텝웨지(20mm-4단)",
    "표준시험편 ST1018스텝웨지(30mm-5단) (KDE21371)",
    "표준시험편 STB G V2",
    "표준시험편 STB G V3",
    "표준시험편 STB GV15-1.4",
    "표준시험편 STB GV15-1",
    "표준시험편 STB GV15-2.8",
    "표준시험편 STB GV15-2",
    "표준시험편 STB GV15-4",
    "표준시험편 STB GV15-5.6",
    "표준시험편 STB GV5",
    "표준시험편 STB GV8",
    "표준시험편 STB N1",
    "표준시험편 STB-A2 (5180)",
    "표준시험편 STB-A2 (J-1)",
    "표준시험편 STB-A2 (293)",
    "표준시험편 STB-A2 (299)",
    "표준시험편 STB-A3 (4543)",
    "표준시험편 V3 BLOCK (STS304) (001-211)",
    "대비시험편 RB-4 (C/S) NO. 1 (16-2577)",
    "대비시험편 RB-4 (C/S) NO. 2 (16-2096)",
    "대비시험편 RB-4 (STS304) NO. 1 (18-3903)",
    "대비시험편 RB-4 (STS304) NO. 2 (19-1518)",
    "대비시험편 RB-A7",
    "대비시험편 RB-RA",
    "대비시험편 RB-RB",
    "대비시험편 RB-RC",
    "대비시험편 RB-RE(1), (2), (3)",
    "접촉매질(Couplant)",
    "ADR, SURVEY, 콜리메타",
    "콜리메타(120˚) TSC-19P (7334)",
    "콜리메타(120˚) TSC-19P (7726)",
    "콜리메타(120˚) TSC-19P (7747)",
    "콜리메타(120˚) TSC-19P (7765)",
    "콜리메타(120˚) TC-120L (L 284)",
    "콜리메타(120˚) TC-120L (L 285)",
    "콜리메타(120˚) TC-120L (L 365)",
    "콜리메타(120˚) TC-120L (L 368)",
    "콜리메타(120˚) TC-120L (L 369)",
    "콜리메타(360˚) TSC-19P (2104)",
    "콜리메타(360˚) TSC-19P (8042)",
    "콜리메타(360˚) TSC-19P (8125)",
    "콜리메타(360˚) TSC-19 (NW074)",
    "Survey Meter RDS-30 (371172)",
    "Survey Meter RDS-30 (371175)",
    "Survey Meter RDS-30 (371174)",
    "Survey Meter RDS-30 (371176)",
    "Survey Meter RDS-30 (371198)",
    "Survey Meter RSM-100 (1001107)",
    "Survey Meter RSM100E (1000248)",
    "Survey Meter RSM-300P (FY12RP10LI)",
    "Survey Meter RSM-300P (FY12RP10LJ)",
    "Survey Meter RSM-100 (1000362)",
    "Survey Meter MKS-05P (0907026)",
    "Survey Meter MKS-05 (1004874)",
    "Survey Meter RSM-100 (1000677)",
    "Alarm Dosimeter PICO CATCH-100 (NW1085)",
    "Alarm Dosimeter PICO CATCH-100 (NW1173)",
    "Alarm Dosimeter PICO CATCH-100 (NW2109)",
    "Alarm Dosimeter PICO CATCH-100 (NW2111)",
    "Alarm Dosimeter PICO CATCH-100 (NW2112)",
    "Alarm Dosimeter PICO CATCH-100 (NW2113)",
    "Alarm Dosimeter PICO CATCH-100 (NW2114)",
    "Alarm Dosimeter PICO CATCH-100 (NW2115)",
    "Alarm Dosimeter PICO CATCH-100 (NW2272)",
    "Alarm Dosimeter PICO CATCH-100 (NW2421)",
    "Alarm Dosimeter PICO CATCH-100 (NW2422)",
    "Alarm Dosimeter PICO CATCH-100 (NW2423)",
    "Alarm Dosimeter PICO CATCH-100 (NW2529)",
    "Alarm Dosimeter PICO CATCH-100 (NW2530)",
    "Alarm Dosimeter PICO CATCH-100 (NW2531)",
    "Alarm Dosimeter PICO CATCH-100 (NW2532)",
    "Alarm Dosimeter PICO CATCH-100 (NW2533)",
    "Alarm Dosimeter SPD-9300 (NB010003)",
    "Alarm Dosimeter SPD-9300 (NB010015)",
    "Alarm Dosimeter SPD-9300 (NB010065)",
    "Alarm Dosimeter SPD-9300 (NB010069)",
    "Alarm Dosimeter SPD-9300 (NB10001)",
    "Alarm Dosimeter SPD-9300 (NB10002)",
    "Alarm Dosimeter CLOVER (DP05AC47DF)",
    "Alarm Dosimeter CLOVER (DQ11AC49HJ)",
    "Alarm Dosimeter CLOVER (FV20AC60FG)",
    "Alarm Dosimeter CLOVER (FV20AC60FF)",
    "Alarm Dosimeter CLOVER (DQ11AC49HL)",
    "Alarm Dosimeter PICO CATCH-100 (NW1081)",
    "Alarm Dosimeter PICO CATCH-100 (NW1087)",
    "Alarm Dosimeter PICO CATCH-100 (NW1090)",
    "Alarm Dosimeter PICO CATCH-100 (NW2780)",
    "Alarm Dosimeter PICO CATCH-100 (NW2127)",
    "R/CONTROL (25ft) NRC-2 (NO.7)",
    "R/CONTROL (25ft) NRC-2 (NW1170)",
    "R/CONTROL (25ft) NRC-2 (NW2750)",
    "R/CONTROL (25ft) KRC-2 (RA23793)",
    "R/CONTROL (25ft) NRC-2 (1510028)",
    "R/CONTROL (25ft) NRC-2 (2509014)",
    "R/CONTROL (25ft) NRC-2 (2509016)",
    "R/CONTROL (25ft) NRC-2 (2509017)",
    "R/CONTROL (50ft) NRC-2 (1510027)",
    "R/CONTROL (50ft) NRC-2 (1602031)",
    "R/CONTROL (기타) NRC-2 (6.5m) (2208006)",
    "R/CONTROL (기타) NRC-2 (8.5m) (2008022)",
    "R/CONTROL (기타) NRC-2 (12m) (2006058)",
    "R/CONTROL (기타) NRC-2 (12m) (2208004)",
    "R/CONTROL (기타) KRC-2 (12m) (RA23406)",
    "R/CONTROL (기타) KRC-2 (12m) (RA50078)",
    "R/CONTROL (25ft) NRC-2 (2607008)",
    "R/CONTROL (50ft) NRC-2 (2601004)",
    "R/CONTROL (50ft) NRC-2 (2601005)",
    "건조기 15형 (NW0413)",
    "건조기 15형 (RD1120)",
    "건조기 25형 (NW0130)",
    "건조기 25형 (NW0219)",
    "건조기 25형 (RD2242)",
    "판독기 P35 (NW0455)",
    "판독기 P35 (NW0401)",
    "판독기 P35 (NW0231)",
    "판독기 P35 (RV6594)",
    "판독기 DF-14L (RV1036)",
    "판독기 FD-2C15 (RV7364)",
    "판독기 P35 (NW0463)",
    "판독기 P35 (NW0522)",
    "판독기 P35 (NW0441)",
    "판독기 LUXA-D35 (063-50027)",
    "판독기 DF-2D55 (RV1751)",
    "판독기 P35 (RV1799)",
    "판독기 FD-2C15 (RV7349)",
    "판독기 P35 (NW0082)",
    "판독기 P35 (NW0129)",
    "농도계 PDA-100 (06011257)",
    "농도계 PDA-100 (06014961)",
    "농도계 PDA-100 (060150183)",
    "농도계 PDA-100 (060150447)",
    "농도계 PDA-100 (060150423)",
    "라이트마킹",
    "영상화장비 P35-14S (NW0401)",
    "판독기 DF-4L (RV-1404)",
    "크롤러 (45B5E8)",
    "농도계 DT-300 PLUS (320644)",
    "조도계 LS-150 (10004333)",
    "자분살포기",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW0096)",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW0083)",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW0113)",
    "Black Light (베터리교체형) NAWOO 2P (1969815)",
    "Black Light (베터리교체형) NAWOO 2P (1969816)",
    "Black Light (베터리교체형) 3P-BS (NW 0007)",
    "Black Light (베터리교체형) EE-365 (2057820)",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0180)",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0065)",
    "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0196)",
    "Black Light (요크부착형) UV-LED YOKE거치식 (1504050)",
    "Black Light (요크부착형) UV-LED YOKE거치식 (1504062)",
    "UV Meter YK-35UV (071550)",
    "조도계 LUX Meter 1330A (180604411)",
    "조도계 LUX Meter 1330A (130310985)",
    "조도계 LUX Meter 1330A (160500222)",
    "표면온도계 1300 (170900753)",
    "표면온도계 TK-307A (E3102035580)",
    "표면온도계 1300 (090204725)",
    "표면온도계 1300 (160701448)",
    "자장계 (바늘형) Model 20 (13-2393)",
    "자장계 (바늘형) Model 20 (13-2464)",
    "KS-표준시험편 (A 형) A1-7/50 (0101002)",
    "팔각시험편",
    "리프팅파워블럭 (AC 형) 10lb(4.5kg) (NW0938)",
    "리프팅파워블럭 (AC 형) 10lb(4.5kg) (MP19574)",
    "리프팅파워블럭 (AC 형) 10lb(4.5kg) (NW0756)",
    "KS-표준시험편 (A 형) A형 대비시험편 (NDI 980058)",
    "KS-표준시험편 (A 형) A형 대비시험편 (614 A/B)",
    "원형통 피 게이지 0.780~15.57mm/10PCS (GCN0002-052-1)",
    "조도계",
    "온습도계",
    "없음"
  ],
  method_defaults: {
    "PMI": {"equip_main": "\ud569\uae08\ubd84\uc11d\uae30(X-ray) XL2 100G(200983)", "aux_equipments": [], "default_aux": ""},
    "RT": {
      "equip_main": "Se-75(880D D10692)",
      "aux_equipments": [
        "ADR, SURVEY, 콜리메타",
        "콜리메타(120˚) TSC-19P (7334)",
        "콜리메타(120˚) TSC-19P (7726)",
        "콜리메타(120˚) TSC-19P (7747)",
        "콜리메타(120˚) TSC-19P (7765)",
        "콜리메타(120˚) TC-120L (L 284)",
        "콜리메타(120˚) TC-120L (L 285)",
        "콜리메타(120˚) TC-120L (L 365)",
        "콜리메타(120˚) TC-120L (L 368)",
        "콜리메타(120˚) TC-120L (L 369)",
        "콜리메타(360˚) TSC-19P (2104)",
        "콜리메타(360˚) TSC-19P (8042)",
        "콜리메타(360˚) TSC-19P (8125)",
        "콜리메타(360˚) TSC-19 (NW074)",
        "Survey Meter RDS-30 (371172)",
        "Survey Meter RDS-30 (371175)",
        "Survey Meter RDS-30 (371174)",
        "Survey Meter RDS-30 (371176)",
        "Survey Meter RDS-30 (371198)",
        "Survey Meter RSM-100 (1001107)",
        "Survey Meter RSM100E (1000248)",
        "Survey Meter RSM-300P (FY12RP10LI)",
        "Survey Meter RSM-300P (FY12RP10LJ)",
        "Survey Meter RSM-100 (1000362)",
        "Survey Meter MKS-05P (0907026)",
        "Survey Meter MKS-05 (1004874)",
        "Survey Meter RSM-100 (1000677)",
        "Alarm Dosimeter PICO CATCH-100 (NW1085)",
        "Alarm Dosimeter PICO CATCH-100 (NW1173)",
        "Alarm Dosimeter PICO CATCH-100 (NW2109)",
        "Alarm Dosimeter PICO CATCH-100 (NW2111)",
        "Alarm Dosimeter PICO CATCH-100 (NW2112)",
        "Alarm Dosimeter PICO CATCH-100 (NW2113)",
        "Alarm Dosimeter PICO CATCH-100 (NW2114)",
        "Alarm Dosimeter PICO CATCH-100 (NW2115)",
        "Alarm Dosimeter PICO CATCH-100 (NW2272)",
        "Alarm Dosimeter PICO CATCH-100 (NW2421)",
        "Alarm Dosimeter PICO CATCH-100 (NW2422)",
        "Alarm Dosimeter PICO CATCH-100 (NW2423)",
        "Alarm Dosimeter PICO CATCH-100 (NW2529)",
        "Alarm Dosimeter PICO CATCH-100 (NW2530)",
        "Alarm Dosimeter PICO CATCH-100 (NW2531)",
        "Alarm Dosimeter PICO CATCH-100 (NW2532)",
        "Alarm Dosimeter PICO CATCH-100 (NW2533)",
        "Alarm Dosimeter SPD-9300 (NB010003)",
        "Alarm Dosimeter SPD-9300 (NB010015)",
        "Alarm Dosimeter SPD-9300 (NB010065)",
        "Alarm Dosimeter SPD-9300 (NB010069)",
        "Alarm Dosimeter SPD-9300 (NB10001)",
        "Alarm Dosimeter SPD-9300 (NB10002)",
        "Alarm Dosimeter CLOVER (DP05AC47DF)",
        "Alarm Dosimeter CLOVER (DQ11AC49HJ)",
        "Alarm Dosimeter CLOVER (FV20AC60FG)",
        "Alarm Dosimeter CLOVER (FV20AC60FF)",
        "Alarm Dosimeter CLOVER (DQ11AC49HL)",
        "Alarm Dosimeter PICO CATCH-100 (NW1081)",
        "Alarm Dosimeter PICO CATCH-100 (NW1087)",
        "Alarm Dosimeter PICO CATCH-100 (NW1090)",
        "Alarm Dosimeter PICO CATCH-100 (NW2780)",
        "Alarm Dosimeter PICO CATCH-100 (NW2127)",
        "R/CONTROL (25ft) NRC-2 (NO.7)",
        "R/CONTROL (25ft) NRC-2 (NW1170)",
        "R/CONTROL (25ft) NRC-2 (NW2750)",
        "R/CONTROL (25ft) KRC-2 (RA23793)",
        "R/CONTROL (25ft) NRC-2 (1510028)",
        "R/CONTROL (25ft) NRC-2 (2509014)",
        "R/CONTROL (25ft) NRC-2 (2509016)",
        "R/CONTROL (25ft) NRC-2 (2509017)",
        "R/CONTROL (50ft) NRC-2 (1510027)",
        "R/CONTROL (50ft) NRC-2 (1602031)",
        "R/CONTROL (기타) NRC-2 (6.5m) (2208006)",
        "R/CONTROL (기타) NRC-2 (8.5m) (2008022)",
        "R/CONTROL (기타) NRC-2 (12m) (2006058)",
        "R/CONTROL (기타) NRC-2 (12m) (2208004)",
        "R/CONTROL (기타) KRC-2 (12m) (RA23406)",
        "R/CONTROL (기타) KRC-2 (12m) (RA50078)",
        "R/CONTROL (25ft) NRC-2 (2607008)",
        "R/CONTROL (50ft) NRC-2 (2601004)",
        "R/CONTROL (50ft) NRC-2 (2601005)",
        "건조기 15형 (NW0413)",
        "건조기 15형 (RD1120)",
        "건조기 25형 (NW0130)",
        "건조기 25형 (NW0219)",
        "건조기 25형 (RD2242)",
        "판독기 P35 (NW0455)",
        "판독기 P35 (NW0401)",
        "판독기 P35 (NW0231)",
        "판독기 P35 (RV6594)",
        "판독기 DF-14L (RV1036)",
        "판독기 FD-2C15 (RV7364)",
        "판독기 P35 (NW0463)",
        "판독기 P35 (NW0522)",
        "판독기 P35 (NW0441)",
        "판독기 LUXA-D35 (063-50027)",
        "판독기 DF-2D55 (RV1751)",
        "판독기 P35 (RV1799)",
        "판독기 FD-2C15 (RV7349)",
        "판독기 P35 (NW0082)",
        "판독기 P35 (NW0129)",
        "농도계 PDA-100 (06011257)",
        "농도계 PDA-100 (06014961)",
        "농도계 PDA-100 (060150183)",
        "농도계 PDA-100 (060150447)",
        "농도계 PDA-100 (060150423)",
        "라이트마킹",
        "영상화장비 P35-14S (NW0401)",
        "판독기 DF-4L (RV-1404)",
        "크롤러 (45B5E8)",
        "농도계 DT-300 PLUS (320644)",
        "조도계 LS-150 (10004333)",
        "없음"
      ],
      "default_aux": "ADR, SURVEY, 콜리메타",
      "unit": "매"
    },
    "UT": {
      "equip_main": "Sitescan 150(I000760)",
      "aux_equipments": [
        "사각 PROBE(70˚) MWB70-2 (56923-5617)",
        "사각 PROBE(70˚) MWB70-2 (56923-5124)",
        "사각 PROBE(70˚) A5013 (572702)",
        "사각 PROBE(70˚) 5C10X10A70 (AFPC017)",
        "사각 PROBE(70˚) 5C10X10A70 (AFPC018)",
        "사각 PROBE(70˚) 2C14X14A70 (ACUC002)",
        "사각 PROBE(70˚) 2C16X16 A70 (ACUC020)",
        "사각 PROBE(70˚) 5C10X10A70 (AFPB11)",
        "사각 PROBE(70˚) D791-RM (556223)",
        "사각 PROBE(70˚) A609-S (514619)",
        "사각 PROBE(70˚) A609-S (16306)",
        "사각 PROBE(70˚) D7908 (1271887)",
        "사각 PROBE(70˚) MWB70-4 (56929-92244)",
        "사각 PROBE(70˚) MWB70-4 (56929-92222)",
        "사각 PROBE(70˚) MWB70-4 (56929-92258)",
        "사각 PROBE(70˚) MWB70-4 (56929-40641)",
        "사각 PROBE(70˚) SWB70-2 (58416-01771)",
        "사각 PROBE(70˚) U-PRO70-4 (H5817)",
        "사각 PROBE(70˚) MWB70-4 (56929-84005)",
        "사각 PROBE(70˚) MWB70-4 (56929-71089)",
        "사각 PROBE(70˚) TKS70-2 (ACUC19)",
        "사각 PROBE(70˚) TKS70-4 (H0006)",
        "사각 PROBE(70˚) KA4-70 (8×9) (H6323)",
        "사각 PROBE(70˚) 5C10×10A70 (AFPC023)",
        "사각 PROBE(70˚) MWB70-4 (8×9) (56929-59715)",
        "사각 PROBE(70˚) WB70-2 (20×22) (57001-51556)",
        "사각 PROBE(60˚) KL5-60 (H7793)",
        "사각 PROBE(60˚) KL5-60 (H7794)",
        "사각 PROBE(60˚) KL5-60 (H7798)",
        "사각 PROBE(60˚) MWB60-2 (56922-05004)",
        "사각 PROBE(60˚) MWB60-2 (57000-51382)",
        "사각 PROBE(60˚) MWB60-4 (56928-17156)",
        "사각 PROBE(60˚) MWB60-4 (56928-17283)",
        "사각 PROBE(60˚) MWB60-4 (56928-35552)",
        "사각 PROBE(60˚) MWB60-4 (56929-52302)",
        "사각 PROBE(60˚) MWB60-4 (56928-17312)",
        "사각 PROBE(60˚) TKS60 (ACUB001)",
        "사각 PROBE(60˚) MWB60-4 (56928-57116)",
        "사각 PROBE(60˚) MWB60-4 (56928-52153)",
        "사각 PROBE(60˚) TKS60 (AFPB54)",
        "사각 PROBE(60˚) TKS60 (ACUB17)",
        "사각 PROBE(60˚) TKS60-2 (16×16) (ACUB016)",
        "사각 PROBE(60˚) TKS5C5×5A60 (AFKA98)",
        "사각 PROBE(45˚) 2Z10X10R (JR1399)",
        "사각 PROBE(45˚) KL5-45 (H6005)",
        "사각 PROBE(45˚) KL5-45 (H6006)",
        "사각 PROBE(45˚) HS-7427-2/45 (H9163)",
        "사각 PROBE(45˚) MWB45-2 (56921-6976)",
        "사각 PROBE(45˚) MWB45-4 (56927-56190)",
        "사각 PROBE(45˚) MWB45-4 (56927-58799)",
        "사각 PROBE(45˚) MWB45-4 (56927-26651)",
        "사각 PROBE(45˚) 2C16X16A45 (ACUA017)",
        "사각 PROBE(45˚) 2C14X14A45 (ACUA001)",
        "사각 PROBE(45˚) 2C10X10A45 (ACPA32)",
        "사각 PROBE(45˚) 2C10X10A45 (ACPA33)",
        "사각 PROBE(45˚) MWB45 (56927-35905)",
        "사각 PROBE(45˚) U-PRO45 (D4065)",
        "사각 PROBE(45˚) TKS45 (ACUC31)",
        "사각 PROBE(45˚) TKS45 (ACUA009)",
        "사각 PROBE(45˚) MWB45 (56927-58233)",
        "사각 PROBE(45˚) MWB45 (56927-52290)",
        "사각 PROBE(45˚) TKS5C5×5A45 (AFKA87)",
        "사각 PROBE(30˚) TKS5C5×5A30 (AFKA79)",
        "사각 PROBE(35˚) TKS5C5×5A35 (AFKA81)",
        "사각 PROBE(40˚) TKS5C5×5A40 (AFKA84)",
        "사각 PROBE(50˚) TKS5C5×5A50 (AFKA92)",
        "사각 PROBE(55˚) TKS5C5×5A55 (AFKA93)",
        "수직 PROBE(Φ5) 1C24N (NAGA67)",
        "수직 PROBE(Φ5) A109-S (57744-54161)",
        "수직 PROBE(Φ5) A109-S (57745-76316)",
        "수직 PROBE(Φ5) 5C5N (57745-76327)",
        "수직 PROBE(Φ10) MB2S (57749-52552)",
        "수직 PROBE(Φ10) MB2S (57749-53039)",
        "수직 PROBE(Φ10) MB2S (57749-75353)",
        "수직 PROBE(Φ10) MB2S (57749-75360)",
        "수직 PROBE(Φ10) MB4S (57749-75378)",
        "수직 PROBE(Φ10) MB4S (502788)",
        "수직 PROBE(Φ10) MB4S (503983)",
        "수직 PROBE(Φ10) MB4S (NFBA 56)",
        "수직 PROBE(Φ10) MB4S (NFGIA 15)",
        "수직 PROBE(Φ20) 5C20N (57746-51530)",
        "수직 PROBE(Φ24) B1S (I 130)",
        "수직 PROBE(Φ24) B2S (I 037)",
        "수직 PROBE(Φ24) B2S (I 138)",
        "수직 PROBE(Φ24) B2S (57748-4571)",
        "수직 PROBE(Φ24) SW2-24 (57748-52567)",
        "수직 PROBE(Φ24) SW2-24 (57748-56177)",
        "수직 PROBE(Φ24) SW2-24 (57748-56178)",
        "수직 PROBE(Φ24) B4S (T029092)",
        "수직 PROBE(Φ24) B4S (57746-51052)",
        "수직 PROBE(Φ25) SLF4-25 (57745-07614)",
        "수직 PROBE(Φ10) MB4S (57749-55037)",
        "수직 PROBE(Φ10) MB4S (57749-74386)",
        "수직 PROBE(Φ10) MB4S (57749-71405)",
        "수직 PROBE(Φ10) MB2S (57748-4757)",
        "수직 PROBE(Φ20) TKS5C20N (NFGIA20)",
        "수직 PROBE(Φ24) B4S (57746-51384)",
        "듀얼-사각 PROBE(70˚) 2C8ⅹ14LAD70 (DLANC006)",
        "듀얼-사각 PROBE(기타) 2C6ⅹ18CRD (DCCRJA032)",
        "듀얼-사각 PROBE(기타) 2C8ⅹ14CRD (DCCNA05)",
        "듀얼-사각 PROBE(기타) TRCr2-AUST (12-1944)",
        "듀얼-수직 PROBE(Φ5) MTD705 (869315)",
        "듀얼-수직 PROBE(Φ10) MSEB4 (57462-78164)",
        "듀얼-수직 PROBE(Φ25) SEB4 (57469-50297)",
        "표준시험편 1mm~30mm 10단",
        "표준시험편 ASME C - BCB",
        "표준시험편 ASME RB FBH (Ø3) (SIS-GI-001~006)",
        "표준시험편 ASME 감도교정",
        "표준시험편 ASME 직진성 (SIS-CAL-2201)",
        "표준시험편 AWS DS (35161)",
        "표준시험편 AWS RC",
        "표준시험편 BEARING 시험편 (4EA)",
        "표준시험편 IIW TYPEⅠ(C/S) (V1-CS-25-D46)",
        "표준시험편 IIW TYPEⅠ(STS304) (16-3060)",
        "표준시험편 IIW TYPEⅠ(STS304) (A19374)",
        "표준시험편 IIW (BS-2704)",
        "표준시험편 ISO19675 PAUT CAL BLOCK",
        "표준시험편 ISO19675 PAUT CAL BLOCK (56340)",
        "표준시험편 SS0902 스텝웨지(20mm-4단)",
        "표준시험편 ST1018스텝웨지(30mm-5단) (KDE21371)",
        "표준시험편 STB G V2",
        "표준시험편 STB G V3",
        "표준시험편 STB GV15-1.4",
        "표준시험편 STB GV15-1",
        "표준시험편 STB GV15-2.8",
        "표준시험편 STB GV15-2",
        "표준시험편 STB GV15-4",
        "표준시험편 STB GV15-5.6",
        "표준시험편 STB GV5",
        "표준시험편 STB GV8",
        "표준시험편 STB N1",
        "표준시험편 STB-A1 (2)",
        "표준시험편 STB-A1 (09148)",
        "표준시험편 STB-A1 (1830)",
        "표준시험편 STB-A1 (356)",
        "표준시험편 STB-A1 (No.09159)",
        "표준시험편 STB-A1 (244)",
        "표준시험편 STB-A1 (257)",
        "표준시험편 STB-A2 (5180)",
        "표준시험편 STB-A2 (J-1)",
        "표준시험편 STB-A2 (293)",
        "표준시험편 STB-A2 (299)",
        "표준시험편 STB-A3 (4543)",
        "표준시험편 V3 BLOCK (STS304) (001-211)",
        "대비시험편 RB-4 (C/S) NO. 1 (16-2577)",
        "대비시험편 RB-4 (C/S) NO. 2 (16-2096)",
        "대비시험편 RB-4 (STS304) NO. 1 (18-3903)",
        "대비시험편 RB-4 (STS304) NO. 2 (19-1518)",
        "대비시험편 RB-A7",
        "대비시험편 RB-RA",
        "대비시험편 RB-RB",
        "대비시험편 RB-RC",
        "대비시험편 RB-RE(1), (2), (3)",
        "접촉매질(Couplant)",
        "없음"
      ],
      "default_aux": "사각 PROBE(70˚) MWB70-2 (56923-5617); 표준시험편 STB-A1 (09148)",
      "unit": "M"
    },
    "PT": {
      "equip_main": "PT 약품",
      "aux_equipments": [
        "UV Meter YK-35UV (071550)",
        "조도계 LUX Meter 1330A (180604411)",
        "조도계 LUX Meter 1330A (130310985)",
        "조도계 LUX Meter 1330A (160500222)",
        "표면온도계 1300 (170900753)",
        "표면온도계 TK-307A (E3102035580)",
        "표면온도계 1300 (090204725)",
        "표면온도계 1300 (160701448)",
        "조도계 LS-150 (10004333)",
        "KS-표준시험편 (A 형) A형 대비시험편 (NDI 980058)",
        "KS-표준시험편 (A 형) A형 대비시험편 (614 A/B)",
        "원형통 피 게이지 0.780~15.57mm/10PCS (GCN0002-052-1)",
        "조도계",
        "온습도계",
        "없음"
      ],
      "default_aux": "조도계 LUX Meter 1330A (180604411)",
      "unit": "M"
    },
    "MT": {
      "equip_main": "Yoke AC(MY-2 NW0453)",
      "aux_equipments": [
        "자분살포기",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW0096)",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW0083)",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW0113)",
        "Black Light (베터리교체형) NAWOO 2P (1969815)",
        "Black Light (베터리교체형) NAWOO 2P (1969816)",
        "Black Light (베터리교체형) 3P-BS (NW 0007)",
        "Black Light (베터리교체형) EE-365 (2057820)",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0180)",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0065)",
        "Black Light (베터리교체형) UV-LED 2P/SPR (NW 0196)",
        "Black Light (요크부착형) UV-LED YOKE거치식 (1504050)",
        "Black Light (요크부착형) UV-LED YOKE거치식 (1504062)",
        "UV Meter YK-35UV (071550)",
        "조도계 LUX Meter 1330A (180604411)",
        "조도계 LUX Meter 1330A (130310985)",
        "조도계 LUX Meter 1330A (160500222)",
        "표면온도계 1300 (170900753)",
        "표면온도계 TK-307A (E3102035580)",
        "표면온도계 1300 (090204725)",
        "표면온도계 1300 (160701448)",
        "자장계 (바늘형) Model 20 (13-2393)",
        "자장계 (바늘형) Model 20 (13-2464)",
        "조도계 LS-150 (10004333)",
        "KS-표준시험편 (A 형) A1-7/50 (0101002)",
        "팔각시험편",
        "리프팅파워블럭 (AC 형) 10lb(4.5kg) (NW0938)",
        "리프팅파워블럭 (AC 형) 10lb(4.5kg) (MP19574)",
        "리프팅파워블럭 (AC 형) 10lb(4.5kg) (NW0756)",
        "없음"
      ],
      "default_aux": "자분살포기",
      "unit": "M"
    },
    "PAUT": {
      "equip_main": "PAUT MX2(OMNI2-103085)",
      "aux_equipments": [
        "PAUT SCANNER (MANUAL)",
        "MICROBE 스캐너 (Body-137331, Encoder-136861)",
        "COBRA 스캐너 (139310)",
        "PAUT 스캐너 (133490)",
        "PAUT PROBE(A2) 2.25L64-0.75*12-D62-U-110-2 (GBA008)",
        "PAUT PROBE(A2) 2.25L64-0.75*12-D62-U-110-2 (GBA009)",
        "PAUT PROBE(A2) 2.25L64-A2 (P3378)",
        "PAUT PROBE(A2) 2.25L64-A2 (R0080)",
        "PAUT PROBE(A2) 5L64-06*10-D62-U-110-2.0-P1 (GBA011)",
        "PAUT PROBE(A2) 5L64-A2 (S1599)",
        "PAUT PROBE(A2) 5L64-A2 (S1694)",
        "PAUT PROBE(A2) 5L64-A2 (S5225)",
        "PAUT PROBE(A15) 7.5CCEV35-16-A15 (V4063)",
        "PAUT PROBE(A15) 7.5CCEV35 (V4109)",
        "PAUT PROBE(A25) XAAB-10882 (W2479)",
        "PAUT PROBE(A25) 4DL16-12X5-A25 (U3188)",
        "PAUT PROBE(A31) 5L32-A31 (P4301)",
        "PAUT PROBE(A25) 5DL16-12X5-A25-P-2.5-OM (Y1193)",
        "PAUT PROBE(A10) 5L16-A10 (S1873)",
        "PAUT PROBE(A32) 5L64-32X10-A32 (U1462)",
        "PAUT PROBE(A32) 5L64-32X10-A32 (U14651)",
        "엔코더(PAUT) COBRA-SP-ENC (25A2505)",
        "엔코더(PAUT) (소형 2EA, 대형 2EA)",
        "엔코더(PAUT) CKS007 (135275)",
        "엔코더(PAUT) CKS007 (136682)",
        "케이블(PAUT) (140828-03)",
        "채널분배기(PAUT) (QC-008086)",
        "채널분배기(PAUT) OMNI-A-ADP05 (974106-03)",
        "채널분배기(PAUT) (1011294-01)",
        "모듈(PAUT) Omni-M-PA32 (OMNI-600336)",
        "TomoView Lock Key (1930372742)",
        "Index Control Bar (N/A(Bar 300mm))",
        "엔코더(PAUT) (145552)",
        "채널분배기(PAUT) (QC003541)",
        "충전기 (70180-0004709)",
        "충전기 (70180-0024164)",
        "빔툴키 PRO H (MN720433)",
        "OMNISCAN PC TOOLKIT / OMNIPC-A (134846330)",
        "WELDSIGHT-A (1522192766)",
        "채널분배기(PAUT) Omni-A2-SPLIT128 4UT (QC-080522)",
        "엔코더(PAUT) (132801)",
        "스캐너 (133490)",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC",
        "PAUT PROBE-WEDGE(SA2) SA2-OL-IHC",
        "PAUT PROBE-WEDGE(SA10) SA10-N55S-IHC",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IHC",
        "PAUT PROBE-WEDGE(SA25) SA25-DL70L 평웻지",
        "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD1.05",
        "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD1.315",
        "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD2.875",
        "PAUT PROBE-WEDGE(SA25) SA25-DL70L AOD3.5",
        "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD8",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD10",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD12",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD14",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD16",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD18",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD20",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD24",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD30",
        "PAUT PROBE-WEDGE(SA2) SA2-N55S-IHC-AOD22 DUAL",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD0.84",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.05",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.315",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.66",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD1.9",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD2.375",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD2.875",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD3.25",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD3.5",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD4",
        "PAUT PROBE-WEDGE(SA15) SA15-N60S-IH-AOD4.5",
        "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC-AOD8.625",
        "PAUT PROBE-WEDGE(SA31) SA31-N55S-IHC-AOD12.75",
        "PAUT PROBE-WEDGE(SD62) SD62-N55S-IHC-DUAL",
        "PAUT PROBE-WEDGE(M15) M10-N55S-IHC(PS)",
        "PAUT PROBE-WEDGE(M15) M15-N60S-IH",
        "PAUT PROBE-WEDGE(M15) M15-N60S-IH-AOD1.66",
        "PAUT PROBE-WEDGE(M15) M15-N60S-IH-AOD2.8",
        "PAUT PROBE-WEDGE(SA31) N55S-I HC-PLATE",
        "표준시험편 IIW TYPEⅠ(C/S) (V1-CS-25-D46)",
        "표준시험편 IIW TYPEⅠ(STS304) (16-3060)",
        "표준시험편 IIW TYPEⅠ(STS304) (A19374)",
        "표준시험편 IIW (BS-2704)",
        "표준시험편 ISO19675 PAUT CAL BLOCK",
        "표준시험편 ISO19675 PAUT CAL BLOCK (56340)",
        "표준시험편 STB-A1 (2)",
        "표준시험편 STB-A1 (09148)",
        "표준시험편 STB-A1 (1830)",
        "표준시험편 STB-A1 (356)",
        "표준시험편 STB-A1 (No.09159)",
        "표준시험편 STB-A1 (244)",
        "표준시험편 STB-A1 (257)",
        "없음"
      ],
      "default_aux": "PAUT SCANNER (MANUAL)",
      "unit": "M"
    }
  },
  pipe_sizes: [
    "15A", "20A", "25A", "32A", "40A", "50A", "65A", "80A", "90A", "100A",
    "125A", "150A", "200A", "250A", "300A", "350A", "400A", "450A", "500A",
    "550A", "600A", "650A", "700A", "750A", "800A", "850A", "900A", "950A",
    "1000A", "1050A", "1100A", "1150A", "1200A"
  ],
  units: ["매", "M", "P,M,I/D", "EA", "Point", "SET", "CAN", "KG", "Inch", "Dia", "I/D", "Meter", "일"],
  test_items: ["Piping", "Tank", "Vessel", "Plate", "Beam", "Structure", "Pipe", "사무실 내근 및 장비점검", "우천 대기", "안전교육", "휴무", "기타"],
  applied_codes: ["KS", "ASME", "AWS", "API", "KEPIC", "기타"],
  materials: [
    'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
    'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
    'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
    'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
    'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
    'Carestream T200-10*12"', 'Carestream T200-14*17"'
  ],
  worktimes: [
    "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00",
    "09:00~22:00", "09:00~23:00", "09:00~24:00", "09:00~익일01:00",
    "09:00~익일02:00", "09:00~익일03:00", "18:00~익일02:00",
    "18:00~익일03:00", "22:00~익일09:00", "12:00~익일08:00"
  ]
};

var HEADERS = [
  'Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime',
  'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타',
  'User', '장비명', '검사량', '단가', '출장비', '일식', '검사비', '검사방법', 'FilmCount',
  'WorkTime', 'OT', 'User2', 'WorkTime2', 'OT2', 'User3', 'WorkTime3', 'OT3',
  'User4', 'WorkTime4', 'OT4', 'User5', 'WorkTime5', 'OT5', 'User6', 'WorkTime6', 'OT6',
  'User7', 'WorkTime7', 'OT7', 'User8', 'WorkTime8', 'OT8', 'User9', 'WorkTime9', 'OT9',
  'User10', 'WorkTime10', 'OT10',
  'NDT_형광자분', 'NDT_흑색자분', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광침투제',
  '차량번호', '주행거리', '차량점검', '차량비고', '품목명', 'RTK_총계', '회사코드', '업체명',
  '검사품명', '적용코드', '성적서번호', 'Unit', 'Material', '작업형태',
  '검사구분', '조인트수', '불량수', '관경(Inch)', '구분', 'Remark', 'DailyMemo', '보조장비', '수량', '검사규격', '검사대상', '재검사개소', '재검사매수'
];

/**
 * 모바일 웹 페이지 접속 (doGet)
 */
function doGet(e) {
  var api = e && e.parameter ? e.parameter.api : null;
  
  if (api === 'config') {
    return createJsonResponse(getConfig());
  } else if (api === 'companies') {
    return createJsonResponse({ status: 'success', companies: getCompaniesList() });
  } else if (api === 'search') {
    return createJsonResponse(searchRecords(e.parameter || {}));
  } else if (api === 'recent') {
    var date = e.parameter.date || '';
    return createJsonResponse(getRecentRecords(date));
  } else if (api === 'sync' || api === 'all_records') {
    return createJsonResponse(getAllRecords());
  }

  // 모바일 웹 화면 반환 (HTML)
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('경인사무소 모바일 현장 입력 (24시간 클라우드)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  return htmlOutput;
}

/**
 * 작업 저장 요청 처리 (doPost)
 */
function doPost(e) {
  try {
    var payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }
    
    var action = payload.action || 'save';
    if (action === 'searchRecords' || action === 'search') {
      return createJsonResponse(searchRecords(payload));
    } else if (action === 'addCompany') {
      return createJsonResponse(addCompany(payload));
    } else if (action === 'deleteCompany') {
      return createJsonResponse(deleteCompany(payload));
    } else if (action === 'syncCompanies') {
      return createJsonResponse(syncCompanies(payload));
    } else if (action === 'batch_sync') {
      var res = batchSyncRecords(payload);
      return createJsonResponse(res);
    } else if (action === 'update' || action === 'updateRecord' || (payload.id !== undefined && payload.is_update)) {
      var res = updateRecord(payload);
      return createJsonResponse(res);
    } else if (action === 'delete' || action === 'deleteRecord') {
      var res = deleteRecord(payload);
      return createJsonResponse(res);
    }

    var res = saveRecord(payload);
    return createJsonResponse(res);
  } catch (err) {
    return createJsonResponse({ status: 'error', detail: err.toString() });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 경인사무소 전용 스프레드시트 객체 획득
 */
function getGyeonginSpreadsheet() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = properties.getProperty('GYEONGIN_SPREADSHEET_ID');
  if (!spreadsheetId) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) {
      throw new Error('경인사무소 전용 스프레드시트를 확인할 수 없습니다.');
    }
    spreadsheetId = activeSpreadsheet.getId();
    properties.setProperty('GYEONGIN_SPREADSHEET_ID', spreadsheetId);
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * 활성 스프레드시트의 'DailyUsage' 시트 획득 또는 자동 생성
 */
function getDailyUsageSheet() {
  var ss = getGyeonginSpreadsheet();
  var sheet = ss.getSheetByName("DailyUsage");
  if (!sheet) {
    sheet = ss.insertSheet("DailyUsage");
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e6f4ea");
    sheet.setFrozenRows(1);
  } else {
    // 헤더가 비어있으면 헤더 작성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#e6f4ea");
      sheet.setFrozenRows(1);
    }
  }
  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['검사대상', '검사규격', '재검사개소', '재검사매수'].forEach(function (field) {
    if (currentHeaders.indexOf(field) < 0) {
      currentHeaders.push(field);
      sheet.getRange(1, currentHeaders.length).setValue(field);
    }
  });
  return sheet;
}

/**
 * 경인사무소 전용 'Companies' 마스터 시트 획득 또는 최초 1회 시드 생성
 */
function getCompaniesSheet() {
  var ss = getGyeonginSpreadsheet();
  var sheet = ss.getSheetByName("Companies");
  if (!sheet || sheet.getLastRow() === 0) {
    if (!sheet) sheet = ss.insertSheet("Companies");
    sheet.getRange(1, 1).setValue("업체명").setFontWeight("bold").setBackground("#e6f4ea");
    sheet.setFrozenRows(1);

    var seed = APP_CONFIG.companies.slice();
    var daily = getDailyUsageSheet();
    var dLastRow = daily.getLastRow();
    if (dLastRow > 1) {
      var dHeaders = daily.getRange(1, 1, 1, daily.getLastColumn()).getValues()[0];
      var compIdx = dHeaders.indexOf("업체명");
      if (compIdx !== -1) {
        var dVals = daily.getRange(2, compIdx + 1, dLastRow - 1, 1).getValues();
        for (var i = 0; i < dVals.length; i++) {
          var v = String(dVals[i][0] || '').trim();
          if (v && v.toLowerCase() !== 'nan' && seed.indexOf(v) === -1) seed.push(v);
        }
      }
    }
    if (seed.length > 0) {
      var rows = seed.map(function (name) { return [name]; });
      sheet.getRange(2, 1, rows.length, 1).setValues(rows);
    }
  }
  return sheet;
}

function getCompaniesList() {
  var sheet = getCompaniesSheet();
  var lastRow = sheet.getLastRow();
  var result = ['경인사무소'];
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var name = String(values[i][0] || '').trim();
      if (name && name.toLowerCase() !== 'nan' && result.indexOf(name) === -1) {
        result.push(name);
      }
    }
  }
  return result;
}

function ensureCompanyRegistered(companyName) {
  var name = String(companyName || '').trim();
  if (!name || name.toLowerCase() === 'nan') return false;
  var existing = getCompaniesList();
  if (existing.indexOf(name) !== -1) return false;
  var sheet = getCompaniesSheet();
  sheet.appendRow([name]);
  return true;
}

function addCompany(req) {
  var name = String((req && req.company) || '').trim();
  if (!name) {
    return { status: 'error', detail: '등록할 업체명을 입력해 주세요.' };
  }
  var added = ensureCompanyRegistered(name);
  return {
    status: 'success',
    added: added,
    company: name,
    companies: getCompaniesList(),
    message: added ? ('[' + name + '] 업체가 등록되었습니다.') : ('[' + name + '] 업체가 이미 목록에 있습니다.')
  };
}

function deleteCompany(req) {
  var name = String((req && req.company) || '').trim();
  if (!name) {
    return { status: 'error', detail: '삭제할 업체명을 지정해 주세요.' };
  }
  if (name === '경인사무소') {
    return { status: 'error', detail: '기본 업체 [경인사무소]는 삭제할 수 없습니다.' };
  }
  var sheet = getCompaniesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'error', detail: '삭제할 업체를 찾을 수 없습니다.' };
  }
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var deleted = false;
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim() === name) {
      sheet.deleteRow(i + 2);
      deleted = true;
    }
  }
  if (!deleted) {
    return { status: 'error', detail: '[' + name + '] 업체를 찾을 수 없습니다.' };
  }
  return {
    status: 'success',
    company: name,
    companies: getCompaniesList(),
    message: '[' + name + '] 업체가 목록에서 삭제되었습니다.'
  };
}

function syncCompanies(req) {
  var incoming = (req && Array.isArray(req.companies)) ? req.companies : [];
  var existing = getCompaniesList();
  var toAdd = [];
  for (var i = 0; i < incoming.length; i++) {
    var name = String(incoming[i] || '').trim();
    if (name && name.toLowerCase() !== 'nan' && existing.indexOf(name) === -1 && toAdd.indexOf(name) === -1) {
      toAdd.push(name);
    }
  }
  if (toAdd.length > 0) {
    var sheet = getCompaniesSheet();
    var startRow = sheet.getLastRow() + 1;
    var rows = toAdd.map(function (c) { return [c]; });
    sheet.getRange(startRow, 1, rows.length, 1).setValues(rows);
  }
  return {
    status: 'success',
    added_count: toAdd.length,
    companies: getCompaniesList()
  };
}

/**
 * 설정 데이터 획득 (google.script.run 및 API 공용)
 */
var COMPANY_NDT_PRESETS = {"D&K켐텍": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["Resin 순환 및 발포라인"], "default_code": "KS B 0816 / SIS H 401"}}, "SG엔지니어링": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["0.49톤 겐트리트레인 비파괴검사"], "default_code": ""}}, "경성기공": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["UNIVERSAL JOINT", "SHAFT", "SPINDLE", "YOKE ASS'Y", "FM SPINDLE"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["BEARING KIT ASS'Y", "FLANGE", "SHAFT", "SPINDLE"], "default_code": "KS D 0248 / SIS H 203 (2류) Rev.2"}}, "글로벌 크레인": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["빔행정, H빔 용접"], "default_code": ""}}, "나라삼양감속기": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT, GEAR", "SHAFT, GEAR , PINION", "SHAFT, GEAR , CARRIER", "PINION 외", "SHAFT 외", "OUTPUT SHAFT", "INPUT BEVEL PINION", "PINION,GEAR,SHAFT", "PINION", "MAIN GEAR", "GEAR", "DRIVE GEAR"], "default_code": "KS D 0213 / SIS H 301"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["PINION, GEAR, SHAFT", "PINION SHAFT, BEVEL GEAR"], "default_code": "KS D 0248 / SIS H 203 (1류) Rev.2"}}, "대성기계": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["W1190 FLIGHT DETAIL", "W1520 FLIGHT DETAIL"], "default_code": "ASME Sec.V Art.7 / SIS H 341 (Rev.13)"}}, "대성테크놀로지": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SAMPLING SYSTEM PIPING", "CHEMICAL DOSEING SYSTEM PIPING", "LIFTING LUG + PAD + RIB"], "default_code": "ASME Sec.V&Ⅷ Div.1(2023Ed) / SIS H 441"}}, "대화유공업": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["君津 熱延 Ｆ２－３用上ＳＴＧリトラクトＣＹＬ（クロス追従上・出）"], "default_code": ""}}, "대화콘베어체인": {"MT": {"target_type": "기계부품", "target_types": ["기계부품", "판재/시험편"], "items": ["HEAD SPROCKET", "C-PIN", "LONG PIN", "PIN LINK PLATE (A)", "PIN LINK PLATE (B)", "BUSH LINK PLATE"], "default_code": "KS0213 / SIS H 301"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["HEAD SPROCKET SHAFT"], "default_code": "KS D 0248 / SIS H 203 (2류) Rev.2"}}, "도원이앤씨": {"MT": {"target_type": "기계부품", "target_types": ["기계부품", "판재/시험편"], "items": ["강구조물 용접부", "용접사 기량테스트 / PLATE"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}}, "동아기업DAC": {"PT": {"target_type": "배관", "target_types": ["배관"], "items": ["SPOOL"], "default_code": "KS B 0816 / SIS H 401 (2급, 2군)"}}, "동해펌프": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SWD-H80-CMKSS-AT-R", "SWE-H125 DCMKSS-TT-R", "SWA-H50 CMKSS-TT-R", "SWA-H40S CMKSS-TT-R", "CHEMICAL DOSEING SYSTEM PIPING", "SAMPLING SYSTEM PIPING", "SWC-H100-CMKSS-AT-R-S14", "LIFTING LUG + PAD + RIB"], "default_code": "ASME Sec.V&Ⅷ Div.1(2021Ed) / SIS H 441"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SWE-H125-DCMKSS-TT-R", "SWA-H50-CMKSS-TT-R", "SWA-H40S-CMKSS-TT-R", "SWD-H80-CMKSS-AT-R", "SWC-H100-CMKSS-AT-R-S14"], "default_code": "ASTM A-388 / SIS H 247 Rev.2"}}, "디에이치건설산업(주)": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["오정대교 현장 샤클 (40톤)"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}}, "롯데건설": {"MT": {"target_type": "판재/시험편", "target_types": ["판재/시험편", "기계부품"], "items": ["용접사 기량테스트", "용접사 기량테스트 / PLATE", "M-BED"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}, "PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["STACK SHELL"], "default_code": "KS B 0816 / SIS H 401(2급,2군) Rev.2"}}, "리니트코리아": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["체인블럭용 러그"], "default_code": "KS B 0816 / SIS H 401 (2급, 2군)"}}, "모나스펌프": {"PMI": {"target_type": "PMI", "target_types": ["PMI"], "items": ["PUMP CASING", "SHAFT", "IMPELLER", "FLANGE", "VALVE"], "default_code": ""}, "PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["CAUSTIC SODA DOSING PUMP (236-J-05A/B)", "ACID GUMS PUMP (236-J-09A/B)", "GUM TRNASFER PUMP (237-J-13A/B)", "LIME DOSING PUMP A/B (00GBN05AP001 / 00GBN06AP001)", "CLARIFIER SLUDGE TRANSFER PUMP A/B (00GNK36AP001 /", "SETTLED SLUDGE TRANSFER PUMP A/B (00GNK55AP001 / 0", "TK-484-13A/B"], "default_code": "A000-ROT-VP235-J108-COMMON"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["CAUSTIC SODA DOSING PUMP (236-J-05A/B)", "ACID GUMS PUMP ( 236-J-09 A/B )", "GUM TRANSFER PUMP ( 237-J-13 A/B )", "SLUDGE FEED PUMP ( P-305841 B )", "SLUDGE FEED PUMP ( P-305841 A/B )", "C-POLYMER INJECTION PUMP ( P-305843 A/B )", "LIME DOSING PUMP A/B (00GBN05AP001 / 00GBN06AP001)", "CLARIFIER SLUDGE TRANSFER PUMP A/B (00GNK36AP001 /", "SETTLED SLUDGE TRANSFER PUMP A/B (00GNK55AP001 / 0"], "default_code": "A000-ROT-VP235-J108-COMMON"}}, "부강풀리": {"UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT"], "default_code": "KS D 0248 / SIS H 203 (1류) Rev.2"}}, "삼남엔지니어링(주)": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["M-BED"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}}, "삼성파워텍": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["외함 이동 받침대"], "default_code": ""}}, "삼승종합건설": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["미래에셋 및 NcSoft 수혈관로 설치공사"], "default_code": ""}}, "삼화테크": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["BELLOWS HOSE"], "default_code": ""}}, "서울발전본부": {"MT": {"target_type": "배관", "target_types": ["배관"], "items": ["PIPE(20A, 80A)"], "default_code": "KS D 0213 / SIS H 301"}, "UT": {"target_type": "배관", "target_types": ["배관"], "items": ["PIPE(80A)"], "default_code": "KS B 0896 / SIS H 201 Rev.3"}}, "서진인스텍": {"PMI": {"target_type": "PMI", "target_types": ["PMI"], "items": ["LEVEL GAUGE", "FLANGE", "VALVE"], "default_code": ""}}, "성헌": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["CROSS BEARING", "추진축 조립체", "UNIVERSAL JOINT"], "default_code": "KDS 0050-0021 / SIS H 301(2급,2군)"}}, "세진공조기계": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["임펠라(IMPELLER)"], "default_code": "KS B 0816 / SIS H 401 (2급2군)"}}, "수성웹툰(주)": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["BATTERY FORK LIFT(3단 MAST) 1.3TON", "BATTERY BABY CRANE 1.0TON CAPA", "전동 포크리프트 수성샐바시온 SWR-1300L"], "default_code": "KS D 0213/SIS H 301(2급,2군)"}}, "스페이스 체인지": {"MT": {"target_type": "구조물/설비", "target_types": ["구조물/설비"], "items": ["상부 슬라브 접합 외 5개"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}}, "신성플랜트": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["Spray Tower Top Cone"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}}, "아이조음": {"MT": {"target_type": "구조물/설비", "target_types": ["구조물/설비"], "items": ["보강대"], "default_code": "KS D 0213 / SIS H 301"}}, "아이피에스": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["CV Seat", "MSV, CV Stem"], "default_code": "KS B 0816 / SIS H 401 (2급, 2군)"}}, "양영": {"MT": {"target_type": "판재/시험편", "target_types": ["판재/시험편", "기계부품"], "items": ["용접사 기량테스트 (Plate,Fillet)", "구조물 용접부"], "default_code": "KS D 0213 /SIS H 301 (2급,2군)"}}, "에버랜드": {"MT": {"target_type": "기계부품", "target_types": ["기계부품", "구조물/설비"], "items": ["롤링X트레인 / C23 레일 파이프 외", "캐치카 / Pusher Dog 외", "페스티벌 트레인 / 차륜 SHAFT", "매직스윙 / 안전레바 PIN", "롤링 X 트레인 / X 텐션바 철판 용접부", "플래쉬팡팡 / 승용물브라켓 외", "피터팬 / 안전랩바 고정핀 외", "롤링 X 트레인 / 레일 용접부", "챔피언쉽로데오 / 안전레버 라쳇 bar 외", "렛츠 트위스트"], "default_code": "ISO 23278_2015_EN level 1 / SIS H 363"}, "PT": {"target_type": "기계부품", "target_types": ["기계부품", "구조물/설비"], "items": ["비룡열차 / 전륜Frame 외", "피터팬 / 안전렙바 용접부", "롤링X트레인(청차) / 가이드 휠 외"], "default_code": "KS B 0816 / SIS H 401 (2급2군)"}, "TG": {"target_type": "구조물/설비", "target_types": ["구조물/설비"], "items": ["롤링 X 트레인 / X 텐션바 철판 두께측정"], "default_code": ""}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["페스티벌 트레인 / 차륜 SHAFT", "매직스윙 / 주행축 외", "콜럼버스 대탐험 / 헤드바 킹핀", "자동차 왕국 / 휠 샤프트 외", "롤링X트레인(홍차) / 휠캐리어 스핀들 외", "더블 락스핀 / 굴절암 샤프트", "플래쉬팡팡 / 실린더로드", "피터팬 / 안전랩바 고정핀 외", "롤링X트레인 / 지주대 베이스볼트", "챔피언쉽 로데오 / 휠 디스크 연결핀 외"], "default_code": "KS D 0248 / SIS H 203 (1류) Rev.2"}}, "에스엠테크": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["VACUUM SYSTEM EXCHANG / WHEEL"], "default_code": "KS D 0213 /포스코 검사기준서 공통 02 / SIS H 301 ("}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["VACUUM SYSTEM EXCHANG / WHEEL", "VACUUM SYSTEM EXCHANG / SHAFT", "Sprocket / Shaft"], "default_code": "KS D 0248 / 포스코 공통 기준서 / 02 SIS H 203 (2"}}, "에이스냉동공조": {"TG": {"target_type": "배관", "target_types": ["배관"], "items": ["MDI IN & OUT 배관"], "default_code": ""}}, "올스톤산업": {"MT": {"target_type": "판재/시험편", "target_types": ["판재/시험편"], "items": ["GFRC 제작공장(지유글로벌) 용접사 기량테스트"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}}, "위드엔지니어링": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["각형강관, 원형강관"], "default_code": ""}}, "유성씨앤에프": {"TG": {"target_type": "구조물/설비", "target_types": ["구조물/설비"], "items": ["내압 두께측정"], "default_code": "KS B 0536"}}, "의성기어": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["REOUCER(R04) / SPIRAL BEVEL GEAR SPEC"], "default_code": "KS D 0213 / SIS H 301"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["REOUCER(R04) / SPIRAL BEVEL GEAR SPEC"], "default_code": "KS D 0248 / SIS H 203 (1류) Rev.3"}}, "이노텍코리아": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["다이크 작업대차"], "default_code": "KS D 0213 / SIS H 301"}}, "자인건축": {"MT": {"target_type": "기계부품", "target_types": ["기계부품", "구조물/설비"], "items": ["UNIVERSAL JOINT", "H-Beam 용접부"], "default_code": "KS D 0213 /SIS H 301 (2급,2군)"}}, "전진이앤씨": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["S"], "default_code": ""}}, "제이에스티": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["WORM SHAFT", "GEAR, PINION, SHAFT"], "default_code": "KS D 0213 / SIS H 301(2급,2군)"}}, "제일푸리": {"PT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT"], "default_code": "ASME Sec.V / SIS H 441"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT", "RC FAN SHAFT"], "default_code": "ASME Sec.Ⅴ (SA-388) / SIS H 247 (2)"}}, "청원기연": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["하부구조물 (가이드레일)", "Buket LUG", "인력케이지 LUG"], "default_code": "KS D 0213/SIS H 301(2급,2군)"}}, "캐논코리아": {"MT": {"target_type": "배관", "target_types": ["배관"], "items": ["Resin 순환 및 발포라인"], "default_code": ""}}, "캐리비안베이": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["캐리비언 사업장 內 유기시설 및 공연시설 NDT 검사"], "default_code": ""}}, "코아베큠": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["BUBBLER CANISTER"], "default_code": ""}}, "코오롱글로벌": {"MT": {"target_type": "기계부품", "target_types": ["기계부품", "구조물/설비"], "items": ["강서IC Ramp-C교 A1 / PHC 말뚝", "오정대교 현장 MT-01호 설치러그", "오정대교 현장 MT-02호 설치러그"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["오정대교 수원방향", "오정대교 문산방향"], "default_code": "KS B 0896 (L검출레벨. 2류) / SIS H 201 Rev.3"}}, "태성기전": {"UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT"], "default_code": "KS D 0248 / SIS H 203 (2류) Rev.2"}}, "태창기계공업": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["Back Up Roll / M43131914(02)", "CPL Cassette 부품(Top Cassette 인양용 러그)/ M43161914(03"], "default_code": "KS D 0213 / SIS H 301"}}, "포스코이앤씨": {"MT": {"target_type": "판재/시험편", "target_types": ["판재/시험편"], "items": ["용접사 기량시험"], "default_code": ""}}, "포스코인터내셔널": {"MT": {"target_type": "배관", "target_types": ["배관"], "items": ["#70 HRSG DRAIN LINE"], "default_code": "KS D 0213 / SIS H 301 (2급, 2군)"}, "TG": {"target_type": "구조물/설비", "target_types": ["구조물/설비", "배관"], "items": ["IP LOWER RISER PIPE", "LP LOWER RISER PIPE", "IP EVA LOWER TUBE", "IP EVA UPPER TUBE", "LP EVA LOWER TUBE", "LP EVA UPPER TUBE", "#70 HRSG IP EVA LOWER #3 HEADER TUBE", "#70 HRSG IP EVA LOWER #2 HEADER TUBE", "#70 HRSG IP EVA LOWER #1 HEADER TUBE", "#70 HRSG IP EVA LOWER #2 RISER MANIFOLD", "#70 HRSG IP EVA LOWER #3 RISER MANIFOLD", "#70 HRSG IP EVA LOWER #1 RISER MANIFOLD"], "default_code": ""}}, "피테크": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["LIFTING LUG"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}}, "한양스틸": {"UT": {"target_type": "배관", "target_types": ["배관"], "items": ["FORGED STEEL"], "default_code": "SEP 1921 Group3 Class D/d / SIS H 265 Re"}}, "함감속기제작소": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["EXCITER 406 / 408", "PT 400HP x 1150RPM x 1/24.75 C=920", "EXCITER 406/408", "KTS 290HP x 470RPM x 1/2.909", "KTS 280HP × 2100RPM × 1/3.667", "EXCITER 306/308", "KPSI 1150HP × 4000RPM × 1/2.4 , 1/1/4.8", "KTD 290HP x 1454RPM x 1/8.992", "KTD 280HP × 2100RPM × 1/5.134", "KTT 112HP × 2262RPM × 1/7.334", "BS 80HP × 1420RPM × 1/4.09", "KTS 230HP × 2100RPM × 1/3.667"], "default_code": "KS D 0213 / SIS H 301"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["EXCITER 406/408", "EXCITER 406 /408", "BS 80HP x 1420RPM x 1/4.09", "EXCITER 206/208", "KPSI 1150HP x 4000RPM x 1/2.4, 1/4.8", "EXCITER 306/308", "KTS 280HP x 2100RPM x 1/3.667", "KTD 280HP x 2100RPM x 1/5.134", "KTT 112HP X 2262RPM X 1/7.334", "KTD 280HP X 2100RPM X 1/5.134", "EXCITER 306 /308", "KTS 280HP × 2100RPM × 1/3.667"], "default_code": "KS D 0248 / SIS H 203 (1류) Rev.2"}}, "현대건설": {"MT": {"target_type": "구조물/설비", "target_types": ["구조물/설비", "기계부품"], "items": ["탄천동로 Top Down 중앙말뚝 이음 용접검사", "누드E/V철골 용접", "추진기지#1 추진관 용접부 검사"], "default_code": "KS D 0213 (2급,2군) / SIS H 301"}, "PMI": {"target_type": "PMI", "target_types": ["PMI"], "items": ["Girder", "Tie-bar", "Beam", "철골"], "default_code": ""}, "UT": {"target_type": "판재/시험편", "target_types": ["판재/시험편"], "items": ["용접사 기량테스트 / PLATE"], "default_code": "KS B 0896 (L검출레벨. 2류) / SIS H 201 Rev.3"}}, "현대산업": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["CONVEYOR PULLEY", "HEAD PULLEY (CV-012)", "TAIL PULLEY (CV-012)", "SNUB PULLEY (CV-012)", "TAIL DRUM (CV-017)", "SNUB&BEND PULLEY(CV-013)", "TAKE-UP PULLEY (CV-013)", "TAIL PULLEY(CV-013)", "HEAD(TRIPPER) PULLEY(CV-013)", "HEAD PULLEY(CV-013)"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}, "UT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["SHAFT"], "default_code": "KS D 0248 / SIS H 203 (2류) Rev.2"}}, "현성펌프": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["BFP Upper casing"], "default_code": "KS D 0213 / SIS H 301 (2급,2군)"}}, "현진산업": {"MT": {"target_type": "기계부품", "target_types": ["기계부품"], "items": ["화성 EVO 조립공장 로봇설치"], "default_code": ""}}, "SK에코플랜트": {"PT": {"target_type": "배관", "target_types": ["배관"], "items": ["강관 Pipe (1500A)", "도수관로", "PIPE"], "default_code": "KS B 0888, KS B 0816 / SIS H 401SKP"}}, "K&M": {"RT": {"target_type": "배관", "target_types": ["배관"], "items": ["PIPE (150A)", "PIPE", "WQT (PIPE)"], "default_code": "ASME"}}, "케이엔엠": {"RT": {"target_type": "배관", "target_types": ["배관"], "items": ["150A PIPE WQT(6G)", "PIPE"], "default_code": "ASME"}}, "동남ENG": {"RT": {"target_type": "배관", "target_types": ["배관"], "items": ["WQT PIPE / 25A", "WQT (PIPE)", "PIPE"], "default_code": "ASME"}}, "한국산업인력공단": {"RT": {"target_type": "판재/시험편", "target_types": ["판재/시험편"], "items": ["PLATE / 301~332", "PLATE"], "default_code": "KS B 0845"}}, "영흥발전": {"PAUT": {"target_type": "배관", "target_types": ["배관"], "items": ["#4 Final RH Bundle tube", "BOILER TUBE", "PIPE"], "default_code": "SIS N 203 KOEN-YH Rev.0"}}};

function getConfig() {
  var cfg = JSON.parse(JSON.stringify(APP_CONFIG));
  cfg.companies = getCompaniesList();
  cfg.company_ndt_presets = COMPANY_NDT_PRESETS;
  return cfg;
}

/**
 * 작업 레코드 시트에 저장
 */
function buildRowMap(req, preserveEntryTime) {
  var nowStr = preserveEntryTime || (req.entry_time ? String(req.entry_time) : Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"));
  
  // PC 프로그램에서 직접 보낸 row_data 맵이 있는 경우 그대로 매핑
  if (req.row_data && typeof req.row_data === 'object') {
    var directMap = {};
    for (var k in req.row_data) {
      if (req.row_data.hasOwnProperty(k)) {
        directMap[k] = req.row_data[k];
      }
    }
    directMap['EntryTime'] = nowStr;
    return directMap;
  }

  var calculatedFee = Number(req.total_fee || 0);
  if (calculatedFee <= 0) {
    calculatedFee = Math.round((Number(req.qty || 0) * Number(req.unit_price || 0)) + Number(req.travel_cost || 0));
  }

  var rowMap = {
    'Date': req.date || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"),
    'Site': req.site || '경인사무소',
    '작업구분': req.loc_type || '경인사무소',
    '구분': req.loc_type || '경인사무소',
    '검사구분': req.loc_type || '경인사무소',
    '업체명': req.company || '경인사무소',
    '회사코드': String(req.company_code || '').trim(),
    '검사방법': req.method || 'RT',
    '검사대상': req.inspection_target || '배관',
    '검사규격': String(req.inspection_spec || '').trim(),
    '규격': req.inspection_target === '배관 외' ? '' : (req.pipe || ''),
    '관경(Inch)': req.inspection_target === '배관 외' ? '' : (req.pipe || ''),
    '재검사개소': Number(req.rep_joints || 0),
    '재검사매수': Number(req.rep_qty || 0),
    '검사품명': req.item_name || 'Piping',
    '자재명': req.material || '',
    'MaterialID': req.material || '',
    'Material': req.material || '',
    '품목명': req.material || '',
    '적용코드': req.applied_code || 'KS',
    '성적서번호': req.report_no || '',
    '조인트수': req.joints || req.ori_joints || 0,
    '불량수': req.rej_qty || 0,
    '수량': Number(req.qty || 0),
    '검사량': Number(req.qty || 0),
    'Usage': Number(req.qty || 0),
    '단위': req.unit || '매',
    'Unit': req.unit || '매',
    '단가': Number(req.unit_price || 0),
    '출장비': Number(req.travel_cost || 0),
    '일식': req.lump_sum || '',
    '검사비': calculatedFee,
    '장비명': req.equip_main || '',
    '보조장비': req.equip_aux || '',
    '작업형태': req.work_shape || '일반',
    '차량번호': req.vehicle_no || '',
    '주행거리': req.mileage || '',
    '차량점검': '양호',
    '차량비고': req.vehicle_note || '',
    'Note': req.note || '',
    'Remark': req.note || '',
    'DailyMemo': req.note || '',
    'EntryTime': nowStr,
    'FilmCount': (req.method === 'RT') ? Number(req.qty || 0) : 0,
    // RTK (8)
    'RTK_센터미스': Number(req.rtk_center_miss || 0),
    'RTK_농도': Number(req.rtk_density || 0),
    'RTK_마킹미스': Number(req.rtk_marking_miss || 0),
    'RTK_필름마크': Number(req.rtk_film_mark || 0),
    'RTK_취급부주의': Number(req.rtk_handling || 0),
    'RTK_고객불만': Number(req.rtk_complaint || 0),
    'RTK_기타': Number(req.rtk_etc || 0),
    'RTK_총계': Number(req.rtk_total || 0),
    // NDT Chemicals (7)
    'NDT_형광자분': Number(req.chem_fl_mag || 0),
    'NDT_흑색자분': Number(req.chem_black_mag || 0),
    'NDT_백색페인트': Number(req.chem_white_paint || 0),
    'NDT_침투제': Number(req.chem_penetrant || 0),
    'NDT_세척제': Number(req.chem_cleaner || 0),
    'NDT_현상제': Number(req.chem_developer || 0),
    'NDT_형광침투제': Number(req.chem_fl_pen || 0)
  };

  // Reset worker columns
  for (var u = 1; u <= 10; u++) {
    var uKey = (u === 1) ? 'User' : ('User' + u);
    var wtKey = (u === 1) ? 'WorkTime' : ('WorkTime' + u);
    var otKey = (u === 1) ? 'OT' : ('OT' + u);
    rowMap[uKey] = '';
    rowMap[wtKey] = '';
    rowMap[otKey] = 0;
  }

  // 참여 작업자 할당 (1 ~ 10)
  if (req.workers && req.workers.length > 0) {
    rowMap['User'] = req.workers[0].user || '';
    for (var i = 0; i < req.workers.length && i < 10; i++) {
      var w = req.workers[i];
      var idx = i + 1;
      var uKey = (idx === 1) ? 'User' : ('User' + idx);
      var wtKey = (idx === 1) ? 'WorkTime' : ('WorkTime' + idx);
      var otKey = (idx === 1) ? 'OT' : ('OT' + idx);
      rowMap[uKey] = w.user || '';
      rowMap[wtKey] = w.time || '';
      var cleanOt = Number(String(w.ot || 0).replace(/,/g, '').trim()) || 0;
      rowMap[otKey] = cleanOt;
    }
  }

  return rowMap;
}

/**
 * 작업 레코드 시트에 저장
 */
function saveRecord(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowMap = buildRowMap(req);

  // 헤더 순서대로 배열 생성
  var newRow = [];
  for (var h = 0; h < headerRow.length; h++) {
    var headerName = headerRow[h];
    newRow.push(rowMap.hasOwnProperty(headerName) ? rowMap[headerName] : "");
  }

  sheet.appendRow(newRow);
  ensureCompanyRegistered(rowMap['업체명']);

  return {
    status: 'success',
    companies: getCompaniesList(),
    message: '구글 드라이브 클라우드에 성공적으로 저장되었습니다.'
  };
}

function formatCellTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  }
  return String(val).split('.')[0].trim();
}

/**
 * 작업 레코드 수정
 */
function updateRecord(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");
  var lastRow = sheet.getLastRow();
  var targetRow = -1;

  if (entryTimeCol !== -1 && req.entry_time && lastRow > 1) {
    var reqTimeNorm = String(req.entry_time).split('.')[0].trim();
    var timeValues = sheet.getRange(2, entryTimeCol + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < timeValues.length; i++) {
      var cellTimeNorm = formatCellTime(timeValues[i][0]);
      if (cellTimeNorm === reqTimeNorm) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1 && req.id !== undefined && req.id !== null) {
    var rowNum = Number(req.id);
    if (rowNum >= 2 && rowNum <= lastRow) {
      targetRow = rowNum;
    }
  }

  if (targetRow === -1) {
    return { status: 'error', detail: '수정할 일보 내역을 찾을 수 없습니다.' };
  }

  var rowMap = buildRowMap(req, req.entry_time);
  var rowValues = [];
  for (var j = 0; j < headerRow.length; j++) {
    var colName = headerRow[j];
    rowValues.push(rowMap.hasOwnProperty(colName) ? rowMap[colName] : '');
  }
  sheet.getRange(targetRow, 1, 1, headerRow.length).setValues([rowValues]);
  ensureCompanyRegistered(rowMap['업체명']);
  return { status: 'success', companies: getCompaniesList(), message: '성공적으로 수정되었습니다.' };
}

/**
 * 작업 레코드 삭제
 */
function deleteRecord(req) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");
  var lastRow = sheet.getLastRow();
  var targetRow = -1;

  if (entryTimeCol !== -1 && req.entry_time && lastRow > 1) {
    var reqTimeNorm = String(req.entry_time).split('.')[0].trim();
    var timeValues = sheet.getRange(2, entryTimeCol + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < timeValues.length; i++) {
      var cellTimeNorm = formatCellTime(timeValues[i][0]);
      if (cellTimeNorm === reqTimeNorm) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1 && req.id !== undefined && req.id !== null) {
    var rowNum = Number(req.id);
    if (rowNum >= 2 && rowNum <= lastRow) {
      targetRow = rowNum;
    }
  }

  if (targetRow === -1) {
    return { status: 'error', detail: '삭제할 일보 내역을 찾을 수 없습니다.' };
  }

  sheet.deleteRow(targetRow);
  return { status: 'success', message: '성공적으로 삭제되었습니다.' };
}

/**
 * PC 프로그램 일괄 동기화 (추가, 수정, 삭제 일괄 처리)
 */
function batchSyncRecords(payload) {
  var sheet = getDailyUsageSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var entryTimeCol = headerRow.indexOf("EntryTime");

  var deletesCount = 0;
  var updatesCount = 0;
  var savesCount = 0;

  // 1. 삭제 처리
  var deletes = payload.deletes || [];
  if (deletes.length > 0 && entryTimeCol !== -1) {
    for (var d = 0; d < deletes.length; d++) {
      var targetTime = String(deletes[d]).split('.')[0].trim();
      var curLastRow = sheet.getLastRow();
      if (curLastRow <= 1) break;
      var timeVals = sheet.getRange(2, entryTimeCol + 1, curLastRow - 1, 1).getValues();
      for (var r = timeVals.length - 1; r >= 0; r--) {
        var cellTime = formatCellTime(timeVals[r][0]);
        if (cellTime === targetTime) {
          sheet.deleteRow(r + 2);
          deletesCount++;
          break;
        }
      }
    }
  }

  // 2. 수정 처리
  var updates = payload.updates || [];
  for (var u = 0; u < updates.length; u++) {
    var resU = updateRecord(updates[u]);
    if (resU.status === 'success') updatesCount++;
  }

  // 3. 신규 추가 처리
  var saves = payload.saves || [];
  for (var s = 0; s < saves.length; s++) {
    var resS = saveRecord(saves[s]);
    if (resS.status === 'success') savesCount++;
  }

  return {
    status: 'success',
    added_count: savesCount,
    updated_count: updatesCount,
    deleted_count: deletesCount,
    message: '일괄 동기화 완료 (추가: ' + savesCount + '건, 수정: ' + updatesCount + '건, 삭제: ' + deletesCount + '건)'
  };
}

function normalizeSheetDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var s = String(val).trim().replace(/[.\/]/g, '-');
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return String(val).substring(0, 10);
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function sheetItemToRecord(item, rowNumber) {
  var rowDateStr = normalizeSheetDateStr(item['Date']);
  var workersText = [];
  var workerObjects = [];
  for (var u = 1; u <= 10; u++) {
    var userKey = u === 1 ? 'User' : 'User' + u;
    var timeKey = u === 1 ? 'WorkTime' : 'WorkTime' + u;
    var otKey = u === 1 ? 'OT' : 'OT' + u;
    var userName = String(item[userKey] || '').trim();
    if (!userName) continue;
    var userTime = String(item[timeKey] || '09:00~18:00').trim();
    var userOt = Number(item[otKey] || 0);
    workersText.push(userName + ' (' + userOt.toLocaleString() + '원)');
    workerObjects.push({
      user: userName,
      shift: String(item['작업형태'] || '일반'),
      time: userTime,
      ot: userOt.toLocaleString()
    });
  }

  var extraList = [];
  var extraColumns = [
    ['NDT_형광자분', '형광자분'], ['NDT_흑색자분', '흑색자분'],
    ['NDT_백색페인트', '백색페인트'], ['NDT_침투제', '침투제'],
    ['NDT_세척제', '세척제'], ['NDT_현상제', '현상제'],
    ['NDT_형광침투제', '형광침투제']
  ];
  for (var e = 0; e < extraColumns.length; e++) {
    var extraValue = Number(item[extraColumns[e][0]] || 0);
    if (extraValue > 0) extraList.push(extraColumns[e][1] + ':' + extraValue);
  }
  var rtkTotal = Number(item['RTK_총계'] || 0);
  if (rtkTotal > 0) extraList.push('RTK:' + rtkTotal);

  return {
    id: rowNumber,
    entry_time: formatCellTime(item['EntryTime']) || String(item['EntryTime'] || ''),
    date: rowDateStr,
    site: String(item['Site'] || item['작업구분'] || item['구분'] || '경인사무소'),
    company: String(item['업체명'] || '경인사무소'),
    company_code: String(item['회사코드'] || ''),
    method: String(item['검사방법'] || ''),
    item_name: String(item['검사품명'] || 'Piping'),
    material: String(item['자재명'] || item['Material'] || item['MaterialID'] || item['품목명'] || ''),
    loc_type: String(item['작업구분'] || item['검사구분'] || item['구분'] || '경인사무소'),
    inspection_target: String(item['검사대상'] || ((item['관경(Inch)'] || item['규격']) ? '배관' : '배관 외')),
    inspection_spec: String(item['검사규격'] || ''),
    pipe: item['검사대상'] === '배관 외' ? '' : String(item['관경(Inch)'] || item['규격'] || ''),
    ori_qty: Number(item['Usage'] || item['수량'] || 0),
    rep_qty: Number(item['재검사매수'] || 0),
    rep_joints: Number(item['재검사개소'] || 0),
    joints: Number(item['조인트수'] || 0),
    ori_joints: Number(item['조인트수'] || 0),
    rej_qty: Number(item['불량수'] || 0),
    qty: Number(item['수량']) || Number(item['Usage']) || Number(item['검사량']) || 0,
    unit: String(item['Unit'] || item['단위'] || '매'),
    unit_price: Number(item['단가'] || 0),
    travel_cost: Number(item['출장비'] || 0),
    lump_sum: String(item['일식'] || ''),
    total_fee: Number(item['검사비'] || 0),
    applied_code: String(item['적용코드'] || ''),
    report_no: String(item['성적서번호'] || ''),
    equip_main: String(item['장비명'] || ''),
    equip_aux: String(item['보조장비'] || ''),
    vehicle_no: String(item['차량번호'] || ''),
    mileage: String(item['주행거리'] || ''),
    vehicle_check: String(item['차량점검'] || ''),
    vehicle_note: String(item['차량비고'] || ''),
    note: String(item['DailyMemo'] || item['Remark'] || item['Note'] || ''),
    workers: workersText.join(', '),
    worker_list: workerObjects,
    rtk_center_miss: Number(item['RTK_센터미스'] || 0),
    rtk_density: Number(item['RTK_농도'] || 0),
    rtk_marking_miss: Number(item['RTK_마킹미스'] || 0),
    rtk_film_mark: Number(item['RTK_필름마크'] || 0),
    rtk_handling: Number(item['RTK_취급부주의'] || 0),
    rtk_complaint: Number(item['RTK_고객불만'] || 0),
    rtk_etc: Number(item['RTK_기타'] || 0),
    rtk_total: rtkTotal,
    chem_fl_mag: Number(item['NDT_형광자분'] || 0),
    chem_black_mag: Number(item['NDT_흑색자분'] || 0),
    chem_white_paint: Number(item['NDT_백색페인트'] || 0),
    chem_penetrant: Number(item['NDT_침투제'] || 0),
    chem_cleaner: Number(item['NDT_세척제'] || 0),
    chem_developer: Number(item['NDT_현상제'] || 0),
    chem_fl_pen: Number(item['NDT_형광침투제'] || 0),
    extra: extraList.join(' ')
  };
}

/**
 * 경인사무소 전용 작업일보 검색.
 * 날짜+현장 그룹을 기준으로 20개씩 반환한다.
 */
function searchRecords(params) {
  params = params || {};
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'success', groups: [], has_more: false, next_offset: 0, total_groups: 0, total_records: 0 };
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
  var startDate = String(params.start_date || '');
  var endDate = String(params.end_date || '');
  var keyword = String(params.keyword || '').trim().toLowerCase();
  var methods = params.methods || [];
  if (typeof methods === 'string') methods = methods ? methods.split(',') : [];
  methods = methods.map(function (method) { return String(method).trim(); }).filter(String);
  var offset = Math.max(0, Number(params.offset || 0));
  var limit = Math.min(50, Math.max(1, Number(params.limit || 20)));
  var records = [];

  for (var r = 0; r < data.length; r++) {
    var item = {};
    for (var c = 0; c < headerRow.length; c++) item[headerRow[c]] = data[r][c];
    var record = sheetItemToRecord(item, r + 2);
    if (!record.date) continue;
    if (startDate && record.date < startDate) continue;
    if (endDate && record.date > endDate) continue;
    if (methods.length && methods.indexOf(record.method) === -1) continue;
    if (keyword) {
      var searchable = [record.site, record.company, record.report_no].join(' ').toLowerCase();
      if (searchable.indexOf(keyword) === -1) continue;
    }
    records.push(record);
  }

  records.sort(function (a, b) {
    var aKey = a.date + ' ' + a.entry_time;
    var bKey = b.date + ' ' + b.entry_time;
    if (aKey === bKey) return b.id - a.id;
    return aKey < bKey ? 1 : -1;
  });

  var groupMap = {};
  var groups = [];
  for (var i = 0; i < records.length; i++) {
    var current = records[i];
    var groupKey = current.date + '|' + current.site;
    if (!groupMap[groupKey]) {
      groupMap[groupKey] = {
        key: groupKey,
        date: current.date,
        site: current.site,
        methods: [],
        record_count: 0,
        total_fee: 0,
        records: []
      };
      groups.push(groupMap[groupKey]);
    }
    var group = groupMap[groupKey];
    group.records.push(current);
    group.record_count++;
    group.total_fee += Number(current.total_fee || 0);
    if (current.method && group.methods.indexOf(current.method) === -1) group.methods.push(current.method);
  }

  var page = groups.slice(offset, offset + limit);
  return {
    status: 'success',
    groups: page,
    has_more: offset + limit < groups.length,
    next_offset: offset + page.length,
    total_groups: groups.length,
    total_records: records.length
  };
}

/**
 * 최근 저장된 내역 조회 (당일 또는 최근 15건)
 */
function getRecentRecords(targetDate) {
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { records: [] };
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var startRow = Math.max(2, lastRow - 30);
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, 1, numRows, headerRow.length).getValues();

  var records = [];
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var item = {};
    for (var c = 0; c < headerRow.length; c++) {
      item[headerRow[c]] = row[c];
    }

    var rowDateStr = "";
    if (item['Date'] instanceof Date) {
      rowDateStr = Utilities.formatDate(item['Date'], "Asia/Seoul", "yyyy-MM-dd");
    } else {
      rowDateStr = String(item['Date'] || '').substring(0, 10);
    }

    if (targetDate && rowDateStr !== targetDate) {
      continue;
    }

    // 작업자 요약
    var workersArr = [];
    var workerObjects = [];
    for (var u = 1; u <= 10; u++) {
      var uName = item[(u === 1) ? 'User' : ('User' + u)];
      if (uName && String(uName).trim()) {
        var uTime = item[(u === 1) ? 'WorkTime' : ('WorkTime' + u)] || '09:00~18:00';
        var otPay = Number(item[(u === 1) ? 'OT' : ('OT' + u)] || 0);
        workersArr.push(String(uName).trim() + " (" + otPay.toLocaleString() + "원)");
        workerObjects.push({
          user: String(uName).trim(),
          shift: '일반',
          time: String(uTime).trim() || '09:00~18:00',
          ot: otPay.toLocaleString()
        });
      }
    }

    // 약품 / RTK 요약
    var chemList = [];
    if (Number(item['NDT_형광자분'] || 0) > 0) chemList.push("형자:" + item['NDT_형광자분']);
    if (Number(item['NDT_흑색자분'] || 0) > 0) chemList.push("흑자:" + item['NDT_흑색자분']);
    if (Number(item['NDT_백색페인트'] || 0) > 0) chemList.push("백페:" + item['NDT_백색페인트']);
    if (Number(item['NDT_침투제'] || 0) > 0) chemList.push("침투:" + item['NDT_침투제']);
    if (Number(item['NDT_세척제'] || 0) > 0) chemList.push("세척:" + item['NDT_세척제']);
    if (Number(item['NDT_현상제'] || 0) > 0) chemList.push("현상:" + item['NDT_현상제']);
    if (Number(item['NDT_형광침투제'] || 0) > 0) chemList.push("형침:" + item['NDT_형광침투제']);

    var rtkTotal = Number(item['RTK_총계'] || 0);
    if (rtkTotal > 0) chemList.push("RTK:" + rtkTotal);

    records.push({
      id: startRow + r,
      entry_time: formatCellTime(item['EntryTime']) || String(item['EntryTime'] || ''),
      date: rowDateStr,
      company: String(item['업체명'] || '경인사무소'),
    company_code: String(item['회사코드'] || ''),
      method: String(item['검사방법'] || ''),
      item_name: String(item['검사품명'] || 'Piping'),
      material: String(item['자재명'] || item['MaterialID'] || item['품목명'] || ''),
      loc_type: String(item['작업구분'] || item['구분'] || '경인사무소'),
      inspection_target: String(item['검사대상'] || ((item['관경(Inch)'] || item['규격']) ? '배관' : '배관 외')),
    inspection_spec: String(item['검사규격'] || ''),
    pipe: item['검사대상'] === '배관 외' ? '' : String(item['관경(Inch)'] || item['규격'] || ''),
    ori_qty: Number(item['Usage'] || item['수량'] || 0),
    rep_qty: Number(item['재검사매수'] || 0),
    rep_joints: Number(item['재검사개소'] || 0),
      qty: Number(item['수량']) || Number(item['Usage']) || Number(item['검사량']) || 0,
      unit: String(item['Unit'] || item['단위'] || '매'),
      unit_price: Number(item['단가'] || 0),
      travel_cost: Number(item['출장비'] || 0),
      total_fee: Number(item['검사비'] || 0),
      applied_code: String(item['적용코드'] || ''),
      report_no: String(item['성적서번호'] || ''),
      extra: chemList.join(" "),
      workers: workersArr.join(", "),
      worker_list: workerObjects
    });
  }

  records.reverse();
  return { records: records.slice(0, 15) };
}

/**
 * 전체 저장 데이터 조회 (PC 프로그램 동기화용 API)
 */
function getAllRecords() {
  var sheet = getDailyUsageSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    return { status: 'success', count: 0, headers: HEADERS, rows: [] };
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  var rows = [];

  for (var r = 1; r < data.length; r++) {
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = data[r][c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      }
      rowObj[headers[c]] = val;
    }
    rows.push(rowObj);
  }

  return {
    status: 'success',
    count: rows.length,
    headers: headers,
    rows: rows
  };
}
