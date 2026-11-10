"""Optional auxiliary equipment selection (with Serial Numbers) for gyeongin daily inspection records."""
import tkinter as tk
from tkinter import ttk
import pandas as pd

METHOD_MAIN_EQUIP_MAP = {
    "PAUT": [
        "PAUT MX2(OMNI2-103085)",
        "PAUT MX2(OMNI2-100055)",
        "PAUT MX2(OMNI2-103918)",
        "PAUT X3(QC-0093084)",
    ],
    "UT": [
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
    ],
    "RT": [
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
    ],
    "MT": [
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
    ],
    "PT": [
        "PT 약품",
        "해당없음",
    ],
    "PMI": [
        "합금분석기(X-ray) XL2 100G(200983)",
        "합금분석기(RI) XLi818(24673)",
        "경도계 MIC10(1-2129)",
        "페이라트함량측정기 FMP30(160002524)",
        "내시경장비 ER400(E30969R)",
    ],
    "내근/대기": [
        "해당없음",
    ],
    "휴무": [
        "해당없음",
    ],
    "기타": [
        "합금분석기(X-ray) XL2 100G(200983)",
        "합금분석기(RI) XLi818(24673)",
        "경도계 MIC10(1-2129)",
        "페이라트함량측정기 FMP30(160002524)",
        "내시경장비 ER400(E30969R)",
        "해당없음",
    ],
}


def classify_main_equipment_method(label):
    u = str(label or '').strip().upper()
    if not u or u == '해당없음':
        return None
    if any(k in u for k in ('합금분석기', 'XRF', 'NITON', 'XL2', 'XLI818')):
        return 'PMI'
    if any(k in u for k in ('PAUT', 'OMNI', 'MX2', 'X3')):
        return 'PAUT'
    if any(k in u for k in ('SITESCAN', 'MASTERSCAN', 'EPOCH', 'USM', '두께측정기', '37DL', 'DMS-2', 'DM-4', 'M-UT')):
        return 'UT'
    if any(k in u for k in ('SE-75', 'IR-192', '880', '300EGM', 'X-RAY', 'X-선', '방사선')):
        return 'RT'
    if any(k in u for k in ('YOKE', 'MY-2', 'MP-A', 'HWM-', 'NWM', 'SYM-', 'MP-100', 'AC POWER', 'DA-PB', '자분탐상기')):
        return 'MT'
    if 'PT' in u:
        return 'PT'
    return '기타'


def get_main_equipments_by_method(app=None, method=None):
    if method is None and app is not None:
        m_widget = getattr(app, 'cb_daily_test_method', None)
        if m_widget and hasattr(m_widget, 'get'):
            method = str(m_widget.get()).strip().upper()
    m = str(method or '').strip().upper()
    if m in METHOD_MAIN_EQUIP_MAP:
        return list(METHOD_MAIN_EQUIP_MAP[m])
    # Fallback: all main equipments across methods
    all_main = []
    for k in ('RT', 'PAUT', 'UT', 'MT', 'PT', '기타'):
        for item in METHOD_MAIN_EQUIP_MAP.get(k, []):
            if item not in all_main:
                all_main.append(item)
    return all_main


METHOD_AUX_MAP = {
    "PAUT": [
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
        "표준시험편 STB-A1 (257)"
    ],
    "UT": [
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
        "접촉매질(Couplant)"
    ],
    "RT": [
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
        "조도계 LS-150 (10004333)"
    ],
    "MT": [
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
        "리프팅파워블럭 (AC 형) 10lb(4.5kg) (NW0756)"
    ],
    "PT": [
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
        "온습도계"
    ]
}


def _clean_str(val):
    if val is None or pd.isna(val):
        return ''
    s = ' '.join(str(val).split()).strip()
    return '' if s.lower() in ('nan', 'none') else s


def _fmt_inv_item(prefix, model, sn):
    parts = [p for p in (prefix, model) if p]
    base = ' '.join(parts).strip()
    return f"{base} ({sn})" if sn else base


def collect_inventory_main_by_method(app):
    """Return method -> list of main equipments (with S/N) for Gyeongin Office."""
    dyn_map = {m: list(items) for m, items in METHOD_MAIN_EQUIP_MAP.items()}
    eq_list = getattr(app, 'equipments', None) or []
    for item in eq_list:
        clean_it = str(item or '').strip()
        if not clean_it:
            continue
        m_key = classify_main_equipment_method(clean_it)
        if m_key and m_key in dyn_map and clean_it not in dyn_map[m_key]:
            dyn_map[m_key].append(clean_it)
    return dyn_map


def collect_inventory_aux_by_method(app):
    """Dynamically collect auxiliary equipments & probes with Serial Numbers from app.materials_df."""
    dyn_map = {m: list(items) for m, items in METHOD_AUX_MAP.items()}
    df = getattr(app, 'materials_df', None)
    if df is None or getattr(df, 'empty', True):
        for m_key, m_items in dyn_map.items():
            with_sn = [x for x in m_items if '(' in x and x.endswith(')') and x != 'PAUT SCANNER (MANUAL)']
            without_sn = [x for x in m_items if x not in with_sn]
            dyn_map[m_key] = with_sn + without_sn
        return dyn_map

    active_mask = pd.to_numeric(df.get('Active', 1), errors='coerce').fillna(1) > 0.5
    for _, row in df[active_mask].iterrows():
        name = _clean_str(row.get('품목명'))
        model = _clean_str(row.get('모델명'))
        sn = _clean_str(row.get('SN'))
        cat = _clean_str(row.get('품목군코드'))
        if not name and not model:
            continue

        def _add(m_key, label):
            if label and label not in dyn_map[m_key]:
                dyn_map[m_key].append(label)

        if name.startswith('PAUT PROBE(') or name.startswith('PAUT PROBE-WEDGE'):
            _add('PAUT', _fmt_inv_item(name, model, sn))
        elif name in ('PAUT 부자재', '채널분배기(PAUT)') and sn:
            lbl = _fmt_inv_item(name, model, sn) if name == '채널분배기(PAUT)' else _fmt_inv_item(model or name, '', sn)
            _add('PAUT', lbl)
        elif 'PROBE' in name and not name.startswith('PAUT'):
            _add('UT', _fmt_inv_item(name, model, sn))
        elif ('표준시험편' in name or '대비시험편' in name) and (sn or model):
            if cat == 'UT자재':
                prefix = '표준시험편' if '표준시험편' in name else '대비시험편'
                lbl = _fmt_inv_item(prefix, model, sn)
                _add('UT', lbl)
                if any(k in model for k in ('PAUT', 'STB-A1', 'IIW')):
                    _add('PAUT', lbl)
            elif cat == 'MT자재':
                _add('MT', _fmt_inv_item(name, model, sn))
            elif cat == 'PT자재':
                _add('PT', _fmt_inv_item(name, model, sn))
        elif '콜리메타' in name or 'R/CONTROL' in name or name in ('Survey Meter', 'Alarm Dosimeter'):
            _add('RT', _fmt_inv_item(name, model, sn))
        elif cat == 'RT장비' and name in ('농도계', '판독기', '건조기', '크롤러', '영상화장비', '조도계', '라이트마킹') and (sn or name == '라이트마킹'):
            lbl = _fmt_inv_item(name, model if model != name else '', sn)
            _add('RT', lbl)
            if name == '조도계':
                _add('MT', lbl)
                _add('PT', lbl)
        elif 'Black Light' in name or '자장계' in name or '리프팅파워블럭' in name or '팔각시험편' in name:
            _add('MT', _fmt_inv_item(name, model if model != name else '', sn))
        elif name == 'UV Meter':
            lbl = _fmt_inv_item(name, model, sn)
            _add('MT', lbl)
            _add('PT', lbl)
        elif name == 'LUX Meter':
            lbl = _fmt_inv_item('조도계 LUX Meter', model, sn)
            _add('MT', lbl)
            _add('PT', lbl)
        elif name == '표면온도계':
            lbl = _fmt_inv_item(name, model, sn)
            _add('MT', lbl)
            _add('PT', lbl)
        elif '피 게이지' in name:
            _add('PT', _fmt_inv_item(name, model, sn))

    for m_key, m_items in dyn_map.items():
        with_sn = [x for x in m_items if '(' in x and x.endswith(')') and x != 'PAUT SCANNER (MANUAL)']
        without_sn = [x for x in m_items if x not in with_sn]
        dyn_map[m_key] = with_sn + without_sn

    return dyn_map


def names(value):
    if value is None or str(value).strip().lower() in ('', 'nan', 'none'):
        return []
    return list(dict.fromkeys(n.strip() for n in str(value).split(';') if n.strip()))


def value(app):
    field = getattr(app, 'daily_aux_equipment_var', None)
    return '; '.join(names(field.get())) if field is not None else ''


def restore(app, text='', method=None):
    field = getattr(app, 'daily_aux_equipment_var', None)
    if field is not None:
        field.set('; '.join(names(text)))
    if method is not None:
        app._daily_equipment_method = str(method).strip().upper()


def report_names(records):
    result = []
    for column in ('장비명', '보조장비'):
        if column in records.columns:
            for text in records[column].dropna():
                result.extend(names(text))
    return list(dict.fromkeys(result))


def _get_target_text(target):
    if target is None:
        return ''
    if hasattr(target, 'get'):
        try:
            return str(target.get() or '')
        except Exception:
            return ''
    return str(target or '')


def _set_target_text(target, text):
    if target is None:
        return
    if isinstance(target, ttk.Combobox) or (hasattr(target, 'delete') and hasattr(target, 'insert')):
        try:
            target.delete(0, tk.END)
            target.insert(0, text)
        except Exception:
            pass
    if hasattr(target, 'set'):
        try:
            target.set(text)
        except Exception:
            pass


def _bind_count_badge(widget_or_var, badge_label):
    """Keep the 'N개 선택' badge synced with the entry/combobox/StringVar value."""
    def update_badge(*_args):
        try:
            if not badge_label.winfo_exists():
                return
            cnt = len([x for x in names(_get_target_text(widget_or_var)) if x != '해당없음'])
            if cnt > 0:
                badge_label.config(text=f'{cnt}개 선택', bg='#dbeafe', fg='#1d4ed8')
            else:
                badge_label.config(text='0개 선택', bg='#e2e8f0', fg='#475569')
        except Exception:
            pass

    if isinstance(widget_or_var, tk.Variable):
        try:
            widget_or_var.trace_add('write', update_badge)
        except Exception:
            pass
    elif hasattr(widget_or_var, 'bind'):
        for ev in ('<KeyRelease>', '<<ComboboxSelected>>', '<FocusOut>'):
            try:
                widget_or_var.bind(ev, update_badge, add='+')
            except Exception:
                pass
    update_badge()
    return update_badge


def add_selector(parent, app, variable, width=28):
    frame = ttk.Frame(parent)
    cb = ttk.Combobox(frame, textvariable=variable, width=width)
    cb.pack(side='left', fill='x', expand=True)
    if app is not None and variable is getattr(app, 'daily_aux_equipment_var', None):
        app.cb_daily_aux_equip = cb

    badge_lbl = tk.Label(
        frame, text='0개 선택', font=('Malgun Gothic', 8, 'bold'),
        bg='#e2e8f0', fg='#475569', padx=5, pady=1
    )
    badge_lbl.pack(side='left', padx=(3, 1))
    refresh_badge = _bind_count_badge(variable, badge_lbl)

    btn_sel = tk.Button(
        frame, text='☑ 선택', font=('Malgun Gothic', 8, 'bold'),
        bg='#eff6ff', fg='#1d4ed8', activebackground='#dbeafe', activeforeground='#1e40af',
        relief='solid', bd=1, padx=6, pady=1, cursor='hand2',
        command=lambda: select(frame, app, variable, on_changed=refresh_badge)
    )
    btn_sel.pack(side='left', padx=2)
    badge_lbl.bind('<Button-1>', lambda e: select(frame, app, variable, on_changed=refresh_badge))
    badge_lbl.config(cursor='hand2')

    if hasattr(app, 'open_list_management_dialog'):
        tk.Button(frame, text="⚙️", font=('Malgun Gothic', 8), bd=0,
                  bg=getattr(app, 'theme_bg', '#f0f0f0'), fg='blue', cursor='hand2',
                  command=lambda: app.open_list_management_dialog('aux_equipments')).pack(side='left', padx=(0, 2))
    return frame


def bind_main_selector(container, app, combobox_widget, method_getter=None):
    """Attach the mobile-style multi-select badge + [☑ 선택] button to a main equipment combobox container."""
    badge_lbl = tk.Label(
        container, text='0개 선택', font=('Malgun Gothic', 8, 'bold'),
        bg='#e2e8f0', fg='#475569', padx=5, pady=1, cursor='hand2'
    )
    badge_lbl.pack(side='left', padx=(3, 1))
    refresh_badge = _bind_count_badge(combobox_widget, badge_lbl)

    def _open():
        select_main(
            container, app, combobox_widget,
            method_getter=method_getter,
            on_changed=refresh_badge
        )

    btn_sel = tk.Button(
        container, text='☑ 선택', font=('Malgun Gothic', 8, 'bold'),
        bg='#eff6ff', fg='#1d4ed8', activebackground='#dbeafe', activeforeground='#1e40af',
        relief='solid', bd=1, padx=6, pady=1, cursor='hand2',
        command=_open
    )
    btn_sel.pack(side='left', padx=2)
    badge_lbl.bind('<Button-1>', lambda e: _open())
    combobox_widget._refresh_multi_badge = refresh_badge
    return btn_sel, badge_lbl


def select(parent, app, variable, method_getter=None, on_changed=None):
    return _open_multi_select_dialog(
        parent=parent,
        app=app,
        target=variable,
        mode='aux',
        method_getter=method_getter,
        on_changed=on_changed,
    )


def select_main(parent, app, target, method_getter=None, on_changed=None):
    return _open_multi_select_dialog(
        parent=parent,
        app=app,
        target=target,
        mode='main',
        method_getter=method_getter,
        on_changed=on_changed,
    )


def _open_multi_select_dialog(parent, app, target, mode='aux', method_getter=None, on_changed=None):
    is_main = (mode == 'main')
    dialog = tk.Toplevel(parent)
    if is_main:
        dialog.title('주 장비 다중 선택 (검사방법별 · 시리얼번호 S/N 복수 체크)')
    else:
        dialog.title('보조장비 · 프로브 · 시험편 다중 선택 (시리얼번호 S/N 복수 체크)')
    dialog.geometry('780x610')
    dialog.configure(bg='#f8fafc')
    dialog.transient(parent.winfo_toplevel())

    selected_set = set(names(_get_target_text(target)))

    if is_main:
        dyn_method_map = collect_inventory_main_by_method(app)
        method_buttons = ['현재방법', '전체', 'PAUT', 'UT', 'RT', 'MT', 'PT', 'PMI', '기타']
        config_attr = 'equipments'
        history_col = '장비명'
    else:
        dyn_method_map = collect_inventory_aux_by_method(app)
        method_buttons = ['현재방법', '전체', 'PAUT', 'UT', 'RT', 'MT', 'PT']
        config_attr = 'aux_equipments'
        history_col = '보조장비'

    def _detect_current_method():
        raw = ''
        if callable(method_getter):
            try:
                raw = str(method_getter() or '').strip()
            except Exception:
                raw = ''
        if not raw:
            method_widget = (
                getattr(app, 'cb_daily_test_method', None)
                or getattr(app, 'daily_method_var', None)
                or getattr(app, 'method_var', None)
            )
            if method_widget and hasattr(method_widget, 'get'):
                try:
                    raw = str(method_widget.get() or '').strip()
                except Exception:
                    raw = ''
        upper = raw.upper()
        if upper in dyn_method_map:
            return upper
        if raw in dyn_method_map:
            return raw
        if is_main and '기타' in dyn_method_map and raw:
            return '기타'
        return '전체'

    detected_method = _detect_current_method()

    # Collect all candidate items
    all_candidates = []
    for m_list in dyn_method_map.values():
        all_candidates.extend(m_list)

    cfg_list = getattr(app, config_attr, None)
    if cfg_list:
        for it in cfg_list:
            all_candidates.extend(names(it))

    history = getattr(app, 'daily_usage_df', None)
    if history is not None and history_col in history.columns:
        for it in history[history_col].dropna():
            all_candidates.extend(names(it))

    all_candidates.extend(selected_set)
    all_candidates = list(dict.fromkeys(
        c.strip() for c in all_candidates
        if c and c.strip() and c.strip() != '없음'
    ))

    def persist_items(items):
        changed = False
        lst = getattr(app, config_attr, None)
        if lst is None:
            lst = []
            setattr(app, config_attr, lst)
        for it in items:
            clean_it = str(it).strip()
            if clean_it and clean_it not in lst:
                lst.append(clean_it)
                changed = True
        if changed and hasattr(app, 'save_tab_config'):
            try:
                app.save_tab_config(force=True)
            except Exception as ex:
                print(f"Error persisting {config_attr}: {ex}")

    # 1. Header bar (Mobile modal style)
    header_frame = tk.Frame(dialog, bg='#ffffff', padx=14, pady=10, highlightthickness=1, highlightbackground='#e2e8f0')
    header_frame.pack(fill='x')

    header_title = (
        '주 장비 선택 (복수 체크 가능)'
        if is_main else
        '보조장비 · 프로브 · 시험편 선택 (복수 체크 가능)'
    )
    tk.Label(
        header_frame, text=header_title,
        font=('Malgun Gothic', 11, 'bold'), bg='#ffffff', fg='#0f172a'
    ).pack(side='left')

    lbl_count = tk.Label(
        header_frame, text=f'선택: {len(selected_set)}개',
        font=('Malgun Gothic', 10, 'bold'), bg='#dbeafe', fg='#1d4ed8', padx=10, pady=3
    )
    lbl_count.pack(side='right')

    # 2. Method Filter Pill Tabs + Search Input (Mobile style)
    filter_container = tk.Frame(dialog, bg='#ffffff', padx=14, pady=8, highlightthickness=1, highlightbackground='#e2e8f0')
    filter_container.pack(fill='x')

    tab_frame = tk.Frame(filter_container, bg='#ffffff')
    tab_frame.pack(fill='x', pady=(0, 6))

    active_method_filter = tk.StringVar(value='현재방법')
    tab_btn_widgets = {}

    def _update_tab_styles():
        cur = active_method_filter.get()
        for m_name, btn in tab_btn_widgets.items():
            if m_name == cur:
                btn.config(bg='#2563eb', fg='#ffffff', relief='flat')
            else:
                btn.config(bg='#e2e8f0', fg='#475569', relief='flat')

    def _select_tab(m_name):
        active_method_filter.set(m_name)
        _update_tab_styles()
        apply_filter()

    for m in method_buttons:
        label_text = f'현재방법({detected_method})' if m == '현재방법' and detected_method != '전체' else m
        b = tk.Button(
            tab_frame, text=label_text, font=('Malgun Gothic', 9, 'bold'),
            padx=10, pady=3, bd=0, cursor='hand2',
            command=lambda m_val=m: _select_tab(m_val)
        )
        b.pack(side='left', padx=(0, 5))
        tab_btn_widgets[m] = b

    search_frame = tk.Frame(filter_container, bg='#ffffff')
    search_frame.pack(fill='x')
    search_var = tk.StringVar()
    tk.Label(
        search_frame, text='🔍 검색:', font=('Malgun Gothic', 9, 'bold'),
        bg='#ffffff', fg='#334155'
    ).pack(side='left', padx=(0, 6))
    search_entry = ttk.Entry(search_frame, textvariable=search_var, font=('Malgun Gothic', 10))
    search_entry.pack(side='left', fill='x', expand=True)
    tk.Button(
        search_frame, text='✕ 지우기', font=('Malgun Gothic', 8),
        bg='#f1f5f9', fg='#475569', relief='flat', padx=6, cursor='hand2',
        command=lambda: search_var.set('')
    ).pack(side='left', padx=(4, 0))

    # 3. Checkbox List Body (Single-click toggles ☑ / ☐ without losing other selections)
    body = tk.Frame(dialog, bg='#f8fafc', padx=14, pady=8)
    body.pack(fill='both', expand=True)
    scrollbar = ttk.Scrollbar(body)
    scrollbar.pack(side='right', fill='y')
    listing = tk.Listbox(
        body,
        selectmode='browse',
        exportselection=False,
        activestyle='none',
        font=('Malgun Gothic', 10),
        bg='#ffffff',
        fg='#1e293b',
        selectbackground='#dbeafe',
        selectforeground='#1e3a8a',
        highlightthickness=1,
        highlightbackground='#cbd5e1',
        relief='flat',
        yscrollcommand=scrollbar.set,
    )
    listing.pack(fill='both', expand=True)
    scrollbar.configure(command=listing.yview)

    displayed_items = []

    def _render_item_row(idx, item):
        checked = item in selected_set
        prefix = '☑  ' if checked else '☐  '
        listing.delete(idx)
        listing.insert(idx, f'{prefix}{item}')
        if checked:
            listing.itemconfig(idx, bg='#eff6ff', fg='#1d4ed8', selectbackground='#dbeafe', selectforeground='#1e3a8a')
        else:
            listing.itemconfig(idx, bg='#ffffff', fg='#1e293b', selectbackground='#f1f5f9', selectforeground='#0f172a')

    def _update_count_label():
        lbl_count.config(text=f'선택: {len(selected_set)}개')

    def apply_filter(*_args):
        nonlocal displayed_items
        keyword = search_var.get().strip().lower()
        method_filter = active_method_filter.get()
        effective_method = detected_method if method_filter == '현재방법' else method_filter

        if effective_method == '전체' or effective_method not in dyn_method_map:
            base_source = list(all_candidates)
        else:
            method_items = list(dyn_method_map.get(effective_method, []))
            # Also keep currently selected items visible at the top if they belong to custom/other
            base_source = method_items

        filtered = []
        for it in base_source:
            if keyword and keyword not in it.lower():
                continue
            filtered.append(it)

        displayed_items = filtered
        yview_pos = listing.yview()
        listing.delete(0, 'end')
        for idx, item in enumerate(displayed_items):
            checked = item in selected_set
            prefix = '☑  ' if checked else '☐  '
            listing.insert('end', f'{prefix}{item}')
            if checked:
                listing.itemconfig(idx, bg='#eff6ff', fg='#1d4ed8', selectbackground='#dbeafe', selectforeground='#1e3a8a')
            else:
                listing.itemconfig(idx, bg='#ffffff', fg='#1e293b', selectbackground='#f1f5f9', selectforeground='#0f172a')
        try:
            listing.yview_moveto(yview_pos[0])
        except Exception:
            pass
        _update_count_label()

    def on_row_click(event):
        if not displayed_items:
            return 'break'
        idx = listing.nearest(event.y)
        if idx < 0 or idx >= len(displayed_items):
            return 'break'
        bbox = listing.bbox(idx)
        if bbox and (event.y < bbox[1] or event.y > bbox[1] + bbox[3] + 4):
            return 'break'
        item = displayed_items[idx]
        if is_main and item == '해당없음':
            if '해당없음' in selected_set:
                selected_set.discard('해당없음')
            else:
                selected_set.clear()
                selected_set.add('해당없음')
            apply_filter()
            return 'break'
        else:
            selected_set.discard('해당없음')

        if item in selected_set:
            selected_set.discard(item)
        else:
            selected_set.add(item)
        _render_item_row(idx, item)
        _update_count_label()
        return 'break'

    def on_space_toggle(event):
        cur = listing.curselection()
        if not cur:
            return 'break'
        idx = cur[0]
        if 0 <= idx < len(displayed_items):
            item = displayed_items[idx]
            if item in selected_set:
                selected_set.discard(item)
            else:
                selected_set.add(item)
            _render_item_row(idx, item)
            _update_count_label()
        return 'break'

    listing.bind('<Button-1>', on_row_click)
    listing.bind('<B1-Motion>', lambda e: 'break')
    listing.bind('<space>', on_space_toggle)
    search_var.trace_add('write', apply_filter)

    _update_tab_styles()
    apply_filter()

    # 4. Direct Input & Action Footer (Mobile style)
    custom_frame = tk.Frame(dialog, bg='#ffffff', padx=14, pady=8, highlightthickness=1, highlightbackground='#e2e8f0')
    custom_frame.pack(fill='x')
    tk.Label(
        custom_frame, text='✏️ 직접 추가:', font=('Malgun Gothic', 9, 'bold'),
        bg='#ffffff', fg='#334155'
    ).pack(side='left', padx=(0, 6))
    entry = ttk.Entry(custom_frame, font=('Malgun Gothic', 9))
    entry.pack(side='left', fill='x', expand=True)

    def add_custom():
        new_items = names(entry.get())
        if not new_items:
            return
        for item in new_items:
            if item not in all_candidates:
                all_candidates.append(item)
            selected_set.add(item)
        persist_items(new_items)
        entry.delete(0, 'end')
        apply_filter()

    def delete_selected():
        if not selected_set:
            return
        to_remove = list(selected_set)
        changed = False
        lst = getattr(app, config_attr, None)
        if lst:
            for r in to_remove:
                if r in lst:
                    lst.remove(r)
                    changed = True
        for r in to_remove:
            if r in all_candidates:
                all_candidates.remove(r)
            selected_set.discard(r)
        if changed and hasattr(app, 'save_tab_config'):
            try:
                app.save_tab_config(force=True)
            except Exception as ex:
                print(f"Error saving tab config on delete: {ex}")
        apply_filter()

    def clear_all():
        selected_set.clear()
        apply_filter()

    tk.Button(
        custom_frame, text='직접 추가', font=('Malgun Gothic', 9, 'bold'),
        bg='#eff6ff', fg='#1d4ed8', relief='solid', bd=1, padx=8, pady=2, cursor='hand2',
        command=add_custom
    ).pack(side='left', padx=3)
    tk.Button(
        custom_frame, text='목록에서 삭제', font=('Malgun Gothic', 8),
        bg='#fef2f2', fg='#dc2626', relief='solid', bd=1, padx=6, pady=2, cursor='hand2',
        command=delete_selected
    ).pack(side='left', padx=2)
    entry.bind('<Return>', lambda event: add_custom())

    footer = tk.Frame(dialog, bg='#ffffff', padx=14, pady=10)
    footer.pack(fill='x')

    def apply():
        add_custom()
        chosen = [it for it in all_candidates if it in selected_set]
        for it in selected_set:
            if it not in chosen:
                chosen.append(it)
        joined = '; '.join(chosen)
        _set_target_text(target, joined)
        persist_items(chosen)
        if callable(on_changed):
            try:
                on_changed()
            except Exception:
                pass
        if is_main and hasattr(target, '_refresh_multi_badge'):
            try:
                target._refresh_multi_badge()
            except Exception:
                pass
        dialog.destroy()

    tk.Button(
        footer, text='초기화', font=('Malgun Gothic', 10, 'bold'),
        bg='#f1f5f9', fg='#334155', relief='flat', padx=16, pady=6, cursor='hand2',
        command=clear_all
    ).pack(side='left')

    tk.Button(
        footer, text='선택 완료', font=('Malgun Gothic', 10, 'bold'),
        bg='#2563eb', fg='#ffffff', activebackground='#1d4ed8', activeforeground='#ffffff',
        relief='flat', padx=24, pady=6, cursor='hand2',
        command=apply
    ).pack(side='right')

    tk.Button(
        footer, text='취소', font=('Malgun Gothic', 10),
        bg='#e2e8f0', fg='#475569', relief='flat', padx=14, pady=6, cursor='hand2',
        command=dialog.destroy
    ).pack(side='right', padx=6)

    dialog.grab_set()
    search_entry.focus_set()
