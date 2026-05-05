# PAUT (Phased Array Ultrasonic Testing) 검사 절차서

## 1. 개요 및 적용 범위 (Introduction and Scope)
본 절차서는 **ASME B31.3 (2024 Edition, Process Piping)** 규정에 따라 제작되는 배관 용접부의 비파괴 검사를 위한 **Phased Array Ultrasonic Testing (PAUT)** 수행 절차를 규정한다. 본 검사는 용접부 내부 결함의 검출, 위치 확인 및 크기 측정을 목적으로 한다.

## 2. 참조 규격 (Reference Standards)
*   **ASME B31.3 (2024 Edition)**: Process Piping
*   **ASME BPVC Section V, Article 4 (2023 Edition)**: Ultrasonic Examination Methods for Welds
*   **ASME BPVC Section V, Article 23, SE-2491**: Standard Guide for for Phased Array Ultrasonic Testing (PAUT)
*   **ASME BPVC Section V, Article 23, SE-2700**: Standard Practice for Contact Ultrasonic Testing of Welds Using Phased Arrays
*   **ASME BPVC Section V, Article 1**: General Requirements
*   **ASNT SNT-TC-1A (2020 Edition)**: Personnel Qualification and Certification in NDT

## 3. 검사원 자격 (Personnel Qualification)
*   모든 검사원은 **ASNT SNT-TC-1A (2020 Edition)** 또는 ISO 9712에 따라 UT Level II 이상의 자격을 보유해야 한다.
*   **ASME 2023 Edition** 요구사항(Table T-631A)에 따라 PAUT 검사원은 최소 **320시간** 이상의 실무 경험을 보유해야 한다.
*   PAUT 장비 조작 및 데이터 분석 교육을 이수한 자여야 한다.

## 4. 장비 및 재료 (Equipment and Materials)
### 4.1 PAUT 장비 (Main Equipment)
*   **제조사/모델**: Olympus (Evident) / **OmniScan MX2** (32:128PR)
*   **분석 소프트웨어**: OmniPC 6.0, Tomoview

### 4.2 탐촉자 및 웨지 (Probe and Wedge)
*   **탐촉자 (Probe)**: 
    *   5L64-A2 (5MHz, 64 Elements)
    *   7.5CCEV-A15 (7.5MHz)
*   **웨지 (Wedge)**: 
    *   SA2-N55S-IHC
    *   SA15-N60S-IHC
*   **스캐너 (Scanner/Encoder)**: Jireh Microbe (Integrated Encoder), SPPS 250E, SPW 400

### 4.3 접촉매질 및 대비 시험편 (Couplant and Calibration Block)
*   **접촉매질**: 소너겔 (Sonagel), 물 또는 전용 젤
*   **대비 시험편 (Reference Block)**: Test Level A Block (Φ2.5), IIW Block, Step Block

## 5. 표면 준비 (Surface Preparation)
*   검사 대상 표면은 스패터, 먼지, 녹, 기름등이 제거되어야 하며, 탐촉자의 원활한 이동을 위해 매끄러워야 한다.

## 6. 교정 및 설정 (Calibration and Setup)
### 6.1 장비 설정 및 성능 점검 (System Setup & Performance Evaluation)
*   **SE-2491** 지침에 따라 PAUT 시스템의 성능 특성(빔 조정 능력, 해상도 등)을 주기적으로 점검한다.
*   OmniScan MX2의 **Calibration Wizard**를 사용하여 다음 순서로 교정을 실시한다.

#### 6.1.1 음속 교정 (Velocity Calibration)
1.  **Wizard > Calibration > Type: Ultrasound > Mode: Velocity** 선택.
2.  교정 블록의 두께(Thickness 1 & 2)를 입력한다.
3.  게이트(Gate A) 내의 신호 강도를 최대화하여 재질 내 음속을 확정한다.

#### 6.1.2 웨지 지연 교정 (Wedge Delay Calibration)
1.  **Mode: Wedge Delay** 선택.
2.  Echo Type(Radius 등)을 선택하고 블록의 형상 치수를 입력한다.
3.  모든 각도(S-Scan 범위)에서 반사 신호가 동일한 거리에 위치하도록 웨지 지연 값을 자동 계산한다.

#### 6.1.3 감도 교정 (Sensitivity Calibration - ACG)
1.  **Mode: Sensitivity** 선택.
2.  모든 A-Scan 신호 강도가 **80% FSH**(Full Screen Height)에 도달하도록 감도를 균일하게 조정(Angle Corrected Gain)한다.

#### 6.1.4 TCG(Time Corrected Gain) 설정
1.  깊이에 따른 감도 저하를 보상하기 위해 여러 깊이의 반사체(SDH 등)를 사용하여 TCG 곡선을 구성한다.
2.  모든 깊이에서 동일한 크기의 반사체가 동일한 신호 강도(80% FSH)를 유지하도록 설정한다.

#### 6.1.5 장비 성능 평가 (Performance Evaluation)
**SE-2491 (ASTM E2491)** 지침에 따라 다음 항목에 대한 성능 점검을 실시하며, 기록을 유지한다.

*   **엘리먼트 활성 점검 (Element Activity Check)**: 
    *   모든 엘리먼트가 정상적으로 파동을 송수신하는지 확인한다.
    *   연속된 2개 이상의 엘리먼트 또는 전체 엘리먼트의 25% 이상이 불량일 경우 해당 탐촉자를 교체해야 한다.
*   **빔 조정 능력 (Beam Steering Capability)**: 
    *   Phased Array Assessment Block(Type B 등)을 사용하여 설정된 최소/최대 각도에서 반사체(SDH)의 신호가 적절한 감도로 검출되는지 확인한다. 
*   **각도 및 선형 해상도 (Angular & Linear Resolution)**: 
    *   인접한 두 개의 반사체 신호가 명확히 분리되어 나타나는지 확인하여 시스템의 공간 해상도를 검증한다.
*   **빔 프로파일 검증 (Beam Profiling)**: 
    *   초점 깊이(Focus Depth)에서 실제 빔의 집중도가 설계한 대로 형성되는지 확인한다.
*   **진폭 및 시간 선형성 (Amplitude & Time Linearity)**: 
    *   장비의 증폭 선형성과 시간 축의 정확성을 주기적으로 점검한다.

#### 6.1.6 진폭 제어 선형성 (Amplitude Control Linearity)
**ASME Section V, Article 4, Mandatory Appendix II**에 따라 디지털 장비의 경우 매 1년마다(또는 수리 후) 진폭 제어 선형성을 검증해야 한다.

*   **검본 절차 (Procedure)**: 
    1. 적절한 대비 시험편에서 나오는 반사 신호를 **80% FSH**에 위치시킨다.
    2. 진폭 제어기(Gain)를 **-6dB**만큼 낮춘다.
    3. 낮아진 신호의 진폭이 **40% FSH**가 되는지 확인한다. 
    4. 다시 **+6dB**를 높여 신호가 **80% FSH**로 복귀하는지 확인한다.
*   **합격 기준 (Acceptance Criteria)**: 
    *   표시된 진폭 신호와 공칭 진폭 비율의 차이가 **±20%** 이내여야 한다. (예: 6dB 감쇄 시 신호가 32% ~ 48% FSH 범위 내에 들어와야 함)

### 6.3 수평/수직 선형성
*   ASME Section V 요구사항에 따라 장비의 선형성을 정기적으로 확인한다.

### 6.4 교정 확인 (Calibration Check)
*   교정은 매 4시간 마다, 또는 검사 작업조가 바뀔 때, 또는 장비의 이상이 의심될 때 재확인해야 한다.
*   교정 값의 오차가 10% 또는 2dB를 초과할 경우, 직전 교정 확인 시점 이후의 모든 검사 데이터는 무효로 하고 재검사한다.

## 7. 검사 수행 (Scanning and Data Acquisition)
### 7.1 스캔 계획 (Scan Plan)
*   **SE-2700**에 따라 맞대기 용접(Butt Weld)은 양면(Both sides) 검사를 원칙으로 하며, 전체 체적(Full volume) 커버리지를 보장해야 한다.
*   **Scan Path (스캔 경로)**: 용접물 중심선으로부터의 거리(Index Offset)를 설정하며, 일반적으로 용접부 두께와 베벨 각도를 고려하여 계산된 위치(-10mm 등)를 준수한다.
*   **Focus Depth (초첨 깊이)**: 용접부의 주요 결함 발생 예상 지점 또는 중심부에 초점을 맞춘다. (예: 10mm 또는 두께의 1/2 지점). 필요 시 다중 초점(Multi-focusing)을 적용한다.

### 7.2 데이터 수집
*   E-Scan 또는 S-Scan 방식을 사용하여 용접부 전체를 커버하도록 스캐닝한다. (S-Scan: 45° ~ 75°)
*   스캔 속도는 데이터 누락이 발생하지 않도록 적절히 유지하며, 스캔 해상도(Scan Resolution)는 보통 1.0mm 이하로 설정한다.
*   인코더(Encoder)를 사용하여 결함의 정확한 위치를 기록한다.

## 8. 결과 분석 및 판정 기준 (Analysis and Acceptance Criteria)
ASME B31.3 Table 341.3.2 (Acceptance Criteria for Welds - Normal Fluid Service, 2024 Edition) 및 관련 규정을 적용한다.

### 8.1 ASME B31.3 Table 341.3.2 합격 기준 (Normal Fluid Service)

| 결함 유형 (Imperfection Type) | 합격 기준 (Acceptance Criteria - Normal Fluid Service) |
| :--- | :--- |
| **균열 (Crack)** | 허용되지 않음 (Not Acceptable) |
| **미용융 (Lack of Fusion)** | 허용되지 않음 (Not Acceptable) |
| **미투과 (Incomplete Penetration)** | 개별 깊이 ≤ 1mm (1/32") 및 0.2Tw 미만<br>150mm 용접당 누적 길이 ≤ 38mm (1.5") |
| **내부 슬래그 / 텅스텐 / 선상 지시<br>(Slag, Tungsten, Elongated)** | 개별 길이 ≤ Tw / 3<br>개별 폭 ≤ 2.4mm (3/32") 또는 Tw / 4 중 작은 값<br>12Tw 용접당 누적 길이 ≤ Tw |
| **기공 (Internal Porosity)** | ASME BPVC Section VIII, Division 1, Appendix 4 기준 준용 |
| **언더컷 (Undercut)** | 깊이 ≤ 1mm (1/32") 또는 Tw / 4 중 작은 값 |
| **표면 기공 / 노출된 슬래그<br>(Surface Porosity / Slag)** | 허용되지 않음 (Not Acceptable) |
| **용접 보강탈 (Weld Reinforcement)** | Tw ≤ 6mm: ≤ 1.5mm<br>6 < Tw ≤ 13mm: ≤ 3.0mm<br>13 < Tw ≤ 25mm: ≤ 4.0mm<br>25mm < Tw: ≤ 5.0mm |

*   **Tw**: 접합부 중 얇은 쪽 부재의 공칭 두께.
### 8.2 결함 판독 및 분류 방법 (Flaw Interpretation & Characterization)
PAUT 데이터를 활용하여 다음과 같은 방식으로 결함을 판독하고 분류한다.

#### 8.2.1 데이터 뷰(Data View) 활용
*   **A-Scan**: 신호의 진폭(Amplitude)과 발생 위치(Time of Flight)를 통해 결함의 기준 높이 대비 강도를 확인한다.
*   **S-Scan (Sectorial Scan)**: 용접부 단면 이미지를 통해 결함의 깊이와 단면상 형상을 확인하며, 특히 선상(Planar) 결함의 기울기를 판별하는 데 유용하다.
*   **C-Scan (Planar View)**: 용접부 상부 평면 이미지를 통해 결함의 평면상 위치와 길이(Length)를 정밀하게 측정한다.

#### 8.2.2 주요 결함별 판독 특징
*   **균열 (Crack)**: 날카롭고 강한 신호를 보이며, 지시선의 끝단(Tip)에서 회절 신호(Diffraction)가 관찰되기도 한다. 수직 또는 경사 방향으로 뚜렷한 선형 지시를 형성한다.
*   **미용융 (Lack of Fusion)**: 용접 베벨면을 따라 형성되는 매끄럽고 강한 반사파가 특징이다. 특정 입사각에서 최대 진폭을 보이며 선형적으로 분포한다.
*   **슬래그 혼입 (Slag Inclusion)**: 불규칙하고 다중 피크(Multi-peak)를 가진 신호가 관찰된다. S-Scan상에서 뭉툭한 형상을 띄며 비금속 개재물의 특성을 보인다.
*   **기공 (Porosity)**: 개별적이고 작은 점 형태의 신호가 산발적으로 나타나거나 군집(Cluster)을 형성한다.

#### 8.2.3 결함의 크기 측정 (Sizing)
*   **길이 측정**: 인코더 데이터를 바탕으로 C-Scan 상에서 진폭이 일정 기준(예: 6dB 하락 지점)으로 떨어지는 지점 사이의 거리를 측정한다.
*   **높이 측정 (Height Sizing)**: S-Scan 상에서 단면 깊이 차이를 이용하여 결함의 수직 높이(h)를 측정한다.
    *   **6dB Drop Method**: 지시 최대 에코가 80% FSH인 경우, 상하로 이동하여 40%(-6dB)가 되는 지점 사이의 높이를 h라 한다.
    *   **만약 최대 에코가 40% FSH인 경우**, 6dB 낮춘 20% 지점 사이의 지시 높이를 h라 한다.

### 8.3 절차서의 인정 (Procedure Qualification)
**ASME Section V, Article 4, Mandatory Appendix VII**의 요구사항에 따라 다음과 같이 절차를 인정(Qualification)한다.

*   **필수 변수 (Essential Variables)**: 장비(MX2), 탐촉자 주파수/엘리먼트 수, 웨지 각도, 스캔 방향(S-Scan 범위), 보정용 시험편 등이 해당하며, 이 변수 중 하나라도 변경될 경우 절차서 재인정(Requalification)을 실시해야 한다.
*   **비 필수 변수 (Non-essential Variables)**: 표면 준비 방법, 접촉매질의 브랜드(종류 변경 제외), 스캐너의 제조사/모델, 케이블 길이 등이 해당한다. 비 필수 변수의 변경은 절차서의 재인정을 요구하지 않으나, 변경 사항은 절차서에 기록/관리되어야 한다.
*   **검증 항목**: 모든 예상 결함(균열, 미용융 등)을 안정적으로 검출하고 크기를 측정할 수 있음을 입증해야 한다.
*   **기록 관리**: 절차 인정 시험에 사용된 모든 데이터와 파라미터는 별도의 인정 보고서(PQR)로 작성하여 관리한다.

## 9. 보고서 작성 (Reporting)
보고서에는 다음 사항이 포함되어야 한다.
1. 고객사 및 프로젝트 명칭
2. 검사 대상 번호 (Line No., Weld No.)
3. 사용 장비 및 탐촉자 정보
4. 교정 데이터 및 설정 값
5. 검사 결과 (합격/불합격) 및 결함 위치/크기
6. 검사원 서명 및 날짜

## 10. 후처리 (Post-Cleanup)
*   검사가 완료된 후, 용접부 및 모재 표면에 남아있는 접촉매질(Couplant)을 깨끗이 제거하여 부식을 방지한다.

## 11. 안전 사항 (Safety)
*   현장 검사 시 안전모, 안전화, 보안경 등 적절한 개인보호구(PPE)를 착용해야 한다.
*   고소 작업 시 안전벨트 착용 및 추락 방지 조치를 취해야 한다.
