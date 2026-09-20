"""비파괴검사 기술사 학습용 핵심용어 사전."""

GLOSSARY = [
    {"term": "매질", "aliases": ["medium"], "category": "UT 기초", "definition": "파동이나 에너지가 전달되는 물질이다.", "relation": "초음파 탐상에서는 시험체, 접촉매질, 물과 공기 등이 매질이다. 매질의 밀도와 탄성특성이 음속·음향임피던스·감쇠를 결정한다.", "formula": "Z = ρc", "related": ["밀도", "음속", "음향임피던스", "접촉매질"], "questions": ["파동과 음향임피던스", "접촉매질의 사용목적과 선정 인자"]},
    {"term": "밀도", "aliases": ["density", "ρ"], "category": "UT 기초", "definition": "단위체적당 질량이며 기호는 ρ, SI 단위는 kg/m³이다.", "relation": "매질의 음향임피던스 Z=ρc를 결정하며 경계면의 반사·투과에 영향을 준다.", "formula": "ρ = m/V", "related": ["매질", "음속", "음향임피던스"], "questions": ["음향임피던스와 경계면 반사"]},
    {"term": "음속", "aliases": ["sound velocity", "c"], "category": "UT 기초", "definition": "탄성파가 매질 안에서 단위시간 동안 진행하는 거리이다.", "relation": "결함 깊이와 두께는 초음파의 왕복시간과 음속으로 계산한다. 재질과 파동모드에 따라 값이 달라진다.", "formula": "d = ct/2", "related": ["매질", "파장", "종파", "횡파"], "questions": ["초음파 속도와 결함 위치", "위상속도와 군속도"]},
    {"term": "음향임피던스", "aliases": ["acoustic impedance", "Z"], "category": "UT 기초", "definition": "초음파 진행에 대한 매질의 저항을 나타내는 물리량이다.", "relation": "경계 양쪽 매질의 음향임피던스 차가 클수록 반사가 커진다. 공기와 금속의 차이가 커 접촉매질이 필요하다.", "formula": "Z = ρc,  R = (Z₂-Z₁)/(Z₂+Z₁)", "related": ["매질", "밀도", "음속", "반사", "접촉매질"], "questions": ["파동과 음향임피던스", "경계면에서의 반사와 투과"]},
    {"term": "파장", "aliases": ["wavelength", "λ"], "category": "UT 기초", "definition": "파동에서 같은 위상인 인접 두 점 사이의 거리이다.", "relation": "파장이 짧을수록 작은 결함 분해능은 좋아지지만 산란과 감쇠가 커질 수 있다.", "formula": "λ = c/f", "related": ["음속", "주파수", "분해능"], "questions": ["주파수 선정", "근거리음장 한계거리"]},
    {"term": "주파수", "aliases": ["frequency", "f"], "category": "UT 기초", "definition": "1초 동안 반복되는 진동 횟수이며 단위는 Hz이다.", "relation": "고주파는 작은 결함에 민감하지만 감쇠가 크고, 저주파는 투과력이 좋지만 분해능이 낮다.", "formula": "f = 1/T", "related": ["파장", "감쇠", "분해능", "탐촉자"], "questions": ["초음파 탐상에서 주파수 선정법"]},
    {"term": "감쇠", "aliases": ["attenuation"], "category": "UT 기초", "definition": "초음파가 진행하면서 진폭 또는 에너지가 감소하는 현상이다.", "relation": "흡수, 산란, 빔 확산과 경계면 손실 때문에 발생하며 검사거리와 감도에 영향을 준다.", "formula": "A = A₀e⁻ᵅˣ", "related": ["흡수", "산란", "주파수", "거리진폭보정"], "questions": ["초음파 감쇠 원인", "반사신호가 감쇠되는 경우"]},
    {"term": "반사", "aliases": ["reflection"], "category": "UT 기초", "definition": "파동이 서로 다른 매질의 경계에서 진행방향을 바꾸어 되돌아오는 현상이다.", "relation": "결함 에코는 결함과 모재 사이의 음향임피던스 차이로 발생한다.", "formula": "R = (Z₂-Z₁)/(Z₂+Z₁)", "related": ["음향임피던스", "투과", "굴절", "에코"], "questions": ["경계면에서 초음파의 반사와 투과"]},
    {"term": "굴절", "aliases": ["refraction"], "category": "UT 기초", "definition": "파동이 음속이 다른 매질의 경계를 비스듬히 통과할 때 진행방향이 변하는 현상이다.", "relation": "사각탐상에서 웨지와 시험체의 음속 차를 이용해 원하는 굴절각의 종파 또는 횡파를 만든다.", "formula": "sinθ₁/c₁ = sinθ₂/c₂", "related": ["스넬의 법칙", "임계각", "사각탐상"], "questions": ["제1·제2 임계각의 중요성"]},
    {"term": "스넬의 법칙", "aliases": ["Snell's law", "Snell law"], "category": "UT 기초", "definition": "서로 다른 매질 경계에서 입사각·반사각·굴절각과 각 파동의 속도 관계를 나타내는 법칙이다.", "relation": "사각탐촉자의 굴절각 계산과 모드변환, 임계각 결정에 사용한다.", "formula": "sinθ₁/c₁ = sinθ₂/c₂", "related": ["굴절", "모드변환", "임계각"], "questions": ["Snell's law와 임계각"]},
    {"term": "근거리음장", "aliases": ["near field", "Fresnel zone"], "category": "UT 기초", "definition": "진동자 각 부분에서 나온 파가 간섭해 음압이 불규칙하게 변하는 탐촉자 인접 영역이다.", "relation": "근거리에서는 동일 결함의 에코가 거리와 단조롭게 대응하지 않아 정량평가가 어렵다.", "formula": "N = D²/(4λ) = D²f/(4c)", "related": ["원거리음장", "진동자 직경", "파장"], "questions": ["근거리음장 한계거리"]},
    {"term": "접촉매질", "aliases": ["couplant"], "category": "UT 기초", "definition": "탐촉자와 시험체 사이의 공기층을 제거해 초음파 에너지를 전달하는 물질이다.", "relation": "물, 글리세린, 오일, 젤 등을 사용하며 재질·표면상태·온도·부식성·제거성을 고려해 선정한다.", "formula": "공기층 제거 → 전달효율 증가", "related": ["매질", "음향임피던스", "투과"], "questions": ["접촉매질의 사용목적과 선정 인자"]},
    {"term": "종파", "aliases": ["longitudinal wave"], "category": "UT 파동", "definition": "입자 진동방향이 파동 진행방향과 평행한 탄성파이다.", "relation": "고체·액체·기체에 전파되며 수직탐상과 두께측정에 주로 사용한다.", "formula": "고체에서 일반적으로 cL > cS", "related": ["횡파", "표면파", "모드변환"], "questions": ["초음파의 진동양식"]},
    {"term": "횡파", "aliases": ["shear wave", "transverse wave"], "category": "UT 파동", "definition": "입자 진동방향이 파동 진행방향과 수직인 탄성파이다.", "relation": "전단력을 전달할 수 있는 고체에서만 전파되며 용접부 사각탐상에 널리 사용한다.", "formula": "cS = √(G/ρ)", "related": ["종파", "굴절", "사각탐상"], "questions": ["초음파의 진동양식"]},
    {"term": "표면파", "aliases": ["surface wave", "Rayleigh wave"], "category": "UT 파동", "definition": "고체 표면을 따라 약 한 파장 깊이까지 에너지가 집중되어 진행하는 파동이다.", "relation": "표면균열 검출에 민감하지만 표면상태, 곡률과 액체 접촉에 영향을 받는다.", "formula": "cR ≈ 0.9cS", "related": ["종파", "횡파", "판파"], "questions": ["표면파의 특징"]},
    {"term": "탐촉자", "aliases": ["probe", "transducer"], "category": "UT 장비", "definition": "전기신호와 초음파를 상호 변환해 송수신하는 장치이다.", "relation": "압전소자, 댐핑재, 보호면, 케이블과 웨지 등으로 구성되며 주파수·크기·형식이 검사성능을 결정한다.", "formula": "압전효과·역압전효과", "related": ["주파수", "분해능", "감도", "PAUT"], "questions": ["탐촉자의 성능항목", "주파수 선정"]},
    {"term": "감도", "aliases": ["sensitivity"], "category": "UT 성능", "definition": "작은 반사원 또는 결함신호를 검출할 수 있는 능력이다.", "relation": "탐촉자, 주파수, 증폭도, 재료 감쇠, 표면상태와 교정방법의 영향을 받는다.", "formula": "기준반사체 에코 높이로 조정", "related": ["분해능", "교정", "DAC", "TCG"], "questions": ["초음파 감도 측정규정", "STB-A2와 RB-4 비교"]},
    {"term": "분해능", "aliases": ["resolution"], "category": "UT 성능", "definition": "서로 가까운 두 반사원을 별개의 신호로 구분하는 능력이다.", "relation": "짧은 펄스와 넓은 대역폭이 유리하며 표면분해능과 원거리분해능으로 나눌 수 있다.", "formula": "짧은 펄스폭 → 높은 분해능", "related": ["감도", "주파수", "대역폭", "불감대"], "questions": ["탐촉자 성능항목"]},
    {"term": "PAUT", "aliases": ["위상배열", "phased array"], "category": "고급 UT", "definition": "다수의 압전소자에 시간지연을 주어 초음파 빔을 전자적으로 조향·집속하는 검사기술이다.", "relation": "선형주사와 섹터주사로 용접부 체적을 빠르게 영상화하고 결함 위치·길이·높이를 평가한다.", "formula": "Delay law → steering·focusing", "related": ["탐촉자", "S-scan", "빔 조향", "집속"], "questions": ["PAUT의 원리와 특성"]},
]

GLOSSARY.append({
    "term": "강도 반사율",
    "aliases": ["음향강도 반사율", "intensity reflection coefficient", "RI", "R_I"],
    "category": "UT 기초",
    "definition": "두 매질의 경계면에 입사한 초음파의 음향강도 중 반사된 음향강도가 차지하는 비율이다. 값은 0~1 또는 0~100%로 나타낸다.",
    "relation": "수직입사에서 두 매질의 음향임피던스 차이가 클수록 강도 반사율이 커진다. 음압 반사계수는 진폭비이므로 부호를 가질 수 있지만, 강도 반사율은 그 제곱인 에너지 비율이다. 손실이 없는 경계면에서는 강도 투과율과의 합이 1이다.",
    "formula": "R_I = I_r/I_i = [(Z₂-Z₁)/(Z₂+Z₁)]² = r_p²\nT_I = 1-R_I  (손실이 없는 수직입사)\n예: Z₁=4, Z₂=6이면 R_I=[(6-4)/(6+4)]²=0.04=4%",
    "related": ["반사", "음향임피던스", "음압 반사계수", "강도 투과율", "매질"],
    "questions": ["음향임피던스와 경계면에서의 반사·투과 관계", "음압 반사계수와 강도 반사율의 차이"],
})

GLOSSARY.append({
    "term": "파동",
    "aliases": ["wave", "초음파 파동", "탄성파"],
    "category": "UT 기초",
    "definition": "진동이나 교란이 공간을 따라 전달되면서 에너지를 운반하는 현상이다. 매질 자체가 이동하는 것이 아니라 매질의 입자가 평형 위치 주위에서 진동한다.",
    "relation": "초음파탐상검사에서 초음파는 매질의 탄성에 의해 전달되는 기계적 파동이다. 따라서 진공에서는 전파되지 않는다. 서로 다른 매질의 경계면에서는 음향임피던스와 입사각에 따라 반사·투과·굴절 및 모드변환이 발생한다.",
    "formula": "c = fλ\nc: 음속, f: 주파수, λ: 파장\n종파: 입자진동∥진행방향\n횡파: 입자진동⊥진행방향",
    "related": ["매질", "음속", "주파수", "파장", "종파", "횡파", "표면파", "반사", "굴절"],
    "questions": ["파동의 정의와 초음파의 전파 특성", "종파·횡파·표면파·판파의 비교", "경계면에서 발생하는 반사·굴절 및 모드변환"],
})

GLOSSARY.append({
    "term": "대비시험편",
    "aliases": ["reference block", "reference specimen", "대비 시험편", "교정시험편"],
    "category": "UT 장비·교정",
    "definition": "실제 시험체와 유사한 재질·형상에 규정된 인공 반사체를 가공하여 탐상감도 설정과 결함 평가의 비교기준으로 사용하는 시험편이다.",
    "relation": "초음파탐상검사에서 탐상감도 설정, 거리별 에코 비교, DAC·TCG 작성, 결함 크기 추정 및 검사절차 검증에 사용한다. 평저공, 횡공·측면공, 노치 등이 인공 반사체로 사용된다.",
    "formula": "표준시험편: 장비·탐촉자의 기본 성능과 시간축 등을 교정\n대비시험편: 실제 검사조건의 감도 설정과 결함 평가\n유사조건: 재질·열처리·두께·곡률·표면상태·탐상거리·결함 방향",
    "related": ["표준시험편", "STB-A1", "STB-A2", "RB-4", "탐상감도", "DAC", "TCG", "인공 반사체"],
    "questions": ["표준시험편과 대비시험편의 용도 비교", "대비시험편의 인공 반사체 종류와 선정조건", "DAC·TCG 작성 시 대비시험편의 역할"],
})
