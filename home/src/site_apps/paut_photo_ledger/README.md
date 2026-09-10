# PAUT Probe 및 Wedge 입고 사진대장

기존 현장 앱과 데이터에 연결되지 않는 독립형 데스크톱 앱입니다.

## 실행

```powershell
python home/src/site_apps/paut_photo_ledger/main.py
```

사진 원본은 수정하지 않습니다. 가져온 사진은 앱의 `photos` 폴더에 복사되고 품목 정보는 `data/items.json`에 저장됩니다.

## 자동 인식

새 사진 폴더를 가져오면 OCR이 모델명, 규격, S/N을 분석하고 같은 정보의 사진을 묶습니다. 결과는 반드시 화면에서 확인한 뒤 저장합니다. 글자가 흐리거나 보이지 않는 사진은 `확인 필요`로 남습니다.

필수 패키지: `openpyxl`, `Pillow`, `easyocr`
