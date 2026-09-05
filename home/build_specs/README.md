# 분리 앱 빌드

기존 `report_manager.spec`와 `material_manager.spec`는 복구 비교용으로 유지한다.
새 빌드는 앱별 spec을 사용하며 서로 다른 실행 파일을 만든다.

```powershell
python -m PyInstaller --noconfirm --clean --workpath home/build/<app> --distpath home/dist home/build_specs/<app>.spec
```

`<app>`은 `central`, `kogas`, `lotte`, `ndt_report` 중 하나다.
검증 중에는 기존 산출물을 덮어쓰지 않도록 별도의 `--workpath`와 `--distpath`를 지정한다.
