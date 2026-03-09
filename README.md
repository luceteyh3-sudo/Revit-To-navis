# Revit → Navisworks 일괄 변환기

Revit(.rvt) 파일을 Navisworks(.nwc) 파일로 일괄 변환하는 독립 실행형 GUI 프로그램입니다.

---

## 요구 사항

- **Autodesk Navisworks Manage 또는 Simulate** 설치 필요 (2019 이상 권장)
  - 변환 시 `FiletoolsTaskRunner.exe` 또는 `roamer.exe`를 사용합니다.
- Python 3.10 이상 (EXE 빌드 시에만 필요)

---

## 사용 방법

### 직접 실행 (Python)

```bash
python main.py
```

### EXE 빌드 후 실행

```
build_exe.bat 실행 → dist\RevitToNavisConverter.exe 생성
```

---

## 기능

| 기능 | 설명 |
|------|------|
| Navisworks 자동 감지 | 설치된 Navisworks 경로를 자동으로 탐색 |
| 재귀 파일 검색 | 선택한 경로의 모든 하위 폴더에서 .rvt 파일 탐색 |
| 폴더 구조 유지 | Revit 파일의 하위 폴더 구조 그대로 저장 경로에 재현 |
| 원본 파일 보호 | 원본 .rvt 파일은 변환하지 않고, .nwc 파일만 생성 |
| 진행률 표시 | 파일별 변환 상태 및 전체 진행률 실시간 표시 |
| 변환 중지 | 진행 중 언제든 중지 가능 |
| 로그 저장 | 변환 결과를 텍스트 파일로 저장 |
| 설정 자동 저장 | Navisworks 경로, 저장 경로를 자동 저장/복원 |

---

## 폴더 구조 예시

```
[Revit 경로]
  구역A/
    건축/
      건물1.rvt
    구조/
      건물1.rvt

[NWC 저장 경로]  ← 동일한 구조로 저장
  구역A/
    건축/
      건물1.nwc
    구조/
      건물1.nwc
```

---

## 변환 도구 우선순위

1. `FiletoolsTaskRunner.exe` (Navisworks 배치 변환 전용 도구, 권장)
2. `roamer.exe` (Navisworks 메인 실행 파일)
"# Revit-To-navis" 
