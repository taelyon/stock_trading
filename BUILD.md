# Stock Trader 빌드 가이드

## 📋 목차
1. [개요](#개요)
2. [필수 요구사항](#필수-요구사항)
3. [빌드 방법](#빌드-방법)
4. [빌드 파일 구조](#빌드-파일-구조)
5. [문제 해결](#문제-해결)

---

## 개요

이 문서는 Stock Trader 자동매매 프로그램을 실행 파일(`.exe`)로 빌드하는 방법을 설명합니다.

### 최근 업데이트 (2025-10-08)
- ✅ 중복 코드 리팩토링 완료 (`strategy_utils.py` 추가)
- ✅ 불필요한 메서드 제거 (코드 간소화)
- ✅ spec 파일 업데이트 (새 모듈 포함)

---

## 필수 요구사항

### 1. Python 환경
- **Python 버전**: 3.8 이상 권장
- **운영체제**: Windows 10/11 (64-bit)

### 2. 필수 라이브러리

#### 설치 명령어
```bash
pip install pyinstaller
pip install PyQt5
pip install pywin32
pip install pandas numpy
pip install matplotlib mplfinance
pip install TA-Lib
pip install requests slacker
pip install openpyxl
pip install pyautogui pygetwindow psutil
```

#### TA-Lib 설치 (중요!)
TA-Lib은 별도의 설치 과정이 필요합니다:

**방법 1**: whl 파일 다운로드
```bash
# https://www.lfd.uci.edu/~gohlke/pythonlibs/#ta-lib 에서 다운로드
pip install TA_Lib-0.4.XX-cp3X-cp3X-win_amd64.whl
```

**방법 2**: conda 사용
```bash
conda install -c conda-forge ta-lib
```

### 3. 필수 파일
빌드 전에 다음 파일들이 있는지 확인하세요:

```
stock_trading/
├── stock_trader.py          ✅ 메인 프로그램
├── strategy_utils.py         ✅ 전략 유틸리티 (리팩토링으로 추가)
├── backtester.py            ✅ 백테스팅 엔진
├── stock_trader.spec        ✅ PyInstaller 설정
├── stock_trader.ico         ✅ 아이콘 파일
├── settings.ini.example     ✅ 설정 예제
├── 빌드.bat                 ✅ 빌드 스크립트
└── 정리.bat                 ✅ 정리 스크립트
```

---

## 빌드 방법

### 방법 1: 배치 파일 사용 (권장)

1. **환경 진단 (선택사항, 권장)**
   ```cmd
   빌드_진단.bat
   ```
   - Python 버전 확인
   - PyInstaller 설치 확인
   - 필수 파일 확인
   - 패키지 설치 확인
   - import 테스트

2. **빌드 실행**
   ```cmd
   빌드.bat
   ```

3. **자동 실행 과정**
   - 필수 파일 확인
   - 이전 빌드 정리
   - PyInstaller 실행
   - 빌드 완료 후 `dist` 폴더 자동 열림

### 방법 2: 수동 빌드

1. **명령 프롬프트 열기**
   ```cmd
   cd C:\MyAPP\stock_trading
   ```

2. **이전 빌드 정리**
   ```cmd
   rmdir /s /q build
   del /q dist\stock_trader.exe
   ```

3. **PyInstaller 실행**
   ```cmd
   pyinstaller stock_trader.spec
   ```

4. **빌드 확인**
   ```cmd
   dir dist\stock_trader.exe
   ```

---

## 빌드 파일 구조

### spec 파일 주요 설정

#### 1. 포함할 Python 스크립트
```python
Analysis([
    'stock_trader.py',      # 메인 프로그램
    'strategy_utils.py',    # 전략 유틸리티
    'backtester.py'         # 백테스팅 엔진
])
```

#### 2. 데이터 파일
```python
datas = [
    ('settings.ini.example', '.'),
]
```

#### 3. 숨겨진 import 모듈
```python
hiddenimports = [
    # 자체 모듈
    'strategy_utils',
    'backtester',
    
    # PyQt5
    'PyQt5.QtWidgets',
    'PyQt5.QtCore',
    # ...
    
    # 크레온 API
    'win32com.client',
    # ...
    
    # 데이터 분석
    'pandas',
    'numpy',
    'talib',
    # ...
]
```

#### 4. 실행 파일 설정
```python
exe = EXE(
    name='stock_trader',
    console=False,          # GUI 모드
    icon='stock_trader.ico',
    uac_admin=True,         # 관리자 권한
    upx=True               # 압축 활성화
)
```

---

## 빌드 결과

### 성공 시
```
dist/
├── stock_trader.exe        # 실행 파일 (단일 파일)
├── settings.ini            # 설정 파일 (사용자 생성)
└── vi_stock_data.db        # 데이터베이스 (자동 생성)
```

### 파일 크기
- **예상 크기**: 약 150-250 MB
- **압축 포함**: UPX 활성화로 크기 최적화

---

## 문제 해결

### 0. 빌드.bat 실패하지만 수동 빌드는 성공하는 경우
**증상**: `pyinstaller stock_trader.spec`는 성공하는데 `빌드.bat`는 실패

**원인**:
1. **변수 지연 확장 문제** - 배치 파일 스크립트 오류
2. **인코딩 문제** - UTF-8 vs EUC-KR 코드페이지 충돌
3. **환경 변수 문제** - 배치 파일 내 변수 설정 오류

**해결책**:
```cmd
# 1. 진단 실행
빌드_진단.bat

# 2. 배치 파일 수정 확인 (이미 수정됨)
# - setlocal enabledelayedexpansion 추가됨
# - 오류 코드 캡처 추가됨

# 3. 여전히 실패 시 수동 빌드 사용
pyinstaller stock_trader.spec
```

**배치 파일 최신 수정사항** (2025-10-08):
- ✅ `setlocal enabledelayedexpansion` 추가 (변수 지연 확장)
- ✅ `BUILD_ERROR=%ERRORLEVEL%` 추가 (종료 코드 캡처)
- ✅ 상세한 오류 메시지 출력

### 1. TA-Lib 오류
**증상**: `ImportError: DLL load failed`

**해결책**:
```bash
# 1. TA-Lib whl 파일 설치
pip uninstall TA-Lib
pip install TA_Lib-0.4.XX-cp3X-cp3X-win_amd64.whl

# 2. 또는 conda 사용
conda install -c conda-forge ta-lib
```

### 2. win32com 오류
**증상**: `ModuleNotFoundError: No module named 'win32com'`

**해결책**:
```bash
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

### 3. PyQt5 오류
**증상**: `ModuleNotFoundError: No module named 'PyQt5'`

**해결책**:
```bash
pip uninstall PyQt5
pip install PyQt5==5.15.9
```

### 4. 빌드 실패 (일반)
**증상**: `Error: failed to execute script`

**해결책**:
```cmd
# 1. 이전 빌드 완전 정리
정리.bat

# 2. Python 캐시 삭제
del /s /q __pycache__
del /s /q *.pyc

# 3. 재빌드
빌드.bat
```

### 5. 실행 파일이 바이러스로 인식됨
**증상**: Windows Defender가 실행 파일을 차단

**해결책**:
```
1. Windows 보안 > 바이러스 및 위협 방지
2. 제외 항목 추가
3. dist\stock_trader.exe 경로 추가
```

### 6. 모듈을 찾을 수 없음
**증상**: 실행 시 특정 모듈을 찾을 수 없다는 오류

**해결책**:
`stock_trader.spec` 파일의 `hiddenimports`에 해당 모듈 추가:
```python
hiddenimports = [
    # ... 기존 목록 ...
    'your_missing_module',
]
```

---

## 빌드 최적화 팁

### 1. 빌드 속도 개선
```python
# stock_trader.spec
exe = EXE(
    # ...
    upx=False,  # UPX 비활성화 (속도 우선)
)
```

### 2. 파일 크기 최소화
```python
# stock_trader.spec
excludes = [
    'tkinter',
    'IPython',
    'jupyter',
    'pytest',
    # 사용하지 않는 모듈 추가
]
```

### 3. 디버그 모드
```python
# stock_trader.spec
exe = EXE(
    # ...
    console=True,   # 콘솔 창 표시 (에러 확인용)
    debug=True,     # 디버그 모드
)
```

---

## 배포 전 체크리스트

- [ ] 모든 필수 파일 존재 확인
- [ ] spec 파일에 새 모듈 포함 확인
- [ ] 빌드 성공 확인
- [ ] 실행 파일 테스트
- [ ] 설정 파일 (`settings.ini.example`) 포함 확인
- [ ] 아이콘 파일 적용 확인
- [ ] 관리자 권한 설정 확인
- [ ] 백신 프로그램 예외 처리

---

## 추가 정보

### PyInstaller 옵션
```bash
# 자세한 빌드 로그
pyinstaller --log-level=DEBUG stock_trader.spec

# 빌드 정보만 보기 (빌드 안 함)
pyinstaller --log-level=INFO stock_trader.spec --dry-run

# 클린 빌드
pyinstaller --clean stock_trader.spec
```

### 참고 문서
- PyInstaller 공식 문서: https://pyinstaller.org/
- TA-Lib 설치 가이드: https://github.com/mrjbq7/ta-lib
- 크레온 API 가이드: https://money2.creontrade.com/

---

## 변경 이력

### 2025-10-08
- ✅ `strategy_utils.py` 모듈 추가 (중복 코드 리팩토링)
- ✅ `backtester.py` 명시적 포함
- ✅ spec 파일 업데이트
- ✅ 빌드 배치 파일 개선
- ✅ 불필요한 메서드 제거 (`_check_momentum_buy`)

### 이전 버전
- 초기 빌드 설정

---

**문의사항이나 문제가 있으면 이슈를 등록해주세요!**
