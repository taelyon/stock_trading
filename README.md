# 🚀 Stock Trader - 대신증권 자동매매 프로그램

대신증권 크레온 API를 활용한 자동매매 프로그램

## ✨ 주요 기능

- ✅ 실시간 차트 모니터링 (30틱/3분/일봉)
- ✅ 기술적 지표 (MA, MACD, RSI, Stochastic, Bollinger Bands, VWAP 등)
- ✅ 조건검색 자동 편입
- ✅ VI 자동 감지
- ✅ 갭상승/급등주 전략
- ✅ 변동성 돌파 전략
- ✅ 자동 손절/익절
- ✅ Slack 알림
- ✅ **크레온 PLUS 자동 로그인**

## 📋 요구사항

### 필수
- Python 3.8+
- 대신증권 크레온 PLUS (필수!)
- Windows 10+

### Python 패키지
```bash
pip install PyQt5 pywin32 pandas numpy matplotlib mplfinance
pip install talib requests slacker openpyxl
pip install psutil pygetwindow pyautogui pyyaml
```

## 🔧 설정

### 1. settings.ini 파일 생성
```bash
copy settings.ini.example settings.ini
```

### 2. 로그인 정보 입력
```ini
[LOGIN]
username = YOUR_ID          # 대신증권 아이디
password = YOUR_PASSWORD    # 비밀번호
certpassword = YOUR_CERT    # 공인인증서 비밀번호
autologin = True            # 자동 로그인 활성화
```

## ▶️ 실행

### Python 스크립트
```bash
# 관리자 권한 CMD에서
python stock_trader.py
```

### 실행 파일 (빌드 후)
```bash
# 관리자 권한으로 실행
dist\stock_trader.exe 우클릭 → 관리자 권한으로 실행
```

## 🏗️ 빌드

### 빌드 방법
```bash
# 1. 정리 (선택사항)
정리.bat

# 2. 빌드
빌드.bat

# 3. 결과
dist\stock_trader.exe
```

### 디버그 빌드 (오류 확인)
```bash
디버그빌드.bat
```

## 📁 프로젝트 구조

```
stock_trading/
├── stock_trader.py              # 메인 프로그램
├── stock_trader.spec            # 빌드 설정
├── stock_trader_debug.spec      # 디버그 빌드 설정
├── stock_trader.ico             # 아이콘
├── settings.ini.example         # 설정 예제
├── backtester.py               # 백테스터
├── vi_stock_data.db            # 데이터베이스
├── 빌드.bat                     # 빌드 스크립트
├── 디버그빌드.bat                # 디버그 빌드
├── 정리.bat                     # 정리 스크립트
├── log/                        # 로그 폴더
└── dist/                       # 빌드 결과
    └── stock_trader.exe
```

## 🎯 사용 방법

1. **크레온 PLUS 로그인** (자동 로그인 설정 시 자동)
2. **프로그램 실행** (관리자 권한 필수)
3. **전략 선택**
4. **자동매매 시작**

## ⚙️ 최적화 설정

- 20개 종목 모니터링 최적화
- API 제한 방어 (15초당 60건)
- 메모리 효율화
- 실시간 데이터 우선

## ⚠️ 주의사항

1. **관리자 권한 필수** (크레온 API 요구사항)
2. **크레온 PLUS 로그인 유지**
3. **API 제한 준수** (자동 제어됨)
4. **settings.ini 보안** (비밀번호 평문 저장)

## 🐛 문제 해결

### 실행 안 됨
```bash
# 디버그 모드로 확인
디버그빌드.bat
→ Y 선택
→ 콘솔 창 오류 확인
```

### 로그 확인
```
log\trading_YYYYMMDD.log
```

## 📊 성능

- 시작 시간: ~10초
- 메모리: ~150-300MB
- CPU: ~5-10%

## 📄 라이선스

개인 사용 목적

---

**⚠️ 면책**: 투자 손실 책임은 사용자에게 있습니다.

