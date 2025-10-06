import sys
import ctypes
from PyQt5.QtWidgets import *
from PyQt5.QtCore import (
    QTimer, pyqtSignal, QProcess, QObject, QThread, Qt, 
    pyqtSlot, QRunnable, QThreadPool, QEventLoop
)
from PyQt5.QtGui import QIcon, QPainter, QFont
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from datetime import datetime, timedelta
import pandas as pd
import win32com.client
import time
import numpy as np

import re
import mplfinance as mpf
import sqlite3
import os
import psutil
import configparser
import pygetwindow as gw
import requests
from io import BytesIO
import warnings
import logging
from openpyxl import Workbook
import json
from slacker import Slacker
import threading
import queue
import copy
import talib
import traceback
import pyautogui
from collections import deque

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# PLUS 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')

def init_plus_check():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        logging.error(f"오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요")
        return False
    if (cpStatus.IsConnect == 0):
        logging.error(f"PLUS가 정상적으로 연결되지 않음")
        return False
    if (cpTrade.TradeInit(0) != 0):
        logging.error(f"주문 초기화 실패")
        return False
    return True

def setup_logging():
    """로그 설정 (PyInstaller 대응)"""
    try:
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        logging.getLogger('matplotlib').setLevel(logging.WARNING)

        # ✅ 실행 파일 경로 확인
        if getattr(sys, 'frozen', False):
            # PyInstaller로 빌드된 경우
            application_path = os.path.dirname(sys.executable)
        else:
            # 일반 Python 실행
            application_path = os.path.dirname(os.path.abspath(__file__))

        # ✅ 로그 디렉토리 생성 (안전하게)
        log_dir = os.path.join(application_path, 'log')
        try:
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
        except Exception as e:
            # 로그 폴더 생성 실패 시 임시 폴더 사용
            log_dir = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'stock_trader_log')
            os.makedirs(log_dir, exist_ok=True)

        # 로그 파일 경로
        log_path = os.path.join(log_dir, f"trading_{datetime.now().strftime('%Y%m%d')}.log")
        
        # 파일 핸들러
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

        # 콘솔 핸들러
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)
        console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)
        
        logging.info(f"로그 초기화 완료: {log_path}")
        
    except Exception as ex:
        # 로그 설정 실패 시에도 프로그램은 계속 실행
        print(f"로그 설정 오류: {ex}")
        traceback.print_exc()

def send_slack_message(login_handler, channel, message):
    if login_handler is None or login_handler.slack is None:
        logging.warning("Slack 설정이 되어 있지 않습니다.")
        return
    sender = SlackMessageSender(login_handler, channel, message)
    QThreadPool.globalInstance().start(sender)

class SlackMessageSender(QRunnable):
    def __init__(self, login_handler, channel, message):
        super().__init__()
        self.login_handler = login_handler
        self.channel = channel
        self.message = message

    def run(self):
        try:
            if self.login_handler.slack:
                self.login_handler.slack.chat.post_message(self.channel, self.message)
        except Exception as ex:
            logging.error(f"Slack 메시지 전송 실패: {ex}")

# ==================== API 제한 관리자 ====================
class APILimiter:
    """API 호출 제한 관리"""
    
    def __init__(self):
        self.request_times = deque(maxlen=15)  # 최근 15개 요청 시간
        self.lock = threading.Lock()
        
    def wait_if_needed(self):
        wait_time = 0
        
        with self.lock:
            now = time.time()
            while len(self.request_times) >= 15:
                oldest = self.request_times[0]
                if now - oldest < 60:
                    wait_time = 60 - (now - oldest) + 0.1
                    break
                else:
                    self.request_times.popleft()
            
            if wait_time == 0 and len(self.request_times) > 0:
                last_request = self.request_times[-1]
                if now - last_request < 1.0:
                    wait_time = 1.0 - (now - last_request)
        
        # 락 밖에서 sleep
        if wait_time > 0:
            logging.debug(f"API 제한: {wait_time:.1f}초 대기")
            time.sleep(wait_time)
        
        with self.lock:
            self.request_times.append(time.time())

# 전역 API 제한자
api_limiter = APILimiter()

# ==================== 데이터 캐시 ====================
class DataCache:
    """종목 정보 캐싱"""
    
    def __init__(self, expire_seconds=300):
        self.cache = {}
        self.expire_seconds = expire_seconds
        self.lock = threading.Lock()
    
    def get(self, key):
        with self.lock:
            if key in self.cache:
                data, timestamp = self.cache[key]
                if time.time() - timestamp < self.expire_seconds:
                    return data
                else:
                    del self.cache[key]
            return None
    
    def set(self, key, value):
        with self.lock:
            self.cache[key] = (value, time.time())
    
    def clear(self):
        with self.lock:
            self.cache.clear()

# 전역 캐시
stock_info_cache = DataCache(expire_seconds=300)  # 5분

# ==================== 급등주 스캐너 (검증용으로만 사용) ====================
class MomentumScanner(QObject):
    """급등주 검증 - 조건검색 편입 종목 재확인용"""
    
    stock_found = pyqtSignal(dict)
    
    def __init__(self, trader):
        super().__init__()
        self.trader = trader
        self.cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
    
    # ⚠️ scan_market() 메서드 삭제 - 더 이상 직접 스캔 안함
    # ⚠️ start_screening() 메서드 삭제
    # ⚠️ stop_screening() 메서드 삭제
    
    def verify_momentum_conditions(self, code):
        """급등주 조건 재확인
        
        조건검색으로 들어온 종목이 실제로 급등주 조건을 만족하는지 검증
        
        Returns:
            (is_valid, score, message): (검증 통과 여부, 점수, 메시지)
        """
        try:
            # 캐시 확인
            cached_score = stock_info_cache.get(f"score_{code}")
            if cached_score is not None:
                return (cached_score >= 70, cached_score, "캐시에서 조회")
            
            # 현재가 정보
            api_limiter.wait_if_needed()
            
            self.cpStock.SetInputValue(0, code)
            self.cpStock.BlockRequest2(1)
            
            current_price = self.cpStock.GetHeaderValue(11)
            open_price = self.cpStock.GetHeaderValue(13)
            high_price = self.cpStock.GetHeaderValue(14)
            low_price = self.cpStock.GetHeaderValue(15)
            volume = self.cpStock.GetHeaderValue(18)
            prev_close = self.cpStock.GetHeaderValue(20)
            prev_volume = self.cpStock.GetHeaderValue(21)
            market_cap = self.cpStock.GetHeaderValue(67)  # 백만원 단위
            
            # 1차 필터링
            if current_price < 2000 or current_price > 50000:
                return (False, 0, f"가격대 미달 ({current_price}원)")
            
            if market_cap < 50000 or market_cap > 500000:
                return (False, 0, f"시가총액 미달 ({market_cap/10000:.0f}억)")
            
            score = 0
            
            # 1. 시가 대비 상승률 (0-30점)
            if open_price > 0:
                price_change_pct = (current_price - open_price) / open_price * 100
                
                if 2.0 <= price_change_pct < 3.5:
                    score += 30
                elif 3.5 <= price_change_pct < 5.0:
                    score += 20
                elif 5.0 <= price_change_pct < 7.0:
                    score += 10
                elif price_change_pct < 0:
                    return (False, 0, "시가 대비 하락")
            
            # 2. 거래량 비율 (0-25점)
            if prev_volume > 0:
                volume_ratio = volume / prev_volume
                
                if volume_ratio >= 5.0:
                    score += 25
                elif volume_ratio >= 3.0:
                    score += 20
                elif volume_ratio >= 2.0:
                    score += 10
                elif volume_ratio < 1.5:
                    return (False, 0, f"거래량 부족 ({volume_ratio:.1f}배)")
            
            # 3. 당일 고가 근처 유지 (0-20점)
            if high_price > 0 and low_price > 0:
                position = (current_price - low_price) / (high_price - low_price) if (high_price - low_price) > 0 else 0
                
                if position >= 0.8:
                    score += 20
                elif position >= 0.6:
                    score += 15
                elif position >= 0.4:
                    score += 10
            
            # 4. 시가 상승 유지 (0-15점)
            if current_price > open_price * 1.015:
                score += 15
            elif current_price > open_price:
                score += 10
            
            # 5. 시간대 가중치 (0-10점)
            now = datetime.now()
            if 9 <= now.hour < 10:
                score += 10
            elif 10 <= now.hour < 12:
                score += 7
            elif 13 <= now.hour < 14:
                score += 5
            
            stock_info_cache.set(f"score_{code}", score)
            
            is_valid = score >= 70
            message = f"급등주 점수: {score}/100"
            
            return (is_valid, score, message)
            
        except Exception as ex:
            logging.error(f"verify_momentum_conditions({code}): {ex}")
            return (False, 0, f"검증 오류: {ex}")
        
# ==================== 갭 상승 스캐너 (검증 + 매수조건) ====================
class GapUpScanner:
    """갭 상승 스캐너 - 검증 + 매수 조건 체크"""
    
    def __init__(self, trader):
        self.trader = trader
        self.cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
    
    def verify_gap_conditions(self, code):
        """갭 상승 조건 재확인
        
        조건검색으로 들어온 종목이 실제로 갭 상승 조건을 만족하는지 검증
        
        HTS 조건검색: 간단한 조건만
        - 갭 상승 1.5% ~ 4.0%
        - 현재가 > 시가
        
        Python 재검증: 정밀한 조건 추가
        - 현재가 >= 시가 * 0.99 (시가 대비 -1% 이내)
        - 거래량 비율 >= 150%
        
        Returns:
            (is_valid, gap_pct, message): (검증 통과 여부, 갭 비율, 메시지)
        """
        try:
            # 일봉 데이터 로드
            if not self.trader.daydata.select_code(code):
                return (False, 0, "일봉 데이터 로드 실패")
            
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('C', [])) < 2:
                return (False, 0, "데이터 부족")
            
            # 1️⃣ 갭 상승 비율 확인
            prev_close = day_data['C'][-2]
            today_open = day_data['O'][-1]
            
            if prev_close == 0:
                return (False, 0, "전일 종가 0")
            
            gap_pct = (today_open - prev_close) / prev_close * 100
            
            # 갭 상승 1.5% ~ 4%
            if not (1.5 <= gap_pct <= 4.0):
                return (False, gap_pct, f"갭 범위 미달 ({gap_pct:.2f}%)")
            
            # 2️⃣ 시가 유지 확인 (HTS에서 못하는 정밀 조건)
            # HTS: 현재가 > 시가
            # Python: 현재가 >= 시가 * 0.99 (시가 대비 -1% 이내)
            
            # 현재가 조회
            api_limiter.wait_if_needed()
            
            self.cpStock.SetInputValue(0, code)
            self.cpStock.BlockRequest2(1)
            current_price = self.cpStock.GetHeaderValue(11)
            
            if current_price < today_open * 0.99:
                return (False, gap_pct, f"시가 미유지 ({current_price}/{today_open}, {(current_price/today_open-1)*100:.2f}%)")
            
            # 3️⃣ 거래량 확인
            if len(day_data.get('V', [])) >= 2:
                today_vol = day_data['V'][-1]
                prev_vol = day_data['V'][-2]
                
                if prev_vol > 0:
                    volume_ratio = today_vol / prev_vol
                    
                    if volume_ratio < 1.5:
                        return (False, gap_pct, f"거래량 부족 ({volume_ratio:.1f}배)")
                    
                    # ✅ 모든 조건 통과
                    return (
                        True, 
                        gap_pct, 
                        f"갭상승 {gap_pct:.2f}%, 거래량 {volume_ratio:.1f}배, "
                        f"시가유지 {(current_price/today_open-1)*100:+.2f}%"
                    )
                else:
                    # 거래량 정보 없어도 갭과 시가 조건만으로 통과
                    return (
                        True, 
                        gap_pct, 
                        f"갭상승 {gap_pct:.2f}%, 시가유지 {(current_price/today_open-1)*100:+.2f}%"
                    )
            else:
                # 거래량 정보 없어도 갭과 시가 조건만으로 통과
                return (
                    True, 
                    gap_pct, 
                    f"갭상승 {gap_pct:.2f}%, 시가유지 {(current_price/today_open-1)*100:+.2f}%"
                )
            
        except Exception as ex:
            logging.error(f"verify_gap_conditions({code}): {ex}")
            return (False, 0, f"검증 오류: {ex}")
    
    def check_gap_hold(self, code):
        """갭 유지 확인 (매수 조건)
        
        매수 시점에 갭이 여전히 유지되고 있는지 확인
        시가 대비 -0.3% 이내면 갭 유지로 판단
        """
        try:
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('O', [])) == 0:
                return False
            
            today_open = day_data['O'][-1]
            
            # 현재가
            tick_latest = self.trader.tickdata.get_latest_data(code)
            current_price = tick_latest.get('C', 0)
            
            if current_price == 0:
                return False
            
            # 시가 대비 -0.3% 이내 (갭 유지)
            if current_price >= today_open * 0.997:
                return True
            
            return False
            
        except Exception as ex:
            logging.error(f"check_gap_hold({code}): {ex}")
            return False
        
# ==================== 변동성 돌파 전략 ====================
class VolatilityBreakout:
    """변동성 돌파 전략"""
    
    def __init__(self, trader):
        self.trader = trader
        self.K_value = 0.5
        self.target_prices = {}
        self.breakout_checked = set()
    
    def calculate_target_price(self, code):
        """목표가 계산"""
        
        try:
            # 일봉 데이터
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('H', [])) < 2:
                return None
            
            # 전일 고가/저가
            prev_high = day_data['H'][-2]
            prev_low = day_data['L'][-2]
            
            # 당일 시가
            today_open = day_data['O'][-1]
            
            # 변동폭
            range_value = prev_high - prev_low
            
            # 목표가
            target = today_open + (range_value * self.K_value)
            
            self.target_prices[code] = target
            
            return target
            
        except Exception as ex:
            logging.error(f"calculate_target_price({code}): {ex}")
            return None
    
    def check_breakout(self, code):
        """돌파 확인"""
        
        try:
            # 이미 체크했으면 스킵
            if code in self.breakout_checked:
                return False
            
            # 목표가 계산
            if code not in self.target_prices:
                target = self.calculate_target_price(code)
                if not target:
                    return False
            else:
                target = self.target_prices[code]
            
            # 현재가
            tick_latest = self.trader.tickdata.get_latest_data(code)
            current_price = tick_latest.get('C', 0)
            
            if current_price == 0:
                return False
            
            # 돌파 확인
            if current_price >= target:
                # 거래량 확인
                volume_ratio = self._get_volume_ratio(code)
                
                if volume_ratio >= 1.5:
                    self.breakout_checked.add(code)
                    
                    logging.info(
                        f"{cpCodeMgr.CodeToName(code)}({code}): "
                        f"변동성 돌파 (목표: {target:.0f}, "
                        f"현재: {current_price:.0f}, "
                        f"거래량비: {volume_ratio:.1f}배)"
                    )
                    
                    return True
            
            return False
            
        except Exception as ex:
            logging.error(f"check_breakout({code}): {ex}")
            return False
    
    def _get_volume_ratio(self, code):
        """거래량 비율"""
        try:
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('V', [])) < 2:
                return 0
            
            today_vol = day_data['V'][-1]
            prev_vol = day_data['V'][-2]
            
            if prev_vol > 0:
                return today_vol / prev_vol
            return 0
        except:
            return 0

# ==================== 기존 CpEvent 클래스 (유지) ====================
class CpEvent:
    def __init__(self):
        self.last_update_time = None

    def set_params(self, client, name, caller):
        self.client = client
        self.name = name
        self.caller = caller
        self.dic = {ord('1'): "종목별 VI", ord('2'): "배분정보", ord('3'): "기준가결정", ord('4'): "임의종료", ord('5'): "종목정보공개", ord('6'): "종목조치", ord('7'): "시장조치"}

    def OnReceived(self):
        if self.name == '9619s':
            time_num = self.client.GetHeaderValue(0)
            flag = self.client.GetHeaderValue(1)
            time_str = datetime.strptime(f"{time_num:06d}", '%H%M%S')
            combined_datetime = datetime.now().replace(hour=time_str.hour, minute=time_str.minute, second=time_str.second)
            time = combined_datetime.strftime('%m/%d %H:%M:%S')

            if self.dic.get(flag) == "종목별 VI":
                code = self.client.GetHeaderValue(3)
                event = self.client.GetHeaderValue(5)
                event2 = self.client.GetHeaderValue(6)
                match1 = re.search(r'^A\d{6}$', code)
                match2 = re.search(r"괴리율:(-?\d+\.\d+)%", event2)

                if (cpCodeMgr.GetStockControlKind(code) == 0 and
                    cpCodeMgr.GetStockSectionKind(code) == 1 and
                    match1 and match2 and "정적" in event):
                        gap_rate = float(match2.group(1))
                        if gap_rate > 0 and combined_datetime < combined_datetime.replace(hour=15, minute=15, second=0):
                            logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> VI 발동")
                            self.caller.monitor_vi(time, code, event2)
            return
        
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)
            timess = self.client.GetHeaderValue(18)
            exFlag = self.client.GetHeaderValue(19)
            cprice = self.client.GetHeaderValue(13)
            cVol = self.client.GetHeaderValue(17)

            if (exFlag == ord('2')):
                item = {'code': code, 'time': timess, 'cur': cprice, 'vol': cVol}
                self.caller.updateCurData(item)
            return
        
        if self.name == 'cssalert':
            stgid = self.client.GetHeaderValue(0)
            stgmonid = self.client.GetHeaderValue(1)
            code = self.client.GetHeaderValue(2)
            inoutflag = self.client.GetHeaderValue(3)
            stgtime = self.client.GetHeaderValue(4)
            stgprice = self.client.GetHeaderValue(5)
            time_str = datetime.strptime(stgtime.zfill(6), '%H%M%S')
            combined_datetime = datetime.now().replace(hour=time_str.hour, minute=time_str.minute, second=time_str.second)
            time = combined_datetime.strftime('%m/%d %H:%M:%S')
            if inoutflag == ord('1') and combined_datetime < combined_datetime.replace(hour=15, minute=15, second=0):
                self.caller.checkRealtimeStg(stgid, stgmonid, code, stgprice, time)
            return
        
        if self.name == 'conclusion':
            conflag = self.client.GetHeaderValue(14)
            ordernum = self.client.GetHeaderValue(5)
            qty = self.client.GetHeaderValue(3)
            price = self.client.GetHeaderValue(4)
            code = self.client.GetHeaderValue(9)
            bs = self.client.GetHeaderValue(12)
            buyprice = self.client.GetHeaderValue(21)
            balance = self.client.GetHeaderValue(23)
            conflags = {"1": "체결", "2": "확인", "3": "거부", "4": "접수"}.get(conflag, "")
            self.caller.monitorOrderStatus(code, ordernum, conflags, price, qty, bs, balance, buyprice)

# ==================== 기존 Publish 클래스들 (유지) ====================
class CpPublish:
    def __init__(self, name, service_id):
        self.name = name
        self.obj = win32com.client.Dispatch(service_id)
        self.bIsSB = False

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if len(var) > 0:
            self.obj.SetInputValue(0, var)
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
            self.bIsSB = False

class CpPBConclusion(CpPublish):
    def __init__(self):
        super().__init__('conclusion', 'DsCbo1.CpConclusion')

class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')

class CpPB9619(CpPublish):
    def __init__(self):
        super().__init__('9619s', 'CpSysDib.CpSvr9619s')

class CpPBCssAlert(CpPublish):
    def __init__(self):
        super().__init__('cssalert', 'CpSysDib.CssAlert')

class CpRequest:
    def __init__(self):
        self.result = None
        self.loop = QEventLoop()
        self.is_finished = False
        self.client = None
        self.name = None
        self.caller = None
        self.params = {}
        self.is_requesting = False

    def set_params(self, client, name, caller, params=None):
        self.client = client
        self.name = name
        self.caller = caller
        self.params = params if params else {}

    def wait(self):
        if not self.is_finished:
            self.loop.exec_()
        return self.result

    def OnReceived(self):
        try:
            if self.name == 'order':
                rqStatus = self.client.GetDibStatus()
                if rqStatus != 0:
                    rqRet = self.client.GetDibMsg1()
                    logging.warning(f"{self.params.get('stock_name')}({self.params.get('code')}) 주문 요청 오류, {rqRet}")
                    self.result = False
                    return
                self.result = True

        except Exception as ex:
            logging.error(f"OnReceived -> {ex}")
            self.result = False

        finally:
            self.is_finished = True
            if self.loop.isRunning():
                self.loop.quit()
            if self.caller and hasattr(self.caller, 'cp_request'):
                self.caller.cp_request.is_requesting = False

# ==================== CpStrategy (조건검색 편입 처리 수정) ====================
class CpStrategy:
    def __init__(self, trader):
        self.monList = {}
        self.trader = trader
        self.stgname = {}
        self.objpb = CpPBCssAlert()
        
        # ✅ 검증용 스캐너 (조건검색 편입 종목 재확인용)
        self.momentum_scanner = None
        self.gap_scanner = None

    def requestList(self):
        retStgList = {}
        objRq = win32com.client.Dispatch("CpSysDib.CssStgList")
        objRq.SetInputValue(0, ord('1'))
        objRq.BlockRequest2(1)

        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            logging.warning(f"나의전략 조회실패, {rqStatus}, {rqRet}")
            return (False, retStgList)

        cnt = objRq.GetHeaderValue(0)
        flag = objRq.GetHeaderValue(1)

        for i in range(cnt):
            item = {}
            item['전략명'] = objRq.GetDataValue(0, i)
            item['ID'] = objRq.GetDataValue(1, i)
            item['평균수익률'] = objRq.GetDataValue(6, i)
            retStgList[item['전략명']] = item
        return retStgList

    def requestStgID(self, id):
        retStgstockList = []
        objRq = None
        objRq = win32com.client.Dispatch("CpSysDib.CssStgFind")
        objRq.SetInputValue(0, id)
        objRq.BlockRequest2(1)
        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            logging.warning(f"전략ID 조회실패, {rqStatus}, {rqRet}")
            return (False, retStgstockList)

        cnt = objRq.GetHeaderValue(0)
        totcnt = objRq.GetHeaderValue(1)
        stime = objRq.GetHeaderValue(2)

        for i in range(cnt):
            item = {}
            item['code'] = objRq.GetDataValue(0, i)
            item['종목명'] = cpCodeMgr.CodeToName(item['code'])
            retStgstockList.append(item)

        return (True, retStgstockList)

    def requestMonitorID(self, id):
        objRq = win32com.client.Dispatch("CpSysDib.CssWatchStgSubscribe")
        objRq.SetInputValue(0, id)
        objRq.BlockRequest2(1)

        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            logging.warning(f"감시번호 조회실패, {rqStatus}, {rqRet}")
            return (False, 0)

        monID = objRq.GetHeaderValue(0)
        if monID == 0:
            logging.warning(f"감시 일련번호 구하기 실패")
            return (False, 0)

        return (True, monID)

    def requestStgControl(self, id, monID, bStart, stgname):
        objRq = win32com.client.Dispatch("CpSysDib.CssWatchStgControl")
        objRq.SetInputValue(0, id)
        objRq.SetInputValue(1, monID)

        if bStart == True:
            objRq.SetInputValue(2, ord('1'))
            self.stgname[id] = stgname
        else:
            objRq.SetInputValue(2, ord('3'))
            if id in self.stgname:
                del self.stgname[id]
        objRq.BlockRequest2(1)

        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            logging.warning(f"감시시작 실패, {rqStatus}, {rqRet}")
            return (False, '')

        status = objRq.GetHeaderValue(0)

        self.objpb.Subscribe('', self)

        if bStart == True:
            self.monList[id] = monID
        else:
            if id in self.monList:
                del self.monList[id]

        return (True, status)

    def checkRealtimeStg(self, stgid, stgmonid, code, stgprice, time):
        """조건검색 편입 시 호출 - 재검증 후 투자대상 추가"""
        
        if stgid not in self.monList:
            return
        if (stgmonid != self.monList[stgid]):
            return
        
        # API 제한 체크
        remain_time0 = cpStatus.GetLimitRemainTime(0)
        remain_time1 = cpStatus.GetLimitRemainTime(1)
        if remain_time0 != 0 or remain_time1 != 0:
            return
        
        # 장 시작 후에만 처리
        if datetime.now() < datetime.now().replace(hour=9, minute=3, second=0, microsecond=0):
            return
        
        # 이미 모니터링 중이면 스킵
        if code in self.trader.monistock_set:
            return
        
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # ✅ 조건검색별 재검증
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        
        stgname = self.stgname.get(stgid, '')
        stock_name = cpCodeMgr.CodeToName(code)
        
        try:
            # 1️⃣ 급등주 조건검색
            if stgname == '급등주':
                if not self.momentum_scanner:
                    logging.warning("MomentumScanner가 초기화되지 않음")
                    return
                
                # 급등주 조건 재확인
                is_valid, score, message = self.momentum_scanner.verify_momentum_conditions(code)
                
                if not is_valid:
                    logging.debug(f"{stock_name}({code}): 급등주 재검증 실패 - {message}")
                    return
                
                logging.info(
                    f"{stock_name}({code}) -> 급등주 조건검색 편입 "
                    f"(검증: {message}, 체결강도 확인중...)"
                )
                
                # 추가 검증: 체결강도 (데이터 로드 후)
                if self.trader.daydata.select_code(code):
                    if self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                        # 잠시 대기 후 체결강도 확인
                        time.sleep(0.5)
                        strength = self.trader.tickdata.get_strength(code)
                        
                        if strength >= 120:
                            # ✅ 투자대상 추가
                            self._add_to_monitoring(code, stgprice, time, f"급등주 (점수: {score}, 체결강도: {strength:.0f})")
                        else:
                            logging.debug(f"{stock_name}({code}): 체결강도 부족 ({strength:.0f})")
                            self.trader.daydata.monitor_stop(code)
                            self.trader.tickdata.monitor_stop(code)
                            self.trader.mindata.monitor_stop(code)
                    else:
                        self.trader.daydata.monitor_stop(code)
                else:
                    pass
            
            # 2️⃣ 갭상승 조건검색
            elif stgname == '갭상승':
                if not self.gap_scanner:
                    logging.warning("GapUpScanner가 초기화되지 않음")
                    return
                
                # 갭상승 조건 재확인
                is_valid, gap_pct, message = self.gap_scanner.verify_gap_conditions(code)
                
                if not is_valid:
                    logging.debug(f"{stock_name}({code}): 갭상승 재검증 실패 - {message}")
                    self.trader.daydata.monitor_stop(code)
                    return
                
                logging.info(
                    f"{stock_name}({code}) -> 갭상승 조건검색 편입 "
                    f"(검증: {message})"
                )
                
                # 데이터 로드
                if self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                    # ✅ 투자대상 추가
                    self._add_to_monitoring(code, stgprice, time, f"갭상승 ({gap_pct:.2f}%)")
                else:
                    self.trader.daydata.monitor_stop(code)
            
            # 3️⃣ 기타 조건검색 (VI 발동 등)
            else:
                # 기존 로직 유지
                if self.trader.daydata.select_code(code):
                    if self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                        if code not in self.trader.starting_time:
                            self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                        self.trader.monistock_set.add(code)
                        self.trader.stock_added_to_monitor.emit(code)
                        self.trader.save_list_db(code, self.trader.starting_time[code], self.trader.starting_price[code], 1)
                        logging.info(f"{stock_name}({code}) -> 투자 대상 종목 추가")
                    else:
                        self.trader.daydata.monitor_stop(code)
                        self.trader.mindata.monitor_stop(code)
                        self.trader.tickdata.monitor_stop(code)
        
        except Exception as ex:
            logging.error(f"checkRealtimeStg({code}, {stgname}): {ex}\n{traceback.format_exc()}")

    def _add_to_monitoring(self, code, price, time, reason):
        """투자대상 종목 추가"""
        try:
            stock_name = cpCodeMgr.CodeToName(code)
            
            self.trader.starting_time[code] = time
            self.trader.starting_price[code] = price
            self.trader.monistock_set.add(code)
            self.trader.stock_added_to_monitor.emit(code)
            self.trader.save_list_db(code, time, price, 1)
            
            logging.info(f"{stock_name}({code}) -> 투자 대상 추가: {reason}")
            
        except Exception as ex:
            logging.error(f"_add_to_monitoring({code}): {ex}")

    def Clear(self):
        delitem = []
        for id, monId in self.monList.items():
            delitem.append((id, monId))

        for item in delitem:
            self.requestStgControl(item[0], item[1], False, "Unknown")

        self.objpb.Unsubscribe()

# ==================== CpIndicators (유지) ====================
class CpIndicators:
    def __init__(self, chart_type):
        self.chart_type = chart_type
        self.params = self._get_default_params()
    
    def _get_default_params(self):
        if self.chart_type == 'T':
            return {
                'MA_PERIODS': [5, 20, 60, 120],
                'RSI_PERIOD': 12,
                'RSI_SIGNAL_PERIOD': 12,
                'MACD': (12, 26, 9),
                'STOCH': (7, 3, 3),
                'ATR_PERIOD': 10,
                'CCI_PERIOD': 12,
                'BB_PERIOD': 20,
                'BB_STD': 2,
                'WILLIAMS_R_PERIOD': 14,  # 추가
                'ROC_PERIOD': 10,
            }
        elif self.chart_type == 'm':
            return {
                'MA_PERIODS': [5, 10, 20],
                'RSI_PERIOD': 9,
                'RSI_SIGNAL_PERIOD': 9,
                'MACD': (12, 26, 9),
                'STOCH': (7, 3, 3),
                'ATR_PERIOD': 10,
                'CCI_PERIOD': 12,
                'BB_PERIOD': 20,
                'BB_STD': 2,
                'WILLIAMS_R_PERIOD': 14,  # 추가
                'ROC_PERIOD': 10,
            }
        elif self.chart_type == 'D':
            return {
                'MA_PERIODS': [5, 10],
                'MACD': (12, 26, 9)
            }
        return {}
    
    def _validate_data(self, chart_data, min_length):
        required_keys = ['C', 'H', 'L', 'V', 'D', 'T']
        for key in required_keys:
            if key not in chart_data:
                return False
            if len(chart_data[key]) < min_length:
                return False
        return True
    
    def _fill_nan(self, data, method='smart'):
        if method == 'smart':
            series = pd.Series(data)
            filled = series.fillna(method='ffill').fillna(method='bfill').fillna(0)
            return filled.tolist()
        else:
            return np.nan_to_num(data, nan=0.0).tolist()

    def calculate_williams_r(self, highs, lows, closes, period=14):
        """Williams %R 계산
        
        과매수/과매도 지표
        -20 이상: 과매수 (매도 신호)
        -80 이하: 과매도 (매수 신호)
        
        Args:
            highs: 고가 리스트
            lows: 저가 리스트
            closes: 종가 리스트
            period: 계산 기간 (기본 14)
        
        Returns:
            Williams %R 리스트 (-100 ~ 0)
        """
        williams_r = []
        
        for i in range(len(closes)):
            if i < period - 1:
                williams_r.append(-50)  # 기본값
                continue
            
            # 최근 N일의 최고가/최저가
            high_max = max(highs[i-period+1:i+1])
            low_min = min(lows[i-period+1:i+1])
            
            if high_max - low_min == 0:
                williams_r.append(-50)
            else:
                # Williams %R = (최고가 - 현재가) / (최고가 - 최저가) * -100
                wr = ((high_max - closes[i]) / (high_max - low_min)) * -100
                williams_r.append(wr)
        
        return williams_r
    
    def calculate_roc(self, closes, period=10):
        """Price Rate of Change (ROC) 계산
        
        가격 변화율 - 모멘텀 지표
        양수: 상승 추세
        음수: 하락 추세
        
        Args:
            closes: 종가 리스트
            period: 계산 기간 (기본 10)
        
        Returns:
            ROC 리스트 (%)
        """
        roc = []
        
        for i in range(len(closes)):
            if i < period:
                roc.append(0)
            else:
                if closes[i-period] != 0:
                    # ROC = (현재가 - N일전 가격) / N일전 가격 * 100
                    roc_value = ((closes[i] - closes[i-period]) / closes[i-period]) * 100
                    roc.append(roc_value)
                else:
                    roc.append(0)
        
        return roc
    
    def calculate_obv(self, closes, volumes):
        """On-Balance Volume (OBV) 계산
        
        거래량 기반 추세 확인
        OBV 상승 + 가격 상승: 강한 상승 추세
        OBV 하락 + 가격 상승: 약한 상승 추세 (다이버전스)
        
        Args:
            closes: 종가 리스트
            volumes: 거래량 리스트
        
        Returns:
            OBV 리스트
        """
        obv = [0]
        
        for i in range(1, len(closes)):
            if closes[i] > closes[i-1]:
                # 상승 시 거래량 더함
                obv.append(obv[-1] + volumes[i])
            elif closes[i] < closes[i-1]:
                # 하락 시 거래량 뺌
                obv.append(obv[-1] - volumes[i])
            else:
                # 보합 시 유지
                obv.append(obv[-1])
        
        return obv
    
    def calculate_obv_ma(self, obv, period=20):
        """OBV 이동평균 계산
        
        Args:
            obv: OBV 리스트
            period: 계산 기간 (기본 20)
        
        Returns:
            OBV MA 리스트
        """
        obv_ma = []
        
        for i in range(len(obv)):
            if i < period - 1:
                obv_ma.append(obv[i])
            else:
                ma = sum(obv[i-period+1:i+1]) / period
                obv_ma.append(ma)
        
        return obv_ma
    
    def calculate_volume_profile(self, closes, volumes, bins=20):
        """Volume Profile 계산
        
        가격대별 거래량 분포 분석
        
        Args:
            closes: 종가 리스트
            volumes: 거래량 리스트
            bins: 가격대 구간 수
        
        Returns:
            (max_volume_price, current_vs_poc): 최대 거래량 가격, 현재가 위치
        """
        if len(closes) == 0 or len(volumes) == 0:
            return 0, 0
        
        # 가격 범위
        price_min = min(closes)
        price_max = max(closes)
        
        if price_max == price_min:
            return closes[-1], 0
        
        # 가격대별 거래량 집계
        bin_size = (price_max - price_min) / bins
        volume_profile = {}
        
        for price, volume in zip(closes, volumes):
            bin_index = int((price - price_min) / bin_size) if bin_size > 0 else 0
            bin_index = min(bin_index, bins - 1)  # 상한선
            
            volume_profile[bin_index] = volume_profile.get(bin_index, 0) + volume
        
        # 최대 거래량 가격대 (POC: Point of Control)
        if volume_profile:
            max_volume_bin = max(volume_profile, key=volume_profile.get)
            max_volume_price = price_min + (max_volume_bin + 0.5) * bin_size
        else:
            max_volume_price = closes[-1]
        
        # 현재가 vs POC
        current_price = closes[-1]
        current_vs_poc = (current_price - max_volume_price) / max_volume_price if max_volume_price > 0 else 0
        
        return max_volume_price, current_vs_poc
    
    def _get_default_result(self, indicator_type, length):
        """기본 결과값 반환 (데이터 부족 시)"""
        default_value = [0] * length
        
        if indicator_type == 'MA':
            if self.chart_type == 'T':
                return {
                    'MAT5': default_value, 'MAT20': default_value,
                    'MAT60': default_value, 'MAT120': default_value,
                    'MAT5_MAT20_DIFF': default_value, 'MAT20_MAT60_DIFF': default_value,
                    'MAT60_MAT120_DIFF': default_value, 'C_MAT5_DIFF': default_value,
                    'MAT5_CHANGE': default_value, 'MAT20_CHANGE': default_value,
                    'MAT60_CHANGE': default_value, 'MAT120_CHANGE': default_value
                }
            elif self.chart_type == 'm':
                return {
                    'MAM5': default_value, 'MAM10': default_value, 'MAM20': default_value,
                    'MAM5_MAM10_DIFF': default_value, 'MAM10_MAM20_DIFF': default_value,
                    'C_MAM5_DIFF': default_value, 'C_ABOVE_MAM5': default_value
                }
            elif self.chart_type == 'D':
                return {'MAD5': default_value, 'MAD10': default_value}
        
        elif indicator_type == 'MACD':
            if self.chart_type == 'T':
                return {
                    'MACDT': default_value, 'MACDT_SIGNAL': default_value,
                    'OSCT': default_value
                }
            else:
                return {
                    'MACD': default_value, 'MACD_SIGNAL': default_value,
                    'OSC': default_value
                }
        
        elif indicator_type == 'RSI':
            if self.chart_type == 'T':
                return {'RSIT': default_value, 'RSIT_SIGNAL': default_value}
            else:
                return {'RSI': default_value, 'RSI_SIGNAL': default_value}
        
        elif indicator_type == 'STOCH':
            return {'STOCHK': default_value, 'STOCHD': default_value}
        
        elif indicator_type == 'ATR':
            return {'ATR': default_value}
        
        elif indicator_type == 'CCI':
            return {'CCI': default_value}
        
        elif indicator_type == 'BBANDS':
            return {
                'BB_UPPER': default_value, 'BB_MIDDLE': default_value,
                'BB_LOWER': default_value, 'BB_POSITION': default_value,
                'BB_BANDWIDTH': default_value
            }
        
        elif indicator_type == 'VWAP':
            return {'VWAP': default_value}
        
        # === 새로운 지표들 ===
        elif indicator_type == 'WILLIAMS_R':
            return {'WILLIAMS_R': default_value}
        
        elif indicator_type == 'ROC':
            return {'ROC': default_value}
        
        elif indicator_type == 'OBV':
            return {'OBV': default_value, 'OBV_MA20': default_value}
        
        elif indicator_type == 'VOLUME_PROFILE':
            return {'VP_POC': 0, 'VP_POSITION': 0}
        
        return {}

    def make_indicator(self, indicator_type, code, chart_data):
        try:
            result = {}
            closes = np.array(chart_data['C'], dtype=np.float64)
            highs = np.array(chart_data['H'], dtype=np.float64)
            lows = np.array(chart_data['L'], dtype=np.float64)
            volumes = np.array(chart_data['V'], dtype=np.float64)
            dates = np.array(chart_data['D'], dtype=np.int64)
            
            desired_length = len(closes)
            
            min_lengths = {
                'MA': max(self.params.get('MA_PERIODS', [5])),
                'MACD': 35,
                'RSI': self.params.get('RSI_PERIOD', 14),
                'STOCH': 14,
                'ATR': self.params.get('ATR_PERIOD', 14),
                'CCI': self.params.get('CCI_PERIOD', 14),
                'BBANDS': self.params.get('BB_PERIOD', 20),
                'VWAP': 1,
                'WILLIAMS_R': self.params.get('WILLIAMS_R_PERIOD', 14),
                'ROC': self.params.get('ROC_PERIOD', 10),
                'OBV': 2,
                'VOLUME_PROFILE': 20
            }
            
            min_required = min_lengths.get(indicator_type, 20)
            
            if not self._validate_data(chart_data, min_required):
                logging.debug(f"{code}: {indicator_type} 데이터 부족 ({len(closes)} < {min_required})")
                return self._get_default_result(indicator_type, desired_length)
            
            if indicator_type == 'MA':
                if self.chart_type == 'T':
                    terms = self.params['MA_PERIODS']
                    index_names = ["MAT5", "MAT20", "MAT60", "MAT120"]
                elif self.chart_type == 'm':
                    terms = self.params['MA_PERIODS']
                    index_names = ["MAM5", "MAM10", "MAM20"]
                elif self.chart_type == 'D':
                    terms = self.params['MA_PERIODS']
                    index_names = ["MAD5", "MAD10"]
                
                for term, name in zip(terms, index_names):
                    sma = talib.SMA(closes, timeperiod=term)
                    result[name] = self._fill_nan(sma)
                
                if self.chart_type == 'T':
                    result['MAT5_MAT20_DIFF'] = [
                        (result['MAT5'][i] - result['MAT20'][i]) / result['MAT20'][i] 
                        if result['MAT20'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['MAT20_MAT60_DIFF'] = [
                        (result['MAT20'][i] - result['MAT60'][i]) / result['MAT60'][i] 
                        if result['MAT60'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['MAT60_MAT120_DIFF'] = [
                        (result['MAT60'][i] - result['MAT120'][i]) / result['MAT120'][i] 
                        if result['MAT120'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['C_MAT5_DIFF'] = [
                        (closes[i] - result['MAT5'][i]) / result['MAT5'][i] 
                        if result['MAT5'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    
                    for name in ['MAT5', 'MAT20', 'MAT60', 'MAT120']:
                        ma_values = result[name]
                        changes = [0] + [
                            (ma_values[i] - ma_values[i-1]) / ma_values[i-1]
                            if ma_values[i-1] != 0 else 0
                            for i in range(1, len(ma_values))
                        ]
                        result[f'{name}_CHANGE'] = changes
                
                elif self.chart_type == 'm':
                    result['MAM5_MAM10_DIFF'] = [
                        (result['MAM5'][i] - result['MAM10'][i]) / result['MAM10'][i] 
                        if result['MAM10'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['MAM10_MAM20_DIFF'] = [
                        (result['MAM10'][i] - result['MAM20'][i]) / result['MAM20'][i] 
                        if result['MAM20'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['C_MAM5_DIFF'] = [
                        (closes[i] - result['MAM5'][i]) / result['MAM5'][i] 
                        if result['MAM5'][i] != 0 else 0 
                        for i in range(desired_length)
                    ]
                    result['C_ABOVE_MAM5'] = [
                        1 if closes[i] > result['MAM5'][i] else 0 
                        for i in range(desired_length)
                    ]

            elif indicator_type == 'MACD':
                fast, slow, signal = self.params['MACD']
                macd_line, signal_line, macd_hist = talib.MACD(
                    closes, fastperiod=fast, slowperiod=slow, signalperiod=signal
                )
                
                if self.chart_type == 'T':
                    result['MACDT'] = self._fill_nan(macd_line)
                    result['MACDT_SIGNAL'] = self._fill_nan(signal_line)
                    result['OSCT'] = self._fill_nan(macd_hist)
                else:
                    result['MACD'] = self._fill_nan(macd_line)
                    result['MACD_SIGNAL'] = self._fill_nan(signal_line)
                    result['OSC'] = self._fill_nan(macd_hist)

            elif indicator_type == 'RSI':
                rsi_period = self.params['RSI_PERIOD']
                rsi_signal_period = self.params['RSI_SIGNAL_PERIOD']
                
                rsi_values = talib.RSI(closes, timeperiod=rsi_period)
                rsi_signal = talib.SMA(rsi_values, timeperiod=rsi_signal_period)
                
                if self.chart_type == 'T':
                    result['RSIT'] = self._fill_nan(rsi_values)
                    result['RSIT_SIGNAL'] = self._fill_nan(rsi_signal)
                else:
                    result['RSI'] = self._fill_nan(rsi_values)
                    result['RSI_SIGNAL'] = self._fill_nan(rsi_signal)

            elif indicator_type == 'STOCH':
                fastk_period, slowk_period, slowd_period = self.params['STOCH']
                stoch_k, stoch_d = talib.STOCH(
                    highs, lows, closes,
                    fastk_period=fastk_period,
                    slowk_period=slowk_period,
                    slowd_period=slowd_period
                )
                result['STOCHK'] = self._fill_nan(stoch_k)
                result['STOCHD'] = self._fill_nan(stoch_d)

            elif indicator_type == 'ATR':
                atr_period = self.params['ATR_PERIOD']
                atr = talib.ATR(highs, lows, closes, timeperiod=atr_period)
                result['ATR'] = self._fill_nan(atr)

            elif indicator_type == 'CCI':
                cci_period = self.params['CCI_PERIOD']
                cci = talib.CCI(highs, lows, closes, timeperiod=cci_period)
                result['CCI'] = self._fill_nan(cci)

            elif indicator_type == 'BBANDS':
                bb_period = self.params['BB_PERIOD']
                bb_std = self.params['BB_STD']
                
                upper, middle, lower = talib.BBANDS(
                    closes, timeperiod=bb_period,
                    nbdevup=bb_std, nbdevdn=bb_std
                )
                
                result['BB_UPPER'] = self._fill_nan(upper)
                result['BB_MIDDLE'] = self._fill_nan(middle)
                result['BB_LOWER'] = self._fill_nan(lower)
                
                bandwidth = upper - lower
                # ✅ 경고 억제하여 안전하게 나누기
                with np.errstate(divide='ignore', invalid='ignore'):
                    bb_position = np.where(
                        bandwidth > 1e-6,
                        (closes - middle) / bandwidth,
                        0.5
                    )
                bb_position = np.clip(bb_position, -2, 2)
                result['BB_POSITION'] = bb_position.tolist()
                
                # ✅ 경고 억제하여 안전하게 나누기
                with np.errstate(divide='ignore', invalid='ignore'):
                    bb_bandwidth = np.where(
                        middle > 1e-6,
                        bandwidth / middle,
                        0
                    )
                result['BB_BANDWIDTH'] = bb_bandwidth.tolist()

            elif indicator_type == 'VWAP':
                vwap = np.zeros_like(closes)
                
                if len(dates) == len(closes):
                    unique_dates = np.unique(dates)
                    
                    for d in unique_dates:
                        mask = dates == d
                        day_closes = closes[mask]
                        day_volumes = volumes[mask]
                        
                        cumsum_pv = np.cumsum(day_closes * day_volumes)
                        cumsum_v = np.cumsum(day_volumes)
                        
                        day_vwap = np.divide(
                            cumsum_pv, cumsum_v,
                            out=np.zeros_like(cumsum_pv),
                            where=cumsum_v != 0
                        )
                        vwap[mask] = day_vwap
                
                result['VWAP'] = vwap.tolist()

            elif indicator_type == 'WILLIAMS_R':
                """Williams %R 계산"""
                period = self.params.get('WILLIAMS_R_PERIOD', 14)
                
                williams_r = self.calculate_williams_r(
                    highs.tolist(), 
                    lows.tolist(), 
                    closes.tolist(), 
                    period
                )
                
                result['WILLIAMS_R'] = williams_r
            
            elif indicator_type == 'ROC':
                """ROC 계산"""
                period = self.params.get('ROC_PERIOD', 10)
                
                roc = self.calculate_roc(closes.tolist(), period)
                
                result['ROC'] = roc
            
            elif indicator_type == 'OBV':
                """OBV 및 OBV MA 계산"""
                obv = self.calculate_obv(closes.tolist(), volumes.tolist())
                obv_ma20 = self.calculate_obv_ma(obv, period=20)
                
                result['OBV'] = obv
                result['OBV_MA20'] = obv_ma20
            
            elif indicator_type == 'VOLUME_PROFILE':
                """Volume Profile 계산"""
                max_volume_price, current_vs_poc = self.calculate_volume_profile(
                    closes.tolist(), 
                    volumes.tolist()
                )
                
                # 단일 값이므로 리스트 대신 스칼라
                result['VP_POC'] = max_volume_price
                result['VP_POSITION'] = current_vs_poc

            else:
                logging.error(f"알 수 없는 지표 유형: {indicator_type}")
                return self._get_default_result(indicator_type, desired_length)

            return result

        except Exception as ex:
            logging.error(f"make_indicator -> {code}, {indicator_type}{self.chart_type} {ex}\n{traceback.format_exc()}")
            return self._get_default_result(indicator_type, len(chart_data.get('C', [])))

# ==================== CpData (체결강도 추가) ====================
class CpData(QObject):
    new_bar_completed = pyqtSignal(str)

    def __init__(self, interval, chart_type, number, trader):
        super().__init__()
        self.interval = interval
        self.number = number
        self.chart_type = chart_type
        self.objCur = {}
        self.stockdata = {}
        self.objIndicators = {}
        self.code = ''
        self.LASTTIME = 1530
        self.trader = trader
        self.is_updating = {}
        self.is_initial_loaded = {}
        self.stockdata_lock = threading.Lock()
        self.last_update_time = {}
        
        self.last_indicator_update = {}
        self.indicator_update_interval = 1.0
        self.latest_snapshot = {}
        
        # 체결강도 계산용
        self.buy_volumes = {}
        self.sell_volumes = {}
        self.strength_cache = {}

        now = time.localtime()
        self.todayDate = now.tm_year * 10000 + now.tm_mon * 100 + now.tm_mday

        self.update_data_timer = QTimer()
        self.update_data_timer.timeout.connect(self.periodic_update_data)
        self.update_data_timer.start(10000)

    def get_strength(self, code):
        """체결강도 반환 (매수세 / 매도세 * 100)"""
        
        # 캐시 확인 (1초)
        if code in self.strength_cache:
            cached_strength, cached_time = self.strength_cache[code]
            if time.time() - cached_time < 1.0:
                return cached_strength
        
        with self.stockdata_lock:
            if code not in self.buy_volumes or len(self.buy_volumes[code]) < 3:
                return 100
            
            total_buy = sum(self.buy_volumes[code])
            total_sell = sum(self.sell_volumes[code])
            
            if total_sell > 0:
                strength = (total_buy / total_sell) * 100
            else:
                strength = 200
            
            # 캐시 저장
            self.strength_cache[code] = (strength, time.time())
            
            return strength

    def periodic_update_data(self):
        """주기적 데이터 업데이트 (수정 버전)"""
        try:
            current_time = time.time()
            with self.stockdata_lock:
                codes = list(self.stockdata.keys())
            
            for code in codes:
                if (code in self.trader.vistock_set and 
                    code not in self.trader.monistock_set and 
                    code not in self.trader.bought_set):
                    continue
                
                interval = 10 if code in self.trader.bought_set else 15

                last_time = self.last_update_time.get(code, 0)
                if current_time - last_time < interval:
                    continue

                with self.stockdata_lock:
                    if code not in self.stockdata:
                        logging.debug(f"{code}: stockdata에서 제거됨, 스킵")
                        continue
                    if self.is_updating.get(code, False):
                        logging.debug(f"{code}: 데이터 업데이트 진행 중, 스킵")
                        continue
                
                self.update_chart_data(code, self.interval, self.number)
                self.last_update_time[code] = current_time

                with self.stockdata_lock:
                    if code not in self.stockdata:
                        logging.debug(f"{code}: 업데이트 후 stockdata에서 제거됨, 스킵")
                        continue
                    if code not in self.objIndicators:
                        self.objIndicators[code] = CpIndicators(self.chart_type)
                        
                        # === 모든 지표 재계산 (새 지표 포함) ===
                        indicator_types = [
                            "MA", "MACD", "RSI", "STOCH", 
                            "ATR", "CCI", "BBANDS", "VWAP",
                            "WILLIAMS_R", "ROC", "OBV", "VOLUME_PROFILE"  # 추가
                        ]
                        
                        results = [
                            self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                            for ind in indicator_types
                        ]
                        
                        if all(results):
                            for result in results:
                                self.stockdata[code].update(result)
                            self._update_snapshot(code)
                        else:
                            logging.warning(f"{code}: 지표 생성 실패")
                            
        except Exception as ex:
            logging.error(f"periodic_update_data -> {ex}")

    def select_code(self, code):
        try:
            if code in self.stockdata:
                return True
            
            self.stockdata[code] = {
                'D': [], 'T': [], 'O': [], 'H': [], 'L': [], 'C': [], 'V': [], 'TV': [], 
                'MAD5': [], 'MAD10': [], 'VWAP': []
            }
            
            self.update_chart_data(code, self.interval, self.number)
            self.is_initial_loaded[code] = False
            
            if code not in self.objIndicators:
                self.objIndicators[code] = CpIndicators(self.chart_type)
                result = self.objIndicators[code].make_indicator("MA", code, self.stockdata[code])

                if result:
                    self.stockdata[code].update(result)
                    self._update_snapshot(code)
                    return True
                else:
                    return False
            return True
        except Exception as ex:
            logging.error(f"select_code -> {ex}")
            return False

    def monitor_code(self, code):
        """종목 모니터링 시작 (수정 버전)"""
        try:
            if code in self.stockdata:
                return True
            
            # === 기본 데이터 구조 초기화 ===
            self.stockdata[code] = {
                'D': [], 'T': [], 'O': [], 'H': [], 'L': [], 'C': [], 'V': [], 'TV': [], 
                
                # 이동평균
                'MAT5': [], 'MAT20': [], 'MAT60': [], 'MAT120': [], 
                'MAM5': [], 'MAM10': [], 'MAM20': [], 
                
                # MACD
                'MACDT': [], 'MACDT_SIGNAL': [], 'OSCT': [], 
                'MACD': [], 'MACD_SIGNAL': [], 'OSC': [], 
                
                # RSI
                'RSIT': [], 'RSIT_SIGNAL': [], 
                'RSI': [], 'RSI_SIGNAL': [], 
                
                # Stochastic
                'STOCHK': [], 'STOCHD': [], 
                
                # 기타
                'ATR': [], 'CCI': [],  
                'BB_UPPER': [], 'BB_MIDDLE': [], 'BB_LOWER': [], 
                'BB_POSITION': [], 'BB_BANDWIDTH': [], 
                'VWAP': [],
                
                # === 새로운 지표들 추가 ===
                'WILLIAMS_R': [],      # Williams %R
                'ROC': [],             # Rate of Change
                'OBV': [],             # On-Balance Volume
                'OBV_MA20': [],        # OBV 이동평균
                'VP_POC': 0,           # Volume Profile POC (단일 값)
                'VP_POSITION': 0,      # VP 현재가 위치 (단일 값)
                
                # 기타
                'TICKS': [], 
                'MAT5_MAT20_DIFF': [], 'MAT20_MAT60_DIFF': [], 
                'MAT60_MAT120_DIFF': [], 'C_MAT5_DIFF': [], 
                'MAM5_MAM10_DIFF': [], 'MAM10_MAM20_DIFF': [], 
                'C_MAM5_DIFF': [], 'C_ABOVE_MAM5': [], 
                'MAT5_CHANGE': [], 'MAT20_CHANGE': [], 
                'MAT60_CHANGE': [], 'MAT120_CHANGE': []
            }
            
            # 체결강도 초기화
            self.buy_volumes[code] = deque(maxlen=10)
            self.sell_volumes[code] = deque(maxlen=10)

            # 데이터 로드
            success = self.update_chart_data_from_market_open(code)
            
            if not success:
                logging.warning(f"{code}: 장 시작 데이터 로드 실패, 개수 기준으로 폴백")
                self.update_chart_data(code, self.interval, self.number)
                self.is_initial_loaded[code] = False
            else:
                self.is_initial_loaded[code] = True

            with self.stockdata_lock:
                if code not in self.objIndicators:
                    self.objIndicators[code] = CpIndicators(self.chart_type)
                    
                    # === 모든 지표 계산 (새 지표 포함) ===
                    indicator_types = [
                        "MA", "MACD", "RSI", "STOCH", 
                        "ATR", "CCI", "BBANDS", "VWAP",
                        "WILLIAMS_R", "ROC", "OBV", "VOLUME_PROFILE"  # 추가
                    ]
                    
                    results = [
                        self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                        for ind in indicator_types
                    ]
                    
                    if all(results):
                        for result in results:
                            self.stockdata[code].update(result)
                        
                        self._update_snapshot(code)
                        
                        if code not in self.objCur:
                            self.objCur[code] = CpPBStockCur()
                            self.objCur[code].Subscribe(code, self)
                        return True
                    else:
                        return False
            
            return True
            
        except Exception as ex:
            logging.error(f"monitor_code -> {code}, {ex}")
            return False

    def monitor_stop(self, code):
        try:
            if self.is_updating.get(code, False):
                logging.debug(f"{code}: 데이터 업데이트 진행 중, 1초 후 재시도")
                QTimer.singleShot(1000, lambda: self.monitor_stop(code))
                return
            with self.stockdata_lock:
                if code in self.objCur:
                    self.objCur[code].Unsubscribe()
                    del self.objCur[code]
                if code in self.stockdata:
                    del self.stockdata[code]
                if code in self.objIndicators:
                    del self.objIndicators[code]
                if code in self.is_updating:
                    del self.is_updating[code]
                if code in self.is_initial_loaded:
                    del self.is_initial_loaded[code]
                if code in self.last_indicator_update:
                    del self.last_indicator_update[code]
                if code in self.latest_snapshot:
                    del self.latest_snapshot[code]
                if code in self.buy_volumes:
                    del self.buy_volumes[code]
                if code in self.sell_volumes:
                    del self.sell_volumes[code]
                if code in self.strength_cache:
                    del self.strength_cache[code]
        except Exception as ex:
            logging.error(f"monitor_stop -> {code}, {ex}")
            return False

    def _request_chart_data(self, code, request_type='count', count=None, start_date=None, end_date=None):
        """공통 차트 데이터 요청 로직"""
        try:
            # API 제한 확인
            api_limiter.wait_if_needed()
            
            objRq = win32com.client.Dispatch("CpSysDib.StockChart")
            objRq.SetInputValue(0, code)
            
            if request_type == 'count':
                objRq.SetInputValue(1, ord('2'))
                objRq.SetInputValue(4, count)
            elif request_type == 'period':
                objRq.SetInputValue(1, ord('1'))
                objRq.SetInputValue(2, end_date)
                objRq.SetInputValue(3, start_date)
            else:
                logging.error(f"잘못된 request_type: {request_type}")
                return None
            
            objRq.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8, 9, 13])
            objRq.SetInputValue(6, ord(self.chart_type))
            objRq.SetInputValue(7, self.interval)
            objRq.SetInputValue(9, ord('1'))
            objRq.BlockRequest2(1)
            
            rqStatus = objRq.GetDibStatus()
            if rqStatus != 0:
                rqRet = objRq.GetDibMsg1()
                logging.warning(f"{code} 데이터 조회 실패, {rqStatus}, {rqRet}")
                return None
            
            len_data = objRq.GetHeaderValue(3)
            lastCount = objRq.GetHeaderValue(4)
            
            new_data = {
                'D': [], 'T': [], 'O': [], 'H': [], 'L': [], 'C': [], 'V': [], 'TV': [], 'TICKS': []
            }
            
            for i in range(len_data):
                new_data['D'].append(objRq.GetDataValue(0, i))
                new_data['T'].append(objRq.GetDataValue(1, i))
                new_data['O'].append(objRq.GetDataValue(2, i))
                new_data['H'].append(objRq.GetDataValue(3, i))
                new_data['L'].append(objRq.GetDataValue(4, i))
                new_data['C'].append(objRq.GetDataValue(5, i))
                new_data['V'].append(objRq.GetDataValue(6, i))
                new_data['TV'].append(objRq.GetDataValue(7, i))
                
                if self.chart_type == 'T':
                    if i == (len_data - 1):
                        new_data['TICKS'].append(lastCount)
                    else:
                        new_data['TICKS'].append(self.interval)
            
            for key in ['D', 'T', 'O', 'H', 'L', 'C', 'V', 'TV']:
                new_data[key].reverse()
            if self.chart_type == 'T':
                new_data['TICKS'].reverse()
            
            return new_data
            
        except Exception as ex:
            logging.error(f"_request_chart_data -> {code}, {ex}")
            return None

    def update_chart_data(self, code, interval, number):
        """실시간 증분 업데이트 (주기적 호출)"""
        try:
            self.is_updating[code] = True
            
            new_data = self._request_chart_data(code, request_type='count', count=number)
            
            if new_data is None:
                self.is_updating[code] = False
                return False
            
            with self.stockdata_lock:
                if code not in self.stockdata:
                    logging.debug(f"{code}: stockdata에 없음, 업데이트 중단")
                    self.is_updating[code] = False
                    return False
                
                if self.is_initial_loaded.get(code, False):
                    if self.stockdata[code].get('T') and len(self.stockdata[code]['T']) > 0:
                        last_time = self.stockdata[code]['T'][-1]
                        last_date = self.stockdata[code]['D'][-1]
                        
                        new_indices = [
                            i for i in range(len(new_data['T']))
                            if (new_data['D'][i] > last_date) or 
                               (new_data['D'][i] == last_date and new_data['T'][i] > last_time)
                        ]
                        
                        if new_indices:
                            for key in new_data:
                                filtered_data = [new_data[key][i] for i in new_indices]
                                self.stockdata[code][key].extend(filtered_data)
                            
                            if self.chart_type == 'T':
                                max_length = 400
                            elif self.chart_type == 'm':
                                max_length = 150
                            else:
                                max_length = 50
                            
                            for key in self.stockdata[code]:
                                if isinstance(self.stockdata[code][key], list):
                                    self.stockdata[code][key] = self.stockdata[code][key][-max_length:]
                    else:
                        for key in new_data:
                            self.stockdata[code][key] = new_data[key]
                else:
                    for key in new_data:
                        self.stockdata[code][key] = new_data[key]
            
            self.is_updating[code] = False
            return True
            
        except Exception as ex:
            logging.error(f"update_chart_data -> {ex}")
            self.is_updating[code] = False
            return False

    def update_chart_data_from_market_open(self, code):
        """장 시작부터 전체 로딩 (초기 1회)"""
        try:
            self.is_updating[code] = True
            
            today = datetime.now().strftime('%Y%m%d')
            
            new_data = self._request_chart_data(
                code, 
                request_type='period',
                start_date=today,
                end_date=today
            )
            
            if new_data is None:
                self.is_updating[code] = False
                return False
            
            with self.stockdata_lock:
                if code not in self.stockdata:
                    logging.debug(f"{code}: stockdata에 없음, 업데이트 중단")
                    self.is_updating[code] = False
                    return False
                
                for key in new_data:
                    self.stockdata[code][key] = new_data[key]
            
            logging.info(f"{code}: 장 시작부터 데이터 로드 완료 ({len(new_data['D'])}개, {self.chart_type})")
            self.is_updating[code] = False
            return True
            
        except Exception as ex:
            logging.error(f"update_chart_data_from_market_open -> {ex}")
            self.is_updating[code] = False
            return False

    def verify_data_coverage(self, code):
        """데이터가 장 시작부터 커버하는지 확인"""
        try:
            with self.stockdata_lock:
                if code not in self.stockdata:
                    return False
                
                stock_data = self.stockdata[code]
                if not stock_data.get('T') or len(stock_data['T']) == 0:
                    return False
                
                first_time = stock_data['T'][0]
                market_open = 900
                
                if first_time > market_open + 30:
                    logging.warning(
                        f"{code}: {self.chart_type} 데이터가 {first_time}부터 시작 "
                        f"(장 시작 데이터 부족, 총 {len(stock_data['T'])}개)"
                    )
                    return False
                
                logging.info(
                    f"{code}: {self.chart_type} 데이터 커버리지 양호 "
                    f"({first_time}부터 시작, 총 {len(stock_data['T'])}개)"
                )
                return True
                
        except Exception as ex:
            logging.error(f"verify_data_coverage -> {ex}")
            return False

    def _update_snapshot(self, code):
        """읽기 전용 스냅샷 업데이트 (수정 버전)"""
        try:
            if code not in self.stockdata:
                return
            
            data = self.stockdata[code]
            
            if self.chart_type == 'T':
                self.latest_snapshot[code] = {
                    # 기본 가격
                    'C': data.get('C', [0])[-1] if data.get('C') else 0,
                    'O': data.get('O', [0])[-1] if data.get('O') else 0,
                    'H': data.get('H', [0])[-1] if data.get('H') else 0,
                    'L': data.get('L', [0])[-1] if data.get('L') else 0,
                    'V': data.get('V', [0])[-1] if data.get('V') else 0,
                    
                    # 이동평균
                    'MAT5': data.get('MAT5', [0])[-1] if data.get('MAT5') else 0,
                    'MAT20': data.get('MAT20', [0])[-1] if data.get('MAT20') else 0,
                    'MAT60': data.get('MAT60', [0])[-1] if data.get('MAT60') else 0,
                    'MAT120': data.get('MAT120', [0])[-1] if data.get('MAT120') else 0,
                    
                    # RSI
                    'RSIT': data.get('RSIT', [0])[-1] if data.get('RSIT') else 0,
                    'RSIT_SIGNAL': data.get('RSIT_SIGNAL', [0])[-1] if data.get('RSIT_SIGNAL') else 0,
                    
                    # MACD
                    'MACDT': data.get('MACDT', [0])[-1] if data.get('MACDT') else 0,
                    'MACDT_SIGNAL': data.get('MACDT_SIGNAL', [0])[-1] if data.get('MACDT_SIGNAL') else 0,
                    'OSCT': data.get('OSCT', [0])[-1] if data.get('OSCT') else 0,
                    
                    # Stochastic
                    'STOCHK': data.get('STOCHK', [0])[-1] if data.get('STOCHK') else 0,
                    'STOCHD': data.get('STOCHD', [0])[-1] if data.get('STOCHD') else 0,
                    
                    # 기타
                    'ATR': data.get('ATR', [0])[-1] if data.get('ATR') else 0,
                    'CCI': data.get('CCI', [0])[-1] if data.get('CCI') else 0,
                    'BB_UPPER': data.get('BB_UPPER', [0])[-1] if data.get('BB_UPPER') else 0,
                    'BB_MIDDLE': data.get('BB_MIDDLE', [0])[-1] if data.get('BB_MIDDLE') else 0,
                    'BB_LOWER': data.get('BB_LOWER', [0])[-1] if data.get('BB_LOWER') else 0,
                    'BB_POSITION': data.get('BB_POSITION', [0])[-1] if data.get('BB_POSITION') else 0,
                    'BB_BANDWIDTH': data.get('BB_BANDWIDTH', [0])[-1] if data.get('BB_BANDWIDTH') else 0,
                    'VWAP': data.get('VWAP', [0])[-1] if data.get('VWAP') else 0,
                    
                    # === 새로운 지표들 ===
                    'WILLIAMS_R': data.get('WILLIAMS_R', [0])[-1] if data.get('WILLIAMS_R') else -50,
                    'ROC': data.get('ROC', [0])[-1] if data.get('ROC') else 0,
                    'OBV': data.get('OBV', [0])[-1] if data.get('OBV') else 0,
                    'OBV_MA20': data.get('OBV_MA20', [0])[-1] if data.get('OBV_MA20') else 0,
                    'VP_POC': data.get('VP_POC', 0),
                    'VP_POSITION': data.get('VP_POSITION', 0),
                    
                    # 최근 추이
                    'C_recent': data.get('C', [0])[-3:] if data.get('C') else [0, 0, 0],
                    'H_recent': data.get('H', [0])[-3:] if data.get('H') else [0, 0, 0],
                    'L_recent': data.get('L', [0])[-3:] if data.get('L') else [0, 0, 0],
                }
            
            elif self.chart_type == 'm':
                self.latest_snapshot[code] = {
                    # 기본 가격
                    'C': data.get('C', [0])[-1] if data.get('C') else 0,
                    'O': data.get('O', [0])[-1] if data.get('O') else 0,
                    'H': data.get('H', [0])[-1] if data.get('H') else 0,
                    'L': data.get('L', [0])[-1] if data.get('L') else 0,
                    'V': data.get('V', [0])[-1] if data.get('V') else 0,
                    
                    # 이동평균
                    'MAM5': data.get('MAM5', [0])[-1] if data.get('MAM5') else 0,
                    'MAM10': data.get('MAM10', [0])[-1] if data.get('MAM10') else 0,
                    'MAM20': data.get('MAM20', [0])[-1] if data.get('MAM20') else 0,
                    
                    # RSI
                    'RSI': data.get('RSI', [0])[-1] if data.get('RSI') else 0,
                    'RSI_SIGNAL': data.get('RSI_SIGNAL', [0])[-1] if data.get('RSI_SIGNAL') else 0,
                    
                    # MACD
                    'MACD': data.get('MACD', [0])[-1] if data.get('MACD') else 0,
                    'MACD_SIGNAL': data.get('MACD_SIGNAL', [0])[-1] if data.get('MACD_SIGNAL') else 0,
                    'OSC': data.get('OSC', [0])[-1] if data.get('OSC') else 0,
                    
                    # Stochastic
                    'STOCHK': data.get('STOCHK', [0])[-1] if data.get('STOCHK') else 0,
                    'STOCHD': data.get('STOCHD', [0])[-1] if data.get('STOCHD') else 0,
                    
                    # 기타
                    'CCI': data.get('CCI', [0])[-1] if data.get('CCI') else 0,
                    'VWAP': data.get('VWAP', [0])[-1] if data.get('VWAP') else 0,
                    
                    # === 새로운 지표들 ===
                    'WILLIAMS_R': data.get('WILLIAMS_R', [0])[-1] if data.get('WILLIAMS_R') else -50,
                    'ROC': data.get('ROC', [0])[-1] if data.get('ROC') else 0,
                    'OBV': data.get('OBV', [0])[-1] if data.get('OBV') else 0,
                    'OBV_MA20': data.get('OBV_MA20', [0])[-1] if data.get('OBV_MA20') else 0,
                    'VP_POC': data.get('VP_POC', 0),
                    'VP_POSITION': data.get('VP_POSITION', 0),
                    
                    # 최근 추이
                    'C_recent': data.get('C', [0])[-2:] if data.get('C') else [0, 0],
                    'O_recent': data.get('O', [0])[-2:] if data.get('O') else [0, 0],
                    'H_recent': data.get('H', [0])[-2:] if data.get('H') else [0, 0],
                    'L_recent': data.get('L', [0])[-2:] if data.get('L') else [0, 0],
                }
            
            elif self.chart_type == 'D':
                self.latest_snapshot[code] = {
                    'C': data.get('C', [0])[-1] if data.get('C') else 0,
                    'V': data.get('V', [0])[-1] if data.get('V') else 0,
                    'MAD5': data.get('MAD5', [0])[-1] if data.get('MAD5') else 0,
                    'MAD10': data.get('MAD10', [0])[-1] if data.get('MAD10') else 0,
                    'VWAP': data.get('VWAP', [0])[-1] if data.get('VWAP') else 0,
                }
                
        except Exception as ex:
            logging.error(f"_update_snapshot -> {code}, {ex}")

    def get_latest_data(self, code):
        """빠른 읽기 - 스냅샷 반환 (락 불필요)"""
        return self.latest_snapshot.get(code, {})
    
    def get_full_data(self, code):
        """전체 데이터 읽기 (락 필요, 차트 그리기용)"""
        with self.stockdata_lock:
            return copy.deepcopy(self.stockdata.get(code, {}))
    
    def get_recent_data(self, code, count=10):
        """최근 N개 데이터 읽기 (락 필요)"""
        with self.stockdata_lock:
            if code not in self.stockdata:
                return {}
            
            data = self.stockdata[code]
            result = {}
            for key, values in data.items():
                if isinstance(values, list) and len(values) > 0:
                    result[key] = values[-count:] if len(values) >= count else values
                else:
                    result[key] = values
            return result

    def updateCurData(self, item):
        """실시간 체결 데이터 업데이트"""
        try:
            if self.is_updating.get(item['code'], False):
                return
            
            code = item['code']
            time_val = item['time']
            cur = item['cur']
            vol = item['vol']
            current_time = time.time()
            
            # 체결강도 업데이트
            with self.stockdata_lock:
                if code in self.buy_volumes:
                    if len(self.stockdata.get(code, {}).get('C', [])) > 0:
                        prev_price = self.stockdata[code]['C'][-1]
                        if cur > prev_price:
                            self.buy_volumes[code].append(vol)
                            self.sell_volumes[code].append(0)
                        elif cur < prev_price:
                            self.buy_volumes[code].append(0)
                            self.sell_volumes[code].append(vol)
                        else:
                            self.buy_volumes[code].append(vol / 2)
                            self.sell_volumes[code].append(vol / 2)

            with self.stockdata_lock:
                bar_completed = False
                
                if self.chart_type == 'T':
                    hh, mm = divmod(time_val, 10000)
                    mm, tt = divmod(mm, 100)
                    if mm == 60:
                        hh += 1
                        mm = 0
                    lCurTime = hh * 100 + mm

                    bFind = False
                    if lCurTime > self.LASTTIME:
                        lCurTime = self.LASTTIME

                    if code in self.stockdata:
                        if len(self.stockdata[code]['T']) > 0:
                            lastCount = self.stockdata[code]['TICKS'][-1]
                            if 1 <= lastCount < self.interval:
                                bFind = True
                                self.stockdata[code]['T'][-1] = lCurTime
                                self.stockdata[code]['C'][-1] = cur
                                if self.stockdata[code]['H'][-1] < cur:
                                    self.stockdata[code]['H'][-1] = cur
                                if self.stockdata[code]['L'][-1] > cur:
                                    self.stockdata[code]['L'][-1] = cur
                                self.stockdata[code]['V'][-1] += vol
                                self.stockdata[code]['TICKS'][-1] += 1

                        if not bFind:
                            bar_completed = True
                            
                            self.stockdata[code]['D'].append(self.todayDate)
                            self.stockdata[code]['T'].append(lCurTime)
                            self.stockdata[code]['O'].append(cur)
                            self.stockdata[code]['H'].append(cur)
                            self.stockdata[code]['L'].append(cur)
                            self.stockdata[code]['C'].append(cur)
                            self.stockdata[code]['V'].append(vol)
                            self.stockdata[code]['TICKS'].append(1)

                        desired_length = 600
                        for key in self.stockdata[code]:
                            self.stockdata[code][key] = self.stockdata[code][key][-desired_length:]

                        self._update_snapshot(code)

                        last_update = self.last_indicator_update.get(code, 0)
                        if current_time - last_update >= self.indicator_update_interval:
                            if code in self.objIndicators:
                                results = [
                                    self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                                    for ind in ["MA", "RSI", "MACD", "STOCH", "ATR", "CCI", "BBANDS", "VWAP",
                                               "WILLIAMS_R", "ROC", "OBV", "VOLUME_PROFILE"]
                                ]
                                if all(results):
                                    for result in results:
                                        self.stockdata[code].update(result)
                                    self._update_snapshot(code)
                            
                            self.last_indicator_update[code] = current_time
                    
                elif self.chart_type == 'm':
                    hh, mm = divmod(time_val, 10000)
                    mm, tt = divmod(mm, 100)
                    convertedMintime = hh * 60 + mm

                    bFind = False
                    a, b = divmod(convertedMintime, self.interval)
                    intervaltime = a * self.interval
                    lChartTime = intervaltime + self.interval
                    hour, minute = divmod(lChartTime, 60)
                    lCurTime = hour * 100 + minute

                    if lCurTime > self.LASTTIME:
                        lCurTime = self.LASTTIME

                    if code in self.stockdata:
                        if len(self.stockdata[code]['T']) > 0:
                            lLastTime = self.stockdata[code]['T'][-1]
                            if lLastTime == lCurTime:
                                bFind = True
                                self.stockdata[code]['C'][-1] = cur
                                if self.stockdata[code]['H'][-1] < cur:
                                    self.stockdata[code]['H'][-1] = cur
                                if self.stockdata[code]['L'][-1] > cur:
                                    self.stockdata[code]['L'][-1] = cur
                                self.stockdata[code]['V'][-1] += vol
        
                        if not bFind:
                            # ✅ 새 봉 생성 = 완성 이벤트
                            bar_completed = True
                            
                            self.stockdata[code]['D'].append(self.todayDate)
                            self.stockdata[code]['T'].append(lCurTime)
                            self.stockdata[code]['O'].append(cur)
                            self.stockdata[code]['H'].append(cur)
                            self.stockdata[code]['L'].append(cur)
                            self.stockdata[code]['C'].append(cur)
                            self.stockdata[code]['V'].append(vol)

                        desired_length = 150
                        for key in self.stockdata[code]:
                            self.stockdata[code][key] = self.stockdata[code][key][-desired_length:]

                        self._update_snapshot(code)

                        last_update = self.last_indicator_update.get(code, 0)
                        if current_time - last_update >= self.indicator_update_interval:
                            if code in self.objIndicators:
                                results = [
                                    self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                                    for ind in ["MA", "MACD", "RSI", "STOCH", "ATR", "CCI", "BBANDS", "VWAP"]
                                ]
                                if all(results):
                                    for result in results:
                                        self.stockdata[code].update(result)
                                    self._update_snapshot(code)
                            
                            self.last_indicator_update[code] = current_time
                
                # ✅ 새 봉 완성 시 signal 발생
                if bar_completed:
                    self.new_bar_completed.emit(code)
        
        except Exception as ex:
            logging.error(f"updateCurData -> {ex}")

# ==================== CTrader (계속) ====================
class CTrader(QObject):
    """트레이더 클래스 (DatabaseWorker 제거, combined_tick_data 단일 사용)"""
    
    stock_added_to_monitor = pyqtSignal(str)
    stock_bought = pyqtSignal(str)
    stock_sold = pyqtSignal(str)

    def __init__(self, obj_cp_trade, cp_balance, cpCodeMgr, cp_cash, cp_order, cp_stock, buycount, window):
        super().__init__()
        self.cpTrade = obj_cp_trade
        self.cpBalance = cp_balance
        self.cpCodeMgr = cpCodeMgr
        self.cpCash = cp_cash
        self.cpOrder = cp_order
        self.cpStock = cp_stock
        self.target_buy_count = buycount

        self.bought_set = set()
        self.database_set = set()
        self.monistock_set = set()
        self.vistock_set = set()
        self.sell_half_set = set()
        self.buyorder_set = set()
        self.buyordering_set = set()
        self.sellorder_set = set()
        
        self.starting_price = {}
        self.starting_time = {}
        self.highest_price = {}
        self.buy_price = {}
        self.buy_qty = {}
        self.buyorder_qty = {}
        self.sell_half_qty = {}
        self.balance_qty = {}
        self.sell_amount = {}
        self.buy_amount = 0
        self.conclusion = CpPBConclusion()
        self.conclusion.Subscribe('', self)
        self.window = window

        self.cp_request = CpRequest()

        self.daydata = CpData(1, 'D', 50, self)
        self.mindata = CpData(3, 'm', 150, self)
        self.tickdata = CpData(30, 'T', 600, self)

        self.db_name = 'vi_stock_data.db'

        # ===== 설정 파일 읽기 (간소화) =====
        config = configparser.ConfigParser(interpolation=None)
        if os.path.exists('settings.ini'):
            config.read('settings.ini', encoding='utf-8')
        
        # combined_tick_data 저장 주기만 사용
        self.combined_save_interval = config.getint('DATA_SAVING', 'interval_seconds', fallback=5)
        
        logging.info(f"데이터 저장 설정: combined_tick_data 간격={self.combined_save_interval}초")

    def init_database(self):
        """데이터베이스 초기화 (combined_tick_data 단일 테이블)"""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            # ===== combined_tick_data (백테스팅용 메인 테이블) =====
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS combined_tick_data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT NOT NULL,
                    timestamp DATETIME NOT NULL,
                    date TEXT NOT NULL,
                    time TEXT NOT NULL,
                    
                    -- 틱 데이터 (현재 시점)
                    tick_C REAL, tick_O REAL, tick_H REAL, tick_L REAL, tick_V INTEGER,
                    tick_MAT5 REAL, tick_MAT20 REAL, tick_MAT60 REAL, tick_MAT120 REAL,
                    tick_RSIT REAL, tick_RSIT_SIGNAL REAL,
                    tick_MACDT REAL, tick_MACDT_SIGNAL REAL, tick_OSCT REAL,
                    tick_STOCHK REAL, tick_STOCHD REAL,
                    tick_ATR REAL, tick_CCI REAL,
                    tick_BB_UPPER REAL, tick_BB_MIDDLE REAL, tick_BB_LOWER REAL,
                    tick_BB_POSITION REAL, tick_BB_BANDWIDTH REAL,
                    tick_VWAP REAL,
                    
                    -- === 새 지표: 틱 ===
                    tick_WILLIAMS_R REAL,
                    tick_ROC REAL,
                    tick_OBV REAL,
                    tick_OBV_MA20 REAL,
                    tick_VP_POC REAL,
                    tick_VP_POSITION REAL,
                    
                    -- 분봉 데이터 (가장 최근 완성된 분봉)
                    min_C REAL, min_O REAL, min_H REAL, min_L REAL, min_V INTEGER,
                    min_MAM5 REAL, min_MAM10 REAL, min_MAM20 REAL,
                    min_RSI REAL, min_RSI_SIGNAL REAL,
                    min_MACD REAL, min_MACD_SIGNAL REAL, min_OSC REAL,
                    min_STOCHK REAL, min_STOCHD REAL,
                    min_CCI REAL, min_VWAP REAL,
                    
                    -- === 새 지표: 분봉 ===
                    min_WILLIAMS_R REAL,
                    min_ROC REAL,
                    min_OBV REAL,
                    min_OBV_MA20 REAL,
                    
                    -- 추가 정보
                    strength REAL,
                    buy_price REAL,
                    position_type TEXT,
                    save_reason TEXT,
                    
                    UNIQUE(code, timestamp)
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_combined_code ON combined_tick_data(code)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_combined_date ON combined_tick_data(date)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_combined_timestamp ON combined_tick_data(timestamp)')
            
            # ===== trades 테이블 (실거래 기록) =====
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS trades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT,
                    stock_name TEXT,
                    date TEXT,
                    time TEXT,
                    action TEXT,
                    price REAL,
                    quantity INTEGER,
                    amount REAL,
                    strategy TEXT,
                    buy_reason TEXT,
                    sell_reason TEXT,
                    buy_price REAL,
                    profit REAL,
                    profit_pct REAL,
                    hold_minutes REAL
                )
            ''')
            
            # ===== daily_summary 테이블 (일별 요약) =====
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS daily_summary (
                    date TEXT PRIMARY KEY,
                    strategy TEXT,
                    total_trades INTEGER DEFAULT 0,
                    win_trades INTEGER DEFAULT 0,
                    lose_trades INTEGER DEFAULT 0,
                    win_rate REAL DEFAULT 0,
                    total_profit REAL DEFAULT 0,
                    avg_profit_pct REAL DEFAULT 0,
                    max_profit_pct REAL DEFAULT 0,
                    max_loss_pct REAL DEFAULT 0,
                    total_buy_amount REAL DEFAULT 0,
                    final_cash REAL DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # ===== backtest_results 테이블 (백테스팅 결과) =====
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS backtest_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    strategy TEXT NOT NULL,
                    start_date TEXT NOT NULL,
                    end_date TEXT NOT NULL,
                    initial_cash REAL NOT NULL,
                    final_cash REAL NOT NULL,
                    total_profit REAL NOT NULL,
                    total_return_pct REAL NOT NULL,
                    total_trades INTEGER NOT NULL,
                    win_trades INTEGER NOT NULL,
                    lose_trades INTEGER NOT NULL,
                    win_rate REAL NOT NULL,
                    avg_profit_pct REAL NOT NULL,
                    max_profit_pct REAL NOT NULL,
                    max_loss_pct REAL NOT NULL,
                    mdd REAL NOT NULL,
                    sharpe_ratio REAL,
                    avg_hold_minutes REAL,
                    parameters TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_backtest_strategy ON backtest_results(strategy)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_backtest_date ON backtest_results(start_date, end_date)')
            
            conn.commit()
            conn.close()
            
            logging.info("데이터베이스 초기화 완료 (새 지표 포함)")
            
        except Exception as ex:
            logging.error(f"init_database -> {ex}\n{traceback.format_exc()}")
            raise

    def save_trade_record(self, code, action, price, quantity, **kwargs):
        """실거래 기록 저장"""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            now = datetime.now()
            date = now.strftime('%Y%m%d')
            time_str = now.strftime('%H%M%S')
            
            stock_name = cpCodeMgr.CodeToName(code)
            amount = price * quantity
            
            # 전략 이름
            strategy = kwargs.get('strategy', '')
            if not strategy and hasattr(self, 'window'):
                strategy = self.window.comboStg.currentText()
            
            cursor.execute('''
                INSERT INTO trades (
                    code, stock_name, date, time, action, price, quantity, amount,
                    strategy, buy_reason, sell_reason, buy_price, profit, profit_pct, hold_minutes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                code, 
                stock_name, 
                date, 
                time_str, 
                action, 
                price, 
                quantity, 
                amount,
                strategy,
                kwargs.get('buy_reason', ''),
                kwargs.get('sell_reason', ''),
                kwargs.get('buy_price', None),
                kwargs.get('profit', None),
                kwargs.get('profit_pct', None),
                kwargs.get('hold_minutes', None)
            ))
            
            conn.commit()
            conn.close()
            
            logging.debug(f"거래 기록 저장: {stock_name}({code}) {action} {quantity}주 @{price:,}원")
            
        except Exception as ex:
            logging.error(f"save_trade_record -> {ex}\n{traceback.format_exc()}")

    def update_daily_summary(self):
        """일별 요약 업데이트"""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            today = datetime.now().strftime('%Y%m%d')
            strategy = self.window.comboStg.currentText() if hasattr(self, 'window') else ''
            
            # 오늘의 매도 거래만 집계
            cursor.execute('''
                SELECT 
                    COUNT(*) as total_trades,
                    SUM(CASE WHEN profit > 0 THEN 1 ELSE 0 END) as win_trades,
                    SUM(CASE WHEN profit <= 0 THEN 1 ELSE 0 END) as lose_trades,
                    AVG(profit_pct) as avg_profit_pct,
                    MAX(profit_pct) as max_profit_pct,
                    MIN(profit_pct) as max_loss_pct,
                    SUM(profit) as total_profit
                FROM trades
                WHERE date = ? AND action = 'SELL'
            ''', (today,))
            
            row = cursor.fetchone()
            
            if row and row[0] > 0:
                total_trades, win_trades, lose_trades, avg_profit_pct, max_profit_pct, max_loss_pct, total_profit = row
                win_rate = (win_trades / total_trades * 100) if total_trades > 0 else 0
                
                # 매수 금액 합계
                cursor.execute('''
                    SELECT SUM(amount) FROM trades
                    WHERE date = ? AND action = 'BUY'
                ''', (today,))
                
                total_buy_amount = cursor.fetchone()[0] or 0
                
                cursor.execute('''
                    INSERT OR REPLACE INTO daily_summary (
                        date, strategy, total_trades, win_trades, lose_trades, win_rate,
                        total_profit, avg_profit_pct, max_profit_pct, max_loss_pct, total_buy_amount
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    today, strategy, total_trades, win_trades, lose_trades, win_rate,
                    total_profit or 0, avg_profit_pct or 0, max_profit_pct or 0, max_loss_pct or 0,
                    total_buy_amount
                ))
                
                conn.commit()
                
                logging.debug(f"일별 요약 업데이트: {today} - 거래 {total_trades}회, 승률 {win_rate:.1f}%")
            
            conn.close()
            
        except Exception as ex:
            logging.error(f"update_daily_summary -> {ex}\n{traceback.format_exc()}")

    def get_stock_balance(self, code, func):
        """계좌 잔고 조회"""
        try:
            stock_name = self.cpCodeMgr.CodeToName(code)
            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)
            self.cpBalance.SetInputValue(0, acc)
            self.cpBalance.SetInputValue(1, accFlag[0])
            self.cpBalance.SetInputValue(2, 50)
            ret = self.cpBalance.BlockRequest2(1)
            if ret != 0:
                logging.warning(f"{stock_name}({code}) 잔고 조회 실패, {ret}")
                return False
            
            if code == 'START':
                logging.info(f"계좌명 : {str(self.cpBalance.GetHeaderValue(0))}")
                logging.info(f"결제잔고수량 : {str(self.cpBalance.GetHeaderValue(1))}")
                logging.info(f"평가금액 : {str(self.cpBalance.GetHeaderValue(3))}")
                logging.info(f"평가손익 : {str(self.cpBalance.GetHeaderValue(4))}")
                logging.info(f"종목수 : {str(self.cpBalance.GetHeaderValue(7))}")
                return

            stocks = []
            for i in range(self.cpBalance.GetHeaderValue(7)):
                stock_code = self.cpBalance.GetDataValue(12, i)
                stock_name = self.cpBalance.GetDataValue(0, i)
                stock_qty = self.cpBalance.GetDataValue(15, i)
                buy_price = self.cpBalance.GetDataValue(17, i)
                stocks.append({'code': stock_code, 'name': stock_name, 'qty': stock_qty, 'buy_price': buy_price})

            if code == 'ALL':
                logging.debug("잔고 전부 조회 성공")
                return stocks

            for s in stocks:
                if s['code'] == code:
                    logging.debug(f"{s['name']}({s['code']}) 잔고 조회 성공")
                    return s['name'], s['qty'], s['buy_price']             

        except Exception as ex:
            logging.error(f"get_stock_balance({func}) -> {ex}")
            return False

    def init_stock_balance(self):
        """시작 시 계좌 잔고 초기화"""
        try:
            stocks = self.get_stock_balance('ALL', 'init_stock_balance')

            for s in stocks:
                if self.daydata.select_code(s['code']) and self.tickdata.monitor_code(s['code']) and self.mindata.monitor_code(s['code']):
                    if s['code'] not in self.starting_time:
                        self.starting_time[s['code']] = datetime.now().strftime('%m/%d 09:00:00')
                    self.monistock_set.add(s['code'])
                    self.stock_added_to_monitor.emit(s['code'])
                    self.bought_set.add(s['code'])
                    self.stock_bought.emit(s['code'])
                    self.buy_price[s['code']] = s['buy_price']
                    self.buy_qty[s['code']] = s['qty']
                    self.balance_qty[s['code']] = s['qty']

            remaining_count = self.target_buy_count - len(stocks)
            self.buy_percent = 1/remaining_count if remaining_count > 0 else 0
            self.total_cash = self.get_current_cash() * 0.9
            self.buy_amount = int(self.total_cash * self.buy_percent)
            
            logging.info(f"주문 가능 금액 : {self.total_cash}")
            logging.info(f"종목별 주문 비율 : {self.buy_percent}")
            logging.info(f"종목별 주문 금액 : {self.buy_amount}")

        except Exception as ex:
            logging.error(f"init_stock_balance -> {ex}")

    def get_current_cash(self):
        """현재 주문 가능 금액 조회"""
        try:
            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)
            self.cpCash.SetInputValue(0, acc)
            self.cpCash.SetInputValue(1, accFlag[0])
            self.cpCash.SetInputValue(5, "Y")
            self.cpCash.BlockRequest2(1)

            rqStatus = self.cpCash.GetDibStatus()
            if rqStatus != 0:
                rqRet = self.cpCash.GetDibMsg1()
                logging.warning(f"주문가능금액 조회실패, {rqStatus}, {rqRet}")
                return (False, '')
            current_cash = self.cpCash.GetHeaderValue(9)
            return current_cash
        except Exception as ex:
            logging.error(f"get_current_cash -> {ex}")
            return False

    def get_current_price(self, code):
        """현재가 조회"""
        try:
            self.cpStock.SetInputValue(0, code)
            ret = self.cpStock.BlockRequest2(1)
            if ret != 0:
                logging.warning(f"{code} 현재가 조회 실패, {ret}")
                return False

            rqStatus = self.cpStock.GetDibStatus()
            if rqStatus != 0:
                rqRet = self.cpStock.GetDibMsg1()
                logging.warning(f"현재가 조회 오류, {rqStatus}, {rqRet}")
                return False
            
            item = {'cur_price': self.cpStock.GetHeaderValue(11), 'ask': self.cpStock.GetHeaderValue(16), 'upper': self.cpStock.GetHeaderValue(8)}
            return item['cur_price'], item['ask'], item['upper']
        
        except Exception as ex:
            logging.error(f"get_current_price -> {ex}")
            return False
    
    def get_trade_profit(self):
        """매매실현손익 조회"""
        try:
            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)

            self.objRq = win32com.client.Dispatch("CpTrade.CpTd6032")
            self.objRq.SetInputValue(0, acc)
            self.objRq.SetInputValue(1, accFlag[0])
            ret = self.objRq.BlockRequest2(1)
            if ret != 0:
                logging.warning(f"매매실현손익 조회 실패, {ret}")
                return
            
            rqStatus = self.objRq.GetDibStatus()
            if rqStatus != 0:
                rqRet = self.objRq.GetDibMsg1()
                logging.warning(f"매매실현손익 조회 오류, {rqStatus}, {rqRet}")
                return False
            
            profit = self.objRq.GetHeaderValue(2)
            logging.info(f"매매실현손익 : {self.objRq.GetHeaderValue(2)}원")
            send_slack_message(self.window.login_handler, "#stock", f"매매실현손익 : {profit}원")

        except Exception as ex:
            logging.error(f"get_trade_profit -> {ex}")
    
    def update_highest_price(self, code, current_price):
        """최고가 업데이트"""
        if code not in self.highest_price:
            self.highest_price[code] = current_price
        elif current_price > self.highest_price[code]:
            self.highest_price[code] = current_price    

    def monitor_vi(self, time, code, event2):
        """VI 발동 모니터링"""
        try:
            if code in self.monistock_set or len(self.monistock_set) >= 10 or code in self.bought_set:
                return

            if not self.daydata.select_code(code):
                logging.debug(f"{code}: 일봉 데이터 로드 실패")
                return
            
            day_data = self.daydata.stockdata[code]
            if not (day_data['MAD5'][-1] > day_data['MAD10'][-1]):
                logging.debug(f"{code}: MAD5 < MAD10, 추세 미달")
                self.daydata.monitor_stop(code)
                return
            
            recent_ma5 = day_data['MAD5'][-3:]
            if len(recent_ma5) >= 3 and not all(recent_ma5[i] < recent_ma5[i+1] for i in range(2)):
                logging.debug(f"{code}: MAD5 추세 약화")
                self.daydata.monitor_stop(code)
                return
            
            if len(day_data['V']) < 30:
                logging.debug(f"{code}: 데이터 부족 ({len(day_data['V'])}일)")
                self.daydata.monitor_stop(code)
                return
            
            Trading_amount = sum(day_data['V'][-30:]) / 30
            Trading_Value = sum(day_data['TV'][-30:]) / 30
            
            MIN_VOLUME = 100000
            MIN_VALUE = 3000000000
            
            if Trading_amount < MIN_VOLUME or Trading_Value < MIN_VALUE:
                logging.debug(
                    f"{code}: 유동성 부족 - "
                    f"거래량 {Trading_amount:.0f}/{MIN_VOLUME}, "
                    f"거래금액 {Trading_Value/100000000:.1f}/{MIN_VALUE/100000000}억"
                )
                self.daydata.monitor_stop(code)
                return
            
            now = datetime.now()
            elapsed_minutes = (now - now.replace(hour=9, minute=0, second=0)).total_seconds() / 60
            
            if elapsed_minutes <= 120:
                volume_threshold = 0.5
            elif elapsed_minutes <= 240:
                volume_threshold = 0.7
            else:
                volume_threshold = 1.0
            
            current_volume = day_data['V'][-1]
            prev_volume = day_data['V'][-2]
            
            if current_volume < prev_volume * volume_threshold:
                logging.debug(
                    f"{code}: 거래량 부족 - "
                    f"{current_volume}/{prev_volume * volume_threshold:.0f} ({volume_threshold*100}%)"
                )
                self.daydata.monitor_stop(code)
                return
            
            current_price = day_data['C'][-1]
            if current_price < 1000 or current_price > 500000:
                logging.debug(f"{code}: 가격 {current_price}원 범위 초과")
                self.daydata.monitor_stop(code)
                return
            
            match_gap = re.search(r"괴리율:(-?\d+\.\d+)%", event2)
            if match_gap:
                gap_rate = float(match_gap.group(1))
                if gap_rate < 3.0:
                    logging.debug(f"{code}: 괴리율 {gap_rate}% (3% 미만)")
                    self.daydata.monitor_stop(code)
                    return
            
            sector = cpCodeMgr.GetStockSectionKind(code)
            sector_count = sum(1 for c in self.monistock_set 
                            if cpCodeMgr.GetStockSectionKind(c) == sector)
            if sector_count >= 2:
                logging.debug(f"{code}: 동일 섹터 종목 {sector_count}개 초과")
                self.daydata.monitor_stop(code)
                return
            
            if not (self.tickdata.monitor_code(code) and self.mindata.monitor_code(code)):
                logging.error(f"{code}: 틱/분 데이터 모니터링 시작 실패")
                self.daydata.monitor_stop(code)
                return
            
            logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> 장 시작부터 데이터 로드 시작")
            
            tick_loaded = self.tickdata.update_chart_data_from_market_open(code)
            min_loaded = self.mindata.update_chart_data_from_market_open(code)
            
            if not tick_loaded or not min_loaded:
                logging.warning(f"{code}: 장 시작 데이터 로드 실패, 개수 기준으로 폴백")
                self.tickdata.update_chart_data(code, 60, 400)
                self.mindata.update_chart_data(code, 3, 150)
            
            self.starting_time[code] = time
            
            match_price = re.search(r"발동가격:\s*(\d+)", event2)
            if match_price:
                self.starting_price[code] = int(match_price.group(1))
            else:
                self.starting_price[code] = current_price
            
            self.monistock_set.add(code)
            self.stock_added_to_monitor.emit(code)
            
            logging.info(
                f"{cpCodeMgr.CodeToName(code)}({code}) -> "
                f"투자 대상 추가 (VI: {time}, 발동가: {self.starting_price[code]:.0f}원, "
                f"괴리율: {gap_rate if match_gap else 'N/A'}%)"
            )
            
            self.save_list_db(code, self.starting_time[code], self.starting_price[code], 1)

        except Exception as ex:
            logging.error(f"monitor_vi -> {code}, {ex}\n{traceback.format_exc()}")

    def save_list_db(self, code, starting_time, starting_price, is_moni=0, db_file='mylist.db'):
        """종목 리스트 DB 저장"""
        conn = sqlite3.connect(db_file)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS items (
                codes TEXT PRIMARY KEY,
                starting_time DATETIME,
                starting_price REAL,
                is_moni INTEGER
            )
        ''')
        c.execute('INSERT OR REPLACE INTO items (codes, starting_time, starting_price, is_moni) VALUES (?, ?, ?, ?)', 
                  (code, starting_time, starting_price, is_moni))
        conn.commit()
        conn.close()

    def delete_list_db(self, code, db_file='mylist.db'):
        """종목 리스트 DB 삭제"""
        conn = sqlite3.connect(db_file)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS items (
                codes TEXT PRIMARY KEY,
                starting_time DATETIME,
                starting_price REAL,
                is_moni INTEGER
            )
        ''')
        c.execute("DELETE FROM items WHERE codes = ?", (code,))
        conn.commit()
        conn.close()

    def clear_list_db(self, db_file='mylist.db'):
        """종목 리스트 DB 전체 삭제"""
        try:
            if os.path.exists(db_file):
                os.remove(db_file)
        except Exception as ex:
            logging.error(f"clear_list_db -> {ex}")
        self.monistock_set = set()
        self.vistock_set = set()
        self.database_set = set()
        self.starting_time = {}
        self.starting_price = {}

    def load_from_list_db(self, db_file='mylist.db'):
        """종목 리스트 DB 로드"""
        if not os.path.exists(db_file):
            self.monistock_set = set()
            self.vistock_set = set()
            self.database_set = set()
            self.stgname = {}
            self.starting_time = {}
            self.starting_price = {}
            return
        
        conn = sqlite3.connect(db_file)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS items (
                codes TEXT PRIMARY KEY,
                starting_time DATETIME,
                starting_price REAL,
                is_moni INTEGER
            )
        ''')
        c.execute('SELECT codes, starting_time, starting_price, is_moni FROM items')
        rows = c.fetchall()
        conn.close()
        if rows:
            self.vistock_set = {row[0] for row in rows if not row[3]}
            self.database_set = {row[0] for row in rows if row[3]}
            self.starting_time = {row[0]: row[1] for row in rows}
            self.starting_price = {row[0]: row[2] for row in rows}

    @pyqtSlot(str, str, str, str)
    def buy_stock(self, code, buy_message, order_condition, order_style):
        """매수 주문"""
        try:
            if code in self.bought_set or code in self.buyorder_set or code in self.buyordering_set:
                return
            self.buyordering_set.add(code)
            stock_name = self.cpCodeMgr.CodeToName(code)
            cur_price, ask_price, upper_price = self.get_current_price(code)

            if not ask_price:
                if ask_price == 0:
                    self.buyorder_set.add(code)
                    logging.info(f"{stock_name}({code}) 상한가 주문 불가")
                else:
                    logging.error(f"{stock_name}({code}) 현재가 조회 실패")
                return
            if (len(self.buyorder_set) + len(self.bought_set)) >= self.target_buy_count:
                return
            buy_qty = self.buy_amount // ask_price
            total_amount = self.total_cash - self.buy_amount * (len(self.buyorder_set) + len(self.bought_set))
            max_buy_qty = total_amount // upper_price
            if max_buy_qty <= 0 or buy_qty <= 0: 
                return
            self.buyorder_qty[code] = int(min(buy_qty, max_buy_qty))

            if self.buyorder_qty[code] >= 0:
                acc = self.cpTrade.AccountNumber[0]
                accFlag = self.cpTrade.GoodsList(acc, 1)
                if self.cp_request.is_requesting:
                    logging.debug("요청 진행 중, 매수 주문 스킵")
                    return None

                self.cp_request.is_requesting = True
                self.cpOrder.SetInputValue(0, "2")
                self.cpOrder.SetInputValue(1, acc)
                self.cpOrder.SetInputValue(2, accFlag[0])
                self.cpOrder.SetInputValue(3, code)
                self.cpOrder.SetInputValue(4, self.buyorder_qty[code])
                if buy_message == '발동가':
                    self.cpOrder.SetInputValue(5, self.starting_price[code])
                else:
                    self.cpOrder.SetInputValue(5, ask_price)
                self.cpOrder.SetInputValue(7, order_condition)
                self.cpOrder.SetInputValue(8, order_style)
                
                remain_count0 = cpStatus.GetLimitRemainCount(0)
                if remain_count0 == 0:
                    logging.error(f"거래 요청 제한")
                    return
                
                logging.info(f"{stock_name}({code}), {buy_message} -> 매수 요청({self.buyorder_qty[code]}주)")
                
                handler = win32com.client.WithEvents(self.cpOrder, CpRequest)
                handler.set_params(self.cpOrder, 'order', self, {'code': code, 'stock_name': stock_name})
                self.cpOrder.Request()          
                self.buyorder_set.add(code)
                if code in self.buyordering_set:
                    self.buyordering_set.remove(code)
                result = handler.wait()
                if not result:
                    logging.warning(f"{stock_name}({code}) 매수주문 실패")
                    if code in self.buyorder_set:
                        self.buyorder_set.remove(code)
                    return            

        except Exception as ex:
            logging.error(f"buy_stock -> {code}, {ex}")

    @pyqtSlot(str, str)
    def sell_stock(self, code, message):
        """매도 주문"""
        try:
            if code in self.sellorder_set:
                return
            
            stock_name = self.cpCodeMgr.CodeToName(code)
            if code in self.buy_qty:
                stock_qty = self.buy_qty[code]
            else:
                return
            
            sell_order_qty = min(stock_qty, self.balance_qty[code])
            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)

            if self.cp_request.is_requesting:
                logging.debug("요청 진행 중, 매도 주문 스킵")
                return None

            self.cp_request.is_requesting = True
            self.cpOrder.SetInputValue(0, "1")
            self.cpOrder.SetInputValue(1, acc)
            self.cpOrder.SetInputValue(2, accFlag[0])
            self.cpOrder.SetInputValue(3, code)
            self.cpOrder.SetInputValue(4, sell_order_qty)
            self.cpOrder.SetInputValue(7, "0")
            self.cpOrder.SetInputValue(8, "03")
        
            remain_count0 = cpStatus.GetLimitRemainCount(0)
            if remain_count0 == 0:
                logging.error(f"거래 요청 제한")
                return
            
            logging.info(f"{stock_name}({code}), {message} -> 매도 요청({sell_order_qty}주)")
            handler = win32com.client.WithEvents(self.cpOrder, CpRequest)
            handler.set_params(self.cpOrder, 'order', self, {'code': code, 'stock_name': stock_name})
            self.cpOrder.Request()
            self.sellorder_set.add(code) 

            result = handler.wait()
            if not result:
                logging.warning(f"{stock_name}({code}) 매도주문 실패")
                if code in self.sellorder_set:
                    self.sellorder_set.remove(code)
                return
                       
        except Exception as ex:
            logging.error(f"sell_stock -> {code}, {ex}")

    @pyqtSlot(str, str)
    def sell_half_stock(self, code, message):
        """분할 매도 주문"""
        try:
            if code in self.sellorder_set:
                return
            
            stock_name = self.cpCodeMgr.CodeToName(code)
            if code in self.buy_qty:
                stock_qty = self.buy_qty[code]
            else:
                return
            self.sell_half_qty[code] = stock_qty - ((stock_qty + 1) // 2)
            sell_half_order_qty = min(((stock_qty + 1) // 2), self.balance_qty.get(code, 0))

            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)

            if self.cp_request.is_requesting:
                logging.debug("요청 진행 중, 분할매도 주문 스킵")
                return None

            self.cp_request.is_requesting = True
            self.cpOrder.SetInputValue(0, "1")
            self.cpOrder.SetInputValue(1, acc)
            self.cpOrder.SetInputValue(2, accFlag[0])
            self.cpOrder.SetInputValue(3, code)
            self.cpOrder.SetInputValue(4, sell_half_order_qty)
            self.cpOrder.SetInputValue(7, "0")
            self.cpOrder.SetInputValue(8, "03")

            remain_count0 = cpStatus.GetLimitRemainCount(0)
            if remain_count0 == 0:
                logging.error(f"거래 요청 제한")
                return
            
            logging.info(f"{stock_name}({code}), {message} -> 분할 매도 요청({sell_half_order_qty}주)")
            handler = win32com.client.WithEvents(self.cpOrder, CpRequest)
            handler.set_params(self.cpOrder, 'order', self, {'code': code, 'stock_name': stock_name})
            self.cpOrder.Request()
            self.sellorder_set.add(code) 

            result = handler.wait()
            if not result:
                logging.warning(f"{stock_name}({code}) 분할매도주문 실패")
                if code in self.sellorder_set:
                    self.sellorder_set.remove(code)
                return
                       
        except Exception as ex:
            logging.error(f"sell_half_stock -> {code}, {ex}")

    @pyqtSlot()
    def sell_all(self):
        """전량 매도"""
        try:
            stocks = self.get_stock_balance('ALL', 'sell_all')
            acc = self.cpTrade.AccountNumber[0]
            accFlag = self.cpTrade.GoodsList(acc, 1)

            for s in stocks:
                if s['qty'] > 0:
                    if self.cp_request.is_requesting:
                        logging.debug("요청 진행 중, 전부매도 주문 스킵")
                        return None

                    self.cp_request.is_requesting = True
                    self.cpOrder.SetInputValue(0, "1")
                    self.cpOrder.SetInputValue(1, acc)
                    self.cpOrder.SetInputValue(2, accFlag[0])
                    self.cpOrder.SetInputValue(3, s['code'])
                    self.cpOrder.SetInputValue(4, s['qty'])
                    self.cpOrder.SetInputValue(7, "0")
                    self.cpOrder.SetInputValue(8, "03")

                    logging.info(f"{s['name']}({s['code']}) -> 매도 요청({s['qty']}주)")
                    handler = win32com.client.WithEvents(self.cpOrder, CpRequest)
                    handler.set_params(self.cpOrder, 'order', self, {'code': s['code'], 'stock_name': s['name']})
                    self.cpOrder.Request()
                    self.sellorder_set.add(s['code'])

                    result = handler.wait()
                    if not result:
                        logging.warning(f"{s['name']}({s['code']}) 전부매도주문 실패")
                        if s['code'] in self.sellorder_set:
                            self.sellorder_set.remove(s['code'])
            return True

        except Exception as ex:
            logging.error(f"sell_all -> {ex}")
            return False

    def monitorOrderStatus(self, code, ordernum, conflags, price, qty, bs, balance, buyprice):
        """주문 체결 모니터링"""
        try:
            stock_name = self.cpCodeMgr.CodeToName(code)

            # ===== 매도 체결 =====
            if bs == '1' and conflags == "체결":
                logging.debug(f"{stock_name}({code}), {price}원, {qty}주 매도, 잔고: {balance}주")
                self.balance_qty[code] = balance

                if code not in self.sell_amount:
                    self.sell_amount[code] = 0
                self.sell_amount[code] += price * qty

                # 분할 매도 완료
                if code in self.sell_half_qty and balance == self.sell_half_qty[code]:
                    logging.info(f"{stock_name}({code}), 분할 매도 완료")                  
                    
                    stock_profit = self.sell_amount[code] * 0.99835 - self.buy_price[code] * (self.buy_qty[code] - balance) * 1.00015
                    stock_rate = (stock_profit / (self.buy_price[code] * (self.buy_qty[code] - balance))) * 100
                    
                    if stock_profit > 0:
                        logging.info(f"{stock_name}({code}), 매매이익({int(stock_profit)}원, {stock_rate:.2f}%)")
                        send_slack_message(self.window.login_handler, "#stock", f"{stock_name}({code}), 매매이익({int(stock_profit)}원, {stock_rate:.2f}%)")
                    else:
                        logging.info(f"{stock_name}({code}), 매매손실({int(stock_profit)}원, {stock_rate:.2f}%)")
                        send_slack_message(self.window.login_handler, "#stock", f"{stock_name}({code}), 매매손실({int(stock_profit)}원, {stock_rate:.2f}%)")
                    
                    self.get_trade_profit()
                    
                    if code in self.sellorder_set:
                        self.sellorder_set.remove(code)
                    if code not in self.sell_half_set:
                        self.sell_half_set.add(code)
                    
                    self.buy_qty[code] = balance
                    self.sell_amount[code] = 0
                    
                    if code in self.sell_half_qty:
                        del self.sell_half_qty[code]

                # 전량 매도 완료
                if balance == 0:
                    logging.info(f"{stock_name}({code}), 매도 완료")

                    stock_profit = self.sell_amount[code] * 0.99835 - self.buy_price[code] * self.buy_qty[code] * 1.00015
                    stock_rate = (stock_profit / (self.buy_price[code] * self.buy_qty[code])) * 100
                    
                    # 보유 시간 계산
                    hold_minutes = 0
                    if code in self.starting_time:
                        try:
                            buy_time = datetime.strptime(
                                f"{datetime.now().year}/{self.starting_time[code]}", 
                                '%Y/%m/%d %H:%M:%S'
                            )
                            hold_minutes = (datetime.now() - buy_time).total_seconds() / 60
                        except:
                            pass
                    
                    # 매도 거래 기록 저장
                    self.save_trade_record(
                        code=code,
                        action='SELL',
                        price=price,
                        quantity=self.buy_qty[code],
                        buy_price=self.buy_price[code],
                        profit=stock_profit,
                        profit_pct=stock_rate,
                        hold_minutes=hold_minutes,
                        sell_reason='매도 완료'
                    )
                    
                    # 일별 요약 업데이트
                    self.update_daily_summary()
                    
                    if stock_profit > 0:
                        logging.info(f"{stock_name}({code}), 매매이익({int(stock_profit)}원, {stock_rate:.2f}%)")
                        send_slack_message(self.window.login_handler, "#stock", f"{stock_name}({code}), 매매이익({int(stock_profit)}원, {stock_rate:.2f}%)")
                    else:
                        logging.info(f"{stock_name}({code}), 매매손실({int(stock_profit)}원, {stock_rate:.2f}%)")
                        send_slack_message(self.window.login_handler, "#stock", f"{stock_name}({code}), 매매손실({int(stock_profit)}원, {stock_rate:.2f}%)")
                    
                    self.get_trade_profit()                    
                                        
                    self.stock_sold.emit(code)                    
                    
                    if code in self.bought_set:
                        self.bought_set.remove(code) 
                    if code in self.sell_half_set:
                        self.sell_half_set.remove(code)
                    if code in self.sellorder_set:
                        self.sellorder_set.remove(code)                 
                    if code in self.buy_price:
                        del self.buy_price[code]
                    if code in self.buy_qty:
                        del self.buy_qty[code]
                    if code in self.sell_amount:
                        del self.sell_amount[code]

            # ===== 매수 체결 =====
            elif bs == '2' and conflags == "체결":
                logging.debug(f"{stock_name}({code}), {qty}주 매수, 잔고: {balance}주")
                self.balance_qty[code] = balance
                
                if code in self.buyorder_qty and balance >= self.buyorder_qty[code]:
                    self.buy_qty[code] = balance
                    self.buy_price[code] = buyprice
                    logging.info(f"{stock_name}({code}), {self.buy_qty[code]}주, 매수 완료({int(self.buy_price[code])}원)")
                    
                    # 매수 거래 기록 저장
                    self.save_trade_record(
                        code=code,
                        action='BUY',
                        price=buyprice,
                        quantity=self.buy_qty[code],
                        buy_reason='매수 완료'
                    )
                    
                    if code not in self.bought_set:
                        self.bought_set.add(code)
                        self.stock_bought.emit(code)
                    if code in self.buyorder_set:
                        self.buyorder_set.remove(code)
                    if code in self.buyorder_qty:
                        del self.buyorder_qty[code]                
                    return

        except Exception as ex:
            logging.error(f"monitorOrderStatus -> {ex}\n{traceback.format_exc()}")

# ==================== AutoTraderThread (통합 전략 적용) ====================
class AutoTraderThread(QThread):
    """자동매매 스레드 - 통합 전략 (DatabaseWorker 제거 반영)"""
    
    buy_signal = pyqtSignal(str, str, str, str)
    sell_signal = pyqtSignal(str, str)
    sell_half_signal = pyqtSignal(str, str)
    sell_all_signal = pyqtSignal()
    stock_removed_from_monitor = pyqtSignal(str)
    counter_updated = pyqtSignal(int)
    stock_data_updated = pyqtSignal(list)

    def __init__(self, trader, window):
        super().__init__()
        
        self.trader = trader
        self.window = window
        
        self.running = True
        self.sell_all_emitted = False
        
        self.counter = 0
        
        self.volatility_strategy = None
        
        self.load_trading_settings()
        
        self.last_evaluation_time = {}
        self.evaluation_lock = threading.Lock()
        
        # DB 저장용
        self.last_save_time = {}
        self.save_lock = threading.Lock()

    def load_trading_settings(self):
        """매매 평가 설정 로드"""
        config = configparser.ConfigParser(interpolation=None)
        if os.path.exists('settings.ini'):
            config.read('settings.ini', encoding='utf-8')
        
        # 기본값 설정
        self.evaluation_interval = config.getint('TRADING', 'evaluation_interval', fallback=5)
        self.event_based_evaluation = config.getboolean('TRADING', 'event_based_evaluation', fallback=True)
        self.min_evaluation_gap = config.getfloat('TRADING', 'min_evaluation_gap', fallback=3.0)
        
        logging.info(
            f"매매 평가 설정: 주기={self.evaluation_interval}초, "
            f"이벤트기반={self.event_based_evaluation}, "
            f"최소간격={self.min_evaluation_gap}초"
        )

    def set_volatility_strategy(self, strategy):
        """변동성 돌파 전략 설정"""
        self.volatility_strategy = strategy

    def connect_bar_signals(self):
        """봉 완성 signal 연결"""
        if self.event_based_evaluation:
            # 틱봉 완성 시
            self.trader.tickdata.new_bar_completed.connect(self.on_tick_bar_completed)
            # 분봉 완성 시
            self.trader.mindata.new_bar_completed.connect(self.on_min_bar_completed)
            logging.info("이벤트 기반 매매 평가 활성화")

    @pyqtSlot(str)
    def on_tick_bar_completed(self, code):
        """틱봉 완성 시 즉시 평가"""
        self._evaluate_code_if_ready(code, "틱봉 완성")

    @pyqtSlot(str)
    def on_min_bar_completed(self, code):
        """분봉 완성 시 즉시 평가"""
        self._evaluate_code_if_ready(code, "분봉 완성")

    def _evaluate_code_if_ready(self, code, reason):
        """종목 평가 (최소 간격 체크) + DB 저장"""
        
        if code not in self.trader.monistock_set:
            return
        
        with self.evaluation_lock:
            now = time.time()
            last_time = self.last_evaluation_time.get(code, 0)
            
            if now - last_time < self.min_evaluation_gap:
                return
            
            self.last_evaluation_time[code] = now
        
        t_now = datetime.now()
        
        if not self._is_trading_hours(t_now):
            return
        
        try:
            # 현재 데이터 조회
            tick_latest = self.trader.tickdata.get_latest_data(code)
            min_latest = self.trader.mindata.get_latest_data(code)
            
            if not tick_latest or not min_latest:
                return
            
            # ✅ 평가 시점마다 DB 저장
            self.save_to_db_if_needed(code, t_now, tick_latest, min_latest, reason)
            
            # 매매 조건 평가
            current_strategy = self.window.comboStg.currentText()
            buy_strategies = [
                stg for stg in self.window.strategies.get(current_strategy, []) 
                if stg['key'].startswith('buy')
            ]
            sell_strategies = [
                stg for stg in self.window.strategies.get(current_strategy, []) 
                if stg['key'].startswith('sell')
            ]
            
            if code not in self.trader.buyorder_set and code not in self.trader.bought_set:
                logging.debug(f"{code}: {reason} - 매수 조건 평가")
                self._evaluate_buy_condition(code, t_now, current_strategy, buy_strategies)
            
            elif (code in self.trader.bought_set and 
                  code not in self.trader.buyorder_set and 
                  code not in self.trader.sellorder_set):
                logging.debug(f"{code}: {reason} - 매도 조건 평가")
                self._evaluate_sell_condition(code, t_now, current_strategy, sell_strategies)
                
        except Exception as ex:
            logging.error(f"{code} 이벤트 기반 평가 오류: {ex}")

    def save_to_db_if_needed(self, code, timestamp, tick_data, min_data, trigger_reason):
        """조건부 DB 저장 (새 지표 포함)"""
        
        should_save = False
        save_reason = ""
        last_save_backup = 0
        
        # === 저장 필요 여부 판단 ===
        with self.save_lock:
            now = time.time()
            last_save = self.last_save_time.get(code, 0)
            last_save_backup = last_save
            
            if now - last_save >= 5.0:
                should_save = True
                save_reason = "주기적 저장"
                self.last_save_time[code] = now
            elif "완성" in trigger_reason and now - last_save >= 1.0:
                should_save = True
                save_reason = trigger_reason
                self.last_save_time[code] = now
            elif (code in self.trader.buyorder_set or code in self.trader.sellorder_set):
                should_save = True
                save_reason = "매매 발생"
                self.last_save_time[code] = now
        
        if not should_save:
            return
        
        # === 실제 DB 저장 ===
        try:
            conn = sqlite3.connect(self.trader.db_name, timeout=5)
            cursor = conn.cursor()
            
            date_str = timestamp.strftime('%Y%m%d')
            time_str = timestamp.strftime('%H%M%S')
            
            # 체결강도
            strength = self.trader.tickdata.get_strength(code)
            
            # 포지션 정보
            if code in self.trader.bought_set:
                position_type = 'BOUGHT'
                buy_price = self.trader.buy_price.get(code, None)
            elif code in self.trader.buyorder_set:
                position_type = 'BUYORDER'
                buy_price = None
            else:
                position_type = 'NONE'
                buy_price = None
            
            # INSERT OR REPLACE
            cursor.execute('''
                INSERT OR REPLACE INTO combined_tick_data (
                    code, timestamp, date, time,
                    tick_C, tick_O, tick_H, tick_L, tick_V,
                    tick_MAT5, tick_MAT20, tick_MAT60, tick_MAT120,
                    tick_RSIT, tick_RSIT_SIGNAL,
                    tick_MACDT, tick_MACDT_SIGNAL, tick_OSCT,
                    tick_STOCHK, tick_STOCHD,
                    tick_ATR, tick_CCI,
                    tick_BB_UPPER, tick_BB_MIDDLE, tick_BB_LOWER,
                    tick_BB_POSITION, tick_BB_BANDWIDTH,
                    tick_VWAP,
                    tick_WILLIAMS_R, tick_ROC, tick_OBV, tick_OBV_MA20,
                    tick_VP_POC, tick_VP_POSITION,
                    min_C, min_O, min_H, min_L, min_V,
                    min_MAM5, min_MAM10, min_MAM20,
                    min_RSI, min_RSI_SIGNAL,
                    min_MACD, min_MACD_SIGNAL, min_OSC,
                    min_STOCHK, min_STOCHD,
                    min_CCI, min_VWAP,
                    min_WILLIAMS_R, min_ROC, min_OBV, min_OBV_MA20,
                    strength, buy_price, position_type,
                    save_reason
                ) VALUES (
                    ?, ?, ?, ?,
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?,
                    ?, ?,
                    ?, ?, ?,
                    ?, ?,
                    ?, ?,
                    ?, ?, ?,
                    ?, ?,
                    ?,
                    ?, ?, ?, ?,
                    ?, ?,
                    ?, ?, ?, ?, ?,
                    ?, ?, ?,
                    ?, ?,
                    ?, ?, ?,
                    ?, ?,
                    ?, ?,
                    ?, ?, ?, ?,
                    ?, ?, ?,
                    ?
                )
            ''', (
                code, timestamp, date_str, time_str,
                # 틱 데이터
                tick_data.get('C', 0), tick_data.get('O', 0), 
                tick_data.get('H', 0), tick_data.get('L', 0), tick_data.get('V', 0),
                tick_data.get('MAT5', 0), tick_data.get('MAT20', 0), 
                tick_data.get('MAT60', 0), tick_data.get('MAT120', 0),
                tick_data.get('RSIT', 0), tick_data.get('RSIT_SIGNAL', 0),
                tick_data.get('MACDT', 0), tick_data.get('MACDT_SIGNAL', 0), 
                tick_data.get('OSCT', 0),
                tick_data.get('STOCHK', 0), tick_data.get('STOCHD', 0),
                tick_data.get('ATR', 0), tick_data.get('CCI', 0),
                tick_data.get('BB_UPPER', 0), tick_data.get('BB_MIDDLE', 0), 
                tick_data.get('BB_LOWER', 0),
                tick_data.get('BB_POSITION', 0), tick_data.get('BB_BANDWIDTH', 0),
                tick_data.get('VWAP', 0),
                # 새 지표 - 틱
                tick_data.get('WILLIAMS_R', -50), tick_data.get('ROC', 0),
                tick_data.get('OBV', 0), tick_data.get('OBV_MA20', 0),
                tick_data.get('VP_POC', 0), tick_data.get('VP_POSITION', 0),
                # 분봉 데이터
                min_data.get('C', 0), min_data.get('O', 0), 
                min_data.get('H', 0), min_data.get('L', 0), min_data.get('V', 0),
                min_data.get('MAM5', 0), min_data.get('MAM10', 0), 
                min_data.get('MAM20', 0),
                min_data.get('RSI', 0), min_data.get('RSI_SIGNAL', 0),
                min_data.get('MACD', 0), min_data.get('MACD_SIGNAL', 0), 
                min_data.get('OSC', 0),
                min_data.get('STOCHK', 0), min_data.get('STOCHD', 0),
                min_data.get('CCI', 0), min_data.get('VWAP', 0),
                # 새 지표 - 분봉
                min_data.get('WILLIAMS_R', -50), min_data.get('ROC', 0),
                min_data.get('OBV', 0), min_data.get('OBV_MA20', 0),
                # 추가 정보
                strength, buy_price, position_type,
                save_reason
            ))
            
            conn.commit()
            conn.close()
            
            logging.debug(f"{code}: DB 저장 완료 ({save_reason})")
            
        except Exception as ex:
            logging.error(f"{code}: DB 저장 실패 - {ex}")
            
            with self.save_lock:
                self.last_save_time[code] = last_save_backup

    def run(self):
        """메인 루프"""
        while self.running:
            self.autotrade()
            # 5초 주기
            self.msleep(self.evaluation_interval * 1000)

    def stop(self):
        """스레드 정지"""
        logging.info("AutoTraderThread 정지 시작...")
        self.running = False
        
        self.quit()
        self.wait()
        logging.info("AutoTraderThread 정지 완료")

    def autotrade(self):
        """자동매매 메인 루프 (주기적 평가)"""
        try:
            t_now = datetime.now()
            
            self.counter += 1
            self.counter_updated.emit(self.counter)
            
            self._update_stock_data_table()
            
            if self._is_trading_hours(t_now):
                # 주기적 평가 + 저장
                self._execute_trading_logic(t_now)
            elif self._is_market_close_time(t_now):
                self._handle_market_close()
                
        except Exception as ex:
            logging.error(f"autotrade 오류: {ex}\n{traceback.format_exc()}")

    def _is_trading_hours(self, t_now):
        """거래 시간인지 확인"""
        t_0903 = t_now.replace(hour=9, minute=3, second=0, microsecond=0)
        t_1515 = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
        return t_0903 < t_now <= t_1515

    def _is_market_close_time(self, t_now):
        """장 종료 시간인지 확인"""
        t_1515 = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
        t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
        return t_1515 < t_now < t_exit and not self.sell_all_emitted

    def _update_stock_data_table(self):
        """주식 데이터 테이블 업데이트"""
        stock_data_list = []
        
        monitored_codes = list(self.trader.monistock_set.copy())
        
        for code in monitored_codes:
            if code not in self.trader.monistock_set:
                continue
            
            tick_latest = self.trader.tickdata.get_latest_data(code)
            current_price = tick_latest.get('C', 0.0) if tick_latest else 0.0
            buy_price = self.trader.buy_price.get(code, 0.0)
            quantity = self.trader.buy_qty.get(code, 0)

            stock_data_list.append({
                'code': code,
                'current_price': float(current_price),
                'upward_probability': 0.0,  # CNN 제거로 0
                'buy_price': float(buy_price),
                'quantity': quantity
            })
        
        self.stock_data_updated.emit(stock_data_list)

    def _execute_trading_logic(self, t_now):
        """거래 로직 실행 (5초 주기)"""
        
        monitored_codes = list(self.trader.monistock_set.copy())
        
        for code in monitored_codes:
            if code not in self.trader.monistock_set:
                continue
            
            try:
                # 5초마다 평가 + 저장
                self._evaluate_code_if_ready(code, "주기적 평가")
                    
            except Exception as ex:
                logging.error(f"{code} 거래 로직 오류: {ex}")

    def _handle_market_close(self):
        """장 종료 처리 (간소화)"""
        
        if self.trader.buyorder_set or self.trader.sellorder_set:
            return
        
        # 보유 주식 전부 매도
        if self.trader.bought_set:
            logging.info("보유 주식 전부 매도")
            self.sell_all_signal.emit()
        
        self.sell_all_emitted = True

    # ===== 헬퍼 함수들 (변경 없음) =====
    
    def get_threshold_by_hour(self):
        """시간대별 임계값 반환"""
        now = datetime.now()
        hour = now.hour
        
        if hour == 9:
            return 65
        elif hour >= 14:
            return 85
        else:
            return 75
    
    def is_after_time(self, hour, minute):
        """특정 시각 이후인지 확인"""
        now = datetime.now()
        return now >= now.replace(hour=hour, minute=minute, second=0)

    # ===== 매수 조건 평가 =====

    def _evaluate_buy_condition(self, code, t_now, strategy, buy_strategies):
        """매수 조건 평가"""
        tick_latest = self.trader.tickdata.get_latest_data(code)
        min_latest = self.trader.mindata.get_latest_data(code)
        
        if not tick_latest or not min_latest:
            return
        
        if tick_latest.get('MAT5', 0) == 0:
            logging.debug(f"{code}: 지표 준비 미완료")
            return
        
        if self._should_remove_from_monitor(code, tick_latest, min_latest, t_now):
            return
        
        # ===== 통합 전략: 텍스트 기반 평가 =====
        if strategy == "통합 전략" and buy_strategies:
            if self._evaluate_integrated_buy(code, buy_strategies, tick_latest, min_latest):
                return  # 매수 신호 발생
        
        # ===== 기타 전략 =====
        elif self._evaluate_strategy_conditions(code, buy_strategies, tick_latest, min_latest):
            self.buy_signal.emit(code, "사용자 전략", "0", "03")

    def _evaluate_integrated_buy(self, code, buy_strategies, tick_latest, min_latest):
        """매수 평가 - 새 지표 포함"""
        
        safe_globals = {
            '__builtins__': {
                'min': min, 'max': max, 'abs': abs, 'round': round,
                'int': int, 'float': float, 'bool': bool, 'str': str,
                'len': len, 'sum': sum, 'all': all, 'any': any,
                'True': True, 'False': False, 'None': None
            }
        }
        
        # === 기존 변수들 ===
        MAT5 = tick_latest.get('MAT5', 0)
        MAT20 = tick_latest.get('MAT20', 0)
        MAT60 = tick_latest.get('MAT60', 0)
        MAT120 = tick_latest.get('MAT120', 0)
        C = tick_latest.get('C', 0)
        VWAP = tick_latest.get('VWAP', 0)
        RSIT = tick_latest.get('RSIT', 50)
        MACDT = tick_latest.get('MACDT', 0)
        OSCT = tick_latest.get('OSCT', 0)
        STOCHK = tick_latest.get('STOCHK', 50)
        STOCHD = tick_latest.get('STOCHD', 50)
        ATR = tick_latest.get('ATR', 0)
        BB_POSITION = tick_latest.get('BB_POSITION', 0)
        BB_BANDWIDTH = tick_latest.get('BB_BANDWIDTH', 0)
        
        MAM5 = min_latest.get('MAM5', 0)
        MAM10 = min_latest.get('MAM10', 0)
        MAM20 = min_latest.get('MAM20', 0)
        min_close = min_latest.get('C', 0)
        min_RSI = min_latest.get('RSI', 50)
        min_STOCHK = min_latest.get('STOCHK', 50)
        min_STOCHD = min_latest.get('STOCHD', 50)
        
        # === 새로운 지표들 ===
        WILLIAMS_R = tick_latest.get('WILLIAMS_R', -50)
        ROC = tick_latest.get('ROC', 0)
        OBV = tick_latest.get('OBV', 0)
        OBV_MA20 = tick_latest.get('OBV_MA20', 0)
        VP_POC = tick_latest.get('VP_POC', 0)
        VP_POSITION = tick_latest.get('VP_POSITION', 0)
        
        min_WILLIAMS_R = min_latest.get('WILLIAMS_R', -50)
        min_ROC = min_latest.get('ROC', 0)
        min_OBV = min_latest.get('OBV', 0)
        min_OBV_MA20 = min_latest.get('OBV_MA20', 0)
        
        # === 기타 변수 ===
        strength = self.trader.tickdata.get_strength(code)
        momentum_score = 0
        if self.window.momentum_scanner:
            momentum_score = self.window.momentum_scanner._calculate_momentum_score(code)
        
        threshold = self.get_threshold_by_hour()
        
        volatility_breakout = False
        if self.volatility_strategy:
            volatility_breakout = self.volatility_strategy.check_breakout(code)
        
        gap_hold = False
        if hasattr(self.window, 'gap_scanner'):
            gap_hold = self.window.gap_scanner.check_gap_hold(code)
        
        # === ROC 최근 추이 ===
        tick_recent = self.trader.tickdata.get_recent_data(code, 5)
        ROC_recent = tick_recent.get('ROC', [0] * 5)
        
        # === Volume Profile 돌파 여부 ===
        volume_profile_breakout = (VP_POSITION > 0)  # 현재가가 POC 위
        
        # === 허용된 변수 딕셔너리 ===
        safe_locals = {
            # 틱 데이터 - 기본
            'MAT5': MAT5, 'MAT20': MAT20, 'MAT60': MAT60, 'MAT120': MAT120,
            'C': C, 'VWAP': VWAP, 'RSIT': RSIT,
            'MACDT': MACDT, 'OSCT': OSCT,
            'STOCHK': STOCHK, 'STOCHD': STOCHD,
            'ATR': ATR, 'BB_POSITION': BB_POSITION, 'BB_BANDWIDTH': BB_BANDWIDTH,
            
            # 틱 데이터 - 새 지표
            'WILLIAMS_R': WILLIAMS_R,
            'ROC': ROC,
            'ROC_recent': ROC_recent,
            'OBV': OBV,
            'OBV_MA20': OBV_MA20,
            'VP_POC': VP_POC,
            'VP_POSITION': VP_POSITION,
            'volume_profile_breakout': volume_profile_breakout,
            
            # 분봉 데이터 - 기본
            'MAM5': MAM5, 'MAM10': MAM10, 'MAM20': MAM20,
            'min_close': min_close, 'min_RSI': min_RSI,
            'min_STOCHK': min_STOCHK, 'min_STOCHD': min_STOCHD,
            
            # 분봉 데이터 - 새 지표
            'min_WILLIAMS_R': min_WILLIAMS_R,
            'min_ROC': min_ROC,
            'min_OBV': min_OBV,
            'min_OBV_MA20': min_OBV_MA20,
            
            # 기타
            'strength': strength,
            'momentum_score': momentum_score,
            'threshold': threshold,
            'volatility_breakout': volatility_breakout,
            'gap_hold': gap_hold,
            'code': code
        }
        
        # === 전략 평가 ===
        for strategy in buy_strategies:
            try:
                condition = strategy.get('content', '')
                
                if eval(condition, safe_globals, safe_locals):
                    buy_reason = strategy.get('name', '통합 전략')
                    logging.info(
                        f"{cpCodeMgr.CodeToName(code)}({code}): {buy_reason} 매수 "
                        f"(체결강도: {strength:.0f}, 점수: {momentum_score}, "
                        f"Williams %R: {WILLIAMS_R:.1f}, ROC: {ROC:.2f}%)"
                    )
                    self.buy_signal.emit(code, buy_reason, "0", "03")
                    return True
                    
            except Exception as ex:
                logging.error(f"{code} 매수 전략 '{strategy.get('name')}' 오류: {ex}")
        
        return False

    def _check_momentum_buy(self, code, tick_latest, min_latest):
        """급등주 모멘텀 매수 조건 (기존 전략용)"""
        
        if len(self.trader.bought_set) >= self.trader.target_buy_count:
            return False
        
        # ===== 1. 체결강도 확인 =====
        strength = self.trader.tickdata.get_strength(code)
        if strength < 120:
            logging.debug(f"{code}: 체결강도 부족 ({strength:.0f})")
            return False
        
        # ===== 2. 모멘텀 점수 계산 =====
        score = 0
        
        # 가격 모멘텀 (0-30점)
        tick_recent = self.trader.tickdata.get_recent_data(code, 5)
        C_recent = tick_recent.get('C', [0]*5)
        
        if len(C_recent) >= 5 and C_recent[0] > 0:
            price_momentum = (C_recent[-1] - C_recent[0]) / C_recent[0] * 100
            
            if price_momentum > 1.0:
                score += 30
            elif price_momentum > 0.5:
                score += 20
            elif price_momentum > 0:
                score += 10
        
        # 이평선 추세 (0-25점)
        MAT5 = tick_latest.get('MAT5', 0)
        MAT20 = tick_latest.get('MAT20', 0)
        MAT60 = tick_latest.get('MAT60', 0)
        C = tick_latest.get('C', 0)
        
        if C > MAT5 > MAT20:
            score += 25
        elif C > MAT5:
            score += 15
        elif MAT5 > MAT20:
            score += 10
        
        # VWAP 대비 위치 (0-20점)
        VWAP = tick_latest.get('VWAP', 0)
        if VWAP > 0:
            if C > VWAP * 1.01:
                score += 20
            elif C > VWAP:
                score += 10
        
        # 체결강도 점수 (0-15점)
        if strength >= 150:
            score += 15
        elif strength >= 130:
            score += 10
        elif strength >= 120:
            score += 5
        
        # RSI (0-10점)
        RSIT = tick_latest.get('RSIT', 50)
        if 50 < RSIT < 70:
            score += 10
        elif 40 < RSIT <= 50:
            score += 5
        
        # ===== 3. 시간대별 임계값 =====
        now = datetime.now()
        hour = now.hour
        
        if hour == 9:
            threshold = 65
        elif hour >= 14:
            threshold = 85
        else:
            threshold = 75
        
        logging.debug(
            f"{code}: 매수점수 {score}/{threshold} "
            f"(체결강도: {strength:.0f})"
        )
        
        return score >= threshold

    def _should_remove_from_monitor(self, code, tick_latest, min_latest, t_now):
        """투자 대상에서 제거해야 하는지 확인"""
        min_close = min_latest.get('C', 0)
        MAM5 = min_latest.get('MAM5', 0)
        MAM10 = min_latest.get('MAM10', 0)
        
        # 급격한 하락
        if code in self.trader.starting_price:
            if min_close < self.trader.starting_price[code] * 0.97:
                if MAM5 < MAM10:
                    try:
                        start_time_str = self.trader.starting_time.get(code, '')
                        if start_time_str:
                            start_time = datetime.strptime(
                                f"{datetime.now().year}/{start_time_str}", 
                                '%Y/%m/%d %H:%M:%S'
                            )
                            if t_now - start_time > timedelta(hours=1):
                                logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> 투자 대상 제거 (하락)")
                                self.stock_removed_from_monitor.emit(code)
                                return True
                    except Exception as ex:
                        logging.error(f"{code} 시작 시각 파싱 오류: {ex}")
        
        # 상한가 체크
        min_high_recent = min_latest.get('H_recent', [0, 0])
        min_low_recent = min_latest.get('L_recent', [0, 0])
        
        if len(min_high_recent) >= 2 and len(min_low_recent) >= 2:
            if all(h == l for h, l in zip(min_high_recent[-2:], min_low_recent[-2:])):
                logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> 투자 대상 제거 (상한가)")
                self.stock_removed_from_monitor.emit(code)
                return True
        
        return False

    def _evaluate_strategy_conditions(self, code, strategies, tick_latest, min_latest):
        """전략별 조건 평가 (기존 전략용)"""
        if not strategies:
            return False
        
        tick_data_full = self.trader.tickdata.get_recent_data(code, 10)
        min_data_full = self.trader.mindata.get_recent_data(code, 10)
        
        tick_close_price = tick_data_full.get('C', [0])
        MAT5 = tick_data_full.get('MAT5', [0])
        MAT20 = tick_data_full.get('MAT20', [0])
        MAT60 = tick_data_full.get('MAT60', [0])
        RSIT = tick_data_full.get('RSIT', [0])
        OSCT = tick_data_full.get('OSCT', [0])
        bb_upper = tick_data_full.get('BB_UPPER', [0])
        
        min_close_price = min_data_full.get('C', [0])
        MAM5 = min_data_full.get('MAM5', [0])
        MAM10 = min_data_full.get('MAM10', [0])
        RSI = min_data_full.get('RSI', [0])
        OSC = min_data_full.get('OSC', [0])
        VWAP = min_data_full.get('VWAP', [0])
        
        min_close_recent = min_data_full.get('C', [0, 0])[-2:]
        min_open_recent = min_data_full.get('O', [0, 0])[-2:]
        positive_candle = all(
            min_close_recent[i] > min_open_recent[i] 
            for i in range(min(2, len(min_close_recent), len(min_open_recent)))
        )
        
        for strategy in strategies:
            try:
                condition = strategy.get('content', '')
                if eval(condition):
                    logging.debug(f"{code}: {strategy.get('name')} 조건 만족")
                    return True
            except Exception as ex:
                logging.error(f"{code} 전략 평가 오류: {ex}")
        
        return False

    # ===== 매도 조건 평가 =====

    def _evaluate_sell_condition(self, code, t_now, strategy, sell_strategies):
        """매도 조건 평가"""
        tick_latest = self.trader.tickdata.get_latest_data(code)
        min_latest = self.trader.mindata.get_latest_data(code)
        
        if not tick_latest or not min_latest:
            return
        
        tick_close = tick_latest.get('C', 0)
        
        self.trader.update_highest_price(code, tick_close)
        
        buy_price = self.trader.buy_price.get(code, 0)
        if buy_price == 0:
            return
        
        # ===== 수익률 계산 =====
        current_profit_pct = (tick_close / buy_price - 1) * 100
        highest_price = self.trader.highest_price.get(code, buy_price)
        from_peak_pct = (tick_close / highest_price - 1) * 100
        
        # ===== 보유 시간 계산 =====
        buy_time_str = self.trader.starting_time.get(code)
        if buy_time_str:
            try:
                buy_time = datetime.strptime(
                    f"{datetime.now().year}/{buy_time_str}", 
                    '%Y/%m/%d %H:%M:%S'
                )
                hold_minutes = (t_now - buy_time).total_seconds() / 60
            except:
                hold_minutes = 0
        else:
            hold_minutes = 0
        
        # ===== 통합 전략: 텍스트 기반 평가 =====
        if strategy == "통합 전략" and sell_strategies:
            if self._evaluate_integrated_sell(code, sell_strategies, tick_latest, min_latest,
                                             current_profit_pct, from_peak_pct, hold_minutes):
                return
        
        # ===== 기타 전략 =====
        elif self._evaluate_strategy_conditions(code, sell_strategies, tick_latest, min_latest):
            self.sell_signal.emit(code, "전략 매도")

    def _evaluate_integrated_sell(self, code, sell_strategies, tick_latest, min_latest,
                              current_profit_pct, from_peak_pct, hold_minutes):
        """매도 평가 - 새 지표 포함"""
        
        safe_globals = {
            '__builtins__': {
                'min': min, 'max': max, 'abs': abs, 'round': round,
                'int': int, 'float': float, 'bool': bool, 'str': str,
                'len': len, 'sum': sum, 'all': all, 'any': any,
                'True': True, 'False': False, 'None': None
            }
        }
        
        # === 기존 변수들 ===
        tick_close = tick_latest.get('C', 0)
        MAT5 = tick_latest.get('MAT5', 0)
        MAT20 = tick_latest.get('MAT20', 0)
        RSIT = tick_latest.get('RSIT', 50)
        OSCT = tick_latest.get('OSCT', 0)
        STOCHK = tick_latest.get('STOCHK', 50)
        STOCHD = tick_latest.get('STOCHD', 50)
        ATR = tick_latest.get('ATR', 0)
        BB_POSITION = tick_latest.get('BB_POSITION', 0)
        CCI = tick_latest.get('CCI', 0)
        
        min_close = min_latest.get('C', 0)
        MAM5 = min_latest.get('MAM5', 0)
        MAM10 = min_latest.get('MAM10', 0)
        min_RSI = min_latest.get('RSI', 50)
        min_STOCHK = min_latest.get('STOCHK', 50)
        min_STOCHD = min_latest.get('STOCHD', 50)
        min_CCI = min_latest.get('CCI', 0)
        
        # === 새로운 지표들 ===
        WILLIAMS_R = tick_latest.get('WILLIAMS_R', -50)
        ROC = tick_latest.get('ROC', 0)
        OBV = tick_latest.get('OBV', 0)
        OBV_MA20 = tick_latest.get('OBV_MA20', 0)
        
        min_WILLIAMS_R = min_latest.get('WILLIAMS_R', -50)
        min_ROC = min_latest.get('ROC', 0)
        min_OBV = min_latest.get('OBV', 0)
        min_OBV_MA20 = min_latest.get('OBV_MA20', 0)
        
        # === 파생 지표 ===
        osct_negative = False
        tick_recent = self.trader.tickdata.get_recent_data(code, 3)
        OSCT_recent = tick_recent.get('OSCT', [0, 0, 0])
        if len(OSCT_recent) >= 2:
            osct_negative = OSCT_recent[-2] < 0 and OSCT_recent[-1] < 0
        
        after_market_close = self.is_after_time(14, 45)
        buy_price = self.trader.buy_price.get(code, 0)
        highest_price = self.trader.highest_price.get(code, buy_price)
        
        # === OBV 다이버전스 감지 ===
        obv_divergence = (OBV < OBV_MA20 and current_profit_pct > 0)
        min_obv_divergence = (min_OBV < min_OBV_MA20 and current_profit_pct > 0)
        
        # === Williams %R 과매수/과매도 ===
        williams_overbought = (WILLIAMS_R > -20)  # 과매수
        williams_oversold = (WILLIAMS_R < -80)    # 과매도
        
        # === 허용된 변수 딕셔너리 ===
        safe_locals = {
            # 틱 데이터 - 기본
            'tick_close': tick_close, 'C': tick_close,
            'MAT5': MAT5, 'MAT20': MAT20,
            'RSIT': RSIT, 'OSCT': OSCT, 'osct_negative': osct_negative,
            'STOCHK': STOCHK, 'STOCHD': STOCHD,
            'ATR': ATR, 'BB_POSITION': BB_POSITION, 'CCI': CCI,
            
            # 틱 데이터 - 새 지표
            'WILLIAMS_R': WILLIAMS_R,
            'williams_overbought': williams_overbought,
            'williams_oversold': williams_oversold,
            'ROC': ROC,
            'OBV': OBV,
            'OBV_MA20': OBV_MA20,
            'obv_divergence': obv_divergence,
            
            # 분봉 데이터 - 기본
            'min_close': min_close,
            'MAM5': MAM5, 'MAM10': MAM10,
            'min_RSI': min_RSI,
            'min_STOCHK': min_STOCHK, 'min_STOCHD': min_STOCHD,
            'min_CCI': min_CCI,
            
            # 분봉 데이터 - 새 지표
            'min_WILLIAMS_R': min_WILLIAMS_R,
            'min_ROC': min_ROC,
            'min_OBV': min_OBV,
            'min_OBV_MA20': min_OBV_MA20,
            'min_obv_divergence': min_obv_divergence,
            
            # 수익률 정보
            'current_profit_pct': current_profit_pct,
            'from_peak_pct': from_peak_pct,
            'hold_minutes': hold_minutes,
            'buy_price': buy_price,
            'highest_price': highest_price,
            
            # 기타
            'after_market_close': after_market_close,
            'code': code,
            'self': self
        }
        
        # === 전략 평가 ===
        for strategy in sell_strategies:
            try:
                condition = strategy.get('content', '')
                
                if eval(condition, safe_globals, safe_locals):
                    sell_reason = strategy.get('name', '통합 전략')
                    
                    logging.info(
                        f"{cpCodeMgr.CodeToName(code)}({code}): {sell_reason} "
                        f"({current_profit_pct:+.2f}%, {hold_minutes:.0f}분 보유, "
                        f"Williams %R: {WILLIAMS_R:.1f}, ROC: {ROC:.2f}%)"
                    )
                    
                    if '분할' in sell_reason:
                        self.sell_half_signal.emit(code, sell_reason)
                    else:
                        self.sell_signal.emit(code, sell_reason)
                    
                    return True
                    
            except Exception as ex:
                logging.error(f"{code} 매도 전략 '{strategy.get('name')}' 오류: {ex}")
        
        return False

# ==================== ChartDrawer 관련 클래스 (유지) ====================
class ChartDrawerThread(QThread):
    data_ready = pyqtSignal(dict)

    def __init__(self, trader, code):
        super().__init__()
        self.trader = trader
        self.code = code
        self.is_running = True

    def run(self):
        while self.is_running:
            if self.code:
                tick_data = self.trader.tickdata.get_full_data(self.code)
                min_data = self.trader.mindata.get_full_data(self.code)

                data = {'tick_data': tick_data, 'min_data': min_data, 'code': self.code}
                self.data_ready.emit(data)
                
            self.msleep(2000)

    def stop(self):
        self.is_running = False
        self.quit()
        self.wait()

class ChartDrawer(QObject):
    def __init__(self, fig, canvas, trader, trader_thread, window):
        super().__init__()
        self.fig = fig
        self.canvas = canvas
        self.trader = trader
        self.trader_thread = trader_thread
        self.window = window
        self.code = None
        self.create_initial_chart()
        self.chart_thread = None
        self.stop_loss = {}
        self.last_chart_update_time = None

    def set_code(self, code):
        self.code = code
        if self.chart_thread:
            self.chart_thread.stop()
            self.chart_thread = None

        if code:
            self.chart_thread = ChartDrawerThread(self.trader, code)
            self.chart_thread.data_ready.connect(self.update_chart)
            self.chart_thread.start()
        else:
            self.create_initial_chart()

    def create_initial_chart(self):
        self.fig.clear()
        self.tick_axes = [self.fig.add_subplot(18, 1, (1, 4), title='Tick Chart'), self.fig.add_subplot(18, 1, (5, 6)), self.fig.add_subplot(18, 1, (7, 8))]
        self.min_axes = [self.fig.add_subplot(18, 1, (10, 13), title='3 Min Chart'), self.fig.add_subplot(18, 1, 14), self.fig.add_subplot(18, 1, (15, 16)), self.fig.add_subplot(18, 1, (17, 18))]

        for ax in self.tick_axes + self.min_axes:
            ax.clear()
            ax.set_xticklabels([])
            ax.set_yticklabels([])
            ax.grid(True)

        self.fig.subplots_adjust(hspace=0)
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        self.canvas.draw()
        self.canvas.flush_events()

    @pyqtSlot(dict)
    def update_chart(self, data):
        self.draw_chart(data)
        self.last_chart_update_time = data['tick_data'].get('T', [None])[-1]

    def draw_chart(self, data):
        try:
            if data['code']:
                code_name = cpCodeMgr.CodeToName(data['code'])
                for ax in self.tick_axes:
                    ax.clear()
                for ax in self.min_axes:
                    ax.clear()

                self.draw_chart_data(data['tick_data'], self.tick_axes, data['code'], data_type='tick')
                self.draw_chart_data(data['min_data'], self.min_axes, data['code'], data_type='min')
                self.fig.suptitle(f"Chart for {data['code']}({code_name})", fontsize=15)
                self.fig.subplots_adjust(hspace=0)
                self.canvas.draw()
            else:
                self.create_initial_chart()

        except Exception as ex:
            logging.error(f"draw_chart -> {ex}")

    def draw_chart_data(self, chart_data, axes, code, data_type):
        try:
            current_strategy = self.window.comboStg.currentText()
            all_strategies = self.window.strategies.get(current_strategy, [])
            
            if chart_data:
                if data_type == 'tick':
                    keys_to_keep = ['O', 'H', 'L', 'C', 'V', 'D', 'T', 'MAT5', 'MAT20', 'MAT60', 'MAT120', 'RSIT', 'RSIT_SIGNAL', 'MACDT', 'MACDT_SIGNAL', 'OSCT']
                elif data_type == 'min':
                    keys_to_keep = ['O', 'H', 'L', 'C', 'V', 'D', 'T', 'MAM5', 'MAM10', 'MAM20', 'RSI', 'RSI_SIGNAL', 'MACD', 'MACD_SIGNAL', 'OSC']

                if data_type == 'tick':
                    filtered_data = {
                        key: [x for x in (chart_data[key][-90:] if len(chart_data[key]) >= 90 
                                         else chart_data[key][:])] 
                        for key in keys_to_keep
                    }

                    df = pd.DataFrame({
                        'Open': filtered_data['O'], 'High': filtered_data['H'], 'Low': filtered_data['L'], 'Close': filtered_data['C'], 'Volume': filtered_data['V'], 
                        'Date': filtered_data['D'], 'Time': filtered_data['T'], 'MAT5': filtered_data['MAT5'], 
                        'MAT20': filtered_data['MAT20'], 'MAT60': filtered_data['MAT60'], 'MAT120': filtered_data['MAT120'], 
                        'RSIT': filtered_data['RSIT'], 'RSIT_SIGNAL': filtered_data['RSIT_SIGNAL'], 'MACDT': filtered_data['MACDT'], 
                        'MACDT_SIGNAL': filtered_data['MACDT_SIGNAL'], 'OSCT': filtered_data['OSCT']
                    })
                    df.index = pd.to_datetime(df['Date'].astype(str) + df['Time'].astype(str), format='%Y%m%d%H%M', errors='coerce')

                    addplots = [
                        mpf.make_addplot(df['MAT5'], color='magenta', label='MAT5', ax=axes[0], width=0.9),
                        mpf.make_addplot(df['MAT20'], color='blue', label='MAT20', ax=axes[0], width=0.9),
                        mpf.make_addplot(df['MAT60'], color='darkorange', label='MAT60', ax=axes[0], width=1.0),                        
                        mpf.make_addplot(df['MAT120'], color='black', label='MAT120', ax=axes[0], width=0.9),
                        mpf.make_addplot(df['RSIT'], color='red', label='RSIT', ax=axes[1], width=0.9),
                        mpf.make_addplot(df['RSIT_SIGNAL'], color='blue', label='RSIT_SIGNAL', ax=axes[1], width=0.9),
                        mpf.make_addplot(df['MACDT'], color='red', label='MACDT', ax=axes[2], width=0.9),
                        mpf.make_addplot(df['MACDT_SIGNAL'], color='blue', label='MACDT_SIGNAL', ax=axes[2], width=0.9),
                        mpf.make_addplot(df['OSCT'], color='purple', type='bar', label='OSCT', ax=axes[2]),
                    ]
                    mpf.plot(df, ax=axes[0], type='candle', style='yahoo', addplot=addplots)

                    current_price = df['Close'].iloc[-1]
                    axes[0].text(
                        1.01, current_price, f'{current_price}',
                        transform=axes[0].get_yaxis_transform(), color='red',
                        verticalalignment='center', bbox=dict(facecolor='white', edgecolor='none')
                    )
                    axes[0].legend(loc='upper left')

                    x_raw = list(range(len(df['Date'])))
                    axes[1].fill_between(x_raw, df['RSIT'], 20, where=(df['RSIT'] <= 20), color='blue', alpha=0.5)
                    axes[1].fill_between(x_raw, df['RSIT'], 80, where=(df['RSIT'] >= 80), color='red', alpha=0.5)
                    axes[1].set_yticks([20, 50, 80])
                    axes[1].legend(loc='upper left')
                    axes[2].legend(loc='upper left')

                    x_labels = self.create_labels(filtered_data)
                    for index, ax in enumerate(axes):
                        ax.set_xlim(-2, len(filtered_data['D']) + 1)
                        ax.grid(True, axis='y')
                        ax.set_xticks(x_raw)
                        if index == 0 or index == 1:
                            ax.set_xticklabels([])
                        if index == 2:
                            ax.set_xticklabels(x_labels)

                elif data_type == 'min':
                    filtered_data = {key: [x for x in (chart_data[key][-50:] if len(chart_data[key]) >= 50 else chart_data[key][:])] for key in keys_to_keep}

                    df = pd.DataFrame({
                        'Open': filtered_data['O'], 'High': filtered_data['H'], 'Low': filtered_data['L'], 'Close': filtered_data['C'], 'Volume': filtered_data['V'],
                        'Date': filtered_data['D'], 'Time': filtered_data['T'], 'MAM5': filtered_data['MAM5'], 'MAM10': filtered_data['MAM10'], 'MAM20': filtered_data['MAM20'],
                        'RSI': filtered_data['RSI'], 'RSI_SIGNAL': filtered_data['RSI_SIGNAL'], 'MACD': filtered_data['MACD'], 'MACD_SIGNAL': filtered_data['MACD_SIGNAL'], 'OSC': filtered_data['OSC']
                    })
                    df.index = pd.to_datetime(df['Date'].astype(str) + df['Time'].astype(str), format='%Y%m%d%H%M')

                    addplots = [
                        mpf.make_addplot(df['MAM5'], color='magenta', label='MAM5', ax=axes[0], width=0.9),
                        mpf.make_addplot(df['MAM10'], color='blue', label='MAM10', ax=axes[0], width=0.9),
                        mpf.make_addplot(df['MAM20'], color='darkorange', label='MAM20', ax=axes[0], width=1.0),
                        mpf.make_addplot(df['RSI'], color='red', label='RSI', ax=axes[2], width=0.9),
                        mpf.make_addplot(df['RSI_SIGNAL'], color='blue', label='RSI_SIGNAL', ax=axes[2], width=0.9),
                        mpf.make_addplot(df['MACD'], color='red', label='MACD', ax=axes[3], width=0.9),
                        mpf.make_addplot(df['MACD_SIGNAL'], color='blue', label='MACD_SIGNAL', ax=axes[3], width=0.9),
                        mpf.make_addplot(df['OSC'], color='purple', type='bar', label='OSC', ax=axes[3])
                    ]
                    mpf.plot(df, ax=axes[0], type='candle', style='yahoo', volume=axes[1], addplot=addplots)

                    current_price = df['Close'].iloc[-1]
                    axes[0].text(
                        1.01, current_price, f'<{current_price}>',
                        transform=axes[0].get_yaxis_transform(), color='red',
                        verticalalignment='center', bbox=dict(facecolor='white', edgecolor='none')
                    )
                    axes[0].legend(loc='upper left')
                    
                    x_raw = list(range(len(df['Date'])))
                    axes[2].fill_between(x_raw, df['RSI'], 30, where=(df['RSI'] <= 30), color='blue', alpha=0.5)
                    axes[2].fill_between(x_raw, df['RSI'], 70, where=(df['RSI'] >= 70), color='red', alpha=0.5)
                    axes[2].set_yticks([30, 50, 70])
                    axes[2].legend(loc='upper left')
                    axes[3].legend(loc='upper left')

                    x_labels = self.create_labels(filtered_data)
                    for index, ax in enumerate(axes):
                        ax.set_xlim(-2, len(filtered_data['D']) + 1)
                        ax.grid(True, axis='y')
                        ax.set_xticks(x_raw)
                        if index == 0 or index == 1 or index == 2:
                            ax.set_xticklabels([])
                        if index == 3:
                            ax.set_xticklabels(x_labels)
                    axes[1].yaxis.set_label_position("right")
                    axes[1].yaxis.tick_right()
                    axes[1].grid(False)

                if code in self.trader.starting_price:
                    starting_price_line = self.trader.starting_price[code]
                    y_min, y_max = axes[0].get_ylim()
                    if data_type == 'tick':
                        if y_min <= starting_price_line <= y_max:
                            axes[0].axhline(y=starting_price_line, color='orangered', linestyle='--', linewidth=1, label='Starting Price')
                    elif data_type == 'min':
                        if starting_price_line < y_min or starting_price_line > y_max:
                            axes[0].set_ylim(min(y_min, starting_price_line), max(y_max, starting_price_line))
                        axes[0].axhline(y=starting_price_line, color='orangered', linestyle='--', linewidth=1, label='Starting Price')

                if code in self.trader.buy_price:
                    buy_price_line = self.trader.buy_price[code]
                    axes[0].axhline(y=buy_price_line, color='red', linestyle='-', linewidth=1, label='Buy Price')

        except Exception as ex:
            logging.error(f"draw_chart_data -> {ex}\n{traceback.format_exc()}")

    def create_labels(self, filtered_data):
        labels = []
        current_hour = None
        mmm_0_added = False
        mmm_15_added = False
        mmm_30_added = False
        mmm_45_added = False

        for i in range(len(filtered_data['D'])):
            date = filtered_data['D'][i]
            time = filtered_data['T'][i]

            yy, mmdd = divmod(date, 10000)
            mm, dd = divmod(mmdd, 100)
            formatted_date = f"{mm:02}/{dd:02} "

            hhh, mmm = divmod(time, 100)
            formatted_time = f"{hhh:02}:{mmm:02}"

            if hhh != current_hour:
                current_hour = hhh
                mmm_0_added = False
                mmm_15_added = False
                mmm_30_added = False
                mmm_45_added = False

            label = ''

            if i == 0:
                label = f"{formatted_date}{formatted_time}"
            elif i == len(filtered_data['D']) - 1:
                label = f"{formatted_date}{formatted_time}"
            else:
                if mmm == 0:
                    if not mmm_0_added:
                        label = formatted_time
                        mmm_0_added = True
                elif mmm == 15:
                    if not mmm_15_added:
                        label = formatted_time
                        mmm_15_added = True
                elif mmm == 30:
                    if not mmm_30_added:
                        label = formatted_time
                        mmm_30_added = True
                elif mmm == 45:
                    if not mmm_45_added:
                        label = formatted_time
                        mmm_45_added = True

            labels.append(label if label else '')

        return labels

# ==================== LoginHandler ====================
class LoginHandler:
    def __init__(self, parent_window):
        self.parent = parent_window
        self.config = configparser.ConfigParser(interpolation=None)
        self.config_file = 'settings.ini'
        self.process = None
        self.slack = None
        self.slack_channel = '#stock'

    def load_settings(self):
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
            self.parent.loginEdit.setText(self.config.get('LOGIN', 'username', fallback=''))
            self.parent.passwordEdit.setText(self.config.get('LOGIN', 'password', fallback=''))
            self.parent.certpasswordEdit.setText(self.config.get('LOGIN', 'certpassword', fallback=''))
            self.parent.autoLoginCheckBox.setChecked(self.config.getboolean('LOGIN', 'autologin', fallback=False))

            self.parent.buycountEdit.setText(self.config.get('BUYCOUNT', 'target_buy_count', fallback='3'))

            if self.config.has_section('SLACK'):
                token = self.config.get('SLACK', 'token', fallback='')
                self.slack_channel = self.config.get('SLACK', 'channel', fallback='#stock')
                if token:
                    self.slack = Slacker(token)

    def save_settings(self):
        if not self.config.has_section('LOGIN'):
            self.config.add_section('LOGIN')
        self.config.set('LOGIN', 'username', self.parent.loginEdit.text())
        self.config.set('LOGIN', 'password', self.parent.passwordEdit.text())
        self.config.set('LOGIN', 'certpassword', self.parent.certpasswordEdit.text())
        self.config.set('LOGIN', 'autologin', str(self.parent.autoLoginCheckBox.isChecked()))
        with open(self.config_file, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def attempt_auto_login(self):
        if self.parent.autoLoginCheckBox.isChecked() and self.parent.loginEdit.text() and self.parent.passwordEdit.text() and self.parent.certpasswordEdit.text():
            self.handle_login()

    def handle_login(self):
        username = self.parent.loginEdit.text()
        password = self.parent.passwordEdit.text()
        certpassword = self.parent.certpasswordEdit.text()

        self.save_settings()
        self.parent.clean_up_processes()

        self.process = QProcess(self.parent)
        creon_path = 'C:\\CREON\\STARTER\\coStarter.exe'
        args = ["/prj:cp", f"/id:{username}", f"/pwd:{password}", f"/pwdcert:{certpassword}"]
        if self.parent.autoLoginCheckBox.isChecked():
            args.append('/autostart')
        self.process.start(creon_path, args)
        self.process.finished.connect(self.init_plus_check_and_continue)
       
    def buycount_setting(self):
        if not self.config.has_section('BUYCOUNT'):
            self.config.add_section('BUYCOUNT')
        self.config.set('BUYCOUNT', 'target_buy_count', self.parent.buycountEdit.text())

        with open(self.config_file, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

        logging.info(f"최대투자 종목수가 업데이트되었습니다.")

    def init_plus_check_and_continue(self):
        if not init_plus_check():
            exit()
        self.parent.post_login_setup()

    def auto_select_creon_popup(self):       
        try:
            button_x, button_y = 960, 500
            pyautogui.moveTo(button_x, button_y, duration=0.5)
            pyautogui.click()
            
            logging.info("모의투자 접속 버튼 클릭 완료")
        except Exception as e:
            logging.error(f"모의투자 접속 버튼 클릭 실패: {e}")

# ==================== MyWindow ====================
class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.is_loading_strategy = False
        self.market_close_emitted = False
        self.login_handler = LoginHandler(self)
        self.init_ui()
        self.login_handler.load_settings()
        self.login_handler.attempt_auto_login()
        self.update_chart_status_timer = None
        
        # 통합 전략 객체들
        self.momentum_scanner = None
        self.gap_scanner = None
        self.volatility_strategy = None

    def __del__(self):
        if hasattr(self, 'objstg'):
            self.objstg.Clear()

    def post_login_setup(self):
        logger = logging.getLogger()
        if not any(isinstance(handler, QTextEditLogger) for handler in logger.handlers):
            text_edit_logger = QTextEditLogger(self.terminalOutput)
            text_edit_logger.setLevel(logging.INFO)
            logger.addHandler(text_edit_logger)
        
        buycount = int(self.buycountEdit.text())
        self.trader = CTrader(cpTrade, cpBalance, cpCodeMgr, cpCash, cpOrder, cpStock, buycount, self)
        self.objstg = CpStrategy(self.trader)
        self.trader_thread = AutoTraderThread(self.trader, self)

        self.chartdrawer = ChartDrawer(self.fig, self.canvas, self.trader, self.trader_thread, self)

        self.code = ''
        self.stocks = []
        self.counter = 0

        self.trader.get_stock_balance('START', 'post_login_setup')
        logging.info(f"시작 시간 : {datetime.now().strftime('%m/%d %H:%M:%S')}")

        self.trader.stock_added_to_monitor.connect(self.on_stock_added)
        self.trader.stock_bought.connect(self.on_stock_bought)
        self.trader.stock_sold.connect(self.on_stock_sold)

        self.close_external_popup()         

        self.load_strategy()

        self.start_timers()
        self.trader_thread.buy_signal.connect(self.trader.buy_stock)
        self.trader_thread.sell_signal.connect(self.trader.sell_stock)
        self.trader_thread.sell_half_signal.connect(self.trader.sell_half_stock)
        self.trader_thread.sell_all_signal.connect(self.trader.sell_all)
        self.trader_thread.stock_removed_from_monitor.connect(self.on_stock_removed)
        self.trader_thread.counter_updated.connect(self.update_counter_label)
        self.trader_thread.stock_data_updated.connect(self.update_stock_table)
        
        # ✅ 봉 완성 signal 연결
        self.trader_thread.connect_bar_signals()

        self.trader_thread.start()

    def start_timers(self):
        """타이머 시작 (간소화)"""
        now = datetime.now()
        start_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
        end_time = now.replace(hour=15, minute=20, second=0, microsecond=0)
        today = datetime.today().weekday()

        if today == 5 or today == 6:
            logging.info(f"Today is {'Saturday.' if today == 5 else 'Sunday.'}")
            logging.info(f"오늘은 장이 쉽니다.")
            return
        
        if now < start_time:
            logging.info(f"자동 매매 시작 대기")
            self.trader.init_database()
            QTimer.singleShot(int((start_time - now).total_seconds() * 1000) + 1000, self.start_timers)

        elif start_time <= now < end_time:
            logging.info(f"자동 매매 시작")
            send_slack_message(self.login_handler, "#stock", f"자동 매매 시작")

            self.update_chart_status_timer = QTimer(self)
            self.update_chart_status_timer.timeout.connect(self.update_chart_status_label)
            self.update_chart_status_timer.start(2000)
            
            QTimer.singleShot(int((end_time - now).total_seconds() * 1000) + 1000, self.start_timers)
            
        elif end_time <= now and not self.market_close_emitted:
            # ===== 타이머 정리 (간소화) =====
            logging.info("=== 장 종료 처리 시작 ===")
                       
            # 차트 업데이트 타이머 정지
            if self.trader.tickdata is not None:
                self.trader.tickdata.update_data_timer.stop()
            if self.trader.mindata is not None:
                self.trader.mindata.update_data_timer.stop()
            if self.trader.daydata is not None:
                self.trader.daydata.update_data_timer.stop()
            if self.update_chart_status_timer is not None:
                self.update_chart_status_timer.stop()
            
            # 미보유 종목 정리
            for code in list(self.trader.monistock_set):
                if code not in self.trader.bought_set:
                    self.on_stock_removed(code)
                    self.trader.delete_list_db(code)
            
            for code in list(self.trader.vistock_set):
                if code not in self.trader.monistock_set and code not in self.trader.bought_set:
                    self.trader.delete_list_db(code)

            self.market_close_emitted = True
            logging.info(f"자동 매매 종료")
            send_slack_message(self.login_handler, "#stock", f"자동 매매 종료")

    def update_chart_status_label(self):
        if hasattr(self, 'chartdrawer') and self.chartdrawer.last_chart_update_time:
            chart_age = int(datetime.now().strftime("%H%M")) - self.chartdrawer.last_chart_update_time
            if chart_age < 2:
                chart_color = "green"
            else:
                chart_color = "red"
            self.chart_status_label.setText(f"Chart: {chart_age}m ago")
            self.chart_status_label.setStyleSheet(f"color: {chart_color}")
        else:
            self.chart_status_label.setText("Chart: None")
            self.chart_status_label.setStyleSheet("color: red")

    def update_stock_table(self, stock_data_list):
        """monistock_set 데이터를 기반으로 표 업데이트"""
        self.stock_table.setRowCount(len(stock_data_list))
        for row, stock_data in enumerate(stock_data_list):
            code = stock_data.get('code', '')
            current_price = stock_data.get('current_price', 0.0)
            upward_prob = stock_data.get('upward_probability', 0.0)
            buy_price = stock_data.get('buy_price', 0.0)
            quantity = stock_data.get('quantity', 0)
            profit_loss = (current_price - buy_price) * quantity
            return_pct = ((current_price - buy_price) / buy_price * 100) if buy_price != 0 else 0.0

            code_item = QTableWidgetItem(code)
            code_item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
            self.stock_table.setItem(row, 0, code_item)

            current_price_item = QTableWidgetItem(f"{current_price:,.0f}")
            current_price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.stock_table.setItem(row, 1, current_price_item)

            upward_prob_item = QTableWidgetItem(f"{upward_prob:.2f}")
            upward_prob_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.stock_table.setItem(row, 2, upward_prob_item)

            buy_price_item = QTableWidgetItem(f"{buy_price:,.0f}")
            buy_price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.stock_table.setItem(row, 3, buy_price_item)

            profit_loss_item = QTableWidgetItem(f"{profit_loss:,.0f}")
            profit_loss_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.stock_table.setItem(row, 4, profit_loss_item)
            profit_loss_item = self.stock_table.item(row, 4)
            if profit_loss > 0:
                profit_loss_item.setForeground(Qt.green)
            elif profit_loss < 0:
                profit_loss_item.setForeground(Qt.red)
            else:
                profit_loss_item.setForeground(Qt.black)

            return_item = QTableWidgetItem(f"{return_pct:.2f}")
            return_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.stock_table.setItem(row, 5, return_item)
            return_item = self.stock_table.item(row, 5)
            if return_pct > 0:
                return_item.setForeground(Qt.green)
            elif return_pct < 0:
                return_item.setForeground(Qt.red)
            else:
                return_item.setForeground(Qt.black)

    @pyqtSlot(str)
    def on_stock_added(self, code):
        if code not in [self.firstListBox.item(i).text() for i in range(self.firstListBox.count())]:
            self.firstListBox.addItem(code)

    @pyqtSlot(str)
    def on_stock_removed(self, code):
        if code in self.trader.vistock_set:
            self.trader.vistock_set.remove(code)
        if code in self.trader.monistock_set:
            self.trader.monistock_set.remove(code)
            self.trader.tickdata.monitor_stop(code)
            self.trader.mindata.monitor_stop(code)
            self.trader.daydata.monitor_stop(code)
            self.trader.delete_list_db(code)

        if code == self.chartdrawer.code:
            self.chartdrawer.set_code(None)

        for index in range(self.firstListBox.count()):
            item = self.firstListBox.item(index)
            if item and item.text() == code:
                self.firstListBox.takeItem(index)
                break
    
    @pyqtSlot(str)
    def on_stock_bought(self, code):
        if code not in [self.secondListBox.item(i).text() for i in range(self.secondListBox.count())]:
            self.secondListBox.addItem(code)

    @pyqtSlot(str)
    def on_stock_sold(self, code):
        for index in range(self.secondListBox.count()):
            item = self.secondListBox.item(index)
            if item and item.text() == code:
                self.secondListBox.takeItem(index)
                break

    @pyqtSlot(int)
    def update_counter_label(self, counter):
        self.counterlabel.setText(f"타이머: {counter}")

    def save_last_stg(self):
        if not self.login_handler.config.has_section('SETTINGS'):
            self.login_handler.config.add_section('SETTINGS')
        self.login_handler.config.set('SETTINGS', 'last_strategy', self.comboStg.currentText())
        with open(self.login_handler.config_file, 'w', encoding='utf-8') as configfile:
            self.login_handler.config.write(configfile)

    def validate_strategy_condition(self, condition, strategy_type='buy'):
        """전략 조건식 검증
        
        Args:
            condition: 검증할 조건식 문자열
            strategy_type: 'buy' 또는 'sell'
        
        Returns:
            (is_valid, message): (True/False, 메시지)
        """
        try:
            import ast
            
            # 빈 문자열 체크
            if not condition or not condition.strip():
                return False, "조건식이 비어있습니다"
            
            # 문법 검증
            try:
                tree = ast.parse(condition, mode='eval')
            except SyntaxError as e:
                return False, f"문법 오류: {e.msg} (라인 {e.lineno})"
            
            # 사용 가능한 변수 정의
            if strategy_type == 'buy':
                available_vars = {
                    # 틱 데이터
                    'MAT5', 'MAT20', 'MAT60', 'MAT120', 'C', 'VWAP', 
                    'RSIT', 'RSIT_SIGNAL', 'MACDT', 'MACDT_SIGNAL', 'OSCT',
                    'STOCHK', 'STOCHD', 'ATR', 'CCI',
                    'BB_UPPER', 'BB_MIDDLE', 'BB_LOWER', 'BB_POSITION', 'BB_BANDWIDTH',
                    
                    # 분봉 데이터
                    'MAM5', 'MAM10', 'MAM20', 'min_close',
                    'RSI', 'RSI_SIGNAL', 'MACD', 'MACD_SIGNAL', 'OSC',
                    
                    # 계산 변수
                    'strength', 'momentum_score', 'threshold',
                    'volatility_breakout', 'gap_hold',
                    
                    # 기타
                    'positive_candle', 'tick_close_price', 'min_close_price'
                }
            else:  # sell
                available_vars = {
                    # 기본 변수
                    'min_close', 'MAM5', 'MAM10',
                    'current_profit_pct', 'from_peak_pct', 'hold_minutes',
                    'code', 'osct_negative', 'after_market_close',
                    
                    # trader 객체 접근 (self.trader)
                    'self'
                }
            
            # 사용된 변수명 추출
            used_vars = set()
            for node in ast.walk(tree):
                if isinstance(node, ast.Name):
                    used_vars.add(node.id)
                elif isinstance(node, ast.Attribute):
                    # self.trader.sell_half_set 같은 경우
                    if isinstance(node.value, ast.Name):
                        used_vars.add(node.value.id)
            
            # 정의되지 않은 변수 체크
            undefined = used_vars - available_vars - {
                # Python 내장 함수 허용
                'True', 'False', 'None', 
                'min', 'max', 'abs', 'round', 'int', 'float', 'len', 'sum', 'all', 'any'
            }
            
            if undefined:
                return False, f"정의되지 않은 변수: {', '.join(sorted(undefined))}"
            
            # 위험한 함수 호출 체크
            dangerous_calls = {'eval', 'exec', 'compile', '__import__', 'open', 'file'}
            for node in ast.walk(tree):
                if isinstance(node, ast.Call):
                    if isinstance(node.func, ast.Name):
                        if node.func.id in dangerous_calls:
                            return False, f"위험한 함수 사용 금지: {node.func.id}"
            
            return True, "검증 성공"
            
        except Exception as ex:
            return False, f"검증 중 오류: {str(ex)}"

    def save_buystrategy(self):
        """매수 전략 저장 (검증 추가)"""
        try:
            investment_strategy = self.comboStg.currentText()
            buy_strategy = self.comboBuyStg.currentText()
            buy_key = self.comboBuyStg.currentData()
            strategy_content = self.buystgInputWidget.toPlainText()

            # ===== 여기에서 검증 =====
            is_valid, message = self.validate_strategy_condition(strategy_content, 'buy')
            if not is_valid:
                QMessageBox.warning(
                    self, 
                    "전략 검증 실패", 
                    f"매수 전략 '{buy_strategy}'의 조건식이 올바르지 않습니다.\n\n{message}"
                )
                return

            if not self.login_handler.config.has_section('STRATEGIES'):
                self.login_handler.config.add_section('STRATEGIES')

            if not self.login_handler.config.has_section(investment_strategy):
                self.login_handler.config.add_section(investment_strategy)

            strategy_data = {
                'name': buy_strategy,
                'content': strategy_content
            }
            strategy_json = json.dumps(strategy_data, ensure_ascii=False)
            self.login_handler.config.set(investment_strategy, str(buy_key), strategy_json)

            with open(self.login_handler.config_file, 'w', encoding='utf-8') as configfile:
                self.login_handler.config.write(configfile)

            existing_strategy = next((stg for stg in self.strategies[investment_strategy] if stg['name'] == buy_strategy), None)
            if existing_strategy:
                existing_strategy['content'] = strategy_content
                logging.info(f"매수전략 '{buy_strategy}'이(가) 업데이트되었습니다.")

            QMessageBox.information(self, "수정 완료", f"매수전략 '{buy_strategy}'이 수정되었습니다.")

        except Exception as ex:
            logging.error(f"save_buystrategy -> {ex}")
            QMessageBox.critical(self, "수정 실패", f"전략 수정 중 오류가 발생했습니다:\n{str(ex)}")

    def save_sellstrategy(self):
        try:
            investment_strategy = self.comboStg.currentText()
            sell_strategy = self.comboSellStg.currentText()
            sell_key = self.comboSellStg.currentData()
            strategy_content = self.sellstgInputWidget.toPlainText()

            # ===== 여기에서 검증 =====
            is_valid, message = self.validate_strategy_condition(strategy_content, 'sell')
            if not is_valid:
                QMessageBox.warning(
                    self, 
                    "전략 검증 실패", 
                    f"매도 전략 '{sell_strategy}'의 조건식이 올바르지 않습니다.\n\n{message}"
                )
                return

            if not self.login_handler.config.has_section('STRATEGIES'):
                self.login_handler.config.add_section('STRATEGIES')

            if not self.login_handler.config.has_section(investment_strategy):
                self.login_handler.config.add_section(investment_strategy)

            strategy_data = {
                'name': sell_strategy,
                'content': strategy_content
            }
            strategy_json = json.dumps(strategy_data, ensure_ascii=False)
            self.login_handler.config.set(investment_strategy, str(sell_key), strategy_json)

            with open(self.login_handler.config_file, 'w', encoding='utf-8') as configfile:
                self.login_handler.config.write(configfile)

            existing_strategy = next((stg for stg in self.strategies[investment_strategy] if stg['name'] == sell_strategy), None)
            if existing_strategy:
                existing_strategy['content'] = strategy_content
                logging.info(f"매도전략 '{sell_strategy}'이(가) 업데이트되었습니다.")

            QMessageBox.information(self, "수정 완료", f"매도전략 '{sell_strategy}'이 수정되었습니다.")

        except Exception as ex:
            logging.error(f"save_strategy -> {ex}")
            QMessageBox.critical(self, "수정 실패", f"전략 수정 중 오류가 발생했습니다:\n{str(ex)}")

    def load_strategy(self):
        try:
            self.dataStg = []
            self.data8537 = {}
            self.strategies = {}

            self.comboStg.clear()
            self.comboBuyStg.clear()
            self.buystgInputWidget.clear()

            if self.login_handler.config.has_section('STRATEGIES'):
                existing_stgnames = set(self.login_handler.config['STRATEGIES'].values())

            self.data8537 = self.objstg.requestList()
            for stgname, v in self.data8537.items():
                if stgname not in existing_stgnames:
                    existing_keys = self.login_handler.config['STRATEGIES'].keys()
                    existing_numbers = []
                    for k in existing_keys:
                        match = re.match(r'stg_?(\d+)', k)
                        if match:
                            existing_numbers.append(int(match.group(1)))
                    next_number = max(existing_numbers, default=0) + 1
                    new_key = f'stg_{next_number}'

                    self.login_handler.config.set('STRATEGIES', new_key, stgname)
                    existing_stgnames.add(stgname)
                    with open(self.login_handler.config_file, 'w', encoding='utf-8') as configfile:
                        self.login_handler.config.write(configfile)

            for investment_strategy in existing_stgnames:
                if self.login_handler.config.has_section(investment_strategy):
                    self.strategies[investment_strategy] = []
                    for buy_key in sorted([k for k in self.login_handler.config[investment_strategy] if k.startswith('buy_stg_')], key=lambda x: int(x.split('_')[-1])):
                        buy_strategy = json.loads(self.login_handler.config.get(investment_strategy, buy_key))
                        buy_strategy['key'] = buy_key
                        self.strategies[investment_strategy].append(buy_strategy)

                    for sell_key in sorted([k for k in self.login_handler.config[investment_strategy] if k.startswith('sell_stg_')], key=lambda x: int(x.split('_')[-1])):
                        sell_strategy = json.loads(self.login_handler.config.get(investment_strategy, sell_key))
                        sell_strategy['key'] = sell_key
                        self.strategies[investment_strategy].append(sell_strategy)

            # 통합 전략 추가
            if "통합 전략" not in existing_stgnames:
                self.login_handler.config.set('STRATEGIES', 'stg_integrated', "통합 전략")
                existing_stgnames.add("통합 전략")
                with open(self.login_handler.config_file, 'w', encoding='utf-8') as configfile:
                    self.login_handler.config.write(configfile)

            self.comboStg.blockSignals(True)
            for stgname in existing_stgnames:
                self.comboStg.addItem(stgname)
            
            last_strategy = self.login_handler.config.get('SETTINGS', 'last_strategy', fallback='VI 발동')
            index = self.comboStg.findText(last_strategy)
            if index != -1:
                self.comboStg.setCurrentIndex(index)
            self.comboStg.blockSignals(False)

            self.is_loading_strategy = True
            self.stgChanged()
            self.is_loading_strategy = False

        except Exception as ex:
            logging.error(f"load_strategy -> {ex}")

    def stgChanged(self, *args):
        stgName = self.comboStg.currentText()
        self.save_last_stg()

        if not self.is_loading_strategy:
            self.sell_all_item()
            self.trader.clear_list_db('mylist.db')
        
        # 기존 전략 객체 정리
        if hasattr(self, 'momentum_scanner') and self.momentum_scanner:
            self.momentum_scanner = None
        
        # ===== 전략별 초기화 =====
        if stgName == 'VI 발동':
            self.objstg.Clear()
            logging.info(f"전략 초기화: VI 발동")
            
            self.trader.init_stock_balance()
            self.trader.load_from_list_db('mylist.db')
            for code in list(self.trader.database_set):
                if code not in self.trader.monistock_set:
                    if self.trader.daydata.select_code(code) and self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                        if code not in self.trader.starting_time:
                            self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                        self.trader.monistock_set.add(code)
                        self.firstListBox.addItem(code)
                        self.trader.save_vi_data(code)
                    else:
                        self.trader.daydata.monitor_stop(code)
                        self.trader.mindata.monitor_stop(code)
                        self.trader.tickdata.monitor_stop(code)
            
            if not hasattr(self, 'pb9619'):
                self.pb9619 = CpPB9619()
                self.pb9619.Subscribe("", self.trader)

        elif stgName == "통합 전략":
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # ✅ 통합 전략: 조건검색 + 매수 조건
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            
            if hasattr(self, 'pb9619'):
                self.pb9619.Unsubscribe()
            self.objstg.Clear()
            
            logging.info(f"=== 통합 전략 시작 (조건검색 기반) ===")
            self.trader.init_stock_balance()
            
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 1️⃣ 종목 추출: 조건검색
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            
            # 검증용 스캐너 초기화
            self.objstg.momentum_scanner = MomentumScanner(self.trader)
            self.objstg.gap_scanner = GapUpScanner(self.trader)
            
            # 조건검색 1: 급등주
            momentum_stg = self.data8537.get("급등주")
            if momentum_stg:
                id = momentum_stg['ID']
                name = momentum_stg['전략명']
                
                ret, monid = self.objstg.requestMonitorID(id)
                if ret:
                    ret, status = self.objstg.requestStgControl(id, monid, True, name)
                    if ret:
                        logging.info(f"✅ 조건검색 '{name}' 감시 시작")
                    else:
                        logging.warning(f"조건검색 '{name}' 시작 실패")
                else:
                    logging.warning(f"조건검색 '{name}' 모니터 ID 획득 실패")
            else:
                logging.warning("조건검색 '급등주'를 찾을 수 없습니다. HTS에서 생성하세요.")
            
            # 조건검색 2: 갭상승
            gap_stg = self.data8537.get("갭상승")
            if gap_stg:
                id = gap_stg['ID']
                name = gap_stg['전략명']
                
                ret, monid = self.objstg.requestMonitorID(id)
                if ret:
                    ret, status = self.objstg.requestStgControl(id, monid, True, name)
                    if ret:
                        logging.info(f"✅ 조건검색 '{name}' 감시 시작")
                    else:
                        logging.warning(f"조건검색 '{name}' 시작 실패")
                else:
                    logging.warning(f"조건검색 '{name}' 모니터 ID 획득 실패")
            else:
                logging.warning("조건검색 '갭상승'을 찾을 수 없습니다. HTS에서 생성하세요.")
            
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 2️⃣ 매수 조건: Python 로직
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            
            # 변동성 돌파 전략 (매수 조건 - buy_stg_9)
            self.volatility_strategy = VolatilityBreakout(self.trader)
            self.trader_thread.set_volatility_strategy(self.volatility_strategy)
            logging.info("✅ 변동성 돌파 전략 활성화 (매수 조건)")
            
            # 갭 상승 유지 체커 (매수 조건 - buy_stg_10)
            self.gap_scanner = self.objstg.gap_scanner  # 동일 객체 사용
            logging.info("✅ 갭 상승 유지 체커 활성화 (매수 조건)")
            
            # 기존 데이터베이스에서 복원
            self.trader.load_from_list_db('mylist.db')
            for code in list(self.trader.database_set):
                if code not in self.trader.monistock_set:
                    if self.trader.daydata.select_code(code) and self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                        if code not in self.trader.starting_time:
                            self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                        self.trader.monistock_set.add(code)
                        self.firstListBox.addItem(code)
                    else:
                        self.trader.daydata.monitor_stop(code)
                        self.trader.mindata.monitor_stop(code)
                        self.trader.tickdata.monitor_stop(code)
            
            logging.info(f"=== 통합 전략 초기화 완료 ===")

        else:
            # 기타 사용자 정의 전략
            if hasattr(self, 'pb9619'):
                self.pb9619.Unsubscribe()

            logging.info(f"전략 초기화: {stgName}")
            self.trader.init_stock_balance()
            self.trader.load_from_list_db('mylist.db')
            for code in list(self.trader.database_set):
                if code not in self.trader.monistock_set:
                    if self.trader.daydata.select_code(code) and self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                        if code not in self.trader.starting_time:
                            self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                        self.trader.monistock_set.add(code)
                        self.firstListBox.addItem(code)
                        self.trader.save_vi_data(code)
                    else:
                        self.trader.daydata.monitor_stop(code)
                        self.trader.mindata.monitor_stop(code)
                        self.trader.tickdata.monitor_stop(code)
            
            item = self.data8537.get(stgName)
            if item:
                id = item['ID']
                name = item['전략명']
                if name != '급등주':
                    ret, self.dataStg = self.objstg.requestStgID(id)
                    if ret == False:
                        return
                    for s in self.dataStg:
                        if self.trader.daydata.select_code(s['code']) and self.trader.tickdata.monitor_code(s['code']) and self.trader.mindata.monitor_code(s['code']):
                            if s['code'] not in self.trader.starting_time:
                                self.trader.starting_time[s['code']] = datetime.now().strftime('%m/%d 09:00:00')
                            self.trader.monistock_set.add(s['code'])
                            self.firstListBox.addItem(s['code'])
                        else:
                            self.trader.daydata.monitor_stop(s['code'])
                            self.trader.mindata.monitor_stop(s['code'])
                            self.trader.tickdata.monitor_stop(s['code'])

                ret, monid = self.objstg.requestMonitorID(id)
                if ret == False:
                    return
                ret, status = self.objstg.requestStgControl(id, monid, True, stgName)
                if ret == False:
                    return
            
        logging.info(f"{stgName} 전략 감시 시작")

        self.comboBuyStg.clear()
        self.comboSellStg.clear()
        if stgName in self.strategies:
            for strategy in self.strategies[stgName]:
                strategy_name = strategy.get('name')
                strategy_key = strategy.get('key')
                if strategy_key.startswith('buy'):
                    self.comboBuyStg.addItem(strategy_name, strategy_key)
                elif strategy_key.startswith('sell'):
                    self.comboSellStg.addItem(strategy_name, strategy_key)

            if self.comboBuyStg.count() > 0:
                self.comboBuyStg.setCurrentIndex(0)
                self.buyStgChanged()

            if self.comboSellStg.count() > 0:
                self.comboSellStg.setCurrentIndex(0)
                self.sellStgChanged()

    @pyqtSlot(dict)
    def on_momentum_stock_found(self, stock):
        """급등주 발견 시 처리"""
        logging.info(f"급등주 발견: {stock['name']}({stock['code']}) - 점수: {stock['score']}")

    def buyStgChanged(self):
        try:
            investment_strategy = self.comboStg.currentText()
            buy_strategy = self.comboBuyStg.currentText()

            if not investment_strategy or not buy_strategy:
                self.buystgInputWidget.clear()
                return

            if investment_strategy in self.strategies:
                selected_strategy = next((stg for stg in self.strategies[investment_strategy] if stg['name'] == buy_strategy), None)
                if selected_strategy:
                    strategy_content = selected_strategy.get('content', '')
                    # ✅ 이스케이프된 \n을 실제 줄바꿈으로 변환
                    strategy_content = strategy_content.replace('\\n', '\n')
                    self.buystgInputWidget.setPlainText(strategy_content)
                else:
                    self.buystgInputWidget.clear()
            else:
                self.buystgInputWidget.clear()
        except Exception as ex:
            logging.error(f"buyStgChanged -> {ex}")

    def sellStgChanged(self):
        try:
            investment_strategy = self.comboStg.currentText()
            sell_strategy = self.comboSellStg.currentText()

            if not investment_strategy or not sell_strategy:
                self.sellstgInputWidget.clear()
                return

            if investment_strategy in self.strategies:
                selected_strategy = next((stg for stg in self.strategies[investment_strategy] if stg['name'] == sell_strategy), None)
                if selected_strategy:
                    strategy_content = selected_strategy.get('content', '')
                    # ✅ 이스케이프된 \n을 실제 줄바꿈으로 변환
                    strategy_content = strategy_content.replace('\\n', '\n')
                    self.sellstgInputWidget.setPlainText(strategy_content)
                else:
                    self.sellstgInputWidget.clear()
            else:
                self.sellstgInputWidget.clear()
        except Exception as ex:
            logging.error(f"sellStgChanged -> {ex}")

    def close_external_popup(self):
        try:
            windows = gw.getWindowsWithTitle('공지사항')
            for window in windows:
                window.close()
        except Exception as ex:
            logging.error(f"close_external_popup -> {ex}")

    def listBoxChanged(self, current):
        if current:
            self.chartdrawer.set_code(current.text())
        else:
            self.chartdrawer.set_code(None)

    def delete_select_item(self):
        selected_items = self.firstListBox.selectedItems()
        if selected_items:
            for item in selected_items:
                self.firstListBox.takeItem(self.firstListBox.row(item))
                code = item.text()
                if code in self.trader.monistock_set:
                    self.trader.monistock_set.remove(code)
                self.trader.tickdata.monitor_stop(code)
                self.trader.mindata.monitor_stop(code)
                self.trader.daydata.monitor_stop(code)
                self.trader.delete_list_db(code)

                if code == self.chartdrawer.code:
                    self.chartdrawer.set_code(None)
                logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> 투자 대상 종목 삭제")

    def buy_item(self):
        selected_items = self.firstListBox.selectedItems()
        if selected_items:
            for item in selected_items:
                self.trader.buy_stock(item.text(), '기타', '0', '03')

    def sell_item(self):
        selected_items = self.secondListBox.selectedItems()
        if selected_items:
            for item in selected_items:
                self.trader.sell_stock(item.text(), '직접 매도')

    def sell_all_item(self):
        if self.trader.sell_all():
            for code in list(self.trader.monistock_set):
                self.on_stock_removed(code)

    def print_chart(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        printer.setOrientation(QPrinter.Portrait)

        printDialog = QPrintDialog(printer, self)
        if printDialog.exec_() == QPrintDialog.Accepted:
            painter = QPainter(printer)
            pageRect = printer.pageRect()
            canvasSize = self.canvas.size()
            xScale = pageRect.width() / canvasSize.width()
            yScale = pageRect.height() / canvasSize.height()
            scale = min(xScale, yScale)
            painter.scale(scale, scale)
            self.canvas.render(painter)
            painter.end()

    def output_current_data(self):
        try:
            tick_data_all = self.trader.tickdata.stockdata
            min_data_all = self.trader.mindata.stockdata
            day_data_all = self.trader.daydata.stockdata

            if not tick_data_all and not min_data_all and not day_data_all:
                QMessageBox.warning(self, "데이터 없음", "현재 모니터링 중인 종목의 데이터가 없습니다.")
                return

            wb = Workbook()
            
            ws_tick = wb.active
            ws_tick.title = "Tick Data"
            for code, tick_data in tick_data_all.items():
                ws_tick.append([f"종목코드: {code}"])
                tick_keys = list(tick_data.keys())
                ws_tick.append(tick_keys)

                max_len_tick = max(len(v) for v in tick_data.values())

                for i in range(max_len_tick):
                    row = []
                    for key in tick_keys:
                        if i < len(tick_data[key]):
                            row.append(tick_data[key][i])
                        else:
                            row.append(None)
                    ws_tick.append(row)
                ws_tick.append([])

            ws_min = wb.create_sheet(title="Minute Data")
            for code, min_data in min_data_all.items():
                ws_min.append([f"종목코드: {code}"])
                min_keys = list(min_data.keys())
                ws_min.append(min_keys)

                max_len_min = max(len(v) for v in min_data.values())

                for i in range(max_len_min):
                    row = []
                    for key in min_keys:
                        if i < len(min_data[key]):
                            row.append(min_data[key][i])
                        else:
                            row.append(None)
                    ws_min.append(row)
                ws_min.append([])

            ws_day = wb.create_sheet(title="Day Data")
            for code, day_data in day_data_all.items():
                ws_day.append([f"종목코드: {code}"])
                day_keys = list(day_data.keys())
                ws_day.append(day_keys)

                max_len_day = max(len(v) for v in day_data.values())

                for i in range(max_len_day):
                    row = []
                    for key in day_keys:
                        if i < len(day_data[key]):
                            row.append(day_data[key][i])
                        else:
                            row.append(None)
                    ws_day.append(row)
                ws_day.append([])

            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            filename, _ = QFileDialog.getSaveFileName(self, "Save Excel File",
                                                    f"stock_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                                    "Excel Files (*.xlsx);;All Files (*)", options=options)
            if not filename:
                QMessageBox.warning(self, "저장 취소", "파일 저장이 취소되었습니다.")
                return

            wb.save(filename)
            QMessageBox.information(self, "성공", f"데이터가 '{filename}'에 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"데이터 저장 중 오류가 발생했습니다:\n{str(e)}")

    def print_terminal(self):
        printer = QPrinter()
        printDialog = QPrintDialog(printer, self)
        if printDialog.exec_() == QPrintDialog.Accepted:
            self.terminalOutput.print_(printer)

    def closeEvent(self, event):
        """프로그램 종료 처리"""
        reply = QMessageBox.question(self, 'Message', "Are you sure you want to quit?", 
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.save_last_stg()

            # 전략 객체 정리
            if self.momentum_scanner:
                self.momentum_scanner.stop_screening()            
            
            # 차트 업데이트 타이머 정지
            if getattr(self.trader, 'tickdata', None):
                self.trader.tickdata.update_data_timer.stop()
            if getattr(self.trader, 'mindata', None):
                self.trader.mindata.update_data_timer.stop()
            if getattr(self.trader, 'daydata', None):
                self.trader.daydata.update_data_timer.stop()
            if getattr(self, 'update_chart_status_timer', None):
                self.update_chart_status_timer.stop()

            if self.chartdrawer.chart_thread:
                self.chartdrawer.chart_thread.stop()

            if self.trader_thread:
                self.trader_thread.stop()

            for code in list(self.trader.monistock_set):
                self.trader.tickdata.monitor_stop(code)
                self.trader.mindata.monitor_stop(code)
                self.trader.daydata.monitor_stop(code)
            self.trader.monistock_set.clear()

            self.objstg.Clear()

            self.clean_up_processes()

            logger = logging.getLogger()
            for handler in logger.handlers[:]:
                handler.close()
                logger.removeHandler(handler)

            cpStatus.PlusDisconnect()

            event.accept()
        else:
            event.ignore()

    def clean_up_processes(self):
        process_names = ['coStarter', 'CpStart', 'DibServer']
        for process_name in process_names:
            for proc in psutil.process_iter(attrs=['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    proc.kill()
                    logging.warning(f"Killed  {proc.info['name']} (PID: {proc.info['pid']})")

    def init_ui(self):
        """UI 초기화 (탭 구조)"""
        self.setWindowTitle("초단타 매매 프로그램 v3.0 - 백테스팅")
        self.setGeometry(0, 0, 1900, 980)

        # ===== 메인 탭 위젯 생성 =====
        self.tab_widget = QTabWidget()
        
        # 탭 1: 실시간 매매
        self.trading_tab = QWidget()
        self.init_trading_tab()
        self.tab_widget.addTab(self.trading_tab, "실시간 매매")
        
        # 탭 2: 백테스팅
        self.backtest_tab = QWidget()
        self.init_backtest_tab()
        self.tab_widget.addTab(self.backtest_tab, "백테스팅")
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tab_widget)
        self.setLayout(main_layout)

    def init_trading_tab(self):
        """실시간 매매 탭 초기화"""
        
        # ===== 로그인 영역 =====
        loginLayout = QVBoxLayout()

        loginLabelLayout = QHBoxLayout()
        loginLabel = QLabel("아이디:")
        loginLabelLayout.addWidget(loginLabel)
        self.loginEdit = QLineEdit()
        loginLabelLayout.addWidget(self.loginEdit)

        passwordLabelLayout = QHBoxLayout()
        passwordLabel = QLabel("비밀번호:")
        passwordLabelLayout.addWidget(passwordLabel)
        self.passwordEdit = QLineEdit()
        self.passwordEdit.setEchoMode(QLineEdit.Password)
        passwordLabelLayout.addWidget(self.passwordEdit)

        certpasswordLabelLayout = QHBoxLayout()
        certpasswordLabel = QLabel("인증번호:")
        certpasswordLabelLayout.addWidget(certpasswordLabel)
        self.certpasswordEdit = QLineEdit()
        self.certpasswordEdit.setEchoMode(QLineEdit.Password)
        certpasswordLabelLayout.addWidget(self.certpasswordEdit)

        label_width = 70
        loginLabel.setFixedWidth(label_width)
        passwordLabel.setFixedWidth(label_width)
        certpasswordLabel.setFixedWidth(label_width)

        loginButtonLayout = QHBoxLayout()
        self.autoLoginCheckBox = QCheckBox("자동 로그인")
        loginButtonLayout.addWidget(self.autoLoginCheckBox)
        self.loginButton = QPushButton("로그인")
        loginButtonLayout.addWidget(self.loginButton)

        loginLayout.addLayout(loginLabelLayout)
        loginLayout.addLayout(passwordLabelLayout)
        loginLayout.addLayout(certpasswordLabelLayout)
        loginLayout.addLayout(loginButtonLayout)

        # ===== 투자 설정 =====
        buycountLayout = QHBoxLayout()
        buycountLabel = QLabel("최대투자 종목수 :")
        buycountLayout.addWidget(buycountLabel)
        self.buycountEdit = QLineEdit()
        buycountLayout.addWidget(self.buycountEdit)
        self.buycountButton = QPushButton("설정")
        self.buycountButton.setFixedWidth(70)
        buycountLayout.addWidget(self.buycountButton)

        # ===== 투자 대상 종목 리스트 =====
        firstListBoxLayout = QVBoxLayout()
        listBoxLabel = QLabel("투자 대상 종목 :")
        firstListBoxLayout.addWidget(listBoxLabel)
        self.firstListBox = QListWidget()
        firstListBoxLayout.addWidget(self.firstListBox, 1)
        firstButtonLayout = QHBoxLayout()
        self.buyButton = QPushButton("매입")
        firstButtonLayout.addWidget(self.buyButton)
        self.deleteFirstButton = QPushButton("삭제")        
        firstButtonLayout.addWidget(self.deleteFirstButton)        
        firstListBoxLayout.addLayout(firstButtonLayout)

        # ===== 투자 종목 리스트 =====
        secondListBoxLayout = QVBoxLayout()
        secondListBoxLabel = QLabel("투자 종목 :")
        secondListBoxLayout.addWidget(secondListBoxLabel)
        self.secondListBox = QListWidget()        
        secondListBoxLayout.addWidget(self.secondListBox, 1)
        secondButtonLayout = QHBoxLayout()
        self.sellButton = QPushButton("매도")
        secondButtonLayout.addWidget(self.sellButton)
        self.sellAllButton = QPushButton("전부 매도")
        secondButtonLayout.addWidget(self.sellAllButton)     
        secondListBoxLayout.addLayout(secondButtonLayout)

        # ===== 출력 버튼 =====
        printLayout = QHBoxLayout()
        self.printChartButton = QPushButton("차트 출력")
        printLayout.addWidget(self.printChartButton)
        self.dataOutputButton2 = QPushButton("차트데이터 저장")
        printLayout.addWidget(self.dataOutputButton2)

        # ===== 왼쪽 영역 통합 =====
        listBoxesLayout = QVBoxLayout()
        listBoxesLayout.addLayout(loginLayout)
        listBoxesLayout.addLayout(buycountLayout)
        listBoxesLayout.addLayout(firstListBoxLayout, 6)
        listBoxesLayout.addLayout(secondListBoxLayout, 4)
        listBoxesLayout.addLayout(printLayout)

        # ===== 차트 영역 =====
        chartLayout = QVBoxLayout()
        self.fig = Figure(figsize=(12, 8))
        self.canvas = FigureCanvas(self.fig)
        chartLayout.addWidget(self.canvas)

        # ===== 차트와 리스트 통합 =====
        chartAndListLayout = QHBoxLayout()
        chartAndListLayout.addLayout(listBoxesLayout, 1)
        chartAndListLayout.addLayout(chartLayout, 4)

        # ===== 전략 및 거래 정보 영역 =====
        strategyAndTradeLayout = QVBoxLayout()

        # 투자 전략
        strategyLayout = QHBoxLayout()
        strategyLabel = QLabel("투자전략:")
        strategyLabel.setFixedWidth(label_width)
        strategyLayout.addWidget(strategyLabel, Qt.AlignLeft)
        self.comboStg = QComboBox()
        self.comboStg.setFixedWidth(200)
        strategyLayout.addWidget(self.comboStg, Qt.AlignLeft)
        strategyLayout.addStretch()
        self.counterlabel = QLabel('타이머: 0')
        self.counterlabel.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        strategyLayout.addWidget(self.counterlabel)
        self.chart_status_label = QLabel("Chart: None")
        self.chart_status_label.setStyleSheet("color: red")
        self.chart_status_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        strategyLayout.addWidget(self.chart_status_label)

        # 매수 전략
        buyStrategyLayout = QHBoxLayout()
        buyStgLabel = QLabel("매수전략:")
        buyStgLabel.setFixedWidth(label_width)
        buyStrategyLayout.addWidget(buyStgLabel, alignment=Qt.AlignLeft)
        self.comboBuyStg = QComboBox()
        self.comboBuyStg.setFixedWidth(200)
        buyStrategyLayout.addWidget(self.comboBuyStg, alignment=Qt.AlignLeft)
        buyStrategyLayout.addStretch()
        self.saveBuyStgButton = QPushButton("수정")
        self.saveBuyStgButton.setFixedWidth(100)
        buyStrategyLayout.addWidget(self.saveBuyStgButton, alignment=Qt.AlignRight)
        self.buystgInputWidget = QTextEdit()
        self.buystgInputWidget.setPlaceholderText("매수전략의 내용을 입력하세요...")
        self.buystgInputWidget.setFixedHeight(80)

        # 매도 전략
        sellStrategyLayout = QHBoxLayout()
        sellStgLabel = QLabel("매도전략:")
        sellStgLabel.setFixedWidth(label_width)
        sellStrategyLayout.addWidget(sellStgLabel, alignment=Qt.AlignLeft)
        self.comboSellStg = QComboBox()
        self.comboSellStg.setFixedWidth(200)
        sellStrategyLayout.addWidget(self.comboSellStg, alignment=Qt.AlignLeft)
        sellStrategyLayout.addStretch()
        self.saveSellStgButton = QPushButton("수정")
        self.saveSellStgButton.setFixedWidth(100)
        sellStrategyLayout.addWidget(self.saveSellStgButton, alignment=Qt.AlignRight)
        self.sellstgInputWidget = QTextEdit()
        self.sellstgInputWidget.setPlaceholderText("매도전략의 내용을 입력하세요...")
        self.sellstgInputWidget.setFixedHeight(63)

        # 주식 현황 테이블
        self.stock_table = QTableWidget()
        self.stock_table.setRowCount(0)
        self.stock_table.setColumnCount(6)
        self.stock_table.setHorizontalHeaderLabels(["종목코드", "현재가", "상승확률(%)", "매수가", "평가손익", "수익률(%)"])
        self.stock_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.stock_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.stock_table.setFixedHeight(220)
        self.stock_table.verticalHeader().setDefaultSectionSize(20)

        strategyAndTradeLayout.addLayout(strategyLayout)
        strategyAndTradeLayout.addLayout(buyStrategyLayout)
        strategyAndTradeLayout.addWidget(self.buystgInputWidget)
        strategyAndTradeLayout.addLayout(sellStrategyLayout)
        strategyAndTradeLayout.addWidget(self.sellstgInputWidget)
        strategyAndTradeLayout.addWidget(self.stock_table)

        # ===== 터미널 출력 =====
        self.terminalOutput = QTextEdit()
        self.terminalOutput.setReadOnly(True)

        counterAndterminalLayout = QVBoxLayout()
        counterAndterminalLayout.addLayout(strategyAndTradeLayout)
        counterAndterminalLayout.addWidget(self.terminalOutput)

        # ===== 메인 레이아웃 =====
        mainLayout = QHBoxLayout()
        mainLayout.addLayout(chartAndListLayout, 70)
        mainLayout.addLayout(counterAndterminalLayout, 30)
        self.trading_tab.setLayout(mainLayout)

        # ===== 이벤트 연결 =====
        self.loginButton.clicked.connect(self.login_handler.handle_login)
        self.buycountButton.clicked.connect(self.login_handler.buycount_setting)

        self.buyButton.clicked.connect(self.buy_item)
        self.deleteFirstButton.clicked.connect(self.delete_select_item)
        self.sellButton.clicked.connect(self.sell_item)
        self.sellAllButton.clicked.connect(self.sell_all_item)

        self.firstListBox.currentItemChanged.connect(self.listBoxChanged)
        self.firstListBox.itemClicked.connect(self.listBoxChanged)
        self.secondListBox.currentItemChanged.connect(self.listBoxChanged)
        self.secondListBox.itemClicked.connect(self.listBoxChanged)

        self.printChartButton.clicked.connect(self.print_chart)
        self.dataOutputButton2.clicked.connect(self.output_current_data)

        self.comboStg.currentIndexChanged.connect(self.stgChanged)
        self.comboBuyStg.currentIndexChanged.connect(self.buyStgChanged)
        self.comboSellStg.currentIndexChanged.connect(self.sellStgChanged)
        self.saveBuyStgButton.clicked.connect(self.save_buystrategy)
        self.saveSellStgButton.clicked.connect(self.save_sellstrategy)

    def init_backtest_tab(self):
        """백테스팅 탭 초기화"""
        
        layout = QVBoxLayout()
        
        # ===== 설정 영역 =====
        settings_group = QGroupBox("백테스팅 설정")
        settings_layout = QGridLayout()
        
        # 기간 선택
        settings_layout.addWidget(QLabel("시작일:"), 0, 0)
        self.bt_start_date = QLineEdit()
        self.bt_start_date.setPlaceholderText("YYYYMMDD (예: 20250101)")
        self.bt_start_date.setFixedWidth(150)
        settings_layout.addWidget(self.bt_start_date, 0, 1)
        
        settings_layout.addWidget(QLabel("종료일:"), 0, 2)
        self.bt_end_date = QLineEdit()
        self.bt_end_date.setPlaceholderText("YYYYMMDD (예: 20250131)")
        self.bt_end_date.setFixedWidth(150)
        settings_layout.addWidget(self.bt_end_date, 0, 3)
        
        # 초기 자금
        settings_layout.addWidget(QLabel("초기 자금:"), 1, 0)
        self.bt_initial_cash = QLineEdit("10000000")
        self.bt_initial_cash.setFixedWidth(150)
        settings_layout.addWidget(self.bt_initial_cash, 1, 1)
        
        # 실행 버튼
        self.bt_run_button = QPushButton("백테스팅 실행")
        self.bt_run_button.setFixedWidth(150)
        self.bt_run_button.clicked.connect(self.run_backtest)
        settings_layout.addWidget(self.bt_run_button, 1, 2)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # ===== 결과 영역 =====
        results_splitter = QSplitter(Qt.Horizontal)
        
        # 왼쪽: 결과 요약
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        
        left_layout.addWidget(QLabel("백테스팅 결과:"))
        self.bt_results_text = QTextEdit()
        self.bt_results_text.setReadOnly(True)
        self.bt_results_text.setMaximumWidth(450)
        left_layout.addWidget(self.bt_results_text)
        
        left_widget.setLayout(left_layout)
        results_splitter.addWidget(left_widget)
        
        # 오른쪽: 차트
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        
        self.bt_fig = Figure(figsize=(10, 8))
        self.bt_canvas = FigureCanvas(self.bt_fig)
        right_layout.addWidget(self.bt_canvas)
        
        right_widget.setLayout(right_layout)
        results_splitter.addWidget(right_widget)
        
        results_splitter.setStretchFactor(0, 1)
        results_splitter.setStretchFactor(1, 2)
        
        layout.addWidget(results_splitter)
        
        self.backtest_tab.setLayout(layout)

    def run_backtest(self):
        """백테스팅 실행"""
        
        try:
            from backtester import Backtester
            
            start_date = self.bt_start_date.text()
            end_date = self.bt_end_date.text()
            initial_cash = int(self.bt_initial_cash.text())
            
            # 입력 검증
            if len(start_date) != 8 or len(end_date) != 8:
                QMessageBox.warning(self, "입력 오류", "날짜 형식: YYYYMMDD (예: 20250101)")
                return
            
            if not hasattr(self, 'trader'):
                QMessageBox.warning(self, "오류", "먼저 로그인해주세요.")
                return
            
            self.bt_results_text.clear()
            self.bt_results_text.append(f"백테스팅 시작: {start_date} ~ {end_date}")
            self.bt_results_text.append(f"초기 자금: {initial_cash:,}원\n")
            self.bt_results_text.append("처리 중...\n")
            
            QApplication.processEvents()
            
            # 백테스팅 실행
            bt = Backtester(
                db_path=self.trader.db_name,
                initial_cash=initial_cash
            )
            
            strategy_name = self.comboStg.currentText() if hasattr(self, 'comboStg') else '통합 전략'
            results = bt.run(start_date, end_date, strategy_name=strategy_name)
            
            # 결과 표시
            result_text = f"""
=== 백테스팅 결과 ===

【 기본 정보 】
전략: {results['strategy']}
기간: {results['start_date']} ~ {results['end_date']}

【 수익 성과 】
초기 자금: {results['initial_cash']:,}원
최종 자금: {results['final_cash']:,}원
총 수익: {results['total_profit']:,.0f}원
수익률: {results['total_return_pct']:.2f}%

【 거래 통계 】
총 거래: {results['total_trades']}회
승리: {results['win_trades']}회
패배: {results['lose_trades']}회
승률: {results['win_rate']:.1f}%

【 손익 분석 】
평균 수익률: {results['avg_profit_pct']:.2f}%
최대 수익: {results['max_profit_pct']:.2f}%
최대 손실: {results['max_loss_pct']:.2f}%
MDD (최대 낙폭): {results['mdd']:.2f}%

【 기타 지표 】
샤프 비율: {results['sharpe_ratio']:.2f}
평균 보유 시간: {results['avg_hold_minutes']:.0f}분

※ 백테스팅 결과는 참고용이며, 실제 매매 결과와 다를 수 있습니다.
"""
            
            self.bt_results_text.setPlainText(result_text)
            
            # 차트 그리기
            bt.plot_results(self.bt_fig)
            self.bt_canvas.draw()
            
            QMessageBox.information(self, "완료", "백테스팅이 완료되었습니다!")
            
        except FileNotFoundError:
            QMessageBox.critical(self, "오류", "backtester.py 파일을 찾을 수 없습니다.\n같은 폴더에 backtester.py가 있는지 확인하세요.")
        except Exception as ex:
            logging.error(f"run_backtest -> {ex}\n{traceback.format_exc()}")
            QMessageBox.critical(self, "오류", f"백테스팅 실패:\n{str(ex)}")

# ==================== QTextEditLogger ====================
class QTextEditLogger(QObject, logging.Handler):
    log_signal = pyqtSignal(str)

    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        self.log_signal.connect(self.append_log)

    @pyqtSlot(str)
    def append_log(self, msg):
        self.text_edit.append(msg)

    def emit(self, record):
        msg = self.format(record)
        if '매매이익' in msg:
            msg = f"<span style='color:green;'>{msg}</span>"
        elif '매매손실' in msg:
            msg = f"<span style='color:red;'>{msg}</span>"
        elif '매매실현손익' in msg:
            msg = f"<span style='font-weight:bold;'>{msg}</span>"
        else:
            msg = f"<span>{msg}</span>"

        self.log_signal.emit(msg)

# ==================== Main ====================
if __name__ == "__main__":
    # ✅ 실행 경로 설정
    if getattr(sys, 'frozen', False):
        # PyInstaller로 빌드된 경우
        application_path = os.path.dirname(sys.executable)
        os.chdir(application_path)
    
    try:
        # 로그 초기화
        setup_logging()
        logging.info("=" * 50)
        logging.info("=== 초단타 매매 프로그램 시작 ===")
        logging.info(f"실행 경로: {os.getcwd()}")
        logging.info(f"Python 버전: {sys.version}")
        logging.info("=" * 50)

        # QApplication 생성
        app = QApplication(sys.argv)
        
        # 폰트 설정
        try:
            app.setFont(QFont("Malgun Gothic", 9))
            logging.info("폰트 설정 완료")
        except Exception as ex:
            logging.warning(f"폰트 설정 실패: {ex}")
        
        # 메인 윈도우 생성
        logging.info("메인 윈도우 생성 중...")
        myWindow = MyWindow()
        
        # 아이콘 설정
        try:
            icon_path = 'stock_trader.ico'
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(application_path, 'stock_trader.ico')
            
            if os.path.exists(icon_path):
                myWindow.setWindowIcon(QIcon(icon_path))
                logging.info(f"아이콘 설정 완료: {icon_path}")
            else:
                logging.warning(f"아이콘 파일 없음: {icon_path}")
        except Exception as ex:
            logging.warning(f"아이콘 설정 실패: {ex}")
        
        # 창 표시
        myWindow.showMaximized()
        logging.info("GUI 초기화 완료")
        
        # 이벤트 루프 실행
        exit_code = app.exec_()
        logging.info(f"프로그램 종료 (exit code: {exit_code})")
        sys.exit(exit_code)
        
    except Exception as ex:
        # 최상위 예외 처리
        error_msg = (
            f"프로그램 실행 중 치명적 오류 발생:\n\n"
            f"{type(ex).__name__}: {ex}\n\n"
            f"상세 정보:\n{traceback.format_exc()}"
        )
        
        # 로그 파일에 기록
        try:
            logging.critical(error_msg)
        except:
            pass
        
        # 오류 파일 생성
        try:
            error_file = os.path.join(os.getcwd(), 'error.txt')
            with open(error_file, 'w', encoding='utf-8') as f:
                f.write(f"발생 시간: {datetime.now()}\n")
                f.write(f"실행 경로: {os.getcwd()}\n")
                f.write(f"Python: {sys.version}\n\n")
                f.write(error_msg)
            print(f"\n오류 정보가 저장되었습니다: {error_file}\n")
        except Exception as e:
            print(f"오류 파일 저장 실패: {e}")
        
        # 메시지 박스 표시
        try:
            from PyQt5.QtWidgets import QMessageBox, QApplication
            app = QApplication.instance()
            if app is None:
                app = QApplication(sys.argv)
            
            QMessageBox.critical(
                None, 
                "프로그램 오류", 
                f"프로그램 실행 중 오류가 발생했습니다.\n\n"
                f"{type(ex).__name__}: {ex}\n\n"
                f"자세한 내용은 error.txt 파일을 확인하세요."
            )
        except:
            # 메시지 박스도 실패하면 콘솔에 출력
            print("\n" + "=" * 60)
            print(error_msg)
            print("=" * 60)
            input("\nEnter 키를 눌러 종료하세요...")
        
        sys.exit(1)