import sys
import ctypes
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QTimer, pyqtSignal, QProcess, QObject, QThread, Qt, pyqtSlot, QRunnable, QThreadPool, QEventLoop
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

import pickle
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
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logging.getLogger('matplotlib').setLevel(logging.WARNING)

    log_dir = os.path.join(os.getcwd(), 'log')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_path = os.path.join(log_dir, f"trading_{datetime.now().strftime('%Y%m%d')}.log")
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

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
        """필요시 대기"""
        with self.lock:
            now = time.time()
            
            # 1분 내 15개 요청 체크
            while len(self.request_times) >= 15:
                oldest = self.request_times[0]
                if now - oldest < 60:
                    wait_time = 60 - (now - oldest) + 0.1
                    logging.debug(f"API 제한: {wait_time:.1f}초 대기")
                    time.sleep(wait_time)
                    now = time.time()
                else:
                    break
            
            # 초당 1회 체크
            if len(self.request_times) > 0:
                last_request = self.request_times[-1]
                if now - last_request < 1.0:
                    time.sleep(1.0 - (now - last_request))
            
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

# ==================== 급등주 스캐너 ====================
class MomentumScanner(QObject):
    """급등주 실시간 스캔"""
    
    stock_found = pyqtSignal(dict)
    
    def __init__(self, trader):
        super().__init__()
        self.trader = trader
        self.scan_timer = QTimer()
        self.scan_timer.timeout.connect(self.scan_market)
        self.candidate_codes = []
        self.scanned_codes = set()
        self.last_scan_time = 0
        self.cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
        
    def start_screening(self):
        """09:05부터 스캔 시작"""
        now = datetime.now()
        start_time = now.replace(hour=9, minute=5, second=0, microsecond=0)
        
        if now < start_time:
            delay = (start_time - now).total_seconds() * 1000
            logging.info(f"급등주 스캔 {int(delay/1000)}초 후 시작 예정")
            QTimer.singleShot(int(delay), self.start_screening)
            return
        
        logging.info("급등주 스캔 시작")
        self.scan_market()
        
        # 5분마다 재스캔
        self.scan_timer.start(300000)
    
    def stop_screening(self):
        """스캔 중지"""
        if self.scan_timer.isActive():
            self.scan_timer.stop()
        logging.info("급등주 스캔 중지")
    
    def scan_market(self):
        """시장 전체 스캔 (API 제한 고려)"""
        
        try:
            now = time.time()
            if now - self.last_scan_time < 60:
                logging.debug("스캔 주기 미달, 스킵")
                return
            
            self.last_scan_time = now
            
            logging.info("=== 급등주 스캔 시작 ===")
            
            # KOSDAQ 전체 종목
            code_list = cpCodeMgr.GetStockListByMarket(2)
            
            # 시가총액으로 사전 필터링 (API 절약)
            filtered_codes = self._pre_filter_by_market_cap(code_list)
            
            candidates = []
            scan_count = 0
            max_scan = 100  # 한 번에 최대 100개만 스캔
            
            for code in filtered_codes:
                if scan_count >= max_scan:
                    break
                
                # 이미 모니터링 중이면 스킵
                if code in self.trader.monistock_set:
                    continue
                
                # 이미 스캔했으면 스킵
                if code in self.scanned_codes:
                    continue
                
                # API 제한 대기
                api_limiter.wait_if_needed()
                
                # 모멘텀 점수 계산
                score = self._calculate_momentum_score(code)
                
                if score >= 70:
                    stock_data = {
                        'code': code,
                        'name': cpCodeMgr.CodeToName(code),
                        'score': score,
                        'price': self._get_cached_price(code),
                        'change_pct': self._get_change_pct(code),
                        'volume_ratio': self._get_volume_ratio(code)
                    }
                    candidates.append(stock_data)
                    self.scanned_codes.add(code)
                
                scan_count += 1
            
            # 점수 순 정렬
            candidates.sort(key=lambda x: x['score'], reverse=True)
            
            # 상위 5개만 모니터링 추가 (메모리 관리)
            max_monitoring = min(5, 10 - len(self.trader.monistock_set))
            
            for stock in candidates[:max_monitoring]:
                self._add_to_monitoring(stock)
            
            logging.info(
                f"급등주 스캔 완료: 검색 {scan_count}개, "
                f"후보 {len(candidates)}개, 추가 {min(max_monitoring, len(candidates))}개"
            )
            
        except Exception as ex:
            logging.error(f"scan_market 오류: {ex}\n{traceback.format_exc()}")
    
    def _pre_filter_by_market_cap(self, code_list):
        """시가총액으로 사전 필터링"""
        filtered = []
        
        for code in code_list:
            # 캐시 확인
            cached = stock_info_cache.get(f"filter_{code}")
            if cached is not None:
                if cached:
                    filtered.append(code)
                continue
            
            # 관리종목 제외
            if cpCodeMgr.GetStockControlKind(code) != 0:
                stock_info_cache.set(f"filter_{code}", False)
                continue
            
            # KOSDAQ만
            if cpCodeMgr.GetStockSectionKind(code) != 2:
                stock_info_cache.set(f"filter_{code}", False)
                continue
            
            # 시가총액 정보는 cpCodeMgr에서 바로 가져올 수 없으므로
            # 일단 통과시키고 나중에 상세 체크
            stock_info_cache.set(f"filter_{code}", True)
            filtered.append(code)
        
        return filtered[:200]  # 최대 200개로 제한
    
    def _calculate_momentum_score(self, code):
        """모멘텀 점수 계산"""
        
        try:
            score = 0
            
            # 캐시 확인
            cached_score = stock_info_cache.get(f"score_{code}")
            if cached_score is not None:
                return cached_score
            
            # 현재가 정보
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
            
            # 가격대 필터 (2,000원 ~ 50,000원)
            if current_price < 2000 or current_price > 50000:
                stock_info_cache.set(f"score_{code}", 0)
                return 0
            
            # 시가총액 필터 (500억 ~ 5000억)
            if market_cap < 50000 or market_cap > 500000:
                stock_info_cache.set(f"score_{code}", 0)
                return 0
            
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
                    stock_info_cache.set(f"score_{code}", 0)
                    return 0
            
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
                    stock_info_cache.set(f"score_{code}", 0)
                    return 0
            
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
            return score
            
        except Exception as ex:
            logging.error(f"_calculate_momentum_score({code}): {ex}")
            return 0
    
    def _get_cached_price(self, code):
        """현재가 조회 (캐시)"""
        cached = stock_info_cache.get(f"price_{code}")
        if cached:
            return cached
        
        try:
            self.cpStock.SetInputValue(0, code)
            self.cpStock.BlockRequest2(1)
            price = self.cpStock.GetHeaderValue(11)
            stock_info_cache.set(f"price_{code}", price)
            return price
        except:
            return 0
    
    def _get_change_pct(self, code):
        """등락률 조회"""
        try:
            self.cpStock.SetInputValue(0, code)
            self.cpStock.BlockRequest2(1)
            
            current = self.cpStock.GetHeaderValue(11)
            prev = self.cpStock.GetHeaderValue(20)
            
            if prev > 0:
                return (current - prev) / prev * 100
            return 0
        except:
            return 0
    
    def _get_volume_ratio(self, code):
        """거래량 비율 조회"""
        try:
            self.cpStock.SetInputValue(0, code)
            self.cpStock.BlockRequest2(1)
            
            volume = self.cpStock.GetHeaderValue(18)
            prev_volume = self.cpStock.GetHeaderValue(21)
            
            if prev_volume > 0:
                return volume / prev_volume
            return 0
        except:
            return 0
    
    def _add_to_monitoring(self, stock):
        """모니터링 대상 추가"""
        
        code = stock['code']
        
        try:
            # 일봉/틱봉/분봉 데이터 로드
            if (self.trader.daydata.select_code(code) and 
                self.trader.tickdata.monitor_code(code) and 
                self.trader.mindata.monitor_code(code)):
                
                # 시작 시간/가격 기록
                self.trader.starting_time[code] = datetime.now().strftime('%m/%d %H:%M:%S')
                self.trader.starting_price[code] = stock['price']
                
                # 모니터링 세트 추가
                self.trader.monistock_set.add(code)
                self.trader.stock_added_to_monitor.emit(code)
                
                logging.info(
                    f"{stock['name']}({code}) -> "
                    f"급등주 추가 (점수: {stock['score']}, "
                    f"상승률: {stock['change_pct']:.2f}%, "
                    f"거래량비: {stock['volume_ratio']:.1f}배)"
                )
                
                self.stock_found.emit(stock)
                
            else:
                logging.warning(f"{code}: 데이터 로드 실패")
                
        except Exception as ex:
            logging.error(f"_add_to_monitoring({code}): {ex}")

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

# ==================== 갭 상승 전략 ====================
class GapUpScanner:
    """갭 상승 종목 스캔"""
    
    def __init__(self, trader):
        self.trader = trader
        self.gap_stocks = []
        self.cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
    
    def scan_gap_up_stocks(self):
        """09:00 갭 상승 종목 스캔"""
        
        try:
            now = datetime.now()
            if now.hour != 9 or now.minute > 5:
                logging.warning("갭 상승 스캔은 09:00-09:05에만 가능")
                return
            
            logging.info("=== 갭 상승 스캔 시작 ===")
            
            code_list = cpCodeMgr.GetStockListByMarket(2)
            
            scan_count = 0
            max_scan = 150
            
            for code in code_list:
                if scan_count >= max_scan:
                    break
                
                # 이미 모니터링 중이면 스킵
                if code in self.trader.monistock_set:
                    continue
                
                # API 제한
                api_limiter.wait_if_needed()
                
                gap_pct = self._calculate_gap(code)
                
                # 갭 상승 1.5% ~ 4%
                if 1.5 <= gap_pct <= 4.0:
                    stock_data = {
                        'code': code,
                        'gap_pct': gap_pct,
                        'name': cpCodeMgr.CodeToName(code)
                    }
                    self.gap_stocks.append(stock_data)
                
                scan_count += 1
            
            # 모니터링 추가
            for stock in self.gap_stocks[:5]:
                self._add_to_monitoring(stock)
            
            logging.info(f"갭 상승 종목: {len(self.gap_stocks)}개, 추가: {min(5, len(self.gap_stocks))}개")
            
        except Exception as ex:
            logging.error(f"scan_gap_up_stocks: {ex}\n{traceback.format_exc()}")
    
    def _calculate_gap(self, code):
        """갭 계산"""
        
        try:
            # 일봉 데이터 로드
            if not self.trader.daydata.select_code(code):
                return 0
            
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('C', [])) < 2:
                return 0
            
            prev_close = day_data['C'][-2]
            today_open = day_data['O'][-1]
            
            if prev_close > 0:
                gap_pct = (today_open - prev_close) / prev_close * 100
                return gap_pct
            
            return 0
            
        except:
            return 0
    
    def check_gap_hold(self, code):
        """갭 유지 확인"""
        
        try:
            day_data = self.trader.daydata.stockdata.get(code, {})
            
            if len(day_data.get('O', [])) == 0:
                return False
            
            today_open = day_data['O'][-1]
            
            # 현재가
            tick_latest = self.trader.tickdata.get_latest_data(code)
            current_price = tick_latest.get('C', 0)
            
            # 시가 대비 -0.3% 이내 (갭 유지)
            if current_price >= today_open * 0.997:
                return True
            
            return False
            
        except:
            return False
    
    def _add_to_monitoring(self, stock):
        """모니터링 추가"""
        
        code = stock['code']
        
        try:
            if (self.trader.tickdata.monitor_code(code) and 
                self.trader.mindata.monitor_code(code)):
                
                self.trader.starting_time[code] = datetime.now().strftime('%m/%d %H:%M:%S')
                
                # 갭 상승가 기록
                day_data = self.trader.daydata.stockdata.get(code, {})
                if len(day_data.get('O', [])) > 0:
                    self.trader.starting_price[code] = day_data['O'][-1]
                
                self.trader.monistock_set.add(code)
                self.trader.stock_added_to_monitor.emit(code)
                
                logging.info(
                    f"{stock['name']}({code}) -> "
                    f"갭 상승 추가 (갭: {stock['gap_pct']:.2f}%)"
                )
                
        except Exception as ex:
            logging.error(f"_add_to_monitoring({code}): {ex}")

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

# ==================== 기존 CpStrategy (유지) ====================
class CpStrategy:
    def __init__(self, trader):
        self.monList = {}
        self.trader = trader
        self.stgname = {}
        self.objpb = CpPBCssAlert()

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
        if stgid not in self.monList:
            return
        if (stgmonid != self.monList[stgid]):
            return
        remain_time0 = cpStatus.GetLimitRemainTime(0)
        remain_time1 = cpStatus.GetLimitRemainTime(1)
        if remain_time0 != 0 or remain_time1 != 0:
            return
        if datetime.now() < datetime.now().replace(hour=9, minute=3, second=0, microsecond=0):
            return
        if code not in self.trader.monistock_set:
            if self.trader.daydata.select_code(code) and self.trader.tickdata.monitor_code(code) and self.trader.mindata.monitor_code(code):
                stgname = self.stgname.get(stgid, '')
                if stgname in ['급등주']:
                    self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                    self.trader.starting_price[code] = stgprice
                else:
                    if code not in self.trader.starting_time:
                        self.trader.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                self.trader.monistock_set.add(code)
                self.trader.stock_added_to_monitor.emit(code)
                self.trader.save_list_db(code, self.trader.starting_time[code], self.trader.starting_price[code], 1)
                logging.info(f"{cpCodeMgr.CodeToName(code)}({code}) -> 투자 대상 종목 추가")
            else:
                self.trader.daydata.monitor_stop(code)
                self.trader.mindata.monitor_stop(code)
                self.trader.tickdata.monitor_stop(code)

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
                'BB_STD': 2
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
                'BB_STD': 2
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
    
    def _get_default_result(self, indicator_type, length):
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
                'VWAP': 1
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

            else:
                logging.error(f"알 수 없는 지표 유형: {indicator_type}")
                return self._get_default_result(indicator_type, desired_length)

            return result

        except Exception as ex:
            logging.error(f"make_indicator -> {code}, {indicator_type}{self.chart_type} {ex}\n{traceback.format_exc()}")
            return self._get_default_result(indicator_type, len(chart_data.get('C', [])))

# ==================== CpData (체결강도 추가) ====================
class CpData(QObject):
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
        try:
            current_time = time.time()
            with self.stockdata_lock:
                codes = list(self.stockdata.keys())
            
            for code in codes:
                if (code in self.trader.vistock_set and code not in self.trader.monistock_set 
                    and code not in self.trader.bought_set):
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
                        results = [
                            self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                            for ind in ["MA", "MACD", "RSI", "STOCH", "ATR", "CCI", "BBANDS", "VWAP"]
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
        try:
            if code in self.stockdata:
                return True
            
            self.stockdata[code] = {
                'D': [], 'T': [], 'O': [], 'H': [], 'L': [], 'C': [], 'V': [], 'TV': [], 
                'MAT5': [], 'MAT20': [], 'MAT60': [], 'MAT120': [], 
                'MAM5': [], 'MAM10': [], 'MAM20': [], 'MACDT': [], 'MACDT_SIGNAL': [], 'OSCT': [], 
                'MACD': [], 'MACD_SIGNAL': [], 'OSC': [], 
                'RSIT': [], 'RSIT_SIGNAL': [], 'RSI': [], 'RSI_SIGNAL': [], 
                'STOCHK': [], 'STOCHD': [], 'ATR': [], 'CCI': [],  
                'BB_UPPER': [], 'BB_MIDDLE': [], 'BB_LOWER': [], 'BB_POSITION': [], 'BB_BANDWIDTH': [], 
                'TICKS': [], 'MAT5_MAT20_DIFF': [], 'MAT20_MAT60_DIFF': [], 'MAT60_MAT120_DIFF': [], 
                'C_MAT5_DIFF': [], 'MAM5_MAM10_DIFF': [], 'MAM10_MAM20_DIFF': [], 
                'C_MAM5_DIFF': [], 'C_ABOVE_MAM5': [], 'VWAP': [], 
                'MAT5_CHANGE': [], 'MAT20_CHANGE': [], 'MAT60_CHANGE': [], 'MAT120_CHANGE': []
            }
            
            # 체결강도 초기화
            self.buy_volumes[code] = deque(maxlen=10)
            self.sell_volumes[code] = deque(maxlen=10)

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
                    results = [
                        self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                        for ind in ["MA", "MACD", "RSI", "STOCH", "ATR", "CCI", "BBANDS", "VWAP"]
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
        """읽기 전용 스냅샷 업데이트 (락 내부에서 호출)"""
        try:
            if code not in self.stockdata:
                return
            
            data = self.stockdata[code]
            
            if self.chart_type == 'T':
                self.latest_snapshot[code] = {
                    'C': data.get('C', [0])[-1] if data.get('C') else 0,
                    'O': data.get('O', [0])[-1] if data.get('O') else 0,
                    'H': data.get('H', [0])[-1] if data.get('H') else 0,
                    'L': data.get('L', [0])[-1] if data.get('L') else 0,
                    'V': data.get('V', [0])[-1] if data.get('V') else 0,
                    'MAT5': data.get('MAT5', [0])[-1] if data.get('MAT5') else 0,
                    'MAT20': data.get('MAT20', [0])[-1] if data.get('MAT20') else 0,
                    'MAT60': data.get('MAT60', [0])[-1] if data.get('MAT60') else 0,
                    'MAT120': data.get('MAT120', [0])[-1] if data.get('MAT120') else 0,
                    'RSIT': data.get('RSIT', [0])[-1] if data.get('RSIT') else 0,
                    'RSIT_SIGNAL': data.get('RSIT_SIGNAL', [0])[-1] if data.get('RSIT_SIGNAL') else 0,
                    'MACDT': data.get('MACDT', [0])[-1] if data.get('MACDT') else 0,
                    'MACDT_SIGNAL': data.get('MACDT_SIGNAL', [0])[-1] if data.get('MACDT_SIGNAL') else 0,
                    'OSCT': data.get('OSCT', [0])[-1] if data.get('OSCT') else 0,
                    'STOCHK': data.get('STOCHK', [0])[-1] if data.get('STOCHK') else 0,
                    'STOCHD': data.get('STOCHD', [0])[-1] if data.get('STOCHD') else 0,
                    'ATR': data.get('ATR', [0])[-1] if data.get('ATR') else 0,
                    'CCI': data.get('CCI', [0])[-1] if data.get('CCI') else 0,
                    'BB_UPPER': data.get('BB_UPPER', [0])[-1] if data.get('BB_UPPER') else 0,
                    'BB_MIDDLE': data.get('BB_MIDDLE', [0])[-1] if data.get('BB_MIDDLE') else 0,
                    'BB_LOWER': data.get('BB_LOWER', [0])[-1] if data.get('BB_LOWER') else 0,
                    'BB_POSITION': data.get('BB_POSITION', [0])[-1] if data.get('BB_POSITION') else 0,
                    'BB_BANDWIDTH': data.get('BB_BANDWIDTH', [0])[-1] if data.get('BB_BANDWIDTH') else 0,
                    'VWAP': data.get('VWAP', [0])[-1] if data.get('VWAP') else 0,
                    'C_recent': data.get('C', [0])[-3:] if data.get('C') else [0, 0, 0],
                    'H_recent': data.get('H', [0])[-3:] if data.get('H') else [0, 0, 0],
                    'L_recent': data.get('L', [0])[-3:] if data.get('L') else [0, 0, 0],
                }
            
            elif self.chart_type == 'm':
                self.latest_snapshot[code] = {
                    'C': data.get('C', [0])[-1] if data.get('C') else 0,
                    'O': data.get('O', [0])[-1] if data.get('O') else 0,
                    'H': data.get('H', [0])[-1] if data.get('H') else 0,
                    'L': data.get('L', [0])[-1] if data.get('L') else 0,
                    'V': data.get('V', [0])[-1] if data.get('V') else 0,
                    'MAM5': data.get('MAM5', [0])[-1] if data.get('MAM5') else 0,
                    'MAM10': data.get('MAM10', [0])[-1] if data.get('MAM10') else 0,
                    'MAM20': data.get('MAM20', [0])[-1] if data.get('MAM20') else 0,
                    'RSI': data.get('RSI', [0])[-1] if data.get('RSI') else 0,
                    'RSI_SIGNAL': data.get('RSI_SIGNAL', [0])[-1] if data.get('RSI_SIGNAL') else 0,
                    'MACD': data.get('MACD', [0])[-1] if data.get('MACD') else 0,
                    'MACD_SIGNAL': data.get('MACD_SIGNAL', [0])[-1] if data.get('MACD') else 0,
                    'OSC': data.get('OSC', [0])[-1] if data.get('OSC') else 0,
                    'STOCHK': data.get('STOCHK', [0])[-1] if data.get('STOCHK') else 0,
                    'STOCHD': data.get('STOCHD', [0])[-1] if data.get('STOCHD') else 0,
                    'CCI': data.get('CCI', [0])[-1] if data.get('CCI') else 0,
                    'VWAP': data.get('VWAP', [0])[-1] if data.get('VWAP') else 0,
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
            
            # 체결강도 업데이트 (임시로 거래량 기반)
            # 실제로는 호가창 데이터가 필요하지만, 여기서는 간단히 구현
            with self.stockdata_lock:
                if code in self.buy_volumes:
                    # 상승시 매수로 간주
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
                            if 1 <= lastCount < 60:
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
                            self.stockdata[code]['D'].append(self.todayDate)
                            self.stockdata[code]['T'].append(lCurTime)
                            self.stockdata[code]['O'].append(cur)
                            self.stockdata[code]['H'].append(cur)
                            self.stockdata[code]['L'].append(cur)
                            self.stockdata[code]['C'].append(cur)
                            self.stockdata[code]['V'].append(vol)
                            self.stockdata[code]['TICKS'].append(1)

                        desired_length = 400
                        for key in self.stockdata[code]:
                            self.stockdata[code][key] = self.stockdata[code][key][-desired_length:]

                        self._update_snapshot(code)

                        last_update = self.last_indicator_update.get(code, 0)
                        if current_time - last_update >= self.indicator_update_interval:
                            if code in self.objIndicators:
                                results = [
                                    self.objIndicators[code].make_indicator(ind, code, self.stockdata[code])
                                    for ind in ["MA", "RSI", "MACD", "STOCH", "ATR", "CCI", "BBANDS", "VWAP"]
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
                
        except Exception as ex:
            logging.error(f"updateCurData -> {ex}")

# ==================== DatabaseWorker (유지) ====================
class DatabaseWorker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, db_name, queue, tickdata, mindata, parent=None):
        super().__init__(parent)
        self.db_name = db_name
        self.queue = queue
        self.tickdata = tickdata
        self.mindata = mindata
        self.running = True
        
        # 마지막 저장 인덱스
        self.last_saved_indices = {}  # {code: {'tick': idx, 'min': idx}}
        
        # ✅ 설정: 안전 마진 (마지막 N개 봉 재저장)
        self.overlap_count = 2  # 마지막 2개 봉 재저장

    def run(self):
        while self.running:
            try:
                item = self.queue.get(timeout=1)
                if item is None:
                    self.queue.task_done()
                    break
                
                code, tick_data_copy, min_data_copy, start_date, start_hhmm, incremental = item
                
                try:
                    self.save_vi_data(code, tick_data_copy, min_data_copy, start_date, start_hhmm, incremental)
                except Exception as ex:
                    error_msg = f"DB Worker 오류 ({code}): {ex}"
                    logging.error(error_msg)
                    self.error.emit(error_msg)
                finally:
                    self.queue.task_done()
            except queue.Empty:
                continue
            except Exception as ex:
                logging.error(f"Queue processing error: {ex}")
        self.finished.emit()

    def stop(self):
        self.running = False
        self.queue.put(None)

    def save_vi_data(self, code, tick_data, min_data, start_date, start_hhmm, incremental=True):
        """VI 데이터 저장 (오버랩 저장 지원)"""
        try:
            with sqlite3.connect(self.db_name, timeout=10) as conn:
                conn.execute("PRAGMA journal_mode=WAL")
                c = conn.cursor()
                c.execute("BEGIN TRANSACTION")

                # 인덱스 초기화
                if code not in self.last_saved_indices:
                    self.last_saved_indices[code] = {'tick': 0, 'min': 0}
                
                last_tick_idx = self.last_saved_indices[code]['tick']
                last_min_idx = self.last_saved_indices[code]['min']

                # ===== 틱 데이터 저장 =====
                if tick_data and tick_data["T"]:
                    dates, times = tick_data["D"], tick_data["T"]
                    values = [tick_data[key] for key in ["C", "V", "MAT5", "MAT20", "MAT60", "MAT120", "RSIT",
                            "RSIT_SIGNAL", "MACDT", "MACDT_SIGNAL", "OSCT", "STOCHK", "STOCHD", "ATR",
                            "CCI", "BB_UPPER", "BB_MIDDLE", "BB_LOWER", "BB_POSITION", "BB_BANDWIDTH",
                            "MAT5_MAT20_DIFF", "MAT20_MAT60_DIFF", "MAT60_MAT120_DIFF",
                            "C_MAT5_DIFF", "VWAP"]]
                    
                    # ✅ 증분 저장: 안전 마진 적용
                    if incremental:
                        # 마지막 N개 봉 뒤로 돌아가기
                        start_idx = max(0, last_tick_idx - self.overlap_count)
                        
                        new_data = [
                            (i, str(date), f"{int(time_val):04d}")
                            for i, (date, time_val) in enumerate(zip(dates, times))
                            if i >= start_idx  # ✅ 오버랩 적용
                            and len(str(date)) == 8 
                            and len(f"{int(time_val):04d}") == 4
                            and str(date) == start_date 
                            and start_hhmm <= f"{int(time_val):04d}" <= "1515"
                        ]
                        
                        overlap_info = f", 오버랩={self.overlap_count}" if start_idx < last_tick_idx else ""
                    else:
                        # 전체 저장
                        new_data = [
                            (i, str(date), f"{int(time_val):04d}")
                            for i, (date, time_val) in enumerate(zip(dates, times))
                            if len(str(date)) == 8 
                            and len(f"{int(time_val):04d}") == 4
                            and str(date) == start_date 
                            and start_hhmm <= f"{int(time_val):04d}" <= "1515"
                        ]
                        overlap_info = ""

                    if new_data:
                        unique_date_times = set((date, time) for _, date, time in new_data)
                        
                        # ✅ 기존 데이터 삭제 (재저장을 위해)
                        for date, time in unique_date_times:
                            c.execute("DELETE FROM tick_data WHERE code = ? AND date = ? AND time = ?", 
                                     (code, date, time))
                        
                        inserted_count = 0
                        updated_count = 0
                        
                        for date, time in unique_date_times:
                            sorted_data = sorted(
                                [(i, d, t) for i, d, t in new_data if d == date and t == time], 
                                key=lambda x: x[2]
                            )
                            
                            for seq, (i, _, _) in enumerate(sorted_data):
                                c.execute("""INSERT INTO tick_data
                                            (code, date, time, sequence, C, V, MAT5, MAT20, MAT60, MAT120, RSIT,
                                            RSIT_SIGNAL, MACDT, MACDT_SIGNAL, OSCT, STOCHK, STOCHD, ATR, CCI,
                                            BB_UPPER, BB_MIDDLE, BB_LOWER, BB_POSITION, BB_BANDWIDTH, MAT5_MAT20_DIFF,
                                            MAT20_MAT60_DIFF, MAT60_MAT120_DIFF, C_MAT5_DIFF, VWAP)
                                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                        (code, date, time, seq, *[val[i] for val in values]))
                                
                                # ✅ 신규 vs 업데이트 구분
                                if i < last_tick_idx:
                                    updated_count += 1
                                else:
                                    inserted_count += 1

                        if inserted_count or updated_count:
                            # ✅ 마지막 인덱스 업데이트 (가장 큰 인덱스 + 1)
                            max_idx = max(i for i, _, _ in new_data)
                            self.last_saved_indices[code]['tick'] = max_idx + 1
                            
                            mode = "증분" if incremental else "전체"
                            logging.debug(
                                f"{code}: 틱 데이터 저장 ({mode}) - "
                                f"신규 {inserted_count}개, 업데이트 {updated_count}개{overlap_info}"
                            )

                # ===== 분봉 데이터 저장 (동일한 방식) =====
                if min_data and min_data["T"]:
                    dates, times = min_data["D"], min_data["T"]
                    values = [min_data[key] for key in ["C", "V", "MAM5", "MAM10", "MAM20", "RSI", "RSI_SIGNAL",
                            "MACD", "MACD_SIGNAL", "OSC", "STOCHK", "STOCHD", "CCI",
                            "MAM5_MAM10_DIFF", "MAM10_MAM20_DIFF", "C_MAM5_DIFF", "C_ABOVE_MAM5", "VWAP"]]
                    
                    # ✅ 증분 저장: 안전 마진 적용
                    if incremental:
                        # 마지막 N개 봉 뒤로 돌아가기
                        start_idx = max(0, last_min_idx - self.overlap_count)
                        
                        new_indices = [
                            (i, str(date), f"{int(time_val):04d}")
                            for i, (date, time_val) in enumerate(zip(dates, times))
                            if i >= start_idx  # ✅ 오버랩 적용
                            and len(str(date)) == 8 
                            and f"{int(time_val):04d}"[2:4].isdigit() 
                            and int(f"{int(time_val):04d}"[2:4]) % 3 == 0
                            and str(date) == start_date 
                            and start_hhmm <= f"{int(time_val):04d}" <= "1515"
                        ]
                        
                        overlap_info = f", 오버랩={self.overlap_count}" if start_idx < last_min_idx else ""
                    else:
                        # 전체 저장
                        new_indices = [
                            (i, str(date), f"{int(time_val):04d}")
                            for i, (date, time_val) in enumerate(zip(dates, times))
                            if len(str(date)) == 8 
                            and f"{int(time_val):04d}"[2:4].isdigit() 
                            and int(f"{int(time_val):04d}"[2:4]) % 3 == 0
                            and str(date) == start_date 
                            and start_hhmm <= f"{int(time_val):04d}" <= "1515"
                        ]
                        overlap_info = ""

                    if new_indices:
                        unique_date_times = set((date, time) for _, date, time in new_indices)
                        
                        # ✅ 기존 데이터 삭제 (재저장을 위해)
                        for date, time in unique_date_times:
                            c.execute("DELETE FROM min_data WHERE code = ? AND date = ? AND time = ?", 
                                     (code, date, time))
                        
                        inserted_count = 0
                        updated_count = 0
                        
                        for i, date, time in new_indices:
                            c.execute("""INSERT INTO min_data
                                        (code, date, time, sequence, C, V, MAM5, MAM10, MAM20, RSI, RSI_SIGNAL,
                                        MACD, MACD_SIGNAL, OSC, STOCHK, STOCHD, CCI, MAM5_MAM10_DIFF,
                                        MAM10_MAM20_DIFF, C_MAM5_DIFF, C_ABOVE_MAM5, VWAP)
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                    (code, date, time, 0, *[val[i] for val in values]))
                            
                            # ✅ 신규 vs 업데이트 구분
                            if i < last_min_idx:
                                updated_count += 1
                            else:
                                inserted_count += 1

                        if inserted_count or updated_count:
                            # ✅ 마지막 인덱스 업데이트
                            max_idx = max(i for i, _, _ in new_indices)
                            self.last_saved_indices[code]['min'] = max_idx + 1
                            
                            mode = "증분" if incremental else "전체"
                            logging.debug(
                                f"{code}: 분 데이터 저장 ({mode}) - "
                                f"신규 {inserted_count}개, 업데이트 {updated_count}개{overlap_info}"
                            )

                conn.commit()

        except sqlite3.OperationalError as e:
            logging.error(f"{code}: DB 저장 오류 (OperationalError) - {e}")
            if conn:
                conn.rollback()
        except Exception as ex:
            logging.error(f"{code}: VI 데이터 저장 오류 - {ex}\n{traceback.format_exc()}")
            if conn:
                conn.rollback()

# ==================== CTrader (계속) ====================
class CTrader(QObject):
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
        self.tickdata = CpData(60, 'T', 400, self)

        self.last_saved_timestamp = {}
        self.db_name = 'vi_stock_data.db'

        # ✅ 설정 파일에서 저장 옵션 읽기
        config = configparser.ConfigParser(interpolation=None)
        if os.path.exists('settings.ini'):
            config.read('settings.ini', encoding='utf-8')
        
        self.save_interval = config.getint('DATA_SAVING', 'interval_seconds', fallback=300)
        self.force_save_on_close = config.getboolean('DATA_SAVING', 'force_save_on_close', fallback=True)
        self.incremental_save = config.getboolean('DATA_SAVING', 'incremental_save', fallback=True)
        overlap_count = config.getint('DATA_SAVING', 'overlap_count', fallback=2)
        
        logging.info(
            f"데이터 저장 설정: 간격={self.save_interval}초, "
            f"증분={self.incremental_save}, 오버랩={overlap_count}, "
            f"장마감저장={self.force_save_on_close}"
        )

        # DatabaseWorker 초기화
        self.db_queue = queue.Queue()
        self.db_thread = QThread()
        self.db_worker = DatabaseWorker(self.db_name, self.db_queue, self.tickdata, self.mindata)
        self.db_worker.overlap_count = overlap_count
        self.db_worker.moveToThread(self.db_thread)
        self.db_thread.started.connect(self.db_worker.run)
        self.db_worker.finished.connect(self.db_thread.quit)
        self.db_worker.error.connect(lambda msg: logging.error(msg))
        self.db_thread.start()

        # ✅ 저장 타이머 (설정값 사용)
        self.save_data_timer = QTimer()
        self.save_data_timer.timeout.connect(self.periodic_save_vi_data)
        self.save_data_timer.start(self.save_interval * 1000)

    def init_database(self):
        """데이터베이스 초기화"""
        try:
            if os.path.exists(self.db_name):
                os.remove(self.db_name)
                logging.debug(f"{self.db_name} 삭제 완료")

            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            # ==================== 차트 데이터 테이블 ====================
            
            # 틱 데이터
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS tick_data (
                    code TEXT, date TEXT, time TEXT, sequence INTEGER, C REAL, V INTEGER,
                    MAT5 REAL, MAT20 REAL, MAT60 REAL, MAT120 REAL, RSIT REAL, RSIT_SIGNAL REAL,
                    MACDT REAL, MACDT_SIGNAL REAL, OSCT REAL, STOCHK REAL, STOCHD REAL, 
                    ATR REAL, CCI REAL, BB_UPPER REAL, BB_MIDDLE REAL, BB_LOWER REAL, 
                    BB_POSITION REAL, BB_BANDWIDTH REAL, MAT5_MAT20_DIFF REAL, 
                    MAT20_MAT60_DIFF REAL, MAT60_MAT120_DIFF REAL, C_MAT5_DIFF REAL, VWAP REAL,
                    PRIMARY KEY (code, date, time, sequence)
                )
            ''')
            
            # 분봉 데이터
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS min_data (
                    code TEXT, date TEXT, time TEXT, sequence INTEGER, C REAL, V INTEGER,
                    MAM5 REAL, MAM10 REAL, MAM20 REAL, RSI REAL, RSI_SIGNAL REAL,
                    MACD REAL, MACD_SIGNAL REAL, OSC REAL, STOCHK REAL, STOCHD REAL,
                    CCI REAL, MAM5_MAM10_DIFF REAL, MAM10_MAM20_DIFF REAL,
                    C_MAM5_DIFF REAL, C_ABOVE_MAM5 REAL, VWAP REAL,
                    PRIMARY KEY (code, date, time, sequence)
                )
            ''')
            
            # ==================== 거래/분석 테이블 ====================
            
            # 실거래 기록
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS trades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT NOT NULL,
                    stock_name TEXT,
                    date TEXT NOT NULL,
                    time TEXT NOT NULL,
                    action TEXT NOT NULL,
                    price REAL NOT NULL,
                    quantity INTEGER NOT NULL,
                    amount REAL NOT NULL,
                    strategy TEXT,
                    buy_reason TEXT,
                    sell_reason TEXT,
                    buy_price REAL,
                    profit REAL,
                    profit_pct REAL,
                    hold_minutes REAL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_trades_date ON trades(date)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_trades_code ON trades(code)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_trades_action ON trades(action)')
            
            # 일별 요약
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
            
            # 백테스팅 결과
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
            
            # ==================== 데이터 정리 ====================
            
            # 중복 제거
            cursor.execute('''
                DELETE FROM tick_data
                WHERE rowid NOT IN (
                    SELECT MAX(rowid)
                    FROM tick_data
                    GROUP BY code, date, time, sequence
                )
            ''')
            
            cursor.execute('''
                DELETE FROM min_data
                WHERE rowid NOT IN (
                    SELECT MAX(rowid)
                    FROM min_data
                    GROUP BY code, date, time
                )
            ''')
            
            conn.commit()
            conn.close()
            
            logging.info("데이터베이스 초기화 완료")
            
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

    def save_vi_data(self, code, force_full_save=False):
        """VI 데이터 저장 (증분/전체 선택 가능)"""
        try:
            start_date, start_hhmm = None, None
            try:
                if code in self.starting_time:
                    current_year = datetime.now().year
                    time_str = self.starting_time[code]
                    start_dt = datetime.strptime(f"{current_year}/{time_str}", '%Y/%m/%d %H:%M:%S')
                    start_date = start_dt.strftime('%Y%m%d')
                    start_hhmm = start_dt.strftime('%H%M')
                else:
                    start_date = datetime.now().strftime('%Y%m%d')
                    start_hhmm = '0900'
            except ValueError as ve:
                logging.error(f"{code}: starting_time 형식 오류 - {self.starting_time.get(code, '없음')}: {ve}")
                start_date = datetime.now().strftime('%Y%m%d')
                start_hhmm = '0900'

            with self.tickdata.stockdata_lock:
                tick_data_copy = copy.deepcopy(self.tickdata.stockdata.get(code, {}))
            with self.mindata.stockdata_lock:
                min_data_copy = copy.deepcopy(self.mindata.stockdata.get(code, {}))

            # ✅ 증분 저장 여부 결정
            incremental = self.incremental_save and not force_full_save
            
            self.db_queue.put((code, tick_data_copy, min_data_copy, start_date, start_hhmm, incremental))
            
        except Exception as ex:
            logging.error(f"save_vi_data -> {code}, {ex}\n{traceback.format_exc()}")

    def periodic_save_vi_data(self):
        """주기적 데이터 저장"""
        try:
            with self.tickdata.stockdata_lock:
                codes = list(self.tickdata.stockdata.keys())
            
            for code in codes:
                self.save_vi_data(code, force_full_save=False)  # 증분 저장
                
        except Exception as ex:
            logging.error(f"periodic_save_vi_data -> {ex}\n{traceback.format_exc()}")

    def force_save_all_data(self):
        """강제 전체 저장 (장마감 시)"""
        try:
            logging.info("=== 장마감 데이터 강제 저장 시작 ===")
            
            with self.tickdata.stockdata_lock:
                codes = list(self.tickdata.stockdata.keys())
            
            for code in codes:
                self.save_vi_data(code, force_full_save=True)  # 전체 저장
            
            # 큐가 비워질 때까지 대기 (최대 30초)
            max_wait = 30
            wait_time = 0
            while not self.db_queue.empty() and wait_time < max_wait:
                time.sleep(0.5)
                wait_time += 0.5
            
            logging.info(f"=== 장마감 데이터 저장 완료 ({len(codes)}개 종목) ===")
            
        except Exception as ex:
            logging.error(f"force_save_all_data -> {ex}\n{traceback.format_exc()}")

    def get_stock_balance(self, code, func):
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
        if code not in self.highest_price:
            self.highest_price[code] = current_price
        elif current_price > self.highest_price[code]:
            self.highest_price[code] = current_price    

    def monitor_vi(self, time, code, event2):
        """VI 발동 모니터링 (VI 전략 전용)"""
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
            self.save_vi_data(code)

        except Exception as ex:
            logging.error(f"monitor_vi -> {code}, {ex}\n{traceback.format_exc()}")

    def download_vi(self):
        """VI 발동 전날 데이터 다운로드"""
        try:
            date = datetime.today() - timedelta(days=1)
            while True:
                date_str = date.strftime('%Y%m%d')

                otp_url = 'http://data.krx.co.kr/comm/fileDn/GenerateOTP/generate.cmd'
                download_url = 'http://data.krx.co.kr/comm/fileDn/download_excel/download.cmd'

                query_str_params = {
                    'locale': 'ko_KR',
                    'mktId': 'ALL',
                    'inqTpCd1': '01',
                    'viKindCd': '1',
                    'tboxisuCd_finder_stkisu1_0': '전체',
                    'isuCd': 'ALL',
                    'isuCd2': 'ALL',
                    'param1isuCd_finder_stkisu1_0': 'ALL',
                    'prcDetailView': '1',
                    'share': '1',
                    'money': '1',
                    'strtDd': date_str,
                    'endDd': date_str,
                    'csvxls_isNo': 'true',
                    'name': 'fileDown',
                    'url': 'dbms/MDC/STAT/issue/MDCSTAT22401'
                }

                headers = {
                    'Referer': 'http://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC02021501',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'
                }

                with requests.Session() as session:
                    r = session.get(otp_url, params=query_str_params, headers=headers)
                    otp_code = r.content
                    r = session.post(download_url, data={'code': otp_code}, headers=headers)
                    df_vi = pd.read_excel(BytesIO(r.content), engine='openpyxl')

                if df_vi['발동가격_가격'].apply(pd.to_numeric, errors='coerce').notnull().all():
                    break
                date -= timedelta(days=1)

            df_vi = df_vi[df_vi['발동가격_괴리율'] > 0]
            df_vi = df_vi[df_vi['거래량'] > 1000000]
            df_vi = df_vi[df_vi['현재가'] > df_vi['시가']]
            df_vi = df_vi[df_vi['현재가'] > df_vi['발동가격_가격']]
            df_vi = df_vi[df_vi['종목코드'].str.match(r'^\d+$')]
            df_vi['종목코드'] = df_vi['종목코드'].apply(lambda x: f"A{int(x):06d}")
            idx = df_vi.groupby('종목코드')['발동가격_가격'].idxmax()
            df_vi = df_vi.loc[idx].reset_index(drop=True)
            vi_list = df_vi['종목코드'].tolist()

            for code in vi_list:
                if self.daydata.select_code(code):
                    if len(self.daydata.stockdata[code].get('MAD5', [])) > 0 and len(self.daydata.stockdata[code].get('MAD10', [])) > 0:
                        if (self.daydata.stockdata[code]['MAD5'][-1] > self.daydata.stockdata[code]['MAD10'][-1]):
                            if (self.daydata.stockdata[code]['O'][-1] > self.daydata.stockdata[code]['C'][-2] * 0.99):
                                if code not in self.monistock_set:
                                    if self.tickdata.monitor_code(code) == True and self.mindata.monitor_code(code) == True:
                                        if code not in self.starting_time:
                                            self.starting_time[code] = datetime.now().strftime('%m/%d 09:00:00')
                                        self.monistock_set.add(code)
                                        self.stock_added_to_monitor.emit(code)
                                    else:
                                        self.daydata.monitor_stop(code)
                                        self.mindata.monitor_stop(code)
                                        self.tickdata.monitor_stop(code)
                            else:
                                self.daydata.monitor_stop(code)
                        else:
                            self.daydata.monitor_stop(code)
                    else:
                        self.daydata.monitor_stop(code)
                else:
                    self.daydata.monitor_stop(code)
        except Exception as ex:
            logging.error(f"download_vi -> {ex}")

    def save_list_db(self, code, starting_time, starting_price, is_moni=0, db_file='mylist.db'):
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
        c.execute('INSERT OR REPLACE INTO items (codes, starting_time, starting_price, is_moni) VALUES (?, ?, ?, ?)', (code, starting_time, starting_price, is_moni))
        conn.commit()
        conn.close()

    def delete_list_db(self, code, db_file='mylist.db'):
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
        """주문 체결 모니터링 (거래 기록 저장 추가)"""
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
                    
                    # ✅ 매도 거래 기록 저장
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
                    
                    # ✅ 매수 거래 기록 저장
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
    """자동매매 스레드 - 통합 전략"""
    
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
        
        # 변동성 돌파 전략 추가
        self.volatility_strategy = None

    def set_volatility_strategy(self, strategy):
        """변동성 돌파 전략 설정"""
        self.volatility_strategy = strategy

    def run(self):
        """메인 루프"""
        while self.running:
            self.autotrade()
            self.msleep(1000)

    def stop(self):
        """스레드 정지"""
        logging.info("AutoTraderThread 정지 시작...")
        self.running = False
        
        self.quit()
        self.wait()
        logging.info("AutoTraderThread 정지 완료")

    def autotrade(self):
        """자동매매 메인 루프"""
        try:
            t_now = datetime.now()
            
            self.counter += 1
            self.counter_updated.emit(self.counter)
            
            self._update_stock_data_table()
            
            if self._is_trading_hours(t_now):
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
        """거래 로직 실행"""
        current_strategy = self.window.comboStg.currentText()
        buy_strategies = [
            stg for stg in self.window.strategies.get(current_strategy, []) 
            if stg['key'].startswith('buy')
        ]
        sell_strategies = [
            stg for stg in self.window.strategies.get(current_strategy, []) 
            if stg['key'].startswith('sell')
        ]
        
        monitored_codes = list(self.trader.monistock_set.copy())
        
        for code in monitored_codes:
            if code not in self.trader.monistock_set:
                continue
            
            try:
                if code not in self.trader.buyorder_set and code not in self.trader.bought_set:
                    self._evaluate_buy_condition(code, t_now, current_strategy, buy_strategies)
                
                elif (code in self.trader.bought_set and 
                      code not in self.trader.buyorder_set and 
                      code not in self.trader.sellorder_set):
                    self._evaluate_sell_condition(code, t_now, current_strategy, sell_strategies)
                    
            except Exception as ex:
                logging.error(f"{code} 거래 로직 오류: {ex}")

    def _handle_market_close(self):
        """장 종료 처리 (강제 저장 추가)"""
        
        if self.trader.buyorder_set or self.trader.sellorder_set:
            return
        
        # ✅ 장마감 강제 저장
        if self.trader.force_save_on_close:
            self.trader.force_save_all_data()
        
        # 보유 주식 전부 매도
        if self.trader.bought_set:
            logging.info("보유 주식 전부 매도")
            self.sell_all_signal.emit()
        
        self.sell_all_emitted = True

    # ===== 헬퍼 함수들 =====
    
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
        """통합 전략 매수 평가 (텍스트 기반)"""
        
        # ===== 기본 변수들 =====
        MAT5 = tick_latest.get('MAT5', 0)
        MAT20 = tick_latest.get('MAT20', 0)
        MAT60 = tick_latest.get('MAT60', 0)
        MAT120 = tick_latest.get('MAT120', 0)
        C = tick_latest.get('C', 0)
        VWAP = tick_latest.get('VWAP', 0)
        RSIT = tick_latest.get('RSIT', 50)
        
        MAM5 = min_latest.get('MAM5', 0)
        MAM10 = min_latest.get('MAM10', 0)
        MAM20 = min_latest.get('MAM20', 0)
        
        # ===== 계산된 값들 (기존 클래스 재사용) =====
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
        
        # ===== 전략 평가 =====
        for strategy in buy_strategies:
            try:
                condition = strategy.get('content', '')
                
                # eval()로 조건 평가
                if eval(condition):
                    buy_reason = strategy.get('name', '통합 전략')
                    logging.info(
                        f"{cpCodeMgr.CodeToName(code)}({code}): {buy_reason} 매수 "
                        f"(체결강도: {strength:.0f}, 점수: {momentum_score})"
                    )
                    self.buy_signal.emit(code, buy_reason, "0", "03")
                    return True
                    
            except Exception as ex:
                logging.error(f"{code} 통합 전략 매수 평가 오류: {ex}\n{traceback.format_exc()}")
        
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
        """통합 전략 매도 평가 (텍스트 기반)"""
        
        # ===== 기본 변수들 =====
        min_close = min_latest.get('C', 0)
        MAM5 = min_latest.get('MAM5', 0)
        MAM10 = min_latest.get('MAM10', 0)
        
        # ===== 계산된 값들 =====
        osct_negative = False
        tick_recent = self.trader.tickdata.get_recent_data(code, 3)
        OSCT_recent = tick_recent.get('OSCT', [0, 0, 0])
        if len(OSCT_recent) >= 2:
            osct_negative = OSCT_recent[-2] < 0 and OSCT_recent[-1] < 0
        
        after_market_close = self.is_after_time(14, 45)
        
        # ===== 전략 평가 =====
        for strategy in sell_strategies:
            try:
                condition = strategy.get('content', '')
                
                # eval()로 조건 평가
                if eval(condition):
                    sell_reason = strategy.get('name', '통합 전략')
                    logging.info(f"{cpCodeMgr.CodeToName(code)}({code}): {sell_reason} ({current_profit_pct:.2f}%)")
                    
                    # 분할 매도 vs 전량 매도
                    if '분할' in sell_reason:
                        self.sell_half_signal.emit(code, sell_reason)
                    else:
                        self.sell_signal.emit(code, sell_reason)
                    return True
                    
            except Exception as ex:
                logging.error(f"{code} 통합 전략 매도 평가 오류: {ex}\n{traceback.format_exc()}")
        
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
                    filtered_data = {key: [x for x in (chart_data[key][-60:] if len(chart_data[key]) >= 60 else chart_data[key][:])] for key in keys_to_keep}

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
        self.trader_thread.start()

    def start_timers(self):
        """타이머 시작 (장마감 처리 개선)"""
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
            # ✅ 타이머 정지 전에 먼저 데이터 저장
            logging.info("=== 장 종료 처리 시작 ===")
            
            # 데이터 저장 타이머 정지
            if self.trader.save_data_timer is not None:
                self.trader.save_data_timer.stop()
                logging.info("데이터 저장 타이머 정지")
            
            # 차트 업데이트 타이머 정지
            if self.trader.tickdata is not None:
                self.trader.tickdata.update_data_timer.stop()
            if self.trader.mindata is not None:
                self.trader.mindata.update_data_timer.stop()
            if self.trader.daydata is not None:
                self.trader.daydata.update_data_timer.stop()
            if self.update_chart_status_timer is not None:
                self.update_chart_status_timer.stop()
            
            # ✅ 장마감 강제 저장은 AutoTraderThread에서 처리
            # (매도 주문 완료 후 실행되도록)
            
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

    def save_buystrategy(self):
        try:
            investment_strategy = self.comboStg.currentText()
            buy_strategy = self.comboBuyStg.currentText()
            buy_key = self.comboBuyStg.currentData()
            strategy_content = self.buystgInputWidget.toPlainText()

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
            logging.error(f"save_strategy -> {ex}")
            QMessageBox.critical(self, "수정 실패", f"전략 수정 중 오류가 발생했습니다:\n{str(ex)}")

    def save_sellstrategy(self):
        try:
            investment_strategy = self.comboStg.currentText()
            sell_strategy = self.comboSellStg.currentText()
            sell_key = self.comboSellStg.currentData()
            strategy_content = self.sellstgInputWidget.toPlainText()

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
        if self.momentum_scanner:
            self.momentum_scanner.stop_screening()
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

        elif stgName == 'VI 발동 D1':
            if hasattr(self, 'pb9619'):
                self.pb9619.Unsubscribe()
            self.objstg.Clear()
            
            logging.info(f"전략 초기화: VI 발동 D1")
            self.trader.init_stock_balance()
            self.trader.download_vi()

        elif stgName == "통합 전략":
            # 통합 전략 초기화
            if hasattr(self, 'pb9619'):
                self.pb9619.Unsubscribe()
            self.objstg.Clear()
            
            logging.info(f"=== 통합 전략 시작 ===")
            self.trader.init_stock_balance()
            
            # 1. 급등주 스캐너
            self.momentum_scanner = MomentumScanner(self.trader)
            self.momentum_scanner.stock_found.connect(self.on_momentum_stock_found)
            self.momentum_scanner.start_screening()
            logging.info("급등주 스캐너 시작")
            
            # 2. 갭 상승 스캐너
            self.gap_scanner = GapUpScanner(self.trader)
            now = datetime.now()
            scan_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
            if now < scan_time:
                delay = (scan_time - now).total_seconds() * 1000
                QTimer.singleShot(int(delay), self.gap_scanner.scan_gap_up_stocks)
                logging.info(f"갭 상승 스캔 {int(delay/1000)}초 후 시작 예정")
            elif now.hour == 9 and now.minute < 10:
                self.gap_scanner.scan_gap_up_stocks()
            
            # 3. 변동성 돌파 전략
            self.volatility_strategy = VolatilityBreakout(self.trader)
            self.trader_thread.set_volatility_strategy(self.volatility_strategy)
            logging.info("변동성 돌파 전략 활성화")
            
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
        reply = QMessageBox.question(self, 'Message', "Are you sure you want to quit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.save_last_stg()

            # 전략 객체 정리
            if self.momentum_scanner:
                self.momentum_scanner.stop_screening()
            
            if hasattr(self.trader, 'db_worker'):
                self.trader.db_worker.stop()
                self.trader.db_thread.quit()
                self.trader.db_thread.wait()

            if getattr(self.trader, 'save_data_timer', None):
                self.trader.save_data_timer.stop()
            if getattr(self, 'update_chart_status_timer', None):
                self.update_chart_status_timer.stop()
            if getattr(self.trader, 'tickdata', None):
                self.trader.tickdata.update_data_timer.stop()
            if getattr(self.trader, 'mindata', None):
                self.trader.mindata.update_data_timer.stop()
            if getattr(self.trader, 'daydata', None):
                self.trader.daydata.update_data_timer.stop()

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
    setup_logging()

    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 9))
    myWindow = MyWindow()
    myWindow.setWindowIcon(QIcon('stock_trader.ico'))
    myWindow.showMaximized()
    app.exec_()