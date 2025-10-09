import sqlite3
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import logging
import json
import copy
import configparser
from strategy_utils import (
    STRATEGY_SAFE_GLOBALS,
    evaluate_strategies,
    build_backtest_buy_locals,
    build_backtest_sell_locals
)

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

class Backtester:
    """백테스팅 엔진 (combined_tick_data 사용, settings.ini 전략 적용)"""
    
    def __init__(self, db_path, config_file='settings.ini', initial_cash=10000000):
        self.db_path = db_path
        self.config_file = config_file
        self.initial_cash = initial_cash
        
        # settings.ini 로드
        # ✅ RawConfigParser 사용 (% 문자 이슈 완전 해결)
        self.config = configparser.RawConfigParser()
        if config_file:
            try:
                self.config.read(config_file, encoding='utf-8')
                logging.info(f"✅ 설정 파일 로드: {config_file}")
            except Exception as ex:
                logging.error(f"설정 파일 로드 실패: {ex}")
        
        # 포트폴리오
        self.cash = initial_cash
        self.holdings = {}
        self.buy_prices = {}
        self.buy_times = {}
        self.highest_prices = {}
        
        # 매매 설정
        self.max_holdings = 3
        self.position_size = 0.3
        
        # 거래 기록
        self.trades = []
        self.equity_curve = []
        
        # 일별 성과 추적
        self.daily_results = []
        self.daily_trades_count = {}
        self.daily_profit = {}
        
        # 전략 설정
        self.strategy_params = {}
        self.tick_interval = 30
        
        # 현재 전략
        self.current_strategy_name = None
        self.buy_strategies = []
        self.sell_strategies = []
    
    def load_codes(self, start_date, end_date):
        """기간 내 거래된 종목 목록"""
        conn = sqlite3.connect(self.db_path)
        
        query = f"""
            SELECT DISTINCT code 
            FROM combined_tick_data 
            WHERE date >= '{start_date}' AND date <= '{end_date}'
            ORDER BY code
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        return df['code'].tolist()
    
    def load_combined_data(self, code, start_date, end_date):
        """결합 데이터 로드 (combined_tick_data)"""
        conn = sqlite3.connect(self.db_path)
        
        query = f"""
            SELECT * FROM combined_tick_data 
            WHERE code = '{code}' 
            AND date >= '{start_date}' 
            AND date <= '{end_date}'
            ORDER BY timestamp
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        if len(df) > 0:
            df['timestamp'] = pd.to_datetime(df['timestamp'])
        
        return df
    
    def load_strategies(self, strategy_name, strategy_type='buy'):
        """settings.ini에서 전략 로드"""
        strategies = []
        
        if not self.config.has_section(strategy_name):
            logging.warning(f"전략 '{strategy_name}' 섹션이 설정 파일에 없음")
            return strategies
        
        prefix = 'buy_stg_' if strategy_type == 'buy' else 'sell_stg_'
        
        for key in self.config[strategy_name]:
            if key.startswith(prefix):
                try:
                    strategy_json = self.config[strategy_name][key]
                    strategy_data = json.loads(strategy_json)
                    strategies.append(strategy_data)
                except Exception as ex:
                    logging.error(f"전략 '{key}' 로드 실패: {ex}")
        
        logging.info(f"✅ {strategy_name} - {strategy_type} 전략 {len(strategies)}개 로드")
        return strategies
    
    def check_buy_condition(self, code, row, previous_rows):
        """매수 조건 확인 (settings.ini 전략 사용)"""
        
        # 최대 보유 종목 수 체크
        if len(self.holdings) >= self.max_holdings:
            return False
        
        # 자금 체크
        if self.cash < self.initial_cash * self.position_size:
            return False
        
        # 전략이 로드되지 않았으면 False
        if not self.buy_strategies:
            return False
        
        # ===== safe_locals 구성 (공통 유틸리티 사용) =====
        safe_locals = build_backtest_buy_locals(row)
        
        # ===== 전략 조건 평가 (공통 함수 사용) =====
        matched, strategy = evaluate_strategies(
            self.buy_strategies, 
            safe_locals, 
            code=code, 
            strategy_type="매수"
        )
        
        if matched:
            logging.debug(f"{code}: 매수 조건 충족 - {strategy['name']}")
            return True
        
        return False
    
    def check_sell_condition(self, code, row, previous_rows):
        """매도 조건 확인 (settings.ini 전략 사용)"""
        
        if code not in self.holdings:
            return False, None
        
        # 전략이 로드되지 않았으면 False
        if not self.sell_strategies:
            return False, None
        
        current_price = row.get('tick_C', 0)
        buy_price = self.buy_prices[code]
        
        # 최고가 업데이트
        if code not in self.highest_prices:
            self.highest_prices[code] = buy_price
        self.highest_prices[code] = max(self.highest_prices[code], current_price)
        
        # ===== safe_locals 구성 (공통 유틸리티 사용) =====
        safe_locals = build_backtest_sell_locals(
            row=row,
            current_price=current_price,
            buy_price=buy_price,
            highest_price=self.highest_prices[code],
            buy_time=self.buy_times[code]
        )
        
        # ===== 전략 조건 평가 (공통 함수 사용) =====
        matched, strategy = evaluate_strategies(
            self.sell_strategies, 
            safe_locals, 
            code=code, 
            strategy_type="매도"
        )
        
        if matched:
            reason = strategy['name']
            logging.debug(f"{code}: 매도 조건 충족 - {reason}")
            return True, reason
        
        return False, None
    
    def execute_buy(self, code, price, timestamp):
        """매수 실행"""
        
        buy_amount = self.cash * self.position_size
        qty = int(buy_amount / price)
        
        if qty <= 0:
            return
        
        cost = price * qty * 1.00015
        
        if cost > self.cash:
            return
        
        self.cash -= cost
        self.holdings[code] = qty
        self.buy_prices[code] = price
        self.buy_times[code] = timestamp
        
        self.trades.append({
            'timestamp': timestamp,
            'code': code,
            'action': 'BUY',
            'price': price,
            'qty': qty,
            'cost': cost,
            'reason': '매수'
        })
        
        logging.debug(f"{timestamp}: {code} 매수 {qty}주 @{price:,.0f}원")
    
    def execute_sell(self, code, price, timestamp, reason):
        """매도 실행"""
        
        if code not in self.holdings:
            return
        
        qty = self.holdings[code]
        revenue = price * qty * 0.99835
        
        buy_price = self.buy_prices[code]
        buy_cost = buy_price * qty * 1.00015
        profit = revenue - buy_cost
        profit_pct = (profit / buy_cost) * 100
        
        hold_minutes = (timestamp - self.buy_times[code]).total_seconds() / 60
        
        self.cash += revenue
        del self.holdings[code]
        del self.buy_prices[code]
        del self.buy_times[code]
        if code in self.highest_prices:
            del self.highest_prices[code]
        
        self.trades.append({
            'timestamp': timestamp,
            'code': code,
            'action': 'SELL',
            'price': price,
            'qty': qty,
            'revenue': revenue,
            'profit': profit,
            'profit_pct': profit_pct,
            'hold_minutes': hold_minutes,
            'reason': reason
        })
        
        # 일별 손익 누적
        date_str = timestamp.strftime('%Y%m%d')
        if date_str not in self.daily_profit:
            self.daily_profit[date_str] = 0
            self.daily_trades_count[date_str] = 0
        self.daily_profit[date_str] += profit
        self.daily_trades_count[date_str] += 1
        
        logging.debug(f"{timestamp}: {code} 매도 {qty}주 @{price:,.0f}원 ({profit_pct:+.2f}%) - {reason}")
    
    def calculate_portfolio_value(self, current_prices):
        """포트폴리오 가치 계산"""
        
        holdings_value = sum(
            current_prices.get(code, 0) * qty 
            for code, qty in self.holdings.items()
        )
        
        return self.cash + holdings_value
    
    def run(self, start_date, end_date, strategy_name='통합 전략'):
        """백테스팅 실행 (combined_tick_data 사용, settings.ini 전략 적용)"""
        
        logging.info(f"=== 백테스팅 시작: {start_date} ~ {end_date} ===")
        logging.info(f"전략: {strategy_name}")
        
        # 초기화
        self.cash = self.initial_cash
        self.holdings = {}
        self.buy_prices = {}
        self.buy_times = {}
        self.highest_prices = {}
        self.trades = []
        self.equity_curve = []
        
        # 일별 성과 초기화
        self.daily_results = []
        self.daily_trades_count = {}
        self.daily_profit = {}
        self.current_date = None
        self.daily_start_cash = self.initial_cash
        
        # ===== 전략 로드 =====
        self.current_strategy_name = strategy_name
        self.buy_strategies = self.load_strategies(strategy_name, 'buy')
        self.sell_strategies = self.load_strategies(strategy_name, 'sell')
        
        if not self.buy_strategies:
            logging.error(f"매수 전략이 없습니다. 백테스팅을 중단합니다.")
            return self.calculate_results(strategy_name, start_date, end_date)
        
        if not self.sell_strategies:
            logging.warning(f"매도 전략이 없습니다. 기본 매도 조건만 사용합니다.")
        
        logging.info(f"매수 전략: {len(self.buy_strategies)}개")
        for idx, stg in enumerate(self.buy_strategies, 1):
            logging.info(f"  {idx}. {stg['name']}")
        
        logging.info(f"매도 전략: {len(self.sell_strategies)}개")
        for idx, stg in enumerate(self.sell_strategies, 1):
            logging.info(f"  {idx}. {stg['name']}")
        
        # 종목 목록
        codes = self.load_codes(start_date, end_date)
        logging.info(f"대상 종목 수: {len(codes)}개")
        
        if len(codes) == 0:
            logging.warning("백테스팅 데이터 없음!")
            return self.calculate_results(strategy_name, start_date, end_date)
        
        # 모든 종목의 데이터 로드
        all_data = {}
        for code in codes:
            df = self.load_combined_data(code, start_date, end_date)
            if len(df) > 0:
                all_data[code] = df
                logging.info(f"{code}: {len(df)}개 데이터 로드")
        
        # 시간순으로 모든 이벤트 정렬
        all_events = []
        for code, df in all_data.items():
            for idx, row in df.iterrows():
                all_events.append((row['timestamp'], code, idx, row))
        
        all_events.sort(key=lambda x: x[0])
        
        logging.info(f"총 {len(all_events)}개 이벤트 처리 시작")
        
        # 이벤트 처리
        processed = 0
        for timestamp, code, idx, row in all_events:
            processed += 1
            
            if processed % 10000 == 0:
                logging.info(f"처리 중: {processed}/{len(all_events)} ({processed/len(all_events)*100:.1f}%)")
            
            # 날짜가 바뀌면 전일 성과 계산
            current_date_str = timestamp.strftime('%Y%m%d')
            if self.current_date and self.current_date != current_date_str:
                # 모든 보유 종목의 현재가 업데이트
                prev_date_prices = {}
                for held_code in self.holdings.keys():
                    if held_code in all_data:
                        code_data = all_data[held_code]
                        recent = code_data[code_data['timestamp'] < timestamp]
                        if len(recent) > 0:
                            prev_date_prices[held_code] = recent.iloc[-1]['tick_C']
                
                # 전일 종료 시점 일별 성과 저장
                self.save_daily_result(self.current_date, prev_date_prices, strategy_name)
                self.daily_start_cash = self.cash
            
            self.current_date = current_date_str
            
            current_price = row['tick_C']
            
            # 현재 가격 정보
            current_prices = {code: current_price}
            
            # 매도 조건 확인 (우선)
            if code in self.holdings:
                # 이전 데이터 (필요시)
                previous_rows = all_data[code].iloc[max(0, idx-10):idx]
                
                should_sell, reason = self.check_sell_condition(code, row, previous_rows)
                if should_sell:
                    self.execute_sell(code, current_price, timestamp, reason)
            
            # 매수 조건 확인
            else:
                previous_rows = all_data[code].iloc[max(0, idx-10):idx]
                
                if self.check_buy_condition(code, row, previous_rows):
                    self.execute_buy(code, current_price, timestamp)
            
            # 포트폴리오 가치 기록 (매 100개 이벤트마다)
            if processed % 100 == 0:
                # 모든 보유 종목의 현재가 업데이트
                for held_code in self.holdings.keys():
                    if held_code in all_data:
                        # 현재 시점 이전 최신 데이터
                        code_data = all_data[held_code]
                        recent = code_data[code_data['timestamp'] <= timestamp]
                        if len(recent) > 0:
                            current_prices[held_code] = recent.iloc[-1]['tick_C']
                
                portfolio_value = self.calculate_portfolio_value(current_prices)
                self.equity_curve.append({
                    'timestamp': timestamp,
                    'value': portfolio_value,
                    'cash': self.cash,
                    'holdings_value': portfolio_value - self.cash
                })
        
        # 최종 청산
        final_timestamp = all_events[-1][0] if all_events else pd.Timestamp.now()
        
        # 최종 가격 수집
        final_prices = {}
        for code in list(self.holdings.keys()):
            code_data = all_data[code]
            final_price = code_data.iloc[-1]['tick_C']
            final_prices[code] = final_price
            self.execute_sell(code, final_price, final_timestamp, "백테스팅 종료")
        
        # 마지막 날 성과 저장
        if self.current_date:
            self.save_daily_result(self.current_date, final_prices, strategy_name)
        
        # 결과 계산
        results = self.calculate_results(strategy_name, start_date, end_date)
        
        # 일별 성과 추가
        results['daily_results'] = self.daily_results
        
        # DB에 저장
        self.save_results(results)
        self.save_daily_results_to_db(strategy_name)
        
        logging.info("=== 백테스팅 완료 ===")
        
        return results
       
    def calculate_results(self, strategy_name, start_date, end_date):
        """결과 계산"""
        
        # ✅ 거래 내역이 없을 때 처리
        if len(self.trades) == 0:
            return {
                'strategy': strategy_name,
                'start_date': start_date,
                'end_date': end_date,
                'initial_cash': self.initial_cash,
                'final_cash': self.cash,
                'total_profit': 0,
                'total_return_pct': 0,
                'total_trades': 0,
                'win_trades': 0,
                'lose_trades': 0,
                'win_rate': 0,
                'avg_profit_pct': 0,
                'max_profit_pct': 0,
                'max_loss_pct': 0,
                'mdd': 0,
                'sharpe_ratio': 0,
                'avg_hold_minutes': 0
            }
        
        df_trades = pd.DataFrame(self.trades)
        sell_trades = df_trades[df_trades['action'] == 'SELL']
        
        if len(sell_trades) == 0:
            return {
                'strategy': strategy_name,
                'start_date': start_date,
                'end_date': end_date,
                'initial_cash': self.initial_cash,
                'final_cash': self.cash,
                'total_profit': 0,
                'total_return_pct': 0,
                'total_trades': 0,
                'win_trades': 0,
                'lose_trades': 0,
                'win_rate': 0,
                'avg_profit_pct': 0,
                'max_profit_pct': 0,
                'max_loss_pct': 0,
                'mdd': 0,
                'sharpe_ratio': 0,
                'avg_hold_minutes': 0
            }
        
        # 기본 통계
        total_trades = len(sell_trades)
        win_trades = len(sell_trades[sell_trades['profit'] > 0])
        lose_trades = len(sell_trades[sell_trades['profit'] <= 0])
        win_rate = (win_trades / total_trades * 100) if total_trades > 0 else 0
        
        total_profit = sell_trades['profit'].sum()
        total_return_pct = (self.cash / self.initial_cash - 1) * 100
        
        avg_profit_pct = sell_trades['profit_pct'].mean()
        max_profit_pct = sell_trades['profit_pct'].max()
        max_loss_pct = sell_trades['profit_pct'].min()
        
        avg_hold_minutes = sell_trades['hold_minutes'].mean()
        
        # MDD 계산
        if len(self.equity_curve) > 0:
            equity_df = pd.DataFrame(self.equity_curve)
            equity_df['peak'] = equity_df['value'].cummax()
            equity_df['drawdown'] = (equity_df['value'] - equity_df['peak']) / equity_df['peak'] * 100
            mdd = equity_df['drawdown'].min()
        else:
            mdd = 0
        
        # 샤프 비율
        returns = sell_trades['profit_pct'] / 100
        sharpe = (returns.mean() / returns.std() * np.sqrt(252)) if returns.std() > 0 else 0
        
        return {
            'strategy': strategy_name,
            'start_date': start_date,
            'end_date': end_date,
            'initial_cash': self.initial_cash,
            'final_cash': self.cash,
            'total_profit': total_profit,
            'total_return_pct': total_return_pct,
            'total_trades': total_trades,
            'win_trades': win_trades,
            'lose_trades': lose_trades,
            'win_rate': win_rate,
            'avg_profit_pct': avg_profit_pct,
            'max_profit_pct': max_profit_pct,
            'max_loss_pct': max_loss_pct,
            'mdd': mdd,
            'sharpe_ratio': sharpe,
            'avg_hold_minutes': avg_hold_minutes
        }
    
    def save_daily_result(self, date_str, current_prices, strategy_name):
        """일별 성과 저장"""
        try:
            # 포트폴리오 가치 계산
            portfolio_value = self.calculate_portfolio_value(current_prices)
            
            # 일별 손익
            daily_profit = self.daily_profit.get(date_str, 0)
            daily_trades = self.daily_trades_count.get(date_str, 0)
            
            # 일일 수익률 계산
            daily_return_pct = ((portfolio_value - self.daily_start_cash) / self.daily_start_cash * 100) if self.daily_start_cash > 0 else 0
            
            # 일별 승/패 계산
            df_trades = pd.DataFrame(self.trades)
            if len(df_trades) > 0:
                daily_sell_trades = df_trades[
                    (df_trades['action'] == 'SELL') & 
                    (df_trades['timestamp'].dt.strftime('%Y%m%d') == date_str)
                ]
                win_trades = len(daily_sell_trades[daily_sell_trades['profit'] > 0])
                lose_trades = len(daily_sell_trades[daily_sell_trades['profit'] <= 0])
            else:
                win_trades = 0
                lose_trades = 0
            
            daily_result = {
                'date': date_str,
                'strategy': strategy_name,
                'portfolio_value': portfolio_value,
                'cash': self.cash,
                'holdings_value': portfolio_value - self.cash,
                'daily_profit': daily_profit,
                'daily_return_pct': daily_return_pct,
                'total_trades': daily_trades,
                'win_trades': win_trades,
                'lose_trades': lose_trades,
                'cumulative_profit': portfolio_value - self.initial_cash
            }
            
            self.daily_results.append(daily_result)
            logging.debug(f"{date_str}: 일별 성과 저장 (수익: {daily_profit:,.0f}원, 수익률: {daily_return_pct:.2f}%)")
            
        except Exception as ex:
            logging.error(f"save_daily_result({date_str}) -> {ex}")
    
    def save_daily_results_to_db(self, strategy_name):
        """일별 성과 DB 저장"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            for daily in self.daily_results:
                cursor.execute('''
                    INSERT OR REPLACE INTO daily_summary (
                        date, strategy, total_trades, win_trades, lose_trades,
                        total_profit, total_return_pct, portfolio_value, cash, holdings_value
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    daily['date'],
                    daily['strategy'],
                    daily['total_trades'],
                    daily['win_trades'],
                    daily['lose_trades'],
                    daily['daily_profit'],
                    daily['daily_return_pct'],
                    daily['portfolio_value'],
                    daily['cash'],
                    daily['holdings_value']
                ))
            
            conn.commit()
            conn.close()
            
            logging.info(f"일별 성과 {len(self.daily_results)}건 DB 저장 완료")
            
        except Exception as ex:
            logging.error(f"save_daily_results_to_db -> {ex}")
    
    def save_results(self, results):
        """결과 DB 저장"""
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO backtest_results (
                    strategy, start_date, end_date, initial_cash, final_cash,
                    total_profit, total_return_pct, total_trades, win_trades, lose_trades,
                    win_rate, avg_profit_pct, max_profit_pct, max_loss_pct,
                    mdd, sharpe_ratio, avg_hold_minutes, parameters
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                results['strategy'], results['start_date'], results['end_date'],
                results['initial_cash'], results['final_cash'],
                results['total_profit'], results['total_return_pct'],
                results['total_trades'], results['win_trades'], results['lose_trades'],
                results['win_rate'], results['avg_profit_pct'],
                results['max_profit_pct'], results['max_loss_pct'],
                results['mdd'], results['sharpe_ratio'], results['avg_hold_minutes'],
                json.dumps(self.strategy_params)
            ))
            
            conn.commit()
            conn.close()
            
            logging.info("백테스팅 결과 DB 저장 완료")
            
        except Exception as ex:
            logging.error(f"save_results -> {ex}")
    
    def plot_results(self, fig=None):
        """결과 시각화"""
        
        if fig is None:
            fig = plt.figure(figsize=(12, 10))
        else:
            fig.clear()
        
        # 4개 차트
        ax1 = fig.add_subplot(4, 1, 1)
        ax2 = fig.add_subplot(4, 1, 2)
        ax3 = fig.add_subplot(4, 1, 3)
        ax4 = fig.add_subplot(4, 1, 4)
        
        # 1. 수익률 곡선
        if len(self.equity_curve) > 0:
            equity_df = pd.DataFrame(self.equity_curve)
            # numpy 배열로 변환하여 플롯
            ax1.plot(equity_df['timestamp'].values, equity_df['value'].values, label='포트폴리오 가치', linewidth=2)
            ax1.axhline(y=self.initial_cash, color='r', linestyle='--', label='초기 자금', alpha=0.5)
            ax1.set_title('포트폴리오 가치 변화', fontsize=12, fontweight='bold')
            ax1.set_ylabel('가치 (원)')
            ax1.legend()
            ax1.grid(True, alpha=0.3)
            ax1.ticklabel_format(style='plain', axis='y')
        
        # 2. Drawdown
        if len(self.equity_curve) > 0:
            equity_df['peak'] = equity_df['value'].cummax()
            equity_df['drawdown'] = (equity_df['value'] - equity_df['peak']) / equity_df['peak'] * 100
            # numpy 배열로 변환
            ax2.fill_between(equity_df['timestamp'].values, 0, equity_df['drawdown'].values, color='red', alpha=0.3)
            ax2.plot(equity_df['timestamp'].values, equity_df['drawdown'].values, color='red', linewidth=1)
            ax2.set_title('Drawdown', fontsize=12, fontweight='bold')
            ax2.set_ylabel('Drawdown (%)')
            ax2.grid(True, alpha=0.3)
        
        # 3. 거래별 손익
        if len(self.trades) > 0:
            df_trades = pd.DataFrame(self.trades)
            sell_trades = df_trades[df_trades['action'] == 'SELL']
        else:
            sell_trades = pd.DataFrame()
        
        if len(sell_trades) > 0:
            colors = ['green' if p > 0 else 'red' for p in sell_trades['profit_pct'].values]
            ax3.bar(range(len(sell_trades)), sell_trades['profit_pct'].values, color=colors, alpha=0.6)
            ax3.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            ax3.set_title('거래별 수익률', fontsize=12, fontweight='bold')
            ax3.set_ylabel('수익률 (%)')
            ax3.set_xlabel('거래 번호')
            ax3.grid(True, alpha=0.3)
        
        # 4. 누적 손익
        if len(sell_trades) > 0:
            cumulative_profit = sell_trades['profit'].cumsum().values
            ax4.plot(range(len(cumulative_profit)), cumulative_profit, color='blue', linewidth=2)
            ax4.axhline(y=0, color='black', linestyle='--', linewidth=0.5)
            ax4.fill_between(range(len(cumulative_profit)), 0, cumulative_profit, 
                            where=(cumulative_profit >= 0), color='green', alpha=0.3)
            ax4.fill_between(range(len(cumulative_profit)), 0, cumulative_profit, 
                            where=(cumulative_profit < 0), color='red', alpha=0.3)
            ax4.set_title('누적 손익', fontsize=12, fontweight='bold')
            ax4.set_ylabel('누적 손익 (원)')
            ax4.set_xlabel('거래 번호')
            ax4.grid(True, alpha=0.3)
            ax4.ticklabel_format(style='plain', axis='y')
        
        fig.tight_layout()
        
        return fig
    
    def plot_daily_results(self, fig=None):
        """일별 성과 시각화"""
        
        if len(self.daily_results) == 0:
            logging.warning("일별 성과 데이터가 없습니다.")
            # 빈 차트 표시
            if fig is not None:
                fig.clear()
                ax = fig.add_subplot(1, 1, 1)
                ax.text(0.5, 0.5, '일별 데이터가 없습니다.', 
                       ha='center', va='center', fontsize=14)
                ax.axis('off')
                fig.tight_layout()
            return fig
        
        if fig is None:
            fig = plt.figure(figsize=(14, 10))
        else:
            fig.clear()
        
        df_daily = pd.DataFrame(self.daily_results)
        df_daily['date'] = pd.to_datetime(df_daily['date'], format='%Y%m%d')
        
        # 4개 차트
        ax1 = fig.add_subplot(4, 1, 1)
        ax2 = fig.add_subplot(4, 1, 2)
        ax3 = fig.add_subplot(4, 1, 3)
        ax4 = fig.add_subplot(4, 1, 4)
        
        # 1. 일별 수익/손실
        colors = ['green' if p > 0 else 'red' for p in df_daily['daily_profit'].values]
        ax1.bar(df_daily['date'].values, df_daily['daily_profit'].values, color=colors, alpha=0.6, width=0.8)
        ax1.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
        ax1.set_title('일별 손익', fontsize=12, fontweight='bold')
        ax1.set_ylabel('손익 (원)')
        ax1.grid(True, alpha=0.3)
        ax1.ticklabel_format(style='plain', axis='y')
        
        # 2. 일별 수익률
        colors = ['green' if p > 0 else 'red' for p in df_daily['daily_return_pct'].values]
        ax2.bar(df_daily['date'].values, df_daily['daily_return_pct'].values, color=colors, alpha=0.6, width=0.8)
        ax2.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
        ax2.set_title('일별 수익률', fontsize=12, fontweight='bold')
        ax2.set_ylabel('수익률 (%)')
        ax2.grid(True, alpha=0.3)
        
        # 3. 누적 수익
        ax3.plot(df_daily['date'].values, df_daily['cumulative_profit'].values, color='blue', linewidth=2, marker='o', markersize=4)
        ax3.axhline(y=0, color='black', linestyle='--', linewidth=0.5)
        cumulative_vals = df_daily['cumulative_profit'].values
        ax3.fill_between(df_daily['date'].values, 0, cumulative_vals, 
                        where=(cumulative_vals >= 0), color='green', alpha=0.3)
        ax3.fill_between(df_daily['date'].values, 0, cumulative_vals, 
                        where=(cumulative_vals < 0), color='red', alpha=0.3)
        ax3.set_title('누적 손익', fontsize=12, fontweight='bold')
        ax3.set_ylabel('누적 손익 (원)')
        ax3.grid(True, alpha=0.3)
        ax3.ticklabel_format(style='plain', axis='y')
        
        # 4. 일별 거래 횟수
        ax4.bar(df_daily['date'].values, df_daily['total_trades'].values, color='steelblue', alpha=0.6, width=0.8)
        ax4.set_title('일별 거래 횟수', fontsize=12, fontweight='bold')
        ax4.set_ylabel('거래 횟수')
        ax4.set_xlabel('날짜')
        ax4.grid(True, alpha=0.3)
        
        # x축 날짜 포맷 설정
        import matplotlib.dates as mdates
        for ax in [ax1, ax2, ax3, ax4]:
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
            ax.tick_params(axis='x', rotation=45)
        
        fig.tight_layout()
        
        return fig