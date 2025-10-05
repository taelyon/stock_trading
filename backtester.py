import sqlite3
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import logging
import json
import copy

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False


class Backtester:
    """백테스팅 엔진"""
    
    def __init__(self, db_path, initial_cash=10000000):
        self.db_path = db_path
        self.initial_cash = initial_cash
        
        # 포트폴리오
        self.cash = initial_cash
        self.holdings = {}  # {code: qty}
        self.buy_prices = {}
        self.buy_times = {}
        self.highest_prices = {}
        
        # 매매 설정
        self.max_holdings = 3
        self.position_size = 0.3  # 종목당 30%
        
        # 거래 기록
        self.trades = []
        self.equity_curve = []
        
        # 전략 설정
        self.strategy_params = {}
        
    def load_codes(self, start_date, end_date):
        """기간 내 거래된 종목 목록"""
        conn = sqlite3.connect(self.db_path)
        
        query = f"""
            SELECT DISTINCT code 
            FROM tick_data 
            WHERE date >= '{start_date}' AND date <= '{end_date}'
            ORDER BY code
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        return df['code'].tolist()
    
    def load_tick_data(self, code, start_date, end_date):
        """틱 데이터 로드"""
        conn = sqlite3.connect(self.db_path)
        
        query = f"""
            SELECT * FROM tick_data 
            WHERE code = '{code}' 
            AND date >= '{start_date}' 
            AND date <= '{end_date}'
            ORDER BY date, time, sequence
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        if len(df) > 0:
            df['timestamp'] = pd.to_datetime(
                df['date'].astype(str) + df['time'].astype(str).str.zfill(4),
                format='%Y%m%d%H%M'
            )
        
        return df
    
    def load_min_data(self, code, start_date, end_date):
        """분봉 데이터 로드"""
        conn = sqlite3.connect(self.db_path)
        
        query = f"""
            SELECT * FROM min_data 
            WHERE code = '{code}' 
            AND date >= '{start_date}' 
            AND date <= '{end_date}'
            ORDER BY date, time
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        if len(df) > 0:
            df['timestamp'] = pd.to_datetime(
                df['date'].astype(str) + df['time'].astype(str).str.zfill(4),
                format='%Y%m%d%H%M'
            )
        
        return df
    
    def check_buy_condition(self, code, tick_df, min_df, idx):
        """매수 조건 확인 (통합 전략)"""
        
        # 현재까지의 데이터만 사용 (미래 정보 누출 방지)
        tick_now = tick_df.iloc[idx]
        min_latest = min_df[min_df['timestamp'] <= tick_now['timestamp']]
        
        if len(min_latest) == 0:
            return False
        
        min_now = min_latest.iloc[-1]
        
        # 최대 보유 종목 수 체크
        if len(self.holdings) >= self.max_holdings:
            return False
        
        # 자금 체크
        if self.cash < self.initial_cash * self.position_size:
            return False
        
        # ===== 간단한 통합 전략 조건 =====
        
        # 1. 이평선 정배열
        MAT5 = tick_now['MAT5']
        MAT20 = tick_now['MAT20']
        C = tick_now['C']
        VWAP = tick_now['VWAP']
        
        if not (MAT5 > MAT20 and C > MAT5):
            return False
        
        # 2. VWAP 위
        if not (C > VWAP):
            return False
        
        # 3. 분봉 이평선 정배열
        MAM5 = min_now['MAM5']
        MAM10 = min_now['MAM10']
        
        if not (MAM5 > MAM10):
            return False
        
        # 4. RSI 조건
        RSIT = tick_now['RSIT']
        if not (40 < RSIT < 80):
            return False
        
        return True
    
    def check_sell_condition(self, code, tick_df, min_df, idx):
        """매도 조건 확인"""
        
        if code not in self.holdings:
            return False, None
        
        tick_now = tick_df.iloc[idx]
        min_latest = min_df[min_df['timestamp'] <= tick_now['timestamp']]
        
        if len(min_latest) == 0:
            return False, None
        
        min_now = min_latest.iloc[-1]
        
        current_price = tick_now['C']
        buy_price = self.buy_prices[code]
        
        # 최고가 업데이트
        if code not in self.highest_prices:
            self.highest_prices[code] = buy_price
        self.highest_prices[code] = max(self.highest_prices[code], current_price)
        
        # 수익률 계산
        profit_pct = (current_price / buy_price - 1) * 100
        from_peak_pct = (current_price / self.highest_prices[code] - 1) * 100
        
        # 보유 시간
        hold_minutes = (tick_now['timestamp'] - self.buy_times[code]).total_seconds() / 60
        
        # ===== 매도 조건 =====
        
        # 1. 손절매 (-0.7%)
        if profit_pct < -0.7:
            return True, "손절매"
        
        # 2. 시간 손절 (60분 이상 + 손실)
        if hold_minutes > 60 and profit_pct < 0:
            return True, "시간 손절"
        
        # 3. 장마감 청산 (14:45 이후)
        hour = tick_now['timestamp'].hour
        minute = tick_now['timestamp'].minute
        if hour >= 14 and minute >= 45:
            return True, "장마감 청산"
        
        # 4. 분할 익절 (+1.5%)
        if profit_pct >= 1.5:
            return True, "분할 익절"
        
        # 5. 추적 손절
        if profit_pct > 0.5 and from_peak_pct < -1.0:
            return True, "추적 손절"
        
        # 6. 데드크로스
        MAM5 = min_now['MAM5']
        MAM10 = min_now['MAM10']
        if MAM5 < MAM10:
            return True, "데드크로스"
        
        return False, None
    
    def execute_buy(self, code, price, timestamp):
        """매수 실행"""
        
        buy_amount = self.cash * self.position_size
        qty = int(buy_amount / price)
        
        if qty <= 0:
            return
        
        cost = price * qty * 1.00015  # 수수료 0.015%
        
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
        revenue = price * qty * 0.99835  # 세금+수수료 0.165%
        
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
        
        logging.debug(f"{timestamp}: {code} 매도 {qty}주 @{price:,.0f}원 ({profit_pct:+.2f}%) - {reason}")
    
    def calculate_portfolio_value(self, current_prices):
        """포트폴리오 가치 계산"""
        
        holdings_value = sum(
            current_prices.get(code, 0) * qty 
            for code, qty in self.holdings.items()
        )
        
        return self.cash + holdings_value
    
    def run(self, start_date, end_date, strategy_name='통합 전략'):
        """백테스팅 실행"""
        
        logging.info(f"=== 백테스팅 시작: {start_date} ~ {end_date} ===")
        
        # 초기화
        self.cash = self.initial_cash
        self.holdings = {}
        self.buy_prices = {}
        self.buy_times = {}
        self.highest_prices = {}
        self.trades = []
        self.equity_curve = []
        
        # 종목 목록
        codes = self.load_codes(start_date, end_date)
        logging.info(f"대상 종목 수: {len(codes)}개")
        
        # 날짜별로 처리
        start_dt = datetime.strptime(start_date, '%Y%m%d')
        end_dt = datetime.strptime(end_date, '%Y%m%d')
        
        current_dt = start_dt
        while current_dt <= end_dt:
            date_str = current_dt.strftime('%Y%m%d')
            
            # 주말 스킵
            if current_dt.weekday() >= 5:
                current_dt += timedelta(days=1)
                continue
            
            logging.info(f"처리 중: {date_str}")
            
            # 해당 날짜 데이터 로드
            daily_data = {}
            for code in codes:
                tick_df = self.load_tick_data(code, date_str, date_str)
                min_df = self.load_min_data(code, date_str, date_str)
                
                if len(tick_df) > 0:
                    daily_data[code] = (tick_df, min_df)
            
            # 시간순으로 처리
            all_timestamps = []
            for code, (tick_df, min_df) in daily_data.items():
                all_timestamps.extend(tick_df['timestamp'].tolist())
            
            all_timestamps = sorted(set(all_timestamps))
            
            for ts in all_timestamps:
                current_prices = {}
                
                # 각 종목별 처리
                for code, (tick_df, min_df) in daily_data.items():
                    # 해당 시각의 데이터
                    tick_at_ts = tick_df[tick_df['timestamp'] == ts]
                    
                    if len(tick_at_ts) == 0:
                        continue
                    
                    idx = tick_df[tick_df['timestamp'] == ts].index[0]
                    current_price = tick_at_ts.iloc[0]['C']
                    current_prices[code] = current_price
                    
                    # 매도 조건 확인 (우선)
                    if code in self.holdings:
                        should_sell, reason = self.check_sell_condition(code, tick_df, min_df, idx)
                        if should_sell:
                            self.execute_sell(code, current_price, ts, reason)
                    
                    # 매수 조건 확인
                    else:
                        if self.check_buy_condition(code, tick_df, min_df, idx):
                            self.execute_buy(code, current_price, ts)
                
                # 포트폴리오 가치 기록
                portfolio_value = self.calculate_portfolio_value(current_prices)
                self.equity_curve.append({
                    'timestamp': ts,
                    'value': portfolio_value,
                    'cash': self.cash,
                    'holdings_value': portfolio_value - self.cash
                })
            
            current_dt += timedelta(days=1)
        
        # 결과 계산
        results = self.calculate_results(strategy_name, start_date, end_date)
        
        # DB에 저장
        self.save_results(results)
        
        logging.info("=== 백테스팅 완료 ===")
        
        return results
    
    def calculate_results(self, strategy_name, start_date, end_date):
        """결과 계산"""
        
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
            ax1.plot(equity_df['timestamp'], equity_df['value'], label='포트폴리오 가치', linewidth=2)
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
            ax2.fill_between(equity_df['timestamp'], 0, equity_df['drawdown'], color='red', alpha=0.3)
            ax2.plot(equity_df['timestamp'], equity_df['drawdown'], color='red', linewidth=1)
            ax2.set_title('Drawdown', fontsize=12, fontweight='bold')
            ax2.set_ylabel('Drawdown (%)')
            ax2.grid(True, alpha=0.3)
        
        # 3. 거래별 손익
        df_trades = pd.DataFrame(self.trades)
        sell_trades = df_trades[df_trades['action'] == 'SELL']
        
        if len(sell_trades) > 0:
            colors = ['green' if p > 0 else 'red' for p in sell_trades['profit_pct']]
            ax3.bar(range(len(sell_trades)), sell_trades['profit_pct'], color=colors, alpha=0.6)
            ax3.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            ax3.set_title('거래별 수익률', fontsize=12, fontweight='bold')
            ax3.set_ylabel('수익률 (%)')
            ax3.set_xlabel('거래 번호')
            ax3.grid(True, alpha=0.3)
        
        # 4. 누적 손익
        if len(sell_trades) > 0:
            cumulative_profit = sell_trades['profit'].cumsum()
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