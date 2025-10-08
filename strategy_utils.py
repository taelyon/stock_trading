"""
전략 평가 및 지표 처리 유틸리티 모듈
중복 코드 제거 및 재사용성 향상을 위한 공통 함수/클래스
"""
import logging

# ==================== 전략 평가용 안전한 globals ====================
STRATEGY_SAFE_GLOBALS = {
    '__builtins__': {
        'min': min, 'max': max, 'abs': abs, 'round': round,
        'int': int, 'float': float, 'bool': bool, 'str': str,
        'len': len, 'sum': sum, 'all': all, 'any': any,
        'True': True, 'False': False, 'None': None
    }
}


# ==================== 전략 평가 공통 함수 ====================
def evaluate_strategies(strategies, safe_locals, code="", strategy_type=""):
    """
    전략 조건들을 평가하고 일치하는 첫 번째 전략을 반환
    
    Args:
        strategies: 평가할 전략 리스트 (각 전략은 'name'과 'content' 필드 포함)
        safe_locals: 평가에 사용할 로컬 변수 딕셔너리
        code: 종목 코드 (로깅용)
        strategy_type: 전략 타입 ("매수", "매도" 등, 로깅용)
    
    Returns:
        (bool, dict or None): (조건 충족 여부, 충족된 전략 또는 None)
    """
    for strategy in strategies:
        try:
            condition = strategy.get('content', '')
            if not condition:
                continue
                
            if eval(condition, STRATEGY_SAFE_GLOBALS, safe_locals):
                strategy_name = strategy.get('name', '전략')
                if code:
                    logging.debug(f"{code}: {strategy_name} 조건 충족")
                return True, strategy
                
        except Exception as ex:
            strategy_name = strategy.get('name', '알 수 없는 전략')
            logging.error(f"{code} {strategy_type} 전략 '{strategy_name}' 평가 오류: {ex}")
    
    return False, None


# ==================== 지표 추출 유틸리티 ====================
class IndicatorExtractor:
    """틱/분봉 데이터로부터 지표를 추출하는 헬퍼 클래스"""
    
    @staticmethod
    def extract_tick_indicators(tick_latest):
        """틱 데이터에서 주요 지표 추출"""
        return {
            # 이동평균
            'MAT5': tick_latest.get('MAT5', 0),
            'MAT20': tick_latest.get('MAT20', 0),
            'MAT60': tick_latest.get('MAT60', 0),
            'MAT120': tick_latest.get('MAT120', 0),
            
            # 가격
            'C': tick_latest.get('C', 0),
            'tick_C': tick_latest.get('C', 0),
            
            # 거래량/강도
            'VWAP': tick_latest.get('VWAP', 0),
            'tick_VWAP': tick_latest.get('VWAP', 0),
            
            # 모멘텀 지표
            'RSIT': tick_latest.get('RSIT', 50),
            'MACDT': tick_latest.get('MACDT', 0),
            'MACDT_SIGNAL': tick_latest.get('MACDT_SIGNAL', 0),
            'OSCT': tick_latest.get('OSCT', 0),
            
            # 스토캐스틱
            'STOCHK': tick_latest.get('STOCHK', 50),
            'STOCHD': tick_latest.get('STOCHD', 50),
            
            # 변동성
            'ATR': tick_latest.get('ATR', 0),
            
            # 볼린저 밴드
            'BB_UPPER': tick_latest.get('BB_UPPER', 0),
            'BB_MIDDLE': tick_latest.get('BB_MIDDLE', 0),
            'BB_LOWER': tick_latest.get('BB_LOWER', 0),
            'BB_POSITION': tick_latest.get('BB_POSITION', 0),
            'BB_BANDWIDTH': tick_latest.get('BB_BANDWIDTH', 0),
            
            # 추가 지표
            'WILLIAMS_R': tick_latest.get('WILLIAMS_R', -50),
            'ROC': tick_latest.get('ROC', 0),
            'OBV': tick_latest.get('OBV', 0),
            'OBV_MA20': tick_latest.get('OBV_MA20', 0),
            'VP_POC': tick_latest.get('VP_POC', 0),
            'VP_POSITION': tick_latest.get('VP_POSITION', 0),
            'CCI': tick_latest.get('CCI', 0),
        }
    
    @staticmethod
    def extract_min_indicators(min_latest):
        """분봉 데이터에서 주요 지표 추출"""
        return {
            # 이동평균
            'MAM5': min_latest.get('MAM5', 0),
            'MAM10': min_latest.get('MAM10', 0),
            'MAM20': min_latest.get('MAM20', 0),
            
            # 가격
            'min_C': min_latest.get('C', 0),
            'min_close': min_latest.get('C', 0),
            'min_O': min_latest.get('O', 0),
            
            # 모멘텀 지표
            'RSI': min_latest.get('RSI', 50),
            'min_RSI': min_latest.get('RSI', 50),
            'MACD': min_latest.get('MACD', 0),
            'MACD_SIGNAL': min_latest.get('MACD_SIGNAL', 0),
            'OSC': min_latest.get('OSC', 0),
            'min_OSC': min_latest.get('OSC', 0),
            
            # 스토캐스틱
            'min_STOCHK': min_latest.get('STOCHK', 50),
            'min_STOCHD': min_latest.get('STOCHD', 50),
            
            # 거래량
            'min_VWAP': min_latest.get('VWAP', 0),
            
            # 추가 지표
            'min_WILLIAMS_R': min_latest.get('WILLIAMS_R', -50),
            'min_ROC': min_latest.get('ROC', 0),
            'min_OBV': min_latest.get('OBV', 0),
            'min_OBV_MA20': min_latest.get('OBV_MA20', 0),
            'min_CCI': min_latest.get('CCI', 0),
        }


# ==================== SafeLocals 빌더 ====================
class SafeLocalsBuilder:
    """전략 평가용 safe_locals 딕셔너리를 단계적으로 구성하는 빌더 클래스"""
    
    def __init__(self):
        self.locals = {}
    
    def add_tick_indicators(self, tick_latest):
        """틱 지표 추가"""
        self.locals.update(IndicatorExtractor.extract_tick_indicators(tick_latest))
        return self
    
    def add_min_indicators(self, min_latest):
        """분봉 지표 추가"""
        self.locals.update(IndicatorExtractor.extract_min_indicators(min_latest))
        return self
    
    def add_custom_vars(self, **kwargs):
        """사용자 정의 변수 추가"""
        self.locals.update(kwargs)
        return self
    
    def add_profit_vars(self, current_profit_pct, from_peak_pct, hold_minutes, 
                       buy_price=0, highest_price=0):
        """수익률 관련 변수 추가 (매도 조건용)"""
        self.locals.update({
            'current_profit_pct': current_profit_pct,
            'from_peak_pct': from_peak_pct,
            'hold_minutes': hold_minutes,
            'buy_price': buy_price,
            'highest_price': highest_price,
        })
        return self
    
    def add_time_vars(self, after_market_close=False):
        """시간 관련 변수 추가"""
        self.locals.update({
            'after_market_close': after_market_close,
        })
        return self
    
    def add_strategy_vars(self, strength=0, gap_hold=False, volatility_breakout=False,
                         volume_profile_breakout=False, positive_candle=False):
        """전략 관련 추가 변수"""
        self.locals.update({
            'strength': strength,
            'gap_hold': gap_hold,
            'volatility_breakout': volatility_breakout,
            'volume_profile_breakout': volume_profile_breakout,
            'positive_candle': positive_candle,
        })
        return self
    
    def add_derived_indicators(self, tick_latest=None, min_latest=None, 
                              tick_recent_data=None, min_recent_data=None):
        """파생 지표 추가 (최근 데이터 기반)"""
        derived = {}
        
        # OSCT 음수 연속 확인
        if tick_recent_data:
            OSCT_recent = tick_recent_data.get('OSCT', [0, 0, 0])
            if len(OSCT_recent) >= 2:
                derived['osct_negative'] = OSCT_recent[-2] < 0 and OSCT_recent[-1] < 0
        
        # OBV 다이버전스
        if tick_latest:
            OBV = tick_latest.get('OBV', 0)
            OBV_MA20 = tick_latest.get('OBV_MA20', 0)
            derived['obv_divergence'] = OBV < OBV_MA20
        
        if min_latest:
            min_OBV = min_latest.get('OBV', 0)
            min_OBV_MA20 = min_latest.get('OBV_MA20', 0)
            derived['min_obv_divergence'] = min_OBV < min_OBV_MA20
        
        # Williams %R 과매수/과매도
        if tick_latest:
            WILLIAMS_R = tick_latest.get('WILLIAMS_R', -50)
            derived['williams_overbought'] = WILLIAMS_R > -20
            derived['williams_oversold'] = WILLIAMS_R < -80
        
        # 양봉 확인
        if min_recent_data:
            min_close_recent = min_recent_data.get('C', [0, 0])[-2:]
            min_open_recent = min_recent_data.get('O', [0, 0])[-2:]
            if len(min_close_recent) >= 2 and len(min_open_recent) >= 2:
                derived['positive_candle'] = all(
                    min_close_recent[i] > min_open_recent[i] 
                    for i in range(len(min_close_recent))
                )
        
        self.locals.update(derived)
        return self
    
    def build(self):
        """최종 딕셔너리 반환"""
        return self.locals
    
    def reset(self):
        """빌더 초기화"""
        self.locals = {}
        return self


# ==================== 백테스팅용 간소화 빌더 ====================
def build_backtest_buy_locals(row):
    """백테스팅 매수용 safe_locals 빌더 (간소화 버전)"""
    builder = SafeLocalsBuilder()
    
    # 틱 데이터
    tick_data = {
        'MAT5': row.get('tick_MAT5', 0),
        'MAT20': row.get('tick_MAT20', 0),
        'MAT60': row.get('tick_MAT60', 0),
        'MAT120': row.get('tick_MAT120', 0),
        'C': row.get('tick_C', 0),
        'VWAP': row.get('tick_VWAP', 0),
        'tick_VWAP': row.get('tick_VWAP', 0),
        'RSIT': row.get('tick_RSIT', 50),
        'MACDT': row.get('tick_MACDT', 0),
        'MACDT_SIGNAL': row.get('tick_MACDT_SIGNAL', 0),
        'OSCT': row.get('tick_OSCT', 0),
        'STOCHK': row.get('tick_STOCHK', 50),
        'STOCHD': row.get('tick_STOCHD', 50),
        'ATR': row.get('tick_ATR', 0),
        'BB_UPPER': row.get('tick_BB_UPPER', 0),
        'BB_MIDDLE': row.get('tick_BB_MIDDLE', 0),
        'BB_LOWER': row.get('tick_BB_LOWER', 0),
        'BB_POSITION': row.get('tick_BB_POSITION', 0),
        'BB_BANDWIDTH': row.get('tick_BB_BANDWIDTH', 0),
        'WILLIAMS_R': row.get('tick_WILLIAMS_R', -50),
        'ROC': row.get('tick_ROC', 0),
        'OBV': row.get('tick_OBV', 0),
        'OBV_MA20': row.get('tick_OBV_MA20', 0),
        'VP_POC': row.get('tick_VP_POC', 0),
        'VP_POSITION': row.get('tick_VP_POSITION', 0),
    }
    
    # 분봉 데이터
    min_data = {
        'MAM5': row.get('min_MAM5', 0),
        'MAM10': row.get('min_MAM10', 0),
        'MAM20': row.get('min_MAM20', 0),
        'min_close': row.get('min_C', 0),
        'RSI': row.get('min_RSI', 50),
        'min_RSI': row.get('min_RSI', 50),
        'MACD': row.get('min_MACD', 0),
        'MACD_SIGNAL': row.get('min_MACD_SIGNAL', 0),
        'OSC': row.get('min_OSC', 0),
        'min_OSC': row.get('min_OSC', 0),
        'min_STOCHK': row.get('min_STOCHK', 50),
        'min_STOCHD': row.get('min_STOCHD', 50),
        'min_VWAP': row.get('min_VWAP', 0),
        'min_WILLIAMS_R': row.get('min_WILLIAMS_R', -50),
        'min_ROC': row.get('min_ROC', 0),
        'min_OBV': row.get('min_OBV', 0),
        'min_OBV_MA20': row.get('min_OBV_MA20', 0),
    }
    
    # 추가 변수
    extra_vars = {
        'strength': row.get('strength', 0),
        'positive_candle': row.get('min_C', 0) > row.get('min_O', 0),
        'gap_hold': True,
        'volatility_breakout': False,
        'volume_profile_breakout': row.get('tick_VP_POSITION', 0) > 0,
    }
    
    builder.locals.update(tick_data)
    builder.locals.update(min_data)
    builder.locals.update(extra_vars)
    
    return builder.build()


def build_backtest_sell_locals(row, current_price, buy_price, highest_price, buy_time):
    """백테스팅 매도용 safe_locals 빌더"""
    # 수익률 계산
    profit_pct = (current_price / buy_price - 1) * 100 if buy_price > 0 else 0
    from_peak_pct = (current_price / highest_price - 1) * 100 if highest_price > 0 else 0
    hold_minutes = (row['timestamp'] - buy_time).total_seconds() / 60
    
    # 장마감 여부
    hour = row['timestamp'].hour
    minute = row['timestamp'].minute
    after_market_close = (hour >= 14 and minute >= 45)
    
    builder = SafeLocalsBuilder()
    
    # 틱 데이터
    tick_data = {
        'MAT5': row.get('tick_MAT5', 0),
        'MAT20': row.get('tick_MAT20', 0),
        'MAT60': row.get('tick_MAT60', 0),
        'C': row.get('tick_C', 0),
        'OSCT': row.get('tick_OSCT', 0),
        'STOCHK': row.get('tick_STOCHK', 50),
        'STOCHD': row.get('tick_STOCHD', 50),
        'RSIT': row.get('tick_RSIT', 50),
        'WILLIAMS_R': row.get('tick_WILLIAMS_R', -50),
    }
    
    # 분봉 데이터
    min_data = {
        'MAM5': row.get('min_MAM5', 0),
        'MAM10': row.get('min_MAM10', 0),
    }
    
    # 매도 조건용 변수들
    sell_vars = {
        'current_profit_pct': profit_pct,
        'from_peak_pct': from_peak_pct,
        'hold_minutes': hold_minutes,
        'after_market_close': after_market_close,
        'gap_hold': True,
    }
    
    builder.locals.update(tick_data)
    builder.locals.update(min_data)
    builder.locals.update(sell_vars)
    
    return builder.build()
