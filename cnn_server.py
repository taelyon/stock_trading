import numpy as np
import pandas as pd
import tensorflow as tf
from tensorflow import keras
from tensorflow.keras import layers
import pickle
import win32pipe
import win32file
import win32api
import winerror
import struct
import logging
import pywintypes
import os
import traceback
import threading
import time
from datetime import datetime

# ==================== 설정 ====================

# 모델 디렉토리 동적 설정
MODEL_DIR = os.getenv('CNN_MODEL_DIR', os.path.join(os.getcwd(), 'models'))
if not os.path.exists(MODEL_DIR):
    os.makedirs(MODEL_DIR)
if not os.access(MODEL_DIR, os.W_OK):
    raise PermissionError(f"쓰기 권한 없음: {MODEL_DIR}")

# 로깅 설정
log_dir = os.path.join(MODEL_DIR, 'log')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"cnn_server_{datetime.now().strftime('%Y%m%d')}.log")

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, mode='a'),
        logging.StreamHandler()
    ]
)

logging.info("CNN Server 시작")

# ==================== 모델 생성 ====================

def create_cnn_model(input_shape, num_classes=3):
    """개선된 CNN 모델 (3클래스 분류)"""
    
    inputs = keras.Input(shape=input_shape)
    
    # ✅ Conv1D + BatchNorm + Dropout
    x = layers.Conv1D(64, 3, padding='same')(inputs)
    x = layers.BatchNormalization()(x)
    x = layers.Activation('relu')(x)
    x = layers.Dropout(0.3)(x)
    
    x = layers.Conv1D(128, 3, padding='same')(x)
    x = layers.BatchNormalization()(x)
    x = layers.Activation('relu')(x)
    x = layers.Dropout(0.3)(x)
    
    x = layers.Conv1D(256, 3, padding='same')(x)
    x = layers.BatchNormalization()(x)
    x = layers.Activation('relu')(x)
    x = layers.Dropout(0.3)(x)
    
    # ✅ Attention 메커니즘
    attention = layers.MultiHeadAttention(num_heads=4, key_dim=64)(x, x)
    x = layers.Add()([x, attention])
    x = layers.LayerNormalization()(x)
    
    # ✅ GRU
    x = layers.GRU(128, return_sequences=False)(x)
    x = layers.Dropout(0.4)(x)
    
    # ✅ Dense
    x = layers.Dense(128, activation='relu')(x)
    x = layers.BatchNormalization()(x)
    x = layers.Dropout(0.3)(x)
    
    x = layers.Dense(64, activation='relu')(x)
    x = layers.Dropout(0.3)(x)
    
    # ✅ 출력층 (3클래스)
    outputs = layers.Dense(num_classes, activation='softmax')(x)
    
    model = keras.Model(inputs=inputs, outputs=outputs)
    
    model.compile(
        optimizer=keras.optimizers.Adam(learning_rate=0.001),
        loss='sparse_categorical_crossentropy',
        metrics=['accuracy']
    )
    
    return model

# ==================== 스케일러 ====================

class IncrementalStandardScaler:
    """개선된 증분 스케일러"""
    
    def __init__(self):
        self.mean_ = None
        self.var_ = None
        self.n_samples_seen_ = 0
        self.version = "1.0"  # 버전 관리

    def partial_fit(self, X):
        X = np.array(X)
        n_samples, n_features = X.shape
        
        if self.mean_ is None:
            self.mean_ = np.zeros(n_features, dtype=np.float64)
            self.var_ = np.zeros(n_features, dtype=np.float64)
            self.n_samples_seen_ = 0
        
        batch_mean = np.mean(X, axis=0)
        batch_var = np.var(X, axis=0)
        batch_count = n_samples
        
        if self.n_samples_seen_ == 0:
            self.mean_ = batch_mean
            self.var_ = batch_var
        else:
            total_count = self.n_samples_seen_ + batch_count
            delta = batch_mean - self.mean_
            self.mean_ += delta * batch_count / total_count
            
            m_a = self.var_ * self.n_samples_seen_
            m_b = batch_var * batch_count
            M2 = m_a + m_b + delta ** 2 * self.n_samples_seen_ * batch_count / total_count
            self.var_ = M2 / total_count
        
        self.n_samples_seen_ += batch_count
        return self

    def transform(self, X):
        X = np.array(X)
        
        if self.mean_ is None or self.var_ is None:
            raise ValueError("스케일러가 피팅되지 않음")
        
        if X.shape[1] != len(self.mean_):
            raise ValueError(f"피처 수 불일치: {X.shape[1]} != {len(self.mean_)}")
        
        # ✅ Robust scaling (극단값 처리)
        std = np.sqrt(self.var_ + 1e-8)
        
        # Clip 적용 (±5 std)
        X_scaled = (X - self.mean_) / std
        X_scaled = np.clip(X_scaled, -5, 5)
        
        return X_scaled

    def fit_transform(self, X):
        return self.partial_fit(X).transform(X)

    def save(self, path):
        """스케일러 저장"""
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                state = {
                    'version': self.version,
                    'mean': self.mean_,
                    'var': self.var_,
                    'n_samples_seen': self.n_samples_seen_
                }
                
                with open(path, 'wb') as f:
                    pickle.dump(state, f)
                
                logging.debug(f"스케일러 저장 완료: {path}")
                return
                
            except OSError as e:
                logging.error(f"스케일러 저장 실패 (시도 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise

    @classmethod
    def load(cls, path):
        """스케일러 로드"""
        with open(path, 'rb') as f:
            state = pickle.load(f)
        
        scaler = cls()
        scaler.version = state.get('version', '1.0')
        scaler.mean_ = state['mean']
        scaler.var_ = state['var']
        scaler.n_samples_seen_ = state['n_samples_seen']
        
        logging.debug(f"스케일러 로드 완료: {path} (버전: {scaler.version})")
        return scaler

# ==================== 전역 변수 ====================

model_scaler_lock = threading.RLock()
shared_model = None
shared_scaler = None
current_seq_length = 5

# ==================== 유틸리티 ====================

def save_model_and_scaler(model, scaler, model_path, scaler_path, seq_length, tick_features, min_features):
    """모델 및 스케일러 저장"""
    with model_scaler_lock:
        try:
            save_dir = os.path.dirname(model_path)
            os.makedirs(save_dir, exist_ok=True)
            
            if not os.access(save_dir, os.W_OK):
                raise PermissionError(f"쓰기 권한 없음: {save_dir}")
            
            # 절대 경로 변환
            model_path = os.path.abspath(model_path)
            scaler_path = os.path.abspath(scaler_path)
            
            # seq_length 저장
            try:
                model.seq_length = seq_length
            except Exception:
                pass
            
            model.save(model_path)
            scaler.save(scaler_path)
            
            logging.info(f"모델/스케일러 저장 완료: {model_path}")
            
        except Exception as ex:
            logging.error(f"모델/스케일러 저장 실패: {ex}\n{traceback.format_exc()}")
            raise

# ==================== 파이프 읽기 ====================

def read_fixed_length(pipe, length, pipe_timeout, desc=""):
    """고정 길이 데이터 읽기"""
    start_time = time.time()
    buffer = b''
    
    while len(buffer) < length:
        if time.time() - start_time > pipe_timeout:
            raise TimeoutError(f"{desc} 읽기 타임아웃")
        
        result = win32file.ReadFile(pipe, length - len(buffer))
        
        if result[0] == winerror.ERROR_MORE_DATA:
            buffer += result[1]
            continue
        elif result[0] != 0 or len(result[1]) == 0:
            raise Exception(f"{desc} 읽기 실패: {result}")
        
        buffer += result[1]
    
    return buffer

def read_message_length(pipe, pipe_timeout):
    """메시지 길이 읽기"""
    return struct.unpack('I', read_fixed_length(pipe, 4, pipe_timeout, "메시지 길이"))[0]

def read_command(pipe, length, pipe_timeout):
    """명령어 읽기"""
    return read_fixed_length(pipe, length, pipe_timeout, "명령어")

def read_data_chunks(pipe, total_chunks, chunk_size, data_len, pipe_timeout):
    """청크 단위 데이터 읽기"""
    data = b''
    received_chunks = set()
    
    for chunk_idx in range(total_chunks):
        # 청크 헤더
        chunk_header_buffer = read_fixed_length(pipe, 4, pipe_timeout, f"청크 {chunk_idx} 헤더")
        chunk_idx_received = struct.unpack('I', chunk_header_buffer)[0]
        
        logging.debug(f"청크 {chunk_idx_received}/{total_chunks} 수신 시작")
        
        # 중복 체크
        if chunk_idx_received in received_chunks:
            logging.error(f"중복 청크: {chunk_idx_received}")
            continue
        
        received_chunks.add(chunk_idx_received)
        
        # 청크 데이터
        chunk_size_to_read = min(chunk_size, data_len - (chunk_idx_received * chunk_size))
        chunk_data = read_fixed_length(pipe, chunk_size_to_read, pipe_timeout, f"청크 {chunk_idx_received} 데이터")
        data += chunk_data
        
        logging.debug(f"청크 {chunk_idx_received}/{total_chunks} 수신 완료 ({len(chunk_data)} 바이트)")
        
        # 확인 응답
        win32file.WriteFile(pipe, struct.pack('I', chunk_idx_received))
    
    return data

# ==================== 훈련 파이프 처리 ====================

def handle_training_pipe(pipe, model_path, scaler_path, stop_event, tick_features, min_features):
    """훈련 파이프 처리"""
    global shared_model, shared_scaler, current_seq_length
    
    chunk_size = 65536
    max_data_size = 52428800
    max_partial_update_size = 10485760
    pipe_timeout = 30
    
    while not stop_event.is_set():
        try:
            logging.debug("훈련 파이프: 클라이언트 연결 대기")
            win32pipe.SetNamedPipeHandleState(pipe, win32pipe.PIPE_READMODE_MESSAGE, None, None)
            
            # 연결 대기
            while not stop_event.is_set():
                try:
                    result = win32pipe.ConnectNamedPipe(pipe, None)
                    if result == 0 or win32api.GetLastError() == winerror.ERROR_PIPE_CONNECTED:
                        break
                    else:
                        time.sleep(5)
                except pywintypes.error as e:
                    logging.warning(f"훈련 파이프 연결 시도 오류: {e}")
                    time.sleep(5)
            
            logging.info("훈련 파이프: 클라이언트 연결 성공")
            
            # 데이터 수신 루프
            while not stop_event.is_set():
                data_len = read_message_length(pipe, pipe_timeout)
                
                if data_len == 0:
                    cmd = read_command(pipe, 4, pipe_timeout)
                    if cmd == b'STOP':
                        logging.info("훈련 파이프: 종료 신호 수신")
                        stop_event.set()
                        break
                    continue
                
                # 크기 제한 체크
                if data_len > max_data_size:
                    logging.error(f"데이터 크기 초과: {data_len} > {max_data_size}")
                    response = pickle.dumps({
                        'request_id': None, 
                        'status': f"TRAINING_FAILED: 데이터 크기 초과 ({data_len} > {max_data_size})"
                    })
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    continue
                
                # 청크 수신
                total_chunks = struct.unpack('I', read_fixed_length(pipe, 4, pipe_timeout, "총 청크 수"))[0]
                command = read_command(pipe, 5, pipe_timeout)
                
                if command != b'TRAIN':
                    logging.error(f"잘못된 명령어: {command}")
                    break
                
                data = read_data_chunks(pipe, total_chunks, chunk_size, data_len, pipe_timeout)
                
                # 크기 검증
                if len(data) != data_len:
                    logging.error(f"데이터 크기 불일치: {len(data)} != {data_len}")
                    response = pickle.dumps({'request_id': None, 'status': "TRAINING_FAILED: 데이터 크기 불일치"})
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    continue
                
                # 훈련 실행
                try:
                    training_data = pickle.loads(data)
                    request_id = training_data.get('request_id')
                    X = np.asarray(training_data['X'])
                    y = training_data['y']
                    scale_pos_weight = training_data['scale_pos_weight']
                    partial_update = training_data['partial_update']
                    seq_length = training_data.get('seq_length', 5)
                    
                    logging.info(f"훈련 시작: samples={len(X)}, features={X.shape[1]}, seq_length={seq_length}")
                    
                    # 데이터 reshape
                    if X.ndim == 3:
                        X = X.reshape(X.shape[0], X.shape[1] * X.shape[2])
                    elif X.ndim != 2:
                        logging.error(f"잘못된 데이터 차원: {X.shape}")
                        response = pickle.dumps({'request_id': request_id, 'status': f"TRAINING_FAILED: 차원 오류 {X.shape}"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue
                    
                    # 데이터 검증
                    if np.any(np.isnan(X)) or np.any(np.isinf(X)):
                        logging.error("X에 NaN/Inf 존재")
                        response = pickle.dumps({'request_id': request_id, 'status': "TRAINING_FAILED: NaN/Inf 존재"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue
                    
                    # 피처 수 검증
                    expected_features = seq_length * (len(tick_features) + len(min_features))
                    if X.shape[1] != expected_features:
                        logging.error(f"피처 수 불일치: {X.shape[1]} != {expected_features}")
                        response = pickle.dumps({
                            'request_id': request_id, 
                            'status': f"TRAINING_FAILED: 피처 수 불일치 (expected {expected_features}, got {X.shape[1]})"
                        })
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue
                    
                    # 부분 업데이트 크기 제한
                    if partial_update and data_len > max_partial_update_size:
                        logging.error(f"부분 업데이트 크기 초과: {data_len}")
                        response = pickle.dumps({'request_id': request_id, 'status': "TRAINING_FAILED: 부분 업데이트 크기 초과"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue
                    
                    # 모델/스케일러 초기화
                    with model_scaler_lock:
                        if shared_model is None:
                            # 기존 모델 로드 시도
                            if os.path.exists(model_path) and os.path.exists(scaler_path):
                                try:
                                    loaded_model = keras.models.load_model(model_path, compile=False)
                                    loaded_model.compile(
                                        optimizer=keras.optimizers.Adam(learning_rate=0.001),
                                        loss='sparse_categorical_crossentropy',
                                        metrics=['accuracy']
                                    )
                                    loaded_scaler = IncrementalStandardScaler.load(scaler_path)
                                    shared_model = loaded_model
                                    shared_scaler = loaded_scaler
                                    logging.info("기존 모델 로드 완료")
                                except Exception as e:
                                    logging.error(f"기존 모델 로드 실패: {e}")
                                    shared_model = create_cnn_model((seq_length, len(tick_features) + len(min_features)))
                                    shared_scaler = IncrementalStandardScaler()
                            else:
                                logging.info("새 모델 생성")
                                shared_model = create_cnn_model((seq_length, len(tick_features) + len(min_features)))
                                shared_scaler = IncrementalStandardScaler()
                        
                        # 스케일링
                        if shared_scaler.n_samples_seen_ == 0:
                            X_scaled = shared_scaler.fit_transform(X)
                        else:
                            shared_scaler.partial_fit(X)
                            X_scaled = shared_scaler.transform(X)
                        
                        # Reshape
                        feature_count_per_seq = len(tick_features) + len(min_features)
                        X_scaled = X_scaled.reshape(-1, seq_length, feature_count_per_seq)
                        
                        # ✅ 시퀀스 가중치 (최근 데이터에 더 큰 가중치)
                        seq_weights = np.exp(np.linspace(0, 1, seq_length))
                        weight_matrix = np.repeat(seq_weights, feature_count_per_seq).reshape(
                            1, seq_length, feature_count_per_seq
                        )
                        X_scaled = X_scaled * weight_matrix
                        
                        # 훈련
                        val_ratio = 0.2
                        indices = np.arange(len(X_scaled))
                        np.random.shuffle(indices)
                        split_idx = int(len(X_scaled) * (1 - val_ratio))
                        train_idx, val_idx = indices[:split_idx], indices[split_idx:]
                        X_train, X_val = X_scaled[train_idx], X_scaled[val_idx]
                        y_train, y_val = np.array(y)[train_idx], np.array(y)[val_idx]
                        
                        # ✅ 클래스 가중치
                        unique_classes = np.unique(y_train)
                        if isinstance(scale_pos_weight, dict):
                            class_weight = scale_pos_weight
                        else:
                            class_weight = {0: 1.0, 1: scale_pos_weight}
                        
                        if partial_update:
                            shared_model.fit(
                                X_train, y_train,
                                epochs=1,
                                batch_size=32,
                                verbose=0,
                                class_weight=class_weight,
                                validation_data=(X_val, y_val)
                            )
                        else:
                            early_stop = keras.callbacks.EarlyStopping(
                                monitor="val_loss",
                                patience=3,
                                restore_best_weights=True,
                                verbose=1
                            )
                            shared_model.fit(
                                X_train, y_train,
                                epochs=50,
                                batch_size=32,
                                verbose=0,
                                class_weight=class_weight,
                                validation_data=(X_val, y_val),
                                callbacks=[early_stop]
                            )
                        
                        logging.info(f"훈련 완료: samples={len(X)}")
                        
                        # ✅ 최적 임계값 계산 (검증 데이터)
                        preds_val = shared_model.predict(X_val, verbose=0)
                        
                        # 3클래스인 경우 클래스 2 확률 사용
                        if preds_val.shape[1] == 3:
                            preds_val_prob = preds_val[:, 2]  # 강한 매수 확률
                        else:
                            preds_val_prob = preds_val.flatten()
                        
                        best_threshold = 0.5
                        best_f1 = 0.0
                        
                        for thr in np.arange(0.4, 0.81, 0.01):
                            # 이진 분류로 변환 (클래스 2 vs 나머지)
                            pred_labels = (preds_val_prob >= thr).astype(int)
                            y_val_binary = (y_val == 2).astype(int)
                            
                            tp = np.sum((pred_labels == 1) & (y_val_binary == 1))
                            fp = np.sum((pred_labels == 1) & (y_val_binary == 0))
                            fn = np.sum((pred_labels == 0) & (y_val_binary == 1))
                            
                            denom = 2 * tp + fp + fn
                            if denom == 0:
                                continue
                            
                            f1 = 2 * tp / denom
                            if f1 > best_f1:
                                best_f1 = f1
                                best_threshold = thr
                        
                        # 모델 저장
                        save_model_and_scaler(shared_model, shared_scaler, model_path, scaler_path, seq_length, tick_features, min_features)
                        current_seq_length = seq_length
                        
                        # 응답
                        response = pickle.dumps({
                            'request_id': request_id,
                            'status': "TRAINING_COMPLETED",
                            'best_threshold': float(best_threshold),
                            'f1': float(best_f1)
                        })
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        logging.info(f"훈련 응답 전송: threshold={best_threshold:.3f}, f1={best_f1:.3f}")
                    
                except Exception as ex:
                    logging.error(f"훈련 처리 실패: {ex}\n{traceback.format_exc()}")
                    response = pickle.dumps({'request_id': None, 'status': f"TRAINING_FAILED: {str(ex)}"})
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    
        except Exception as ex:
            logging.error(f"훈련 파이프 오류: {ex}\n{traceback.format_exc()}")
            time.sleep(1)

# ==================== 예측 파이프 처리 ====================

def handle_prediction_pipe(pipe, stop_event, tick_features, min_features):
    """예측 파이프 처리"""
    global shared_model, shared_scaler, current_seq_length
    
    pipe_timeout = 30
    
    while not stop_event.is_set():
        try:
            logging.debug("예측 파이프: 클라이언트 연결 대기")
            win32pipe.SetNamedPipeHandleState(pipe, win32pipe.PIPE_READMODE_MESSAGE, None, None)
            
            # 연결 대기
            while not stop_event.is_set():
                try:
                    result = win32pipe.ConnectNamedPipe(pipe, None)
                    if result == 0 or win32api.GetLastError() == winerror.ERROR_PIPE_CONNECTED:
                        break
                    else:
                        time.sleep(5)
                except pywintypes.error as e:
                    logging.warning(f"예측 파이프 연결 시도 오류: {e}")
                    time.sleep(5)
            
            logging.info("예측 파이프: 클라이언트 연결 성공")
            
            # 예측 루프
            while not stop_event.is_set():
                data_len = read_message_length(pipe, pipe_timeout)
                
                if stop_event.is_set():
                    break
                
                if data_len == 0:
                    cmd = read_command(pipe, 4, pipe_timeout)
                    if cmd == b'STOP':
                        logging.info("예측 파이프: 종료 신호 수신")
                        stop_event.set()
                        break
                
                command = read_command(pipe, 5, pipe_timeout)
                if command != b'PREDI':
                    logging.error(f"잘못된 명령어: {command}")
                    break
                
                data = read_fixed_length(pipe, data_len, pipe_timeout, "예측 데이터")
                
                try:
                    prediction_data = pickle.loads(data)
                    request_id = prediction_data.get('request_id')
                    X = prediction_data.get('data')
                    
                    with model_scaler_lock:
                        if shared_model is None or shared_scaler is None:
                            logging.error("모델/스케일러 미초기화")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': '모델 미초기화'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue
                        
                        # 피처 수 검증
                        expected_features = current_seq_length * (len(tick_features) + len(min_features))
                        if X.shape[1] != expected_features:
                            logging.error(f"피처 수 불일치: {X.shape[1]} != {expected_features}")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': f'피처 수 불일치 (expected {expected_features}, got {X.shape[1]})'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue
                        
                        # 데이터 검증
                        if np.any(np.isnan(X)) or np.any(np.isinf(X)):
                            logging.error("예측 데이터에 NaN/Inf 존재")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': 'NaN/Inf 존재'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue
                        
                        # 스케일링
                        X_scaled = shared_scaler.transform(X)
                        
                        # Reshape + 가중치
                        feature_count_per_seq = len(tick_features) + len(min_features)
                        seq_weights = np.exp(np.linspace(0, 1, current_seq_length))
                        weight_matrix = np.repeat(seq_weights, feature_count_per_seq).reshape(
                            1, current_seq_length, feature_count_per_seq
                        )
                        X_scaled = X_scaled.reshape(-1, current_seq_length, feature_count_per_seq)
                        X_scaled = X_scaled * weight_matrix
                        
                        # 예측
                        prediction = shared_model.predict(X_scaled, verbose=0)
                        
                        # 3클래스인 경우 클래스 2 확률 반환
                        if prediction.shape[1] == 3:
                            prediction_value = float(prediction[0][2])  # 강한 매수 확률
                        else:
                            prediction_value = float(prediction[0][0])
                        
                        response = pickle.dumps({
                            'request_id': request_id,
                            'prediction': prediction_value
                        })
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    
                except Exception as ex:
                    logging.error(f"예측 처리 실패: {ex}\n{traceback.format_exc()}")
                    response = pickle.dumps({
                        'request_id': request_id,
                        'prediction': 0.5,
                        'error': str(ex)
                    })
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    
        except Exception as ex:
            logging.error(f"예측 파이프 오류: {ex}\n{traceback.format_exc()}")
            time.sleep(1)

# ==================== 메인 ====================

def main():
    """메인 함수"""
    logging.info("CNN Server 메인 시작")
    
    global shared_model, shared_scaler, current_seq_length
    
    model_path = os.path.join(MODEL_DIR, 'cnn_model.keras')
    scaler_path = os.path.join(MODEL_DIR, 'cnn_scaler.pkl')
    
    training_pipe_name = r'\\.\pipe\CnnTrainingPipe'
    prediction_pipe_name = r'\\.\pipe\CnnPredictionPipe'
    
    # 기존 모델 로드 시도
    try:
        if os.path.exists(model_path) and os.path.exists(scaler_path):
            with model_scaler_lock:
                try:
                    shared_model = keras.models.load_model(model_path, compile=False)
                    shared_model.compile(
                        optimizer=keras.optimizers.Adam(learning_rate=0.001),
                        loss='sparse_categorical_crossentropy',
                        metrics=['accuracy']
                    )
                    shared_scaler = IncrementalStandardScaler.load(scaler_path)
                    current_seq_length = getattr(shared_model, 'seq_length', 5)
                    
                    logging.info(f"기존 모델 로드 완료: seq_length={current_seq_length}")
                except Exception as e:
                    logging.warning(f"모델 로드 실패: {e}")
                    shared_model = None
                    shared_scaler = None
        else:
            logging.info("모델 파일 없음, 새로 생성 예정")
    except Exception as ex:
        logging.error(f"모델 초기화 오류: {ex}\n{traceback.format_exc()}")
        raise
    
    # 특징 정의
    tick_features = [
        'C', 'V', 'MAT5', 'MAT20', 'MAT60', 'MAT120', 'RSIT', 'RSIT_SIGNAL',
        'MACDT', 'MACDT_SIGNAL', 'OSCT', 'STOCHK', 'STOCHD', 'ATR', 'CCI',
        'BB_UPPER', 'BB_MIDDLE', 'BB_LOWER', 'BB_POSITION', 'BB_BANDWIDTH',
        'MAT5_MAT20_DIFF', 'MAT20_MAT60_DIFF', 'MAT60_MAT120_DIFF',
        'C_MAT5_DIFF', 'MAT5_CHANGE', 'MAT20_CHANGE', 'MAT60_CHANGE', 
        'MAT120_CHANGE', 'VWAP'
    ]
    
    min_features = [
        'MAM5', 'MAM10', 'MAM20', 'RSI', 'RSI_SIGNAL', 'MACD', 'MACD_SIGNAL',
        'OSC', 'STOCHK', 'STOCHD', 'CCI', 'MAM5_MAM10_DIFF', 'MAM10_MAM20_DIFF',
        'C_MAM5_DIFF', 'C_ABOVE_MAM5', 'VWAP'
    ]
    
    stop_event = threading.Event()
    
    try:
        # 파이프 생성
        training_pipe = win32pipe.CreateNamedPipe(
            training_pipe_name,
            win32pipe.PIPE_ACCESS_DUPLEX,
            win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
            1, 52428800, 52428800, 0, None
        )
        
        prediction_pipe = win32pipe.CreateNamedPipe(
            prediction_pipe_name,
            win32pipe.PIPE_ACCESS_DUPLEX,
            win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
            1, 52428800, 52428800, 0, None
        )
        
        logging.info("파이프 생성 완료")
        
        # 스레드 시작
        training_thread = threading.Thread(
            target=handle_training_pipe,
            args=(training_pipe, model_path, scaler_path, stop_event, tick_features, min_features)
        )
        
        prediction_thread = threading.Thread(
            target=handle_prediction_pipe,
            args=(prediction_pipe, stop_event, tick_features, min_features)
        )
        
        training_thread.start()
        prediction_thread.start()
        
        logging.info("훈련/예측 스레드 시작")
        
        training_thread.join()
        prediction_thread.join()
        
        logging.info("서버 정상 종료")
        
    except Exception as ex:
        logging.error(f"서버 실행 오류: {ex}\n{traceback.format_exc()}")
        raise
    finally:
        stop_event.set()
        try:
            if 'training_pipe' in locals():
                win32file.CloseHandle(training_pipe)
            if 'prediction_pipe' in locals():
                win32file.CloseHandle(prediction_pipe)
        except Exception as ex:
            logging.error(f"파이프 닫기 오류: {ex}")

if __name__ == "__main__":
    main()