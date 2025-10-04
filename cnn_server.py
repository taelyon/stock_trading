import numpy as np
import pandas as pd
import tensorflow as tf
from tensorflow import keras
import pickle
import win32pipe
import win32file
import win32api
import winerror
from winerror import ERROR_MORE_DATA
import struct
import logging
import pywintypes
import os
import traceback
import threading
import time
from datetime import datetime

# Validate paths at startup
MODEL_DIR = r"C:\MyAPP\day_trading"
if not os.path.exists(MODEL_DIR):
    os.makedirs(MODEL_DIR)
if not os.access(MODEL_DIR, os.W_OK):
    raise PermissionError(f"No write access to {MODEL_DIR}")

# 로깅 설정
log_dir = os.path.join(MODEL_DIR, 'log')
log_file = os.path.join(log_dir, f"cnn_server_{datetime.now().strftime('%Y%m%d')}.log")
try:
    log_dir = os.path.dirname(log_file)
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    if not os.access(log_dir, os.W_OK):
        logging.warning(f"로그 디렉토리 {log_dir}에 쓰기 권한 없음, 현재 디렉토리로 변경")
        log_file = os.path.join(os.getcwd(), f"cnn_server_{datetime.now().strftime('%Y%m%d')}.log")
    
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='a'),
            logging.StreamHandler()
        ]
    )
    logging.info("로깅 설정 완료, 서버 시작")
except Exception as e:
    print(f"로깅 설정 실패: {e}")
    raise

# Reentrant lock ensures save_model_and_scaler can be called while already
# holding the lock in the training thread without causing a deadlock
model_scaler_lock = threading.RLock()

# shared state accessible across threads
shared_model = None
shared_scaler = None
current_seq_length = 5

def create_cnn_model(input_shape):
    model = keras.Sequential([
        keras.layers.Conv1D(32, 3, activation='relu', padding='same', input_shape=input_shape),
        keras.layers.Dropout(0.2),
        keras.layers.Conv1D(64, 3, activation='relu', padding='same'),
        keras.layers.Dropout(0.2),
        keras.layers.GRU(64),
        keras.layers.Dropout(0.2),
        keras.layers.Dense(64, activation='relu'),
        keras.layers.Dense(1, activation='sigmoid')
    ])
    model.compile(optimizer='adam', loss='binary_crossentropy', metrics=['accuracy'])
    return model

class IncrementalStandardScaler:
    def __init__(self):
        self.mean_ = None
        self.var_ = None
        self.n_samples_seen_ = 0

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
            raise ValueError("Scaler has not been fitted yet.")
        if X.shape[1] != len(self.mean_):
            raise ValueError(f"Feature dimension mismatch. Expected {len(self.mean_)}, got {X.shape[1]}")
        return (X - self.mean_) / np.sqrt(self.var_ + 1e-8)

    def fit_transform(self, X):
        return self.partial_fit(X).transform(X)

    def save(self, path):
        max_retries = 3
        retry_delay = 2
        for attempt in range(max_retries):
            try:
                with open(path, 'wb') as f:
                    pickle.dump({'mean': self.mean_, 'var': self.var_, 'n_samples_seen': self.n_samples_seen_}, f)
                logging.debug(f"Scaler saved successfully: {path}")
                return
            except OSError as e:
                logging.error(f"스케일러 저장 실패 (시도 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise

    @classmethod
    def load(cls, path):
        with open(path, 'rb') as f:
            state = pickle.load(f)
        scaler = cls()
        scaler.mean_ = state['mean']
        scaler.var_ = state['var']
        scaler.n_samples_seen_ = state['n_samples_seen']
        logging.debug(f"Scaler loaded successfully: {path}")
        return scaler

def generate_feature_names(seq_length, tick_features, min_features):
    feature_names = []
    for i in range(seq_length):
        feature_names.extend([f"{f}_tick_seq{i}" for f in tick_features])
        feature_names.extend([f"{f}_min_seq{i}" for f in min_features])
    return feature_names

def save_model_and_scaler(model, scaler, model_path, scaler_path, seq_length, tick_features, min_features):
    with model_scaler_lock:
        try:
            save_dir = os.path.dirname(model_path)
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
                logging.info(f"디렉토리 생성: {save_dir}")
            if not os.access(save_dir, os.W_OK):
                logging.error(f"디렉토리 쓰기 권한 없음: {save_dir}")
                raise PermissionError(f"디렉토리 쓰기 권한 없음: {save_dir}")
            else:
                logging.debug(f"디렉토리 쓰기 권한 확인 완료: {save_dir}")

            if not os.path.isabs(model_path):
                logging.warning(f"상대 경로 감지, 절대 경로로 변환: {model_path}")
                model_path = os.path.abspath(model_path)
            if not os.path.isabs(scaler_path):
                logging.warning(f"상대 경로 감지, 절대 경로로 변환: {scaler_path}")
                scaler_path = os.path.abspath(scaler_path)

            try:
                model.seq_length = seq_length
            except Exception:
                pass
            model.save(model_path)
            scaler.save(scaler_path)
            logging.info(f"모델 및 스케일러 저장 완료: {model_path}, {scaler_path}")
        except Exception as ex:
            logging.error(f"모델 및 스케일러 저장 실패: {ex}\n{traceback.format_exc()}")
            raise

def read_fixed_length(pipe, length, pipe_timeout, desc=""):
    start_time = time.time()
    buffer = b''
    while len(buffer) < length:
        if time.time() - start_time > pipe_timeout:
            raise TimeoutError(f"{desc} 읽기 타임아웃")
        result = win32file.ReadFile(pipe, length - len(buffer))
        if result[0] == winerror.ERROR_MORE_DATA:
            if desc:
                logging.debug(f"{desc}: ERROR_MORE_DATA 발생, 추가 읽기")
            buffer += result[1]
            continue
        elif result[0] != 0 or len(result[1]) == 0:
            raise Exception(f"{desc} 읽기 실패: {result}")
        buffer += result[1]
    return buffer

def read_message_length(pipe, pipe_timeout):
    return struct.unpack('I', read_fixed_length(pipe, 4, pipe_timeout, "데이터 길이"))[0]

def read_command(pipe, length, pipe_timeout):
    return read_fixed_length(pipe, length, pipe_timeout, "명령어")

def read_data_chunks(pipe, total_chunks, chunk_size, data_len, pipe_timeout):
    data = b''
    received_chunks = set()
    for chunk_idx in range(total_chunks):
        chunk_header_buffer = read_fixed_length(pipe, 4, pipe_timeout, f"청크 {chunk_idx} 헤더")
        chunk_idx_received = struct.unpack('I', chunk_header_buffer)[0]
        logging.debug(f"훈련 파이프: 청크 {chunk_idx_received}/{total_chunks} 수신 시작")

        if chunk_idx_received in received_chunks:
            logging.error(f"중복 청크 수신: {chunk_idx_received}")
            continue
        received_chunks.add(chunk_idx_received)

        chunk_size_to_read = min(chunk_size, data_len - (chunk_idx_received * chunk_size))
        chunk_data = read_fixed_length(pipe, chunk_size_to_read, pipe_timeout, f"청크 {chunk_idx_received} 데이터")
        data += chunk_data
        logging.debug(f"훈련 파이프: 청크 {chunk_idx_received}/{total_chunks} 수신 완료, 크기: {len(chunk_data)} 바이트")

        win32file.WriteFile(pipe, struct.pack('I', chunk_idx_received))
        logging.debug(f"훈련 파이프: 청크 {chunk_idx_received} 수신확인 응답 전송")
    return data

def handle_training_pipe(pipe, model_path, scaler_path, stop_event, tick_features, min_features):
    global shared_model, shared_scaler, current_seq_length
    chunk_size = 65536
    max_buffer_size = 52428800
    max_data_size = 52428800
    max_partial_update_size = 10485760
    pipe_timeout = 30  # 파이프 읽기 타임아웃 (초)
    while not stop_event.is_set():
        try:
            logging.debug("훈련 파이프: 클라이언트 연결 대기")
            win32pipe.SetNamedPipeHandleState(pipe, win32pipe.PIPE_READMODE_MESSAGE, None, None)
            while not stop_event.is_set():
                try:
                    result = win32pipe.ConnectNamedPipe(pipe, None)
                    if result == 0 or win32api.GetLastError() == winerror.ERROR_PIPE_CONNECTED:
                        break
                    else:
                        error_code = win32api.GetLastError()
                        logging.warning(f"훈련 파이프 연결 시도 실패, 에러 코드: {error_code}, 5초 후 재시도")
                        time.sleep(5)
                except pywintypes.error as e:
                    logging.warning(f"훈련 파이프 연결 시도 중 예외 발생: {e}, 5초 후 재시도")
                    time.sleep(5)
            logging.info("훈련 파이프: 클라이언트 연결 성공")

            while not stop_event.is_set():
                data_len = read_message_length(pipe, pipe_timeout)
                logging.debug(f"훈련 파이프: 수신 데이터 길이 {data_len} 바이트, 최대 버퍼 {max_buffer_size} 바이트")
                if data_len == 0:
                    cmd = read_command(pipe, 4, pipe_timeout)
                    if cmd == b'STOP':
                        logging.info("훈련 파이프: 종료 신호 수신")
                        stop_event.set()
                        break
                    continue

                if data_len > max_data_size:
                    logging.error(f"훈련 파이프: 데이터 크기 제한 초과: {data_len} 바이트 (최대 {max_data_size} 바이트)")
                    response = pickle.dumps({'request_id': None, 'status': f"TRAINING_FAILED: Data size {data_len} exceeds {max_data_size} bytes limit"})
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    continue

                total_chunks = struct.unpack('I', read_fixed_length(pipe, 4, pipe_timeout, "총 청크 수"))[0]
                logging.debug(f"훈련 파이프: 총 청크 수: {total_chunks}")

                command = read_command(pipe, 5, pipe_timeout)
                if command != b'TRAIN':
                    logging.error(f"훈련 파이프: 잘못된 명령어 {command}, 연결 종료")
                    break

                data = read_data_chunks(pipe, total_chunks, chunk_size, data_len, pipe_timeout)

                if len(data) != data_len:
                    logging.error(f"데이터 크기 불일치: 예상 {data_len}, 실제 {len(data)}")
                    response = pickle.dumps({'request_id': None, 'status': "TRAINING_FAILED: Data size mismatch"})
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    continue

                try:
                    logging.debug("훈련 데이터 디코딩 시작")
                    training_data = pickle.loads(data)
                    request_id = training_data.get('request_id')
                    X = np.asarray(training_data['X'])
                    y = training_data['y']
                    scale_pos_weight = training_data['scale_pos_weight']
                    partial_update = training_data['partial_update']
                    seq_length = training_data.get('seq_length', 5)

                    if X.ndim == 3:
                        X = X.reshape(X.shape[0], X.shape[1] * X.shape[2])
                    elif X.ndim != 2:
                        logging.error(f"잘못된 훈련 데이터 차원: {X.shape}")
                        response = pickle.dumps({'request_id': request_id, 'status': f"TRAINING_FAILED: Invalid training data dimension {X.shape}"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue

                    logging.info(
                        f"훈련 시작: request_id={request_id}, samples={len(X)}, features={X.shape[1]}, partial_update={partial_update}, seq_length={seq_length}"
                    )

                    # 데이터 유효성 검증
                    if np.any(np.isnan(X)) or np.any(np.isinf(X)):
                        logging.error(f"훈련 데이터에 NaN 또는 Inf 값 존재: shape={X.shape}")
                        response = pickle.dumps({'request_id': request_id, 'status': "TRAINING_FAILED: Invalid training data (NaN/Inf)"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue
                    if len(y) != len(X):
                        logging.error(f"레이블 데이터 크기 불일치: X={len(X)}, y={len(y)}")
                        response = pickle.dumps({'request_id': request_id, 'status': "TRAINING_FAILED: Label size mismatch"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue

                    expected_features = len(generate_feature_names(seq_length, tick_features, min_features))
                    logging.debug(f"예상 피처 수: {expected_features}")
                    if X.shape[1] != expected_features:
                        logging.error(f"훈련 데이터 피처 수 불일치: 예상 {expected_features}, found {X.shape[1]}")
                        response = pickle.dumps({'request_id': request_id, 'status': f"TRAINING_FAILED: Feature count mismatch: expected {expected_features}, got {X.shape[1]}"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue

                    if partial_update and data_len > max_partial_update_size:
                        logging.error(f"부분 업데이트에서 데이터 크기 {data_len} 바이트 초과 (최대 {max_partial_update_size} 바이트)")
                        response = pickle.dumps({'request_id': request_id, 'status': f"TRAINING_FAILED: Data size exceeds {max_partial_update_size//1024}KB limit for partial update"})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        continue

                    with model_scaler_lock:
                        if shared_model is not None and hasattr(shared_model, 'input_shape') and (
                            shared_model.input_shape[1] * shared_model.input_shape[2] != expected_features
                        ):
                            logging.warning(
                                f"기존 모델 피처 수 불일치: 예상 {expected_features}, "
                                f"실제 {shared_model.input_shape[1] * shared_model.input_shape[2]}. 새 모델 초기화"
                            )
                            shared_model = None
                            shared_scaler = IncrementalStandardScaler()

                        if shared_model is None:
                            load_success = False
                            if os.path.exists(model_path) and os.path.exists(scaler_path):
                                try:
                                    loaded_model = keras.models.load_model(model_path, compile=False)
                                    loaded_model.compile(optimizer='adam', loss='binary_crossentropy', metrics=['accuracy'])
                                    loaded_scaler = IncrementalStandardScaler.load(scaler_path)
                                    shared_model = loaded_model
                                    shared_scaler = loaded_scaler
                                    logging.info("기존 모델 로드 후 업데이트 진행")
                                    load_success = True
                                except Exception as e:
                                    logging.error(f"기존 모델 로드 실패: {e}. 새 모델 초기화")
                                    shared_model = create_cnn_model((seq_length, len(tick_features) + len(min_features)))
                                    shared_scaler = IncrementalStandardScaler()
                            else:
                                logging.info("새로운 CNN 모델 초기화")
                                shared_model = create_cnn_model((seq_length, len(tick_features) + len(min_features)))
                                shared_scaler = IncrementalStandardScaler()

                            logging.debug("스케일러 피팅 시작")
                            if load_success:
                                shared_scaler.partial_fit(X)
                                X_scaled = shared_scaler.transform(X)
                            else:
                                X_scaled = shared_scaler.fit_transform(X)
                        else:
                            if not partial_update:
                                logging.info("기존 CNN 모델 전체 데이터로 업데이트")
                            else:
                                logging.debug("기존 스케일러에 부분 학습 및 데이터 변환")
                            shared_scaler.partial_fit(X)
                            X_scaled = shared_scaler.transform(X)

                        feature_count_per_seq = len(tick_features) + len(min_features)
                        seq_weights = np.exp(np.linspace(0, 1, seq_length))
                        weight_matrix = np.repeat(seq_weights, feature_count_per_seq).reshape(
                            1, seq_length, feature_count_per_seq
                        )
                        X_scaled = X_scaled.reshape(-1, seq_length, feature_count_per_seq)
                        X_scaled = X_scaled * weight_matrix
                            

                        logging.debug(
                            f"모델 학습 시작: X_scaled shape={X_scaled.shape}, y shape={y.shape}"
                        )
                        class_weight = {0: 1.0, 1: scale_pos_weight}
                        val_ratio = 0.2
                        indices = np.arange(len(X_scaled))
                        np.random.shuffle(indices)
                        split_idx = int(len(X_scaled) * (1 - val_ratio))
                        train_idx, val_idx = indices[:split_idx], indices[split_idx:]
                        X_train, X_val = X_scaled[train_idx], X_scaled[val_idx]
                        y_train, y_val = np.array(y)[train_idx], np.array(y)[val_idx]

                        if partial_update:
                            shared_model.fit(
                                X_train,
                                y_train,
                                epochs=1,
                                batch_size=32,
                                verbose=0,
                                class_weight=class_weight,
                                validation_data=(X_val, y_val),
                            )
                        else:
                            early_stop = keras.callbacks.EarlyStopping(
                                monitor="val_loss",
                                patience=3,
                                restore_best_weights=True,
                                verbose=1,
                            )
                            shared_model.fit(
                                X_train,
                                y_train,
                                epochs=50,
                                batch_size=32,
                                verbose=0,
                                class_weight=class_weight,
                                validation_data=(X_val, y_val),
                                callbacks=[early_stop],
                            )
                        logging.info(f"훈련 완료: samples={len(X)}, features={X.shape[1]}")

                        # 성능 평가 및 최적 임계치 계산 (검증 데이터 사용)
                        preds_val = shared_model.predict(X_val, verbose=0).flatten()
                        best_threshold = 0.5
                        best_f1 = 0.0
                        for thr in np.arange(0.4, 0.81, 0.01):
                            pred_labels = (preds_val >= thr).astype(int)
                            tp = np.sum((pred_labels == 1) & (y_val == 1))
                            fp = np.sum((pred_labels == 1) & (y_val == 0))
                            fn = np.sum((pred_labels == 0) & (y_val == 1))
                            denom = 2 * tp + fp + fn
                            if denom == 0:
                                continue
                            f1 = 2 * tp / denom
                            if f1 > best_f1:
                                best_f1 = f1
                                best_threshold = thr

                        save_model_and_scaler(shared_model, shared_scaler, model_path, scaler_path, seq_length, tick_features, min_features)

                        # update globals so prediction thread sees new model
                        current_seq_length = seq_length

                        response = pickle.dumps({'request_id': request_id, 'status': "TRAINING_COMPLETED", 'best_threshold': float(best_threshold), 'f1': float(best_f1)})
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        logging.info(f"훈련 응답 전송 완료: request_id={request_id}, best_threshold={best_threshold:.3f}, f1={best_f1:.3f}")

                except Exception as ex:
                    logging.error(f"훈련 처리 실패: {ex}\n{traceback.format_exc()}")
                    response = pickle.dumps({
                        'request_id': request_id,
                        'status': f"TRAINING_FAILED: {str(ex)}"
                    })
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                    continue

        except Exception as ex:
            logging.error(f"훈련 파이프 처리 오류: {ex}\n{traceback.format_exc()}")
            time.sleep(1)

def handle_prediction_pipe(pipe, stop_event, tick_features, min_features):
    global shared_model, shared_scaler, current_seq_length
    pipe_timeout = 30
    while not stop_event.is_set():
        try:
            logging.debug("예측 파이프: 클라이언트 연결 대기")
            win32pipe.SetNamedPipeHandleState(pipe, win32pipe.PIPE_READMODE_MESSAGE, None, None)
            while not stop_event.is_set():
                try:
                    result = win32pipe.ConnectNamedPipe(pipe, None)
                    if result == 0 or win32api.GetLastError() == winerror.ERROR_PIPE_CONNECTED:
                        break
                    else:
                        error_code = win32api.GetLastError()
                        logging.warning(f"예측 파이프 연결 시도 실패, 에러 코드: {error_code}, 5초 후 재시도")
                        time.sleep(5)
                except pywintypes.error as e:
                    logging.warning(f"예측 파이프 연결 시도 중 예외 발생: {e}, 5초 후 재시도")
                    time.sleep(5)
            logging.info("예측 파이프: 클라이언트 연결 성공")

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
                    logging.error(f"예측 파이프: 잘못된 명령어 {command}, 연결 종료")
                    break

                data = read_fixed_length(pipe, data_len, pipe_timeout, "데이터")

                try:
                    prediction_data = pickle.loads(data)
                    request_id = prediction_data.get('request_id')
                    X = prediction_data.get('data')
                    # logging.debug(f"예측 요청 수신: request_id={request_id}, 데이터 shape={X.shape}")

                    with model_scaler_lock:
                        if shared_model is None or shared_scaler is None:
                            logging.error("예측 실패: 모델 또는 스케일러 초기화되지 않음")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': 'Model or scaler not initialized'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue

                        expected_features = len(generate_feature_names(current_seq_length, tick_features, min_features))
                        if X.shape[1] != expected_features:
                            logging.error(f"예측 데이터 피처 수 불일치: 예상 {expected_features}, 실제 {X.shape[1]}")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': f'Feature count mismatch: expected {expected_features}, got {X.shape[1]}'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue

                        if np.any(np.isnan(X)) or np.any(np.isinf(X)):
                            logging.error(f"예측 데이터에 NaN 또는 Inf 값 존재: shape={X.shape}")
                            response = pickle.dumps({
                                'request_id': request_id,
                                'prediction': 0.5,
                                'error': 'Invalid prediction data (NaN/Inf)'
                            })
                            win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                            continue

                        X_scaled = shared_scaler.transform(X)
                        feature_count_per_seq = len(tick_features) + len(min_features)
                        seq_weights = np.exp(np.linspace(0, 1, current_seq_length))
                        weight_matrix = np.repeat(seq_weights, feature_count_per_seq).reshape(
                            1, current_seq_length, feature_count_per_seq
                        )
                        X_scaled = X_scaled.reshape(-1, current_seq_length, feature_count_per_seq)
                        X_scaled = X_scaled * weight_matrix
                        prediction = float(shared_model.predict(X_scaled)[0][0])
                        # logging.info(f"예측 완료: request_id={request_id}, 확률={prediction:.4f}")

                        response = pickle.dumps({
                            'request_id': request_id,
                            'prediction': prediction
                        })
                        win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)
                        # logging.debug(f"예측 응답 전송 완료: request_id={request_id}")

                except Exception as ex:
                    logging.error(f"예측 처리 실패: {ex}\n{traceback.format_exc()}")
                    response = pickle.dumps({
                        'request_id': request_id,
                        'prediction': 0.5,
                        'error': str(ex)
                    })
                    win32file.WriteFile(pipe, struct.pack('I', len(response)) + response)

        except Exception as ex:
            logging.error(f"예측 파이프 처리 오류: {ex}\n{traceback.format_exc()}")
            time.sleep(1)

def main():
    logging.info("Starting CNN server...")
    global shared_model, shared_scaler, current_seq_length
    model_path = os.path.join(MODEL_DIR, 'cnn_model.keras')
    scaler_path = os.path.join(MODEL_DIR, 'cnn_scaler.pkl')
    training_pipe_name = r'\\.\pipe\CnnTrainingPipe'
    prediction_pipe_name = r'\\.\pipe\CnnPredictionPipe'

    try:
        if os.path.exists(model_path) and os.path.exists(scaler_path):
            with model_scaler_lock:
                try:
                    shared_model = keras.models.load_model(model_path, compile=False)
                    shared_model.compile(optimizer='adam', loss='binary_crossentropy', metrics=['accuracy'])
                    shared_scaler = IncrementalStandardScaler.load(scaler_path)
                    current_seq_length = getattr(shared_model, 'seq_length', 5)
                    input_features = shared_model.input_shape[1] * shared_model.input_shape[2]
                    logging.info(f"기존 모델 및 스케일러 로드 성공: features={input_features}, seq_length={current_seq_length}")
                except Exception as e:
                    logging.warning(f"모델 또는 스케일러 로드 실패: {e}. 새 모델로 시작")
                    shared_model = None
                    shared_scaler = None
        else:
            logging.info("모델 또는 스케일러 파일 없음, 새 모델로 시작")

    except Exception as ex:
        logging.error(f"모델 초기화 오류: {ex}\n{traceback.format_exc()}")
        raise

    tick_features = [
        'C', 'V', 'MAT5', 'MAT20', 'MAT60', 'MAT120', 'RSIT', 'RSIT_SIGNAL',
        'MACDT', 'MACDT_SIGNAL', 'OSCT', 'STOCHK', 'STOCHD', 'ATR', 'CCI',
        'BB_UPPER', 'BB_MIDDLE', 'BB_LOWER', 'BB_POSITION', 'BB_BANDWIDTH',
        'MAT5_MAT20_DIFF', 'MAT20_MAT60_DIFF', 'MAT60_MAT120_DIFF',
        'C_MAT5_DIFF', 'MAT5_CHANGE', 'MAT20_CHANGE', 'MAT60_CHANGE', 'MAT120_CHANGE',
        'VWAP'
    ]
    min_features = [
        'MAM5', 'MAM10', 'MAM20', 'RSI', 'RSI_SIGNAL', 'MACD', 'MACD_SIGNAL',
        'OSC', 'STOCHK', 'STOCHD', 'CCI', 'MAM5_MAM10_DIFF', 'MAM10_MAM20_DIFF',
        'C_MAM5_DIFF', 'C_ABOVE_MAM5', 'VWAP'
    ]

    stop_event = threading.Event()

    try:
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
        logging.info("명명된 파이프 생성 완료")

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
        logging.info("훈련 및 예측 스레드 시작")

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
                logging.debug("훈련 파이프 종료")
            if 'prediction_pipe' in locals():
                win32file.CloseHandle(prediction_pipe)
                logging.debug("예측 파이프 종료")
        except Exception as ex:
            logging.error(f"파이프 닫기 오류: {ex}")

if __name__ == "__main__":
    main()