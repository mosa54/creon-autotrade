import sys
from typing import Optional, Tuple, Dict, List
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread, QObject
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView, QCheckBox, QHBoxLayout, QVBoxLayout,
    QSpinBox, QComboBox, QSplitter, QScrollArea, QFrame, QGroupBox,
    QGridLayout, QTabWidget, QProgressBar, QTextEdit, QLineEdit,
    QDoubleSpinBox, QSpacerItem, QSizePolicy, QMessageBox, QFileDialog,
    QInputDialog, QCompleter, QDialog
)
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon
from datetime import datetime
import json
import os
import logging

try:
    import win32com.client
    import pythoncom
except ImportError:
    logging.critical(
        "win32com.client 또는 pythoncom 모듈을 사용할 수 없습니다. "
        "이 프로그램은 Windows에서만 동작하며 Creon Plus와 pywin32가 설치되어 있어야 합니다."
    )
    if QApplication.instance() is not None:
        QMessageBox.critical(
            None,
            "환경 오류",
            "win32com.client 또는 pythoncom 모듈을 사용할 수 없습니다.\n"
            "Windows용 pywin32와 Creon Plus가 설치되어 있는지 확인하세요.",
        )
    else:
        print(
            "[오류] win32com.client 또는 pythoncom 모듈을 불러올 수 없습니다. "
            "Windows에서만 실행 가능합니다.",
            file=sys.stderr,
        )
    sys.exit(1)
import threading
import queue
import time
com_lock = threading.Lock()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ComContext:
    def __enter__(self):
        pythoncom.CoInitialize()
    def __exit__(self, exc_type, exc_val, exc_tb):
        pythoncom.CoUninitialize()

# Lock을 이용해 스레드 안전 처리
from threading import Lock

# COM 오류를 잡기 위한 예외 클래스
from pywintypes import com_error

# 교체할 클래스: CreonManager
class CreonManager(QObject):
    # <<< [추가] UI 업데이트를 위한 시그널
    ui_update_signal = pyqtSignal(dict)

    def _get_trade_value_corrected(self, code: str, raw_trade_value: int) -> int:
        """
        [수정됨] 거래대금 단위를 보정합니다.
        실시간 이벤트 핸들러 내에서 COM 객체를 매번 생성하는 대신,
        초기화된 객체를 사용하고 결과를 캐싱하여 안정성과 속도를 향상시킵니다.
        """
        # 1. 캐시에서 시장 정보 조회
        if code in self.market_cache:
            market = self.market_cache[code]
        else:
            # 2. 캐시에 없으면 스레드에 안전한 방식으로 COM 객체를 통해 조회하고 결과를 캐시에 저장
            try:
                # com_lock을 사용해 여러 스레드에서의 동시 접근을 방지합니다.
                with com_lock:
                    # NOTE: COM 객체는 호출되는 스레드에서 생성해야 하므로 여기서 새로 생성한다
                    local_mgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
                    market = local_mgr.GetMarketKind(code)
                    self.market_cache[code] = market
            except Exception as e:
                logging.error(f"[{code}] 시장 구분 정보 조회 실패: {e}")
                # 조회 실패 시 원본 값을 그대로 반환하여 오류를 방지합니다.
                return raw_trade_value

        # 3. 시장 구분에 따라 거래대금 단위 보정
        if market == 1:       # 코스피(거래소)
            return raw_trade_value * 10_000
        elif market == 2:     # 코스닥
            return raw_trade_value * 1_000
        else:                 # 기타(K-OTC, 채권 등)
            return raw_trade_value

    def __init__(self):
        super().__init__()
        self.cp_cybos = None
        self.cp_util = None
        self.cp_order = None
        self.cp_stock = None
        self.stock_chart = None
        self.cp_code_mgr = None
        self.account = None
        self.acc_flag = None
        self.is_initialized = False
        self.market_cache = {}
        self.sub_lock = Lock()
        
        # <<< [추가] 실시간 구독 관리를 위한 딕셔너리 { stock_code: com_object }
        self.realtime_subscribers = {}

    def initialize(self) -> bool:
        try:
            pythoncom.CoInitialize() # 메인 스레드 COM 초기화
            self.cp_cybos = win32com.client.Dispatch("CpUtil.CpCybos")
            if self.cp_cybos.IsConnect != 1:
                logging.error("Creon Plus에 연결되지 않았습니다.")
                return False

            self.cp_util = win32com.client.Dispatch("CpTrade.CpTdUtil")
            self.cp_util.TradeInit(0)
            self.cp_order = win32com.client.Dispatch("CpTrade.CpTd0311")
            self.cp_stock = win32com.client.Dispatch("DsCbo1.StockMst")
            self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
            self.cp_code_mgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

            accounts = self.cp_util.AccountNumber
            if not accounts:
                logging.error("사용 가능한 계좌가 없습니다.")
                return False
            self.account = accounts[0]
            goods_list = self.cp_util.GoodsList(self.account, 1)
            if not goods_list:
                logging.error("주식 계좌 정보를 찾을 수 없습니다.")
                return False
            self.acc_flag = goods_list[0]

            self.is_initialized = True
            logging.info("Creon 초기화 완료")
            return True
        except Exception as e:
            logging.error(f"Creon 초기화 실패: {e}")
            return False

    # <<< [추가] 실시간 데이터 수신을 위한 이벤트 핸들러 클래스
    class CpEventClass:
        
        def __init__(self):
            self.manager = None
            self.obj     = None

        def set_manager(self, manager):
            """manager 객체를 설정하는 메서드"""
            self.manager = manager

        def OnReceived(self):
            # COM 초기화/해제를 자동으로 1:1로 처리
            with ComContext():
                logging.info(f"[{self.obj.GetHeaderValue(0)}] OnReceived triggered")
                try:
                    if not self.manager:
                        logging.error("Manager가 초기화되지 않았습니다.")
                        return

                    code = self.obj.GetHeaderValue(0)
                    data = {
                        "code":          code,
                        "current_price": self.obj.GetHeaderValue(13),
                        "high_price":    self.obj.GetHeaderValue(5),
                        "low_price":     self.obj.GetHeaderValue(6),
                        "volume":        self.obj.GetHeaderValue(9),
                        "trade_value":   self.manager._get_trade_value_corrected(
                                            code,
                                            self.obj.GetHeaderValue(10)
                                        ),
                        "tick_time":     self.obj.GetHeaderValue(18),
                    }

                    # 큐에 안전하게 삽입
                    try:
                        if code in self.manager.realtime_subscribers:
                            q = self.manager.realtime_subscribers[code]['queue']
                            q.put_nowait(data)
                            logging.info(f"[{code}] 데이터 삽입 – 큐 사이즈: {q.qsize()}")
                    except queue.Full:
                        logging.warning(f"[{code}] 데이터 누락 – 큐가 가득 찼습니다.")

                    # UI 업데이트
                    self.manager.ui_update_signal.emit(data)

                except com_error as ce:
                    logging.error(f"COM 예외 발생({code}): {ce}")
                except Exception as e:
                    logging.error(f"실시간 데이터 처리 중 예외: {e}")

    # <<< [추가] 실시간 시세 구독 메소드
    def subscribe_realtime(self, code: str, data_queue: queue.Queue) -> bool:
        """실시간 시세 구독을 시작합니다."""
        with self.sub_lock:
            if not self.is_initialized:
                return False
            if code in self.realtime_subscribers:
                self.unsubscribe_realtime(code)

            try:
                # 1) COM 객체 생성 및 이벤트 핸들러 연결
                obj = win32com.client.Dispatch("DsCbo1.StockCur")
                handler = win32com.client.WithEvents(obj, self.CpEventClass)
                handler.obj = obj
                handler.set_manager(self)

                # 2) 종목 코드 설정 및 구독 시작
                obj.SetInputValue(0, code)
                obj.Subscribe()

                # 3) 구독 정보 저장
                self.realtime_subscribers[code] = {
                    "obj": obj,
                    "handler": handler,
                    "queue": data_queue
                }
                logging.info(f"[{code}] 실시간 시세 구독 시작")

                # 4) 연속 구독 방지
                time.sleep(0.1)
                return True

            except Exception as e:
                logging.error(f"[{code}] 실시간 시세 구독 실패: {e}")
                return False

    # <<< [추가] 실시간 시세 구독 해지 메소드
    def unsubscribe_realtime(self, code: str):
        with self.sub_lock:
            if code in self.realtime_subscribers:
                self.realtime_subscribers[code]['obj'].Unsubscribe()
                del self.realtime_subscribers[code]
                logging.info(f"[{code}] 실시간 시세 구독 해지")

    # 기존 메소드들은 그대로 유지 (get_stock_info, get_stock_name 등)
    # ... (기존 CreonManager의 다른 메소드들은 여기에 그대로 복사) ...
    def get_safe_close_price(self, code: str, days: int = 3) -> int:
        """전일종가가 0이면, 과거 N일간 0이 아닌 마지막 종가 반환"""
        if not self.is_initialized:
            return 0
        with com_lock:
            if not code.startswith("A"):
                code = "A" + code
            try:
                chart = win32com.client.Dispatch("CpSysDib.StockChart")
                chart.SetInputValue(0, code)
                chart.SetInputValue(1, ord('2'))     # 기간
                chart.SetInputValue(2, days)         # 최근 days개
                chart.SetInputValue(3, ord('2'))     # 종가만
                chart.SetInputValue(5, [0])          # 0: 종가
                chart.SetInputValue(6, ord('D'))     # 일봉
                chart.SetInputValue(9, ord('1'))     # 수정주가
                chart.BlockRequest()
                cnt = chart.GetHeaderValue(3)
                for i in range(cnt):
                    close = chart.GetDataValue(0, i)
                    if close > 0:
                        return close
                return 0
            except Exception as e:
                logging.error(f"종가 조회 오류({code}): {e}")
                return 0

    def get_stock_name(self, code: str) -> str:
        if not self.is_initialized:
            return ""
        return self.cp_code_mgr.CodeToName(code)

    def get_stock_info(self, code: str):
        if not self.is_initialized:
            return None
        with com_lock:
            if not code.startswith("A"):
                code = "A" + code
            stock = win32com.client.Dispatch("DsCbo1.StockMst")
            stock.SetInputValue(0, code)
            stock.BlockRequest()

            # StockCur가 주는 데이터와 필드명을 최대한 일치시킴
            return {
                "code": code,
                "current_price": stock.GetHeaderValue(11),
                "high_price": stock.GetHeaderValue(14),
                "low_price": stock.GetHeaderValue(15),
                "close_price": stock.GetHeaderValue(11), # 종가
                "volume": stock.GetHeaderValue(18),
                "trade_value": stock.GetHeaderValue(19)
            }

    def get_stock_balance_and_avg_price(self, code: str) -> Tuple[int, int]:
        try:
            if not self.is_initialized: return 0, 0
            
            with com_lock:
                obj = win32com.client.Dispatch("CpTrade.CpTd6033")
                obj.SetInputValue(0, self.account)
                obj.SetInputValue(1, self.acc_flag)
                obj.BlockRequest()

            cnt = obj.GetHeaderValue(7)
            for i in range(cnt):
                full_code = obj.GetDataValue(12, i) # A005930
                if code == full_code or "A" + code == full_code:
                    qty = obj.GetDataValue(7, i)
                    avg_price = obj.GetDataValue(17, i)
                    return qty, avg_price
            return 0, 0
        except Exception as e:
            logging.error(f"평균단가/잔고 조회 오류({code}): {e}")
            return 0, 0

    def get_high_price_for_days(self, code: str, days: int) -> int:
        """최근 N일간(당일제외) 고가 중 최고값 반환"""
        if not self.is_initialized:
            return 0
        with com_lock:
            if not code.startswith("A"):
                code = "A" + code
            try:
                chart = win32com.client.Dispatch("CpSysDib.StockChart")
                chart.SetInputValue(0, code)
                chart.SetInputValue(1, ord('2'))     # 기간으로 요청
                chart.SetInputValue(2, days + 1)     # N+1개(오늘 포함)
                chart.SetInputValue(3, ord('1'))     # 1: 고가만
                chart.SetInputValue(5, [2])          # 2: 고가
                chart.SetInputValue(6, ord('D'))     # 일봉
                chart.SetInputValue(9, ord('1'))     # 수정주가
                chart.BlockRequest()
                cnt = chart.GetHeaderValue(3)
                highs = [chart.GetDataValue(0, i) for i in range(cnt)]
                if len(highs) > 1:
                    return max(highs[1:days+1])
                return 0
            except Exception as e:
                logging.error(f"N일고점 조회 오류({code}): {e}")
                return 0
            
    def get_low_price_for_days(self, code: str, days: int) -> int:
        """최근 N일간(당일제외) 저가 중 최저값 반환"""
        if not self.is_initialized:
            return 0
        with com_lock:
            if not code.startswith("A"):
                code = "A" + code
            try:
                chart = win32com.client.Dispatch("CpSysDib.StockChart")
                chart.SetInputValue(0, code)
                chart.SetInputValue(1, ord('2'))         # 기간 기준
                chart.SetInputValue(2, days + 1)         # 오늘 포함 N+1일 조회
                chart.SetInputValue(3, ord('3'))         # 요청 필드: 저가
                chart.SetInputValue(5, [3])              # 필드 코드: 저가
                chart.SetInputValue(6, ord('D'))         # 일봉
                chart.SetInputValue(9, ord('1'))         # 수정주가
                chart.BlockRequest()

                cnt = chart.GetHeaderValue(3)
                lows = [chart.GetDataValue(0, i) for i in range(cnt)]
                if len(lows) > 1:
                    return min(lows[1:days+1])  # 오늘 제외한 N일 중 최저가
                return 0
            except Exception as e:
                logging.error(f"N일저점 조회 오류({code}): {e}")
                return 0

    def place_order(self, stock_code, qty, price, is_buy=True) -> bool:
        logging.info(f"주문 요청: {stock_code} / 수량: {qty} / 가격: {price} / {'매수' if is_buy else '매도'}")
        try:
            with com_lock:
                o = win32com.client.Dispatch("CpTrade.CpTd0311")
                o.SetInputValue(0, "2" if is_buy else "1")  # 1:매도, 2:매수
                o.SetInputValue(1, self.account)
                o.SetInputValue(2, self.acc_flag)
                o.SetInputValue(3, stock_code)
                o.SetInputValue(4, qty)
                o.SetInputValue(5, 0)
                o.SetInputValue(7, "1") # 0:보통, 1:IOC, 2:FOK
                o.SetInputValue(8, "03") # 01:지정가, 03:시장가
                ret = o.BlockRequest()
                time.sleep(0.3) # 연속 주문 시 API 차단을 피하기 위한 지연시간
                if ret != 0:
                    logging.error(f"주문 실패. 응답 코드: {ret}, 종목: {stock_code}, 수량: {qty}")
                    # 추가적인 오류 정보 조회
                    msg = win32com.client.Dispatch("CpUtil.CpCybos").GetDibStatus()
                    if msg:
                        logging.error(f"주문 실패 메시지: {msg}")
                    return False
                logging.info(f"시장가 IOC 주문 성공. 종목: {stock_code}, 수량: {qty}")
                return True
        except Exception as e:
            logging.error(f"주문 실행 중 예외 발생({stock_code}): {e}")
            return False

# Stylesheet definition
DARK_STYLE = """
QMainWindow {
    background-color: #1e1e1e;
    color: #ffffff;
}
QWidget {
    background-color: #1e1e1e;
    color: #ffffff;
    font-family: 'Malgun Gothic', 'Segoe UI', Arial, sans-serif;
    font-size: 10pt;
}
QTableWidget {
    background-color: #2d2d2d;
    alternate-background-color: #353535;
    gridline-color: #555555;
    selection-background-color: #094771;
    border: 1px solid #555555;
}
QTableWidget::item:selected {
    background-color: #094771;
    color: white;
}
QTableWidget::item:!selected {
    color: white;
}
QHeaderView::section {
    background-color: #404040;
    color: white;
    padding: 8px;
    border: 1px solid #555555;
    font-weight: bold;
}
QPushButton {
    background-color: #0d7377;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #14a085;
}
QPushButton:pressed {
    background-color: #0a5d61;
}
QPushButton.danger {
    background-color: #dc3545;
}
QPushButton.danger:hover {
    background-color: #c82333;
}
QPushButton.success {
    background-color: #28a745;
}
QPushButton.success:hover {
    background-color: #218838;
}
QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit {
    background-color: #404040;
    color: white;
    border: 1px solid #555555;
    padding: 6px;
    border-radius: 4px;
}
QComboBox::drop-down {
    border: none;
}
QComboBox::down-arrow {
    width: 12px;
    height: 12px;
}
QGroupBox {
    font-weight: bold;
    border: 2px solid #555555;
    border-radius: 8px;
    margin-top: 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px 0 5px;
}
QTabWidget::pane {
    border: 1px solid #555555;
    background-color: #2d2d2d;
}
QTabBar::tab {
    background-color: #404040;
    color: white;
    padding: 10px 20px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected {
    background-color: #0d7377;
}
QProgressBar {
    border: 1px solid #555555;
    border-radius: 4px;
    text-align: center;
    color: white;
}
QProgressBar::chunk {
    background-color: #0d7377;
    border-radius: 3px;
}
QTextEdit {
    background-color: #2d2d2d;
    color: white;
    border: 1px solid #555555;
    border-radius: 4px;
}
QLabel {
    color: white;
}
QCheckBox {
    color: white;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border-radius: 3px;
    background-color: #404040;
    border: 1px solid #555;
}
QCheckBox::indicator:checked {
    background-color: #0d7377;
    border: 1px solid #14a085;
}
/* === Styles for Buy/Sell toggles === */
QCheckBox.buy_toggle::indicator, QCheckBox.sell_toggle::indicator {
    width: 20px;
    height: 20px;
    border: 2px solid #555;
    border-radius: 4px;
}
QCheckBox.buy_toggle::indicator:checked {
    background-color: #28a745; /* Green for Buy */
    border-color: #1e7e34;
}
QCheckBox.sell_toggle::indicator:checked {
    background-color: #dc3545; /* Red for Sell */
    border-color: #b21f2d;
}
QScrollBar:vertical {
    border: none;
    background: #2d2d2d;
    width: 10px;
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:vertical {
    background: #555555;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QPushButton.delete {
    background-color: #ff5656;      /* 산뜻한 토마토 레드 */
    color: white;
    border: none;
    border-radius: 4px;
    font-size: 16px;                /* “➖” 기호가 잘 보이도록 */
}
QPushButton.delete:hover {
    background-color: #ff3b3b;
}
"""

# Base classes
class _BaseStrategyRow(QWidget):
    strategy_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_base()

    def _build_base(self):
        self.lay = QHBoxLayout(self)
        self.lay.setContentsMargins(5, 5, 5, 5)
        self.lay.setSpacing(8)
        
        # Delete button
        self.btn_del = QPushButton("✕")
        self.btn_del.setFixedSize(28, 28)
        self.btn_del.setProperty("class", "danger")
        self.btn_del.setToolTip("전략 삭제")
        self.btn_del.setStyleSheet("QPushButton { background-color: #dc3545; font-size: 14px; font-weight: bold; }")

    def _add_del_button(self):
        self.lay.addStretch()
        self.lay.addWidget(self.btn_del)

    def get_config(self):
        raise NotImplementedError

    def set_config(self, cfg: dict):
        raise NotImplementedError

class BuyStrategyRow(_BaseStrategyRow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        # Strategy Type
        self.lbl_strategy = QLabel("전략:")
        self.cbo_strategy = QComboBox()
        self.cbo_strategy.addItems(["N일고점돌파", "특정가격돌파"])
        self.cbo_strategy.currentTextChanged.connect(self._update_param)
        self.cbo_strategy.currentTextChanged.connect(self.strategy_changed.emit)
        
        # Parameter
        self.lbl_param = QLabel("값:")
        self.spn_param = QSpinBox()
        
        # Buy Amount
        self.lbl_amount = QLabel("금액:")
        self.spn_amount = QSpinBox()
        self.spn_amount.setRange(10_000, 1_000_000_000)
        self.spn_amount.setSingleStep(10_000)
        self.spn_amount.setSuffix(" 원")
        self.spn_amount.setValue(1_000_000)
        self.spn_amount.setGroupSeparatorShown(True)
        
        # Additional Condition
        self.lbl_cond = QLabel("조건:")
        self.cbo_cond = QComboBox()
        self.cbo_cond.addItems(["조건없음", "거래량", "거래대금"])
        self.cbo_cond.currentTextChanged.connect(self._update_cond)
        
        self.spn_cond_val = QDoubleSpinBox()
        self.spn_cond_val.setGroupSeparatorShown(True)
        
        # Layout
        widgets = [
            (self.lbl_strategy, 0), (self.cbo_strategy, 2),
            (self.lbl_param, 0), (self.spn_param, 1),
            (self.lbl_amount, 0), (self.spn_amount, 2),
            (self.lbl_cond, 0), (self.cbo_cond, 1),
            (self.spn_cond_val, 2)
        ]
        
        for widget, stretch in widgets:
            self.lay.addWidget(widget)
            self.lay.setStretchFactor(widget, stretch)
            
        self._add_del_button()
        
        # Initial state setup
        self._update_param(self.cbo_strategy.currentText())
        self._update_cond(self.cbo_cond.currentText())

    def _update_param(self, text):
        if text == "특정가격돌파":
            self.spn_param.setRange(100, 1_000_000_000)
            self.spn_param.setSingleStep(100)
            self.spn_param.setSuffix(" 원")
            self.spn_param.setValue(50_000)
        elif text == "N일고점돌파":
            self.spn_param.setRange(1, 365)
            self.spn_param.setSingleStep(1)
            self.spn_param.setSuffix(" 일")
            self.spn_param.setValue(20)
        else:  # 조건없음 또는 기타 전략
            self.spn_param.setRange(0, 0)
            self.spn_param.setSingleStep(0)
            self.spn_param.setSuffix("")
            self.spn_param.setValue(0)

    def _update_cond(self, text):
        if text == "거래량":
            self.spn_cond_val.setEnabled(True)
            self.spn_cond_val.setRange(1_000, 100_000_000)
            self.spn_cond_val.setSingleStep(1_000)
            self.spn_cond_val.setSuffix(" 주")
            self.spn_cond_val.setDecimals(0)
            self.spn_cond_val.setValue(500_000)
        elif text == "거래대금":
            self.spn_cond_val.setEnabled(True)
            self.spn_cond_val.setRange(100_000, 10_000_000_000)
            self.spn_cond_val.setSingleStep(100_000)
            self.spn_cond_val.setSuffix(" 원")
            self.spn_cond_val.setDecimals(0)
            self.spn_cond_val.setValue(1_000_000_000)
        else:
            self.spn_cond_val.setEnabled(False)
            self.spn_cond_val.setSuffix("")
            self.spn_cond_val.setValue(0)

    def get_config(self):
        return {
            "strategy": self.cbo_strategy.currentText(),
            "param": self.spn_param.value(),
            "amount": self.spn_amount.value(),
            "cond_type": self.cbo_cond.currentText(),
            "cond_value": self.spn_cond_val.value(),
        }

    def set_config(self, cfg: dict):
        self.cbo_strategy.setCurrentText(cfg.get("strategy", "N일고점돌파"))
        self.spn_param.setValue(cfg.get("param", 20))
        self.spn_amount.setValue(cfg.get("amount", 1_000_000))
        self.cbo_cond.setCurrentText(cfg.get("cond_type", "조건없음"))
        self.spn_cond_val.setValue(cfg.get("cond_value", 0))

class SellStrategyRow(_BaseStrategyRow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        # Strategy Type
        self.lbl_strategy = QLabel("전략:")
        self.cbo_strategy = QComboBox()
        self.cbo_strategy.addItems(["N일저점이탈", "수익률매도", "손절매", "특정가격이탈", "트레일링스탑"])
        self.cbo_strategy.currentTextChanged.connect(self._update_param)
        self.cbo_strategy.currentTextChanged.connect(self.strategy_changed.emit)

        # ── Single Parameter (General) ──
        self.lbl_param = QLabel("값:")
        self.spn_param = QDoubleSpinBox()
        
        # ── Trailing Stop Widgets ──
        self.lbl_trail_base = QLabel("기준:") # <<< [추가] 기준가 라벨
        self.cbo_trail_base = QComboBox()    # <<< [추가] 기준가 선택 콤보박스
        self.cbo_trail_base.addItems(["현재가", "매수평단가", "전일종가"])

        self.lbl_raise = QLabel("상승:")
        self.spn_raise = QDoubleSpinBox()
        self.spn_raise.setRange(0.1, 100.0)
        self.spn_raise.setSingleStep(0.1)
        self.spn_raise.setSuffix(" %")
        self.spn_raise.setDecimals(2)

        self.lbl_trail = QLabel("하락:")
        self.spn_trail = QDoubleSpinBox()
        self.spn_trail.setRange(0.1, 100.0)
        self.spn_trail.setSingleStep(0.1)
        self.spn_trail.setSuffix(" %")
        self.spn_trail.setDecimals(2)

        # Hide initially
        for w in (self.lbl_trail_base, self.cbo_trail_base, self.lbl_raise, self.spn_raise, self.lbl_trail, self.spn_trail): # <<< [수정] 숨길 위젯에 추가
            w.hide()

        # Sell Method
        self.lbl_method = QLabel("방식:")
        self.cbo_method = QComboBox()
        self.cbo_method.addItems(["비중", "금액", "전량"])
        self.cbo_method.currentTextChanged.connect(self._update_value)

        # Sell Quantity/Amount
        self.spn_value = QSpinBox()
        self.spn_value.setGroupSeparatorShown(True)

        # Layout
        widgets = [
            (self.lbl_strategy, 0), (self.cbo_strategy, 2),
            (self.lbl_param, 0), (self.spn_param, 1),
            (self.lbl_trail_base, 0), (self.cbo_trail_base, 1), # <<< [추가] 레이아웃에 추가
            (self.lbl_raise, 0), (self.spn_raise, 1),
            (self.lbl_trail, 0), (self.spn_trail, 1),
            (self.lbl_method, 0), (self.cbo_method, 1),
            (self.spn_value, 2)
        ]
        for widget, stretch in widgets:
            self.lay.addWidget(widget)
            self.lay.setStretchFactor(widget, stretch)

        self._add_del_button()

        # Initial state setup
        self._update_param(self.cbo_strategy.currentText())
        self._update_value(self.cbo_method.currentText())

    def _update_param(self, text):
        is_trail = (text == "트레일링스탑")
        # Show/hide for trailing stop only
        self.lbl_param.setVisible(not is_trail)
        self.spn_param.setVisible(not is_trail)
        for w in (self.lbl_trail_base, self.cbo_trail_base, self.lbl_raise, self.spn_raise, self.lbl_trail, self.spn_trail): # <<< [수정] 보이고 숨길 위젯에 추가
            w.setVisible(is_trail)

        if is_trail:
            self.spn_raise.setValue(3.0)   # Default raise percentage
            self.spn_trail.setValue(1.0)   # Default trail percentage
            return

        # Handle other single parameters
        if text == "N일저점이탈":
            self.spn_param.setRange(1, 365)
            self.spn_param.setSingleStep(1)
            self.spn_param.setSuffix(" 일")
            self.spn_param.setDecimals(0)
            self.spn_param.setValue(10)
        elif text in ("수익률매도", "손절매"):
            self.spn_param.setRange(0.1, 100.0)
            self.spn_param.setSingleStep(0.1)
            self.spn_param.setSuffix(" %")
            self.spn_param.setDecimals(2)
            if text == "수익률매도":
                self.spn_param.setValue(10.0)
            else:  # 손절매
                self.spn_param.setValue(5.0)
        elif text == "트레일링스탑":
            pass  # Already handled above
        else:  # 특정가격이탈
            self.spn_param.setRange(100, 1_000_000_000)
            self.spn_param.setSingleStep(100)
            self.spn_param.setSuffix(" 원")
            self.spn_param.setDecimals(0)
            self.spn_param.setValue(50_000)

    def _update_value(self, method):
        if method == "비중":
            self.spn_value.setEnabled(True)
            self.spn_value.setRange(1, 100)
            self.spn_value.setSingleStep(1)
            self.spn_value.setSuffix(" %")
            self.spn_value.setValue(50)
        elif method == "금액":
            self.spn_value.setEnabled(True)
            self.spn_value.setRange(10_000, 1_000_000_000)
            self.spn_value.setSingleStep(10_000)
            self.spn_value.setSuffix(" 원")
            self.spn_value.setValue(1_000_000)
        else:  # 전량
            self.spn_value.setEnabled(False)
            self.spn_value.setSuffix(" %") # Still show suffix for consistency even if disabled
            self.spn_value.setValue(100)

    def get_config(self):
        cfg = {
            "strategy": self.cbo_strategy.currentText(),
            "method": self.cbo_method.currentText(),
            "value": self.spn_value.value(),
        }
        if self.cbo_strategy.currentText() == "트레일링스탑":
            cfg.update({
                "trail_base": self.cbo_trail_base.currentText(), # <<< [추가] 설정 저장
                "raise_pct": self.spn_raise.value(),
                "trail_pct": self.spn_trail.value()
            })
        else:
            cfg["param"] = self.spn_param.value()
        return cfg

    def set_config(self, cfg: dict):
        strat = cfg.get("strategy", "N일저점이탈")
        self.cbo_strategy.setCurrentText(strat)
        if strat == "트레일링스탑":
            self.cbo_trail_base.setCurrentText(cfg.get("trail_base", "현재가")) # <<< [추가] 설정 불러오기
            self.spn_raise.setValue(cfg.get("raise_pct", 3.0))
            self.spn_trail.setValue(cfg.get("trail_pct", 1.0))
        else:
            self.spn_param.setValue(cfg.get("param", 10))
        self.cbo_method.setCurrentText(cfg.get("method", "비중"))
        self.spn_value.setValue(cfg.get("value", 50))

class StrategySection(QWidget):
    def __init__(self, kind: str = "buy", parent=None):
        super().__init__(parent)
        self.kind = kind
        self.rows = []
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        
        group_title = "매수 전략" if self.kind == "buy" else "매도 전략"
        self.group_box = QGroupBox(group_title)
        group_layout = QVBoxLayout(self.group_box)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(300)
        self.container = QWidget()
        self.container_layout = QVBoxLayout(self.container)
        self.container_layout.setContentsMargins(5, 5, 5, 5)
        self.container_layout.setSpacing(8)
        self.container_layout.addStretch()
        scroll.setWidget(self.container)
        group_layout.addWidget(scroll)

        btn_layout = QHBoxLayout()
        txt_add = f"+ {group_title} 추가"
        self.btn_add = QPushButton(txt_add)
        self.btn_add.setProperty("class", "success")
        self.btn_add.clicked.connect(self.add_row)
        
        self.btn_clear = QPushButton("전체 삭제")
        self.btn_clear.setProperty("class", "danger")
        self.btn_clear.clicked.connect(self.clear_all_confirm)
        
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        group_layout.addLayout(btn_layout)
        
        outer.addWidget(self.group_box)

    def add_row(self, cfg: Optional[dict] = None):
        row = BuyStrategyRow() if self.kind == "buy" else SellStrategyRow()
        # Insert before the stretch factor
        self.container_layout.insertWidget(self.container_layout.count() - 1, row)
        self.rows.append(row)
        row.btn_del.clicked.connect(lambda: self.remove_row(row))
        if cfg:
            row.set_config(cfg)
        return row

    def remove_row(self, row):
        if row in self.rows:
            self.rows.remove(row)
            row.setParent(None)
            row.deleteLater() # Ensure widget is properly deleted

    def clear_all_confirm(self):
        if self.rows:
            reply = QMessageBox.question(self, '확인', 
                                         f'{self.group_box.title()}을 모두 삭제하시겠습니까?',
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.clear_all()

    def clear_all(self):
        for r in self.rows[:]: # Iterate over a copy to allow modification during loop
            self.remove_row(r)

    def get_configs(self):
        return [r.get_config() for r in self.rows]

    def set_configs(self, cfg_list):
        self.clear_all()
        for cfg in cfg_list:
            self.add_row(cfg)

class StatusPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()
        self._setup_timer()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        
        status_group = QGroupBox("시스템 상태")
        status_layout = QGridLayout(status_group)
        
        self.lbl_connection = QLabel("크레온 연결:")
        self.lbl_connection_status = QLabel("🔴 미연결")
        
        self.lbl_auto_trade = QLabel("자동매매:")
        self.lbl_auto_status = QLabel("🟡 대기중")
        
        self.lbl_last_update = QLabel("현재 시간:")
        self.lbl_update_time = QLabel("--:--:--")
        
        status_layout.addWidget(self.lbl_connection, 0, 0)
        status_layout.addWidget(self.lbl_connection_status, 0, 1)
        status_layout.addWidget(self.lbl_auto_trade, 1, 0)
        status_layout.addWidget(self.lbl_auto_status, 1, 1)
        status_layout.addWidget(self.lbl_last_update, 2, 0)
        status_layout.addWidget(self.lbl_update_time, 2, 1)
        status_layout.setColumnStretch(1, 1)
        
        layout.addWidget(status_group)
        
        log_group = QGroupBox("실행 로그")
        log_layout = QVBoxLayout(log_group)
        
        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)
        log_layout.addWidget(self.text_log)
        
        layout.addWidget(log_group)

    def _setup_timer(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self._update_time)
        self.timer.start(1000)

    def _update_time(self):
        current_time = datetime.now().strftime("%H:%M:%S")
        self.lbl_update_time.setText(current_time)

    def add_log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        color = {
            "INFO": "#ffffff",   # White
            "SUCCESS": "#28a745", # Green
            "WARN": "#ffc107",   # Yellow
            "ERROR": "#dc3545"    # Red
        }.get(level, "#ffffff")
        
        log_entry = f'<p style="color:{color}; margin: 2px;">[{timestamp}] [{level}] {message}</p>'
        self.text_log.insertHtml(log_entry)
        self.text_log.verticalScrollBar().setValue(self.text_log.verticalScrollBar().maximum())

class AddSymbolDialog(QDialog):
    def __init__(self, creon_mgr, parent=None):
        super().__init__(parent)
        self.setWindowTitle("종목 추가")
        self.resize(400, 100)
        self.creon = creon_mgr
        self.selected_code = None

        vbox = QVBoxLayout(self)
        self.input = QLineEdit(self)
        self.input.setPlaceholderText("종목코드 또는 종목명을 입력하세요")
        vbox.addWidget(self.input)

        # Get all stock codes from Creon
        all_codes = []
        if self.creon.is_initialized:
            all_codes.extend(self.creon.cp_code_mgr.GetStockListByMarket(1)) # KOSPI
            all_codes.extend(self.creon.cp_code_mgr.GetStockListByMarket(2)) # KOSDAQ
        
        self.stock_list = [(c, self.creon.get_stock_name(c)) for c in all_codes]
        self.stock_list = [(c, n) for c, n in self.stock_list if n] # Filter out empty names

        # Create completer list
        completer_list = [f"{name} ({code})" for code, name in self.stock_list]
        self.completer = QCompleter(completer_list, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.input.setCompleter(self.completer)

        def select_code():
            text = self.input.text()
            # If using autocomplete, the format is "Stock Name (Code)"
            if "(" in text and text.endswith(")"):
                try:
                    code_in_text = text[text.rfind("(")+1:-1]
                    for code, name in self.stock_list:
                        if code == code_in_text:
                            self.selected_code = code
                            self.accept()
                            return
                except:
                    pass # Ignore malformed input
            
            # If directly typed
            for code, name in self.stock_list:
                if text == code or text == name:
                    self.selected_code = code
                    self.accept()
                    return
            
            QMessageBox.warning(self, "오류", "유효한 종목 코드 또는 종목명을 입력해주세요.")

        self.input.returnPressed.connect(select_code)
        self.btn_add = QPushButton("추가", self)
        vbox.addWidget(self.btn_add)
        self.btn_add.clicked.connect(select_code)

    @staticmethod
    def get_code(creon_mgr, parent=None):
        dlg = AddSymbolDialog(creon_mgr, parent)
        if dlg.exec_() == QDialog.Accepted and dlg.selected_code:
            return dlg.selected_code
        return None        

# 교체할 클래스: TradingWorker
class TradingWorker(QThread):
    log_signal = pyqtSignal(str, str)
    trade_signal = pyqtSignal(str)

    def __init__(self, creon, code, strategies):
        super().__init__()
        self.creon = creon
        self.code = code
        # 전략 및 토글 설정
        self.buy_strategies = strategies.get("buy", [])
        self.sell_strategies = strategies.get("sell", [])
        self.buy_enabled = strategies.get("buy_flag", True)
        self.sell_enabled = strategies.get("sell_flag", True)

        # 내부 상태
        self._stop_event = threading.Event()
        self.data_queue = queue.Queue()
        self.trailing_stop_active = False
        self.trailing_peak_price = 0
        self.trailing_base_price_met = False

        # 초기 잔고/평단 및 이익/손절 정보
        self.prev_close_price = 0
        self.avg_buy_price = 0
        self.quantity_held = 0

        # N일 고점/저점 갱신 상태
        self.nday_high_targets = {}
        self.last_high_refresh_date = None
        self.nday_low_targets = {}
        self.last_low_refresh_date = None

    def refresh_nday_high_targets(self):
        today = datetime.now().strftime("%Y%m%d")
        if self.last_high_refresh_date == today:
            return
        self.nday_high_targets.clear()
        for cfg in self.buy_strategies:
            if cfg.get("strategy") == "N일고점돌파":
                n = cfg.get("param", 0)
                high = self.creon.get_high_price_for_days(self.code, n)
                if high > 0:
                    self.nday_high_targets[n] = high
                    self.log_signal.emit(f"[{self.code}] {n}일 고점(갱신): {high:,}원", "INFO")
                else:
                    self.log_signal.emit(f"[{self.code}] {n}일 고점 조회 실패; 전략 건너뜀", "WARN")
        self.last_high_refresh_date = today

    def refresh_nday_low_targets(self):
        today = datetime.now().strftime("%Y%m%d")
        if self.last_low_refresh_date == today:
            return
        self.nday_low_targets.clear()
        for cfg in self.sell_strategies:
            if cfg.get("strategy") == "N일저점이탈":
                n = cfg.get("param", 0)
                low = self.creon.get_low_price_for_days(self.code, n)
                if low > 0:
                    self.nday_low_targets[n] = low
                    self.log_signal.emit(f"[{self.code}] {n}일 저점(갱신): {low:,}원", "INFO")
                else:
                    self.log_signal.emit(f"[{self.code}] {n}일 저점 조회 실패; 전략 건너뜀", "WARN")
        self.last_low_refresh_date = today

    def stop(self):
        self._stop_event.set()
        self.log_signal.emit(f"[{self.code}] 자동매매 스레드 중지 요청", "INFO")

    def run(self):
        pythoncom.CoInitialize()
        self.log_signal.emit(f"[{self.code}] 자동매매 스레드 시작", "INFO")

        # 초기 잔고/평단 및 전일 종가 조회
        self.quantity_held, self.avg_buy_price = self.creon.get_stock_balance_and_avg_price(self.code)
        info = self.creon.get_stock_info(self.code)
        if info:
            self.prev_close_price = info.get("close_price", 0)

        # N일 고점/저점 최초 갱신
        self.refresh_nday_high_targets()
        self.refresh_nday_low_targets()
        
        try:
            while not self._stop_event.is_set():
                # 장 시작 전이나 장 마감 후에는 불필요한 루프 방지
                current_time = datetime.now().time()
                if not (datetime.strptime("09:00", "%H:%M").time() <= current_time <= datetime.strptime("15:30", "%H:%M").time()):
                    time.sleep(1)
                    continue

                # 매 5분마다 N일 고점/저점 갱신
                if datetime.now().minute % 5 == 0 and datetime.now().second < 5:
                    self.refresh_nday_high_targets()
                    self.refresh_nday_low_targets()

                try:
                    # 큐에서 가장 최신 데이터 하나만 사용
                    data = self.data_queue.get(timeout=1.0)
                    while not self.data_queue.empty():
                        data = self.data_queue.get_nowait()
                except queue.Empty:
                    continue # 데이터 없으면 다음 루프로

                # --- 수정된 매수 로직 ---
                if self.buy_enabled and self.quantity_held <= 0:
                    for cfg in self.buy_strategies:
                        if self._check_buy_condition(cfg, data):
                            qty = self._calculate_buy_qty(cfg, data.get("current_price", 0))
                            if qty > 0 and self.creon.place_order(self.code, qty, 0, is_buy=True):
                                self.log_signal.emit(
                                    f"[{self.code}] 매수 체결: {cfg['strategy']} – {qty}주", "SUCCESS"
                                )
                                self.trade_signal.emit(self.code)
                                time.sleep(0.5)
                                # 체결 후 즉시 잔고/평단가 업데이트
                                self.quantity_held, self.avg_buy_price = self.creon.get_stock_balance_and_avg_price(self.code)
                                # 트레일링 스탑 관련 상태 초기화
                                self.trailing_stop_active = False
                                self.trailing_peak_price = 0
                                self.trailing_base_price_met = False
                                break # 매수 성공 시 다른 매수 전략은 더 이상 확인하지 않음

                # --- 수정된 매도 로직 ---
                elif self.sell_enabled and self.quantity_held > 0:
                    data['close_price'] = self.prev_close_price
                    for cfg in self.sell_strategies:
                        # 트레일링 스탑은 자체적으로 매도 주문까지 처리하므로 별도 핸들링
                        if cfg.get("strategy") == "트레일링스탑":
                            self._execute_trailing_stop(cfg, data, self.quantity_held, self.avg_buy_price)
                            if self.quantity_held <= 0: break # 전량 매도되었다면 루프 탈출
                            continue # 트레일링 스탑 조건이 아니면 다음 매도 전략으로

                        # 기타 매도 전략 확인
                        should_sell = False
                        if cfg.get("strategy") == "N일저점이탈":
                            low_target = self.nday_low_targets.get(cfg.get("param", 0))
                            current_price = data.get("current_price", 0)
                            if low_target and current_price > 0 and current_price < low_target:
                                self.log_signal.emit(f"[{self.code}] N일저점이탈: 현재가 {current_price:,} < 목표 {low_target:,}", "INFO")
                                should_sell = True
                        elif self._check_sell_condition(cfg, data, self.quantity_held, self.avg_buy_price):
                            should_sell = True

                        if should_sell:
                            sell_qty = self._calculate_sell_qty(cfg, self.quantity_held, data.get("current_price", 0))
                            if sell_qty > 0 and self.creon.place_order(self.code, sell_qty, 0, is_buy=False):
                                self.log_signal.emit(f"[{self.code}] 매도 체결: {cfg['strategy']} – {sell_qty}주", "SUCCESS")
                                self.trade_signal.emit(self.code)
                                time.sleep(0.5)
                                self.quantity_held, self.avg_buy_price = self.creon.get_stock_balance_and_avg_price(self.code)
                                if self.quantity_held <= 0: self.stop() # 전량 매도 시 스레드 종료
                                break # 매도 성공 시 루프 탈출
                
        except Exception as e:
            self.log_signal.emit(f"[{self.code}] 처리 오류: {e}", "ERROR")
            logging.exception(f"[{self.code}] 예외 발생")
        finally:
            pythoncom.CoUninitialize()
            self.log_signal.emit(f"[{self.code}] 스레드 종료", "INFO")

    def _check_buy_condition(self, cfg, info):
        strat = cfg.get("strategy")
        cur = info.get("current_price", 0)
        logging.info(f"[{self.code}] 전략: {strat}, 현재가: {cur}, 전략 파라미터: {cfg.get('param')}")
        logging.info(f"[{self.code}] 매수 조건 체크 중 – 전략: {cfg}, 데이터: {info}")

        if cur is None or cur <= 0:
            logging.warning(f"[{self.code}] 현재가 없음: cur={cur} → 매수 건너뜀")
            return False

        if strat == "N일고점돌파":
            high = self.nday_high_targets.get(cfg.get("param", 0))
            logging.info(f"[{self.code}] 현재가 {cur}, 고점 {high} / 조건 확인 중 (전략: {strat})")
            if high and cur > high:
                logging.info(f"[{self.code}] 매수 조건 통과! (N일고점돌파) → 현재가: {cur}, 고점: {high}")
                self.log_signal.emit(
                    f"[{self.code}] N일고점돌파: 현재가 {cur:,} > 고점 {high:,}",
                    "INFO"
                )
                return True

        elif strat == "특정가격돌파":
            tgt = cfg.get("param", 0)
            logging.info(f"[{self.code}] 전략 조건 확인 중 – 현재가: {cur}, 목표: {tgt}")
            if cur < tgt:
                return False
            if cfg.get("cond_type") == "조건없음":
                return True
            if cfg.get("cond_type") == "거래량" and info.get("volume", 0) >= cfg.get("cond_value", 0):
                return True
            if cfg.get("cond_type") == "거래대금" and info.get("trade_value", 0) >= cfg.get("cond_value", 0):
                return True

        return False

    def _calculate_buy_qty(self, cfg, cur):
        amt = cfg.get("amount", 0)
        return amt // cur if cur > 0 else 0

    def _check_sell_condition(self, cfg, info, qty, avg):
        # N일저점이탈은 run()에서 바로 처리하므로 여기선 나머지 전략만
        cur = info.get("current_price", 0)
        if avg <= 0:
            return False
        strat = cfg.get("strategy")
        param = cfg.get("param", 0)
        if strat == "수익률매도":
            if (cur - avg) / avg * 100 >= param:
                self.log_signal.emit(
                    f"[{self.code}] 수익률매도: {((cur-avg)/avg*100):.2f}% >= {param}%",
                    "INFO"
                )
                return True
        elif strat == "손절매":
            if (avg - cur) / avg * 100 >= param:
                self.log_signal.emit(
                    f"[{self.code}] 손절매: {((avg-cur)/avg*100):.2f}% >= {param}%",
                    "WARN"
                )
                return True
        elif strat == "특정가격이탈":
            if cur <= param:
                self.log_signal.emit(
                    f"[{self.code}] 특정가격이탈: {cur:,} <= {param:,}",
                    "INFO"
                )
                return True
        return False

    def _execute_trailing_stop(self, cfg, info, qty, avg):
        cur = info.get("current_price", 0)
        if not self.trailing_base_price_met:
            base = cfg.get("trail_base", "현재가")
            if base == "매수평단가" and avg > 0:
                peak = avg
            elif base == "전일종가":
                peak = self.prev_close_price if self.prev_close_price > 0 else cur
            else:
                peak = cur
            self.trailing_peak_price = peak
            self.trailing_base_price_met = True
            self.log_signal.emit(
                f"[{self.code}] 트레일링 초기 기준가: {peak:,}원", "INFO"
            )
            return
        raise_pct = cfg.get("raise_pct", 0)
        trail_pct = cfg.get("trail_pct", 0)
        if not self.trailing_stop_active:
            if cur >= self.trailing_peak_price * (1 + raise_pct/100):
                self.trailing_stop_active = True
                self.trailing_peak_price = cur
                self.log_signal.emit(
                    f"[{self.code}] 트레일링 활성화, 고점 갱신: {cur:,}원", "INFO"
                )
            return
        if cur > self.trailing_peak_price:
            self.trailing_peak_price = cur
            self.log_signal.emit(
                f"[{self.code}] 트레일링 고점 갱신: {cur:,}원", "INFO"
            )
        elif cur <= self.trailing_peak_price * (1 - trail_pct/100):
            sell_qty = self._calculate_sell_qty(cfg, qty, cur)
            if sell_qty > 0:
                if self.creon.place_order(self.code, sell_qty, 0, is_buy=False):
                    self.log_signal.emit(
                        f"[{self.code}] 트레일링 매도: {sell_qty}주", "SUCCESS"
                    )
                    self.trade_signal.emit(self.code)
                    time.sleep(0.5)
                    self.quantity_held, self.avg_buy_price = self.creon.get_stock_balance_and_avg_price(self.code)
                    if self.quantity_held == 0 and cfg.get("method") == "전량":
                        self.stop()
                else:
                    self.log_signal.emit(
                        f"[{self.code}] 트레일링 매도 주문 실패", "ERROR"
                    )

    def _calculate_sell_qty(self, cfg, qty, cur):
        method = cfg.get("method")
        val = cfg.get("value", 0)
        if method == "비중":
            return int(qty * val / 100)
        if method == "금액":
            return int(val / cur) if cur > 0 else 0
        if method == "전량":
            return qty
        return 0

# 교체할 클래스: TradingManager
class TradingManager:
    """Manages multiple TradingWorker threads."""
    def __init__(self, creon_mgr: CreonManager, status_panel: StatusPanel, refresh_callback=None):
        self.creon = creon_mgr
        self.status_panel = status_panel
        self.refresh_callback = refresh_callback
        self.workers = {}  # {code: TradingWorker instance}
        self.is_auto_trading_active = False

    def start_trading(self, strategy_data: dict, selected_codes: list):
        if self.is_auto_trading_active:
            self.status_panel.add_log("자동매매가 이미 실행 중입니다.", "WARN")
            return

        if not self.creon.is_initialized:
            self.status_panel.add_log("크레온 API에 연결되지 않아 자동매매를 시작할 수 없습니다.", "ERROR")
            return

        self.is_auto_trading_active = True
        self.status_panel.lbl_auto_status.setText("🟢 실행 중")
        self.status_panel.add_log("자동매매를 시작합니다.", "INFO")

        for code in selected_codes:
            if code not in self.workers:
                strategies_for_code = strategy_data.get(code, {})
                strategies_for_code["buy_flag"] = True
                strategies_for_code["sell_flag"] = True
                worker = TradingWorker(self.creon, code, strategies_for_code)
                worker.log_signal.connect(self.status_panel.add_log)
                worker.trade_signal.connect(self._handle_trade_signal)
                
                # <<< [추가] 스레드를 시작하기 전 실시간 구독 요청
                self.creon.subscribe_realtime(code, worker.data_queue)

                self.workers[code] = worker
                worker.start()
            else:
                self.status_panel.add_log(f"[{code}] 이미 실행 중인 스레드가 있습니다.", "WARN")

    def stop_trading(self):
        if not self.is_auto_trading_active:
            self.status_panel.add_log("자동매매가 실행 중이 아닙니다.", "WARN")
            return

        self.is_auto_trading_active = False
        self.status_panel.lbl_auto_status.setText("🟡 대기중")
        self.status_panel.add_log("자동매매 중지를 요청합니다. 스레드 종료 대기 중...", "INFO")

        for code, worker in self.workers.items():
            # <<< [추가] 스레드 종료 전 실시간 구독 해지
            self.creon.unsubscribe_realtime(code)
            worker.stop()
            worker.wait(2000) # 2초간 기다림

        self.workers.clear()
        self.status_panel.add_log("모든 자동매매 스레드가 중지되었습니다.", "INFO")

    def _handle_trade_signal(self, code: str):
        # 매매 체결 후 잔고/수익률 갱신을 위해 호출되는 콜백
        if callable(self.refresh_callback):
            self.refresh_callback(code)


# 교체할 클래스: MainWindow
class MainWindow(QMainWindow):
    CONFIG_FILE = "user_config.json"
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Creon Auto Trader Pro v9.0 (Realtime)") # <<< 버전명 변경
        self.setWindowIcon(QIcon()) 
        self.resize(2200, 1200)
        self.setStyleSheet(DARK_STYLE)
        
        self.strategy_data = {}
        self.current_code = None
        self.creon = CreonManager()
        self.trading_manager = None

        self._build_ui()
        self._setup_connections()
        
        self.trading_manager = TradingManager(self.creon, self.status_panel, self.refresh_prices_for_code)

        self.connect_creon()
        # <<< [추가] 크레온 연결 후 실시간 시그널 연결
        if self.creon.is_initialized:
            self.creon.ui_update_signal.connect(self.update_row_from_realtime)

        self.load_config() 
        self.status_panel.add_log("프로그램이 시작되었습니다.", "INFO")

    # <<< [추가] 실시간 데이터로 UI의 특정 행을 업데이트하는 메소드
    def update_row_from_realtime(self, data: dict):
        code = data.get("code")
        if not code: return
        
        for row in range(self.table.rowCount()):
            if self.table.item(row, 1) and self.table.item(row, 1).text() == code:
                # 현재가 업데이트
                # [수정] current_price가 None 타입으로 들어오는 경우를 대비한 방어 코드 추가
                current_price = data.get("current_price") 
                if current_price is None:
                    current_price = 0 # None일 경우 0으로 처리하여 에러 방지

                price_item = QTableWidgetItem(f"{current_price:,}")
                price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, 3, price_item)

                # 수익률 실시간 계산 및 업데이트
                try:
                    # [수정] avg_price를 가져오기 전에 item 존재 여부 확인
                    avg_price_item = self.table.item(row, 4) # 잔고(수량)가 아닌 평단가가 필요합니다. TradingWorker의 평단가를 사용합니다.
                    worker = self.trading_manager.workers.get(code)
                    avg_price = worker.avg_buy_price if worker else 0
                    
                    balance_str = self.table.item(row, 4).text().replace(",", "") if self.table.item(row, 4) else "0"
                    balance = int(balance_str) if balance_str else 0
                    
                    if balance > 0 and avg_price > 0:
                        pnl = ((current_price - avg_price) / avg_price) * 100
                    else:
                        pnl = 0.0

                    pnl_item = QTableWidgetItem(f"{pnl:+.2f}%")
                    # [수정] 수익률에 따라 색상 변경 (가독성 향상)
                    if pnl > 0:
                        pnl_item.setForeground(QColor("#dc3545")) # Red for profit
                    elif pnl < 0:
                        pnl_item.setForeground(QColor("#0d7377")) # Blue for loss
                    else:
                        pnl_item.setForeground(QColor("#ffffff")) # White for neutral

                    pnl_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.table.setItem(row, 5, pnl_item)
                except Exception:
                    # 수익률 계산 중 오류가 발생해도 프로그램이 멈추지 않도록 예외 처리
                    pass 
                
                # 찾았으므로 루프 종료
                return

    def on_delete_clicked(self):
        btn = self.sender()
        if not isinstance(btn, QPushButton):
            return

        # 버튼이 들어 있는 행 번호 찾기 (컨테이너 QFrame의 자식까지 탐색)
        target_row = -1
        for row in range(self.table.rowCount()):
            cell_widget = self.table.cellWidget(row, 8)
            if cell_widget:
                for child in cell_widget.children():
                    if child is btn:
                        target_row = row
                        break
            if target_row != -1:
                break

        if target_row < 0:
            return

        self.remove_symbol_row(target_row)

    def remove_symbol_row(self, row):
        if row < 0 or row >= self.table.rowCount(): return

        code = self.table.item(row, 1).text()
        name = self.table.item(row, 2).text()
        
        reply = QMessageBox.question(self, '종목 삭제 확인', f'\'{name}({code})\' 종목을 목록에서 삭제하시겠습니까?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.No: return

        # 자동매매가 실행 중이면 해당 종목 중지 및 구독 해지
        if self.trading_manager and code in self.trading_manager.workers:
            self.creon.unsubscribe_realtime(code)
            self.trading_manager.workers[code].stop()
            self.trading_manager.workers[code].wait(1000)
            del self.trading_manager.workers[code]
            self.status_panel.add_log(f"'{name}' 종목의 자동매매가 중지되고 구독이 해지되었습니다.", "INFO")

        self.table.removeRow(row)
        if code in self.strategy_data:
            del self.strategy_data[code]

        if self.current_code == code:
            self.current_code = None
            self.section_buy.clear_all()
            self.section_sell.clear_all()
            self.lbl_stock.setText('<h2><i style="color:#aaa;">종목을 선택하세요</i></h2>')
            
        self.status_panel.add_log(f"종목 '{name}'이(가) 삭제되었습니다.", "WARN")

    # ... (기존 MainWindow의 다른 메소드들은 여기에 그대로 복사) ...
    # _build_ui, _create_left_panel, toggle_auto_trading, _create_right_panel 등등
    # refresh_prices_for_code는 잔고/수익률 등 BlockRequest가 필요한 정보 업데이트를 위해 유지
    def _build_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        
        left_panel = self._create_left_panel()
        right_panel = self._create_right_panel()
        
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([950, 1250])
        splitter.setChildrenCollapsible(False)
        
        main_layout.addWidget(splitter)

    def _create_left_panel(self):
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        btn_layout = QHBoxLayout()
        self.btn_add_symbol = QPushButton("➕ 종목 추가")
        self.btn_add_symbol.setProperty("class", "success")
        self.btn_refresh = QPushButton("🔄 잔고/수익률 갱신") # <<< 버튼 이름 변경
        
        self.btn_toggle_trade = QPushButton("▶ 자동매매 시작")
        self.btn_toggle_trade.setCheckable(True)
        self.btn_toggle_trade.setProperty("class", "success")
        self.btn_toggle_trade.clicked.connect(self.toggle_auto_trading)

        btn_layout.addWidget(self.btn_add_symbol)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_toggle_trade)
 
        left_layout.addLayout(btn_layout)
        
        self.table = QTableWidget(0, 9)
        headers = ["ON", "코드", "종목명", "현재가", "잔고", "수익률", "매수", "매도", "삭제"]
        self.table.setHorizontalHeaderLabels(headers)
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch) 
        
        self.table.setColumnWidth(0, 40)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 60)
        self.table.setColumnWidth(7, 60)
        self.table.setColumnWidth(8, 50)
        
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        left_layout.addWidget(self.table)
         
        return left_widget

    def toggle_auto_trading(self):
        if self.btn_toggle_trade.isChecked():
            self.btn_toggle_trade.setText("■ 자동매매 중지")
            self.btn_toggle_trade.setProperty("class", "danger")
            self.btn_toggle_trade.setStyle(self.style())
            self.start_auto_trading()
        else:
            self.btn_toggle_trade.setText("▶ 자동매매 시작")
            self.btn_toggle_trade.setProperty("class", "success")
            self.btn_toggle_trade.setStyle(self.style())
            self.stop_auto_trading()

    def _create_right_panel(self):
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        self.lbl_stock = QLabel('<h2><i style="color:#aaa;">종목을 선택하세요</i></h2>')
        right_layout.addWidget(self.lbl_stock)
        
        tab_widget = QTabWidget()
        
        strategy_tab = QWidget()
        strategy_layout = QVBoxLayout(strategy_tab)
        
        self.section_buy = StrategySection("buy")
        self.section_sell = StrategySection("sell")
        
        strategy_layout.addWidget(self.section_buy)
        strategy_layout.addWidget(self.section_sell)
        strategy_layout.addStretch()
        
        status_tab = StatusPanel()
        
        tab_widget.addTab(strategy_tab, "📈 전략 설정")
        tab_widget.addTab(status_tab, "🖥️ 상태 모니터링")
        
        self.status_panel = status_tab
        
        right_layout.addWidget(tab_widget)
        
        return right_widget

    def _setup_connections(self):
        self.table.cellClicked.connect(self.on_row_selected)
        self.btn_add_symbol.clicked.connect(self.add_symbol_dialog)
        self.btn_refresh.clicked.connect(self.refresh_prices)
 
    def _make_centered_checkbox(self, checked: bool):
        chk = QCheckBox()
        chk.setChecked(checked)

        container = QFrame()
        # ★ 컨테이너가 셀 크기에 맞춰 늘어나도록
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        # ★ 가로·세로 중앙 정렬
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(chk)

        chk.stateChanged.connect(self.on_checkbox_changed)
        return container, chk

    def connect_creon(self):
        self.status_panel.add_log("크레온 API 연결을 시도합니다...", "INFO")
        is_connected = self.creon.initialize()
        if is_connected:
            self.status_panel.lbl_connection_status.setText("🟢 연결됨")
            self.status_panel.add_log(f"크레온 연결 성공. (계좌: {self.creon.account})", "SUCCESS")
        else:
            self.status_panel.lbl_connection_status.setText("🔴 미연결")
            self.status_panel.add_log("크레온 연결에 실패했습니다. 크레온 플러스가 실행 중인지 확인하세요.", "ERROR")

    def add_symbol_row(self, data, configs=None):
        row_pos = self.table.rowCount()
        self.table.insertRow(row_pos)
        is_on, code, name, price, balance, pnl, buy_flag, sell_flag = data
        cell_widget, chk = self._make_centered_checkbox(is_on)
        self.table.setCellWidget(row_pos, 0, cell_widget)
        self.table.setItem(row_pos, 1, QTableWidgetItem(code))
        self.table.setItem(row_pos, 2, QTableWidgetItem(name))
        self.table.setItem(row_pos, 3, QTableWidgetItem(str(price)))
        self.table.setItem(row_pos, 4, QTableWidgetItem(str(balance)))
        self.table.setItem(row_pos, 5, QTableWidgetItem(str(pnl)))
        for col in [1, 3, 4, 5]: self.table.item(row_pos, col).setTextAlignment(Qt.AlignCenter)
        buy_container, buy_chk = self._make_centered_checkbox(buy_flag); buy_chk.setProperty("class", "buy_toggle"); self.table.setCellWidget(row_pos, 6, buy_container)
        sell_container, sell_chk = self._make_centered_checkbox(sell_flag); sell_chk.setProperty("class", "sell_toggle"); self.table.setCellWidget(row_pos, 7, sell_container)
        btn_del = QPushButton("➖")
        btn_del.setProperty("class", "delete")
        btn_del.setFixedSize(28, 28)
        btn_del.setToolTip("종목 삭제")
        btn_del.clicked.connect(self.on_delete_clicked)

        del_container = QFrame()
        # 컨테이너가 셀 크기에 맞춰 늘어나도록
        del_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        del_layout = QHBoxLayout(del_container)
        del_layout.setContentsMargins(0, 0, 0, 0)
        # 가로·세로 중앙 정렬
        del_layout.setAlignment(Qt.AlignCenter)
        del_layout.addWidget(btn_del)

        self.table.setCellWidget(row_pos, 8, del_container)
        if code not in self.strategy_data:
            self.strategy_data[code] = configs if configs else {"buy": [], "sell": [], "on": is_on, "buy_flag": buy_flag, "sell_flag": sell_flag}

    def on_checkbox_changed(self):
        widget = self.sender()
        if not isinstance(widget, QCheckBox): return
        for row in range(self.table.rowCount()):
            for col in [0, 6, 7]:
                cell_widget = self.table.cellWidget(row, col)
                if cell_widget and widget in cell_widget.children():
                    code = self.table.item(row, 1).text()
                    if code not in self.strategy_data: return
                    key = {0: "on", 6: "buy_flag", 7: "sell_flag"}[col]
                    self.strategy_data[code][key] = widget.isChecked()
                    return

    def save_current_strategies(self):
        if self.current_code and self.current_code in self.strategy_data:
            self.strategy_data[self.current_code]["buy"] = self.section_buy.get_configs()
            self.strategy_data[self.current_code]["sell"] = self.section_sell.get_configs()
    
    def load_strategies(self, code):
        data = self.strategy_data.get(code)
        if not data: 
            self.section_buy.clear_all(); self.section_sell.clear_all(); return
        self.section_buy.set_configs(data.get("buy", [])); self.section_sell.set_configs(data.get("sell", []))
        current_row = self.table.currentRow()
        if current_row < 0: return
        name = self.table.item(current_row, 2).text()
        self.lbl_stock.setText(f'<h2>{name} <span style="font-size: 12pt; color: #aaa;">({code})</span></h2>')

    def on_row_selected(self, row, col):
        if col == 8 or row < 0: return
        code = self.table.item(row, 1).text()
        if code == self.current_code: return
        self.save_current_strategies()
        self.current_code = code
        self.load_strategies(code)

    def add_symbol_dialog(self):
        if not self.creon.is_initialized:
            QMessageBox.warning(self, "연결 오류", "Creon Plus가 연결되지 않았습니다.")
            return

        code = AddSymbolDialog.get_code(self.creon, self)
        if code:
            for row in range(self.table.rowCount()):
                if self.table.item(row, 1).text() == code:
                    QMessageBox.warning(self, "중복 종목", f"'{self.creon.get_stock_name(code)}' 종목은 이미 목록에 있습니다.")
                    return

            stock_name = self.creon.get_stock_name(code)
            stock_info = self.creon.get_stock_info(code)

            # [수정] stock_info.get('price', ...) -> stock_info.get('current_price', ...) 로 키 이름 변경
            current_price_val = stock_info.get('current_price', 0) if stock_info else 0
            current_price = f"{current_price_val:,}"

            new_data = (
                True,            # is_active
                code,            # 종목코드
                stock_name,      # 종목명
                current_price,   # 현재가
                "0",             # 잔고 (초기값)
                "0.00%",         # 수익률 (초기값)
                True,            # 자동매수
                True             # 자동매도
            )
            self.add_symbol_row(new_data)
            self.table.setCurrentCell(self.table.rowCount() - 1, 0)
            self.on_row_selected(self.table.rowCount() - 1, 0)

            # 종목 추가 후 바로 잔고/수익률을 갱신하여 정확한 정보를 표시
            self.refresh_prices_for_code(code)

    def refresh_prices_for_code(self, code: str):
        """특정 종목 코드의 잔고, 현재가, 수익률을 모두 새로고침합니다."""
        if not self.creon.is_initialized: return
        
        row_to_update = -1
        for row in range(self.table.rowCount()):
            if self.table.item(row, 1) and self.table.item(row, 1).text() == code:
                row_to_update = row
                break
        if row_to_update == -1: return

        # [개선] 1. API를 통해 잔고, 평단가, 현재가 정보를 모두 새로 조회
        qty, avg_price = self.creon.get_stock_balance_and_avg_price(code)
        stock_info = self.creon.get_stock_info(code)
        cur = stock_info.get("current_price", 0) if stock_info else 0
        
        # worker가 있다면 평단가 및 잔고 정보 업데이트
        if code in self.trading_manager.workers:
            self.trading_manager.workers[code].avg_buy_price = avg_price
            self.trading_manager.workers[code].quantity_held = qty

        # [개선] 2. 새로 조회한 정보로 테이블의 모든 관련 셀 업데이트
        # 현재가
        price_item = QTableWidgetItem(f"{cur:,}")
        price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row_to_update, 3, price_item)
        
        # 잔고
        balance_item = QTableWidgetItem(f"{qty:,}")
        balance_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row_to_update, 4, balance_item)
        
        # 수익률
        pnl_item = QTableWidgetItem("0.00%")
        if qty > 0 and avg_price > 0 and cur > 0:
            pnl = ((cur - avg_price) / avg_price) * 100
            pnl_item = QTableWidgetItem(f"{pnl:+.2f}%")
        else:
            pnl_item = QTableWidgetItem("0.00%")

        pnl_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(row_to_update, 5, pnl_item)
                
    def refresh_prices(self):
        if not self.creon.is_initialized:
            QMessageBox.warning(self, "연결 오류", "Creon Plus가 연결되지 않았습니다."); return
        total = self.table.rowCount()
        if total == 0: return

        self.status_panel.add_log(f"전체 잔고/수익률 갱신 시작 ({total}종목)", "INFO")
        for row in range(total):
            code = self.table.item(row, 1).text()
            if code: self.refresh_prices_for_code(code)
        self.status_panel.add_log("전체 잔고/수익률 갱신 완료", "SUCCESS")

    def start_auto_trading(self):
        self.save_current_strategies()
        codes_to_trade = []
        for row in range(self.table.rowCount()):
            code = self.table.item(row, 1).text()
            stock_data = self.strategy_data.get(code)
            if stock_data and stock_data.get("on"):
                codes_to_trade.append(code)
        if not codes_to_trade:
            QMessageBox.warning(self, "자동매매 시작 불가", "자동매매 대상 종목이 없습니다."); return
        self.trading_manager.start_trading(self.strategy_data, codes_to_trade)

    def stop_auto_trading(self):
        self.trading_manager.stop_trading()
        if hasattr(self, "btn_toggle_trade"):
            self.btn_toggle_trade.blockSignals(True)
            self.btn_toggle_trade.setChecked(False)
            self.btn_toggle_trade.setText("▶ 자동매매 시작"); self.btn_toggle_trade.setProperty("class", "success")
            self.btn_toggle_trade.setStyle(self.style())
            self.btn_toggle_trade.blockSignals(False)

    def save_config(self):
        self.save_current_strategies()
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.strategy_data, f, indent=4, ensure_ascii=False)
            self.status_panel.add_log(f"설정이 '{self.CONFIG_FILE}'에 저장되었습니다.", "INFO")
        except Exception as e:
            self.status_panel.add_log(f"설정 저장 실패: {e}", "ERROR")

    def load_config(self):
        if not os.path.exists(self.CONFIG_FILE): return
        try:
            with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                self.strategy_data = json.load(f)
            self.table.setRowCount(0)
            for code, data in self.strategy_data.items():
                stock_name = self.creon.get_stock_name(code) or "이름 조회 실패"
                new_row_data = (data.get("on", True), code, stock_name, "0", "0", "0.00%", data.get("buy_flag", True), data.get("sell_flag", True))
                self.add_symbol_row(new_row_data, data)
            self.status_panel.add_log(f"'{self.CONFIG_FILE}'에서 설정을 불러왔습니다.", "SUCCESS")
            self.refresh_prices() # 설정 로드 후 잔고/수익률 갱신
        except Exception as e:
            self.status_panel.add_log(f"설정 불러오기 실패: {e}", "ERROR")
    
    def closeEvent(self, event):
        self.stop_auto_trading()
        self.save_config()
        event.accept()

# main 함수는 변경 없음
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())