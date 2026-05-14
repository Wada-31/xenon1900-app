"""
バーコードスキャナー制御モジュール
Honeywell Xenon1900 との通信を担当
"""

import serial
import serial.tools.list_ports
import threading
from typing import Callable, Optional


class BarcodeScanner:
    """Xenon1900 バーコードスキャナー制御クラス"""
    
    def __init__(self, baudrate: int = 9600, callback: Optional[Callable] = None):
        """
        初期化
        
        Args:
            baudrate: ボーレート（デフォルト: 9600）
            callback: バーコード読み込み時のコールバック関数
        """
        self.serial_port = None
        self.baudrate = baudrate
        self.callback = callback
        self.is_connected = False
        self.reading_thread = None
        self.stop_reading = False
    
    @staticmethod
    def get_available_ports() -> list:
        """
        利用可能なシリアルポートを取得
        
        Returns:
            ポート名のリスト（例：['COM3', 'COM4']）
        """
        ports = []
        for port in serial.tools.list_ports.comports():
            ports.append(port.device)
        return sorted(ports)
    
    def connect(self, port: str) -> bool:
        """
        シリアルポートに接続
        
        Args:
            port: ポート名（例：'COM3'）
            
        Returns:
            接続成功フラグ
        """
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=self.baudrate,
                timeout=1
            )
            self.is_connected = True
            
            # バーコード読み込みスレッドを開始
            self.stop_reading = False
            self.reading_thread = threading.Thread(target=self._read_barcodes, daemon=True)
            self.reading_thread.start()
            
            return True
        except Exception as e:
            print(f"接続エラー: {str(e)}")
            self.is_connected = False
            return False
    
    def disconnect(self) -> bool:
        """
        シリアルポートから切断
        
        Returns:
            切断成功フラグ
        """
        try:
            self.stop_reading = True
            
            if self.reading_thread:
                self.reading_thread.join(timeout=2)
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            
            self.is_connected = False
            return True
        except Exception as e:
            print(f"切断エラー: {str(e)}")
            return False
    
    def _read_barcodes(self) -> None:
        """
        バーコードを読み込むメソッド（スレッド内で実行）
        """
        while not self.stop_reading and self.is_connected:
            try:
                if self.serial_port and self.serial_port.is_open:
                    # シリアルデータを読み込み
                    if self.serial_port.in_waiting > 0:
                        line = self.serial_port.readline()
                        if line:
                            # バーコード文字列をデコード
                            barcode = line.decode('utf-8').strip()
                            
                            # コールバック関数を呼び出し
                            if self.callback:
                                self.callback(barcode)
            except Exception as e:
                print(f"読み込みエラー: {str(e)}")
                break
    
    def send_command(self, command: str) -> bool:
        """
        スキャナーにコマンドを送信
        
        Args:
            command: 送信コマンド
            
        Returns:
            送信成功フラグ
        """
        try:
            if self.serial_port and self.is_connected:
                self.serial_port.write((command + '\r\n').encode('utf-8'))
                return True
            return False
        except Exception as e:
            print(f"送信エラー: {str(e)}")
            return False
