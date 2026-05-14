"""
Tkinter UIアプリケーション
Xenon1900バーコードスキャナーのメインUIを提供
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import threading
from typing import Optional

import sys
from pathlib import Path

# 親ディレクトリをPythonパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from barcode_scanner import BarcodeScanner
from excel_handler import ExcelHandler


class XenonApp:
    """Xenon1900 Barcode Scanner Appのメインウィンドウクラス"""
    
    def __init__(self):
        """初期化"""
        self.root = tk.Tk()
        self.root.title("Xenon1900 Barcode Scanner to Excel")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # モジュールの初期化
        self.scanner = BarcodeScanner(callback=self.on_barcode_read)
        self.excel_handler = ExcelHandler()
        
        # UI要素の初期化
        self.setup_ui()
        
        # ウィンドウ閉じる時の処理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_ui(self) -> None:
        """UIの構成要素を設定"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # グリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # ===== 1. 接続設定フレーム =====
        connection_frame = ttk.LabelFrame(main_frame, text="接続設定", padding="10")
        connection_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        connection_frame.columnconfigure(1, weight=1)
        
        ttk.Label(connection_frame, text="シリアルポート:").grid(row=0, column=0, sticky=tk.W)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(
            connection_frame,
            textvariable=self.port_var,
            values=BarcodeScanner.get_available_ports(),
            state="readonly",
            width=15
        )
        self.port_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Button(
            connection_frame,
            text="ポート更新",
            command=self.refresh_ports
        ).grid(row=0, column=2, padx=5)
        
        self.connect_button = ttk.Button(
            connection_frame,
            text="接続",
            command=self.toggle_connection
        )
        self.connect_button.grid(row=0, column=3, padx=5)
        
        # ステータス表示
        ttk.Label(connection_frame, text="ステータス:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.status_var = tk.StringVar(value="未接続")
        self.status_label = ttk.Label(
            connection_frame,
            textvariable=self.status_var,
            foreground="red",
            font=("Arial", 10, "bold")
        )
        self.status_label.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # ===== 2. バーコード読み込みフレーム =====
        input_frame = ttk.LabelFrame(main_frame, text="バーコード読み込み", padding="10")
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="最新バーコード:").grid(row=0, column=0, sticky=tk.W)
        self.barcode_var = tk.StringVar(value="---")
        barcode_label = ttk.Label(
            input_frame,
            textvariable=self.barcode_var,
            font=("Arial", 12, "bold"),
            foreground="blue"
        )
        barcode_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(input_frame, text="読み込み日時:").grid(row=1, column=0, sticky=tk.W)
        self.time_var = tk.StringVar(value="---")
        ttk.Label(input_frame, textvariable=self.time_var).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(input_frame, text="メッセージ:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.message_var = tk.StringVar(value="待機中...")
        self.message_label = ttk.Label(
            input_frame,
            textvariable=self.message_var,
            font=("Arial", 9),
            foreground="green"
        )
        self.message_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # ===== 3. 統計情報フレーム =====
        stats_frame = ttk.LabelFrame(main_frame, text="統計情報", padding="10")
        stats_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        stats_frame.columnconfigure(1, weight=1)
        stats_frame.columnconfigure(3, weight=1)
        
        ttk.Label(stats_frame, text="総件数:").grid(row=0, column=0, sticky=tk.W)
        self.total_var = tk.StringVar(value="0")
        ttk.Label(stats_frame, textvariable=self.total_var, font=("Arial", 11, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(stats_frame, text="重複件数:").grid(row=0, column=2, sticky=tk.W)
        self.duplicate_var = tk.StringVar(value="0")
        ttk.Label(stats_frame, textvariable=self.duplicate_var, font=("Arial", 11, "bold"), foreground="red").grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # ===== 4. 登録済みバーコード一覧フレーム =====
        list_frame = ttk.LabelFrame(main_frame, text="登録済みバーコード一覧", padding="10")
        list_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # ツリービューのスクロールバー
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # ツリービュー
        self.tree = ttk.Treeview(
            list_frame,
            columns=("No", "Barcode", "DateTime"),
            height=10,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.tree.yview)
        
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("No", anchor=tk.CENTER, width=50)
        self.tree.column("Barcode", anchor=tk.CENTER, width=300)
        self.tree.column("DateTime", anchor=tk.CENTER, width=200)
        
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading("No", text="No.", anchor=tk.CENTER)
        self.tree.heading("Barcode", text="バーコード", anchor=tk.CENTER)
        self.tree.heading("DateTime", text="読み込み日時", anchor=tk.CENTER)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ===== 5. 操作ボタンフレーム =====
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=10)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)
        button_frame.columnconfigure(3, weight=1)
        
        ttk.Button(
            button_frame,
            text="Excelに保存",
            command=self.save_to_excel
        ).grid(row=0, column=0, padx=5, sticky=(tk.W, tk.E))
        
        ttk.Button(
            button_frame,
            text="データクリア",
            command=self.clear_data
        ).grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
        
        ttk.Button(
            button_frame,
            text="一覧更新",
            command=self.update_treeview
        ).grid(row=0, column=2, padx=5, sticky=(tk.W, tk.E))
        
        ttk.Button(
            button_frame,
            text="終了",
            command=self.on_closing
        ).grid(row=0, column=3, padx=5, sticky=(tk.W, tk.E))
        
        # 初期表示
        self.refresh_ports()
        self.update_statistics()
    
    def refresh_ports(self) -> None:
        """利用可能なシリアルポートを更新"""
        ports = BarcodeScanner.get_available_ports()
        self.port_combo['values'] = ports
        if ports:
            self.port_combo.current(0)
    
    def toggle_connection(self) -> None:
        """接続/切断をトグル"""
        if self.scanner.is_connected:
            self.disconnect_scanner()
        else:
            self.connect_scanner()
    
    def connect_scanner(self) -> None:
        """スキャナーに接続"""
        port = self.port_var.get()
        if not port:
            messagebox.showerror("エラー", "シリアルポートを選択してください")
            return
        
        if self.scanner.connect(port):
            self.connect_button.config(text="切断")
            self.status_var.set("接続済み")
            self.status_label.config(foreground="green")
            self.message_var.set("スキャナーに接続しました。バーコードをスキャンしてください。")
            self.message_label.config(foreground="green")
        else:
            messagebox.showerror("エラー", f"ポート {port} に接続できません")
    
    def disconnect_scanner(self) -> None:
        """スキャナーから切断"""
        if self.scanner.disconnect():
            self.connect_button.config(text="接続")
            self.status_var.set("未接続")
            self.status_label.config(foreground="red")
            self.message_var.set("スキャナーから切断しました。")
            self.message_label.config(foreground="orange")
    
    def on_barcode_read(self, barcode: str) -> None:
        """
        バーコード読み込みコールバック
        
        Args:
            barcode: 読み込まれたバーコード
        """
        # メインスレッドでUIを更新
        self.root.after(0, self._update_ui_barcode_read, barcode)
    
    def _update_ui_barcode_read(self, barcode: str) -> None:
        """UIをバーコード読み込み結果で更新"""
        # 最新バーコード表示
        self.barcode_var.set(barcode)
        self.time_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        # Excelハンドラーに追加
        success, message = self.excel_handler.add_barcode(barcode)
        
        if success:
            self.message_var.set(message)
            self.message_label.config(foreground="green")
        else:
            self.message_var.set(message)
            self.message_label.config(foreground="red")
            messagebox.showwarning("重複警告", message)
        
        # 統計情報を更新
        self.update_statistics()
        self.update_treeview()
    
    def update_statistics(self) -> None:
        """統計情報を更新"""
        stats = self.excel_handler.get_statistics()
        self.total_var.set(str(stats["total"]))
        self.duplicate_var.set(str(stats["duplicates"]))
    
    def update_treeview(self) -> None:
        """ツリービューを更新"""
        # 既存のアイテムをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # バーコードをツリービューに追加
        barcodes = self.excel_handler.get_barcodes()
        for idx, barcode in enumerate(barcodes, 1):
            self.tree.insert(
                "",
                "end",
                values=(idx, barcode, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
    
    def save_to_excel(self) -> None:
        """Excelファイルに保存"""
        if not self.excel_handler.get_barcodes():
            messagebox.showwarning("警告", "保存するバーコードがありません")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
            initialfile=f"barcode_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if file_path:
            success, message = self.excel_handler.save_to_excel(file_path)
            if success:
                messagebox.showinfo("成功", message)
                self.message_var.set(message)
                self.message_label.config(foreground="green")
            else:
                messagebox.showerror("エラー", message)
                self.message_var.set(message)
                self.message_label.config(foreground="red")
    
    def clear_data(self) -> None:
        """データをクリア"""
        if messagebox.askyesno("確認", "すべてのバーコードデータをクリアしてもよろしいですか？"):
            self.excel_handler.clear_all()
            self.barcode_var.set("---")
            self.time_var.set("---")
            self.message_var.set("データをクリアしました。")
            self.message_label.config(foreground="green")
            self.update_statistics()
            self.update_treeview()
    
    def on_closing(self) -> None:
        """ウィンドウを閉じる処理"""
        if self.scanner.is_connected:
            self.scanner.disconnect()
        self.root.destroy()
    
    def run(self) -> None:
        """アプリケーションを実行"""
        self.root.mainloop()


if __name__ == "__main__":
    app = XenonApp()
    app.run()
