#!/usr/bin/env python3
"""
音声テキスト変換ツール（GUI版）
進行状況を表示しながら音声をテキストへ変換する PyQt6 アプリケーション。
"""

import sys
import os
import threading
import json
from pathlib import Path
from datetime import datetime
from typing import Optional
from dotenv import load_dotenv
try:
    import torch
except ImportError:  # torch is optional for GPU detection with faster-whisper
    torch = None

from faster_whisper import WhisperModel

load_dotenv()

# OpenAI API キーを環境変数から取得
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QWidget, QPushButton, QLabel, QFileDialog, QProgressBar,
                            QTextEdit, QComboBox, QGroupBox, QMessageBox,
                            QListWidget, QListWidgetItem, QSplitter, QFrame, QDialog, QDialogButtonBox, QCheckBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QIcon, QPixmap

# Import our custom modules
try:
    from summarizer import Summarizer
    from audio import ConversionWorker
except ImportError as e:
    print(f"警告: カスタムモジュールを読み込めませんでした: {e}")
    
    Summarizer = None
    ConversionWorker = None


SUMMARY_CONFIGS = [
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::特定事業所加算　保存様式",
        "run_method": "run_sheet1",
        "insert_method": "insert_sheet1",
        "status_label": "特定事業所加算 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::入院時情報提供書",
        "run_method": "run_sheet2",
        "insert_method": "insert_sheet2",
        "status_label": "入院時情報提供書"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::退院・退所加算　保存様式",
        "run_method": "run_sheet3",
        "insert_method": "insert_sheet3",
        "status_label": "退院・退所加算 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::居宅介護支援事業所等連携加算　保存様式",
        "run_method": "run_sheet4",
        "insert_method": "insert_sheet4",
        "status_label": "居宅介護支援事業所等連携加算 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::医療・保育・教育連携加算　保存様式",
        "run_method": "run_sheet5",
        "insert_method": "insert_sheet5",
        "status_label": "医療・保育・教育連携加算 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::サービス担当者会議記録　保存様式",
        "run_method": "run_sheet6",
        "insert_method": "insert_sheet6",
        "status_label": "サービス担当者会議記録 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::サービス提供時モニタリング記録　保存様式",
        "run_method": "run_sheet7",
        "insert_method": "insert_sheet7",
        "status_label": "サービス提供時モニタリング記録 保存様式"
    },
    {
        "summary_key": "サービス担当者会議記録（保存用）標準様式　各種加算標準様式（Excel形式：92KB）.xlsx::体制加算　記録",
        "run_method": "run_sheet8",
        "insert_method": "insert_sheet8",
        "status_label": "体制加算 記録"
    },
    {
        "summary_key": "様式11　モニタリング報告書（Excel形式：45KB）.xlsx",
        "run_method": "run_sheet_monitor",
        "insert_method": "insert_monitor_sheet",
        "status_label": "モニタリング報告書"
    },
    {
        "summary_key": "様式4　サービス等利用計画案・障害児支援利用計画案（Excel形式：45KB）.xlsx",
        "run_method": "run_sheet_proposedPlan",
        "insert_method": "insert_proposedPlan_sheet",
        "status_label": "サービス等利用計画案"
    },
    {
        "summary_key": "様式8　サービス等利用計画・障害児支援利用計画（Excel形式：46KB）.xlsx",
        "run_method": "run_sheet_plan",
        "insert_method": "insert_Plan_sheet",
        "status_label": "サービス等利用計画"
    },
    {
        "summary_key": "様式2、3　アセスメント票（訪問票兼生活支援アセスメント票）（Excel形式：44KB）.xlsx",
        "run_method": "run_sheet_assessment",
        "insert_method": "insert_assessment_sheet",
        "status_label": "アセスメント票"
    },
]

SUMMARY_KEYS = [cfg["summary_key"] for cfg in SUMMARY_CONFIGS]
SUMMARY_KEY_SET = set(SUMMARY_KEYS)

# Mapping from document type checkboxes to SUMMARY_CONFIGS indices
DOCUMENT_TYPE_MAPPING = {
    'service_meeting': list(range(0, 8)),  # Indices 0-7 (sheets 1-8)
    'monitoring': [8],  # Index 8 (モニタリング報告書)
    'proposed_plan': [9],  # Index 9 (サービス等利用計画案)
    'plan': [10],  # Index 10 (サービス等利用計画)
    'assessment': [11],  # Index 11 (アセスメント票)
}

class StatusDialog(QDialog):
    """ターミナルの代わりに処理状況を表示するダイアログ。"""
    
    def __init__(self, parent=None, title="処理中"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setFixedSize(400, 200)
        
        layout = QVBoxLayout()
        
        # ステータスラベル
        self.status_label = QLabel("初期化しています…")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("QLabel { font-size: 12px; padding: 10px; }")
        layout.addWidget(self.status_label)
        
        # 進捗バー
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 詳細表示欄
        self.details_text = QTextEdit()
        self.details_text.setMaximumHeight(80)
        self.details_text.setReadOnly(True)
        self.details_text.setVisible(False)
        layout.addWidget(self.details_text)
        
        # ボタン
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        
        self.setLayout(layout)
        
    def update_status(self, message, show_progress=False, progress_value=None):
        """ステータスメッセージを更新し、必要に応じて進捗を表示する。"""
        self.status_label.setText(message)
        if show_progress:
            self.progress_bar.setVisible(True)
            if progress_value is not None:
                self.progress_bar.setValue(progress_value)
        else:
            self.progress_bar.setVisible(False)
        QApplication.processEvents()
        
    def add_detail(self, detail):
        """詳細欄にメッセージを追加する。"""
        self.details_text.setVisible(True)
        self.details_text.append(detail)
        QApplication.processEvents()
        
    def show_success(self, message):
        """成功メッセージを表示し、OK ボタンを有効化する。"""
        self.status_label.setText(f"✅ {message}")
        self.status_label.setStyleSheet("QLabel { color: green; font-weight: bold; font-size: 12px; padding: 10px; }")
        self.button_box.clear()
        self.button_box.addButton(QDialogButtonBox.StandardButton.Ok)
        self.button_box.accepted.connect(self.accept)
        
    def show_error(self, message):
        """エラーメッセージを表示し、OK ボタンを有効化する。"""
        self.status_label.setText(f"❌ {message}")
        self.status_label.setStyleSheet("QLabel { color: red; font-weight: bold; font-size: 12px; padding: 10px; }")
        self.button_box.clear()
        self.button_box.addButton(QDialogButtonBox.StandardButton.Ok)
        self.button_box.accepted.connect(self.accept)


class SummarizationWorker(QThread):
    """Worker thread for text summarization using OpenAI."""
    
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    summarization_completed = pyqtSignal(str)  # summarized text
    summarization_failed = pyqtSignal(str)  # error message
    
    def __init__(self, text, api_key=None, language='ja-JP', selected_doc_types=None):
        super().__init__()
        self.text = text
        self.api_key = api_key
        self.language = language
        self.selected_doc_types = selected_doc_types or []  # List of selected document type keys
    
    def run(self):
        try:
            self.status_updated.emit("OpenAI クライアントを初期化しています…")
            self.progress_updated.emit(5)
            
            if not Summarizer:
                raise Exception("OpenAI client not available. Please check dependencies.")
            
            
            
            # Filter SUMMARY_CONFIGS based on selected document types
            if not self.selected_doc_types:
                # If no document types are selected, raise an error
                raise ValueError("出力する帳票が選択されていません。少なくとも1つの帳票を選択してください。")
            
            # Get all indices for selected document types
            selected_indices = set()
            for doc_type in self.selected_doc_types:
                if doc_type in DOCUMENT_TYPE_MAPPING:
                    selected_indices.update(DOCUMENT_TYPE_MAPPING[doc_type])
            
            if not selected_indices:
                # If no valid indices found, raise an error
                raise ValueError("選択された帳票に対応する設定が見つかりませんでした。")
            
            # Filter configs to only include selected ones
            filtered_configs = [cfg for idx, cfg in enumerate(SUMMARY_CONFIGS) if idx in selected_indices]
            
            client = Summarizer(self.api_key, self.text, self.language)
            self.progress_updated.emit(5)
            
            self.status_updated.emit("Excel 用の情報抽出手順を分析しています…")
            self.progress_updated.emit(10)
            
            self.status_updated.emit("Excel の構造に沿って回答を抽出しています…")
            self.progress_updated.emit(15)
            
            sections = []
            total_sections = len(filtered_configs)
            # Cache for shared results (サービス利用計画案とサービス利用計画は同じ内容を使用)
            proposed_plan_result = None

            for idx, cfg in enumerate(filtered_configs, start=1):
                status_label = cfg["status_label"]
                run_method_name = cfg["run_method"]
                
                # サービス利用計画案とサービス利用計画は同じ内容を使用
                # run_sheet_planの場合は、常にrun_sheet_proposedPlanの結果を再利用
                if run_method_name == "run_sheet_plan":
                    if proposed_plan_result is not None:
                        # 既にrun_sheet_proposedPlanが実行済みの場合、その結果を再利用
                        # フィールド名を変換: 「計画案::」→「計画::」、「計画案週::」→「週間計画::」
                        section_content = proposed_plan_result.replace('"計画案::', '"計画::').replace('"計画案週::', '"週間計画::')
                        self.status_updated.emit(f"{status_label} を解析しています... (サービス利用計画案の結果を再利用)")
                    else:
                        # run_sheet_proposedPlanを先に実行して結果を取得
                        proposed_plan_method = getattr(client, "run_sheet_proposedPlan", None)
                        if callable(proposed_plan_method):
                            try:
                                self.status_updated.emit("サービス等利用計画案 を解析しています... (サービス利用計画でも使用)")
                                proposed_plan_result = proposed_plan_method().strip()
                                # フィールド名を変換: 「計画案::」→「計画::」、「計画案週::」→「週間計画::」
                                section_content = proposed_plan_result.replace('"計画案::', '"計画::').replace('"計画案週::', '"週間計画::')
                            except Exception as exc:
                                section_content = json.dumps({"error": str(exc)}, ensure_ascii=False)
                                self.status_updated.emit(f"⚠️ サービス等利用計画案 の解析に失敗しました: {exc}")
                        else:
                            # フォールバック: run_sheet_planを実行
                            run_method = getattr(client, run_method_name, None)
                            if not callable(run_method):
                                section_content = json.dumps({"error": f"Method {run_method_name} not found"}, ensure_ascii=False)
                                self.status_updated.emit(f"⚠️ {status_label} の処理メソッドが見つかりませんでした")
                            else:
                                try:
                                    self.status_updated.emit(f"{status_label} を解析しています...")
                                    section_content = run_method().strip()
                                except Exception as exc:
                                    section_content = json.dumps({"error": str(exc)}, ensure_ascii=False)
                                    self.status_updated.emit(f"⚠️ {status_label} の解析に失敗しました: {exc}")
                else:
                    run_method = getattr(client, run_method_name, None)

                    if not callable(run_method):
                        section_content = json.dumps({"error": f"Method {run_method_name} not found"}, ensure_ascii=False)
                        self.status_updated.emit(f"⚠️ {status_label} の処理メソッドが見つかりませんでした")
                    else:
                        try:
                            self.status_updated.emit(f"{status_label} を解析しています...")
                            section_content = run_method().strip()
                            # run_sheet_proposedPlanの結果をキャッシュに保存
                            if run_method_name == "run_sheet_proposedPlan":
                                proposed_plan_result = section_content
                        except Exception as exc:
                            section_content = json.dumps({"error": str(exc)}, ensure_ascii=False)
                            self.status_updated.emit(f"⚠️ {status_label} の解析に失敗しました: {exc}")

                sections.append(f"{cfg['summary_key']}:\n{section_content}")
                progress = 15 + int((idx / total_sections) * 75)
                self.progress_updated.emit(min(progress, 95))

            summary = "\n--------------------------------\n".join(sections)
            self.progress_updated.emit(100)
            
            self.status_updated.emit("✓ 要約処理が完了しました！")
            self.summarization_completed.emit(summary)
            
        except Exception as e:
            self.summarization_failed.emit(str(e))


class ClassificationWorker(QThread):
    """Worker thread for text classification and Excel insertion."""
    
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    classification_completed = pyqtSignal(dict)  # classification results
    classification_failed = pyqtSignal(str)  # error message
    
    def __init__(self, summarized_text, api_key=None, output_dir: Optional[Path] = None, selected_doc_types=None):
        super().__init__()
        self.summarized_text = summarized_text
        self.api_key = api_key
        self.output_dir = Path(output_dir) if output_dir else None
        self.selected_doc_types = selected_doc_types or []  # List of selected document type keys
    
    def _extract_separate_texts(self, summary_text):
        """Extract separate texts from summary by splitting on separators and parse JSON to match clarify_sheet1 format."""
        extracted_texts = {}
        
        if not summary_text or not summary_text.strip():
            return extracted_texts
        
        # Split by separator line
        sections = summary_text.split("--------------------------------")
        
        for section in sections:
            lines = [line.rstrip() for line in section.split("\n") if line.strip()]
            if not lines:
                continue

            header_line = lines[0]
            if header_line.endswith(":"):
                header_key = header_line[:-1].strip()
            else:
                header_key = header_line.strip()

            if header_key not in SUMMARY_KEY_SET:
                continue

            content = "\n".join(lines[1:]).strip()

            if not content:
                parsed_data = {}
            else:
                try:
                    parsed_data = json.loads(content)
                except json.JSONDecodeError:
                    parsed_data = {"content": content}

            extracted_texts[header_key] = parsed_data
        
        return extracted_texts
    
    def run(self):
        try:
            self.status_updated.emit("OpenAI クライアントを初期化しています…")
            self.progress_updated.emit(0)
            
            # Extract separate texts from summary
            self.status_updated.emit("要約から帳票ごとのテキストを切り出しています…")
            text = self.summarized_text
            extracted_texts = self._extract_separate_texts(text)
            self.progress_updated.emit(3)
            
            # Log extracted texts for debugging
            if extracted_texts:
                self.status_updated.emit(f"{len(extracted_texts)} 件の帳票用テキストを抽出しました: {', '.join(extracted_texts.keys())}")
            
            if not Summarizer:
                raise Exception("OpenAI クライアントを利用できません。依存関係を確認してください。")
            
            client = Summarizer(api_key=self.api_key, output_dir=self.output_dir)
            self.progress_updated.emit(8)
            
            self.status_updated.emit("Excel テンプレートへの転記を準備しています…")
            self.progress_updated.emit(10)

            # Filter SUMMARY_CONFIGS based on selected document types
            if not self.selected_doc_types:
                # If no document types are selected, raise an error
                raise ValueError("出力する帳票が選択されていません。少なくとも1つの帳票を選択してください。")
            
            # Get all indices for selected document types
            selected_indices = set()
            for doc_type in self.selected_doc_types:
                if doc_type in DOCUMENT_TYPE_MAPPING:
                    selected_indices.update(DOCUMENT_TYPE_MAPPING[doc_type])
            
            if not selected_indices:
                # If no valid indices found, raise an error
                raise ValueError("選択された帳票に対応する設定が見つかりませんでした。")
            
            # Filter configs to only include selected ones
            filtered_configs = [cfg for idx, cfg in enumerate(SUMMARY_CONFIGS) if idx in selected_indices]
            
            insertion_results = {}
            total_sections = len(filtered_configs)

            for idx, cfg in enumerate(filtered_configs, start=1):
                summary_key = cfg["summary_key"]
                status_label = cfg["status_label"]
                insert_method_name = cfg["insert_method"]
                insert_method = getattr(client, insert_method_name, None)

                data = extracted_texts.get(summary_key)
                if isinstance(data, dict):
                    payload = data
                    has_data = bool(data)
                elif data is None:
                    payload = {}
                    has_data = False
                else:
                    payload = {"content": data}
                    has_data = True

                if not callable(insert_method):
                    insertion_results[summary_key] = {"success": False, "error": f"Method {insert_method_name} not found"}
                    self.status_updated.emit(f"⚠️ {status_label} の挿入メソッドが見つかりませんでした")
                    continue

                try:
                    action_msg = "データ付きテンプレートに挿入" if has_data else "テンプレートをコピー"
                    self.status_updated.emit(f"{status_label} ({action_msg}) を処理しています...")
                    saved_path = insert_method(payload)
                    insertion_results[summary_key] = {"success": True, "path": saved_path, "has_data": has_data}
                    if has_data:
                        self.status_updated.emit(f"✓ {status_label} への挿入が完了しました！")
                    else:
                        self.status_updated.emit(f"✓ {status_label} のテンプレートを出力しました (データなし)")
                except Exception as exc:
                    insertion_results[summary_key] = {"success": False, "error": str(exc)}
                    self.status_updated.emit(f"⚠️ {status_label} への挿入に失敗しました: {exc}")

                progress = 10 + int((idx / total_sections) * 90)
                self.progress_updated.emit(min(progress, 100))

            self.progress_updated.emit(100)
            self.status_updated.emit("✓ 分類と転記が完了しました！")

            # Only include classification results for selected document types
            selected_summary_keys = [cfg["summary_key"] for cfg in filtered_configs]
            ordered_classification = {key: extracted_texts.get(key) for key in selected_summary_keys if key in extracted_texts}
            results_payload = {
                "classification": ordered_classification,
                "insertion": insertion_results,
                "output_dir": str(self.output_dir) if self.output_dir else None
            }

            self.classification_completed.emit(results_payload)
            
        except Exception as e:
            self.classification_failed.emit(str(e))


class AudioToTextGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_worker = None
        self.summarization_worker = None
        self.classification_worker = None
        self.summarized_text = ""
        self.init_ui()
        self.model = self._load_whisper_model()

    def _load_whisper_model(self):
        """Initialise the faster-whisper model with environment overrides."""
        model_size = os.getenv("WHISPER_MODEL_SIZE", "small")
        default_device = "cpu"
        if torch is not None and torch.cuda.is_available():
            default_device = "cuda"
        device = os.getenv("WHISPER_DEVICE", default_device)

        default_compute_type = "float16" if device == "cuda" else "int8"
        compute_type = os.getenv("WHISPER_COMPUTE_TYPE", default_compute_type)

        return WhisperModel(model_size, device=device, compute_type=compute_type)
        
    def init_ui(self):
        self.setWindowTitle("音声テキスト変換 & AI 分析ツール")
        self.setGeometry(100, 100, 1000, 700)
        self.setFixedSize(1000, 700)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # Left panel - File selection and controls
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)
        
        # Right panel - Progress and results
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # Set splitter proportions
        splitter.setSizes([350, 650])
        
        # Status bar
        self.statusBar().showMessage("音声ファイルの変換を待機しています")
        
    def create_left_panel(self):
        """Create the left panel with file selection and controls."""
        panel = QFrame()
        panel.setFrameStyle(QFrame.Shape.StyledPanel)
        layout = QVBoxLayout(panel)
        
        # File selection group
        file_group = QGroupBox("ファイル選択")
        file_layout = QVBoxLayout(file_group)
        
        # File path display
        self.file_path_label = QLabel("ファイルが選択されていません")
        self.file_path_label.setWordWrap(True)
        self.file_path_label.setStyleSheet("QLabel { background-color: #f0f0f0; padding: 5px; border: 1px solid #ccc; color : black; }")
        file_layout.addWidget(self.file_path_label)
        
        # File selection button
        self.browse_button = QPushButton("音声ファイルを選択")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.browse_button)
        
        layout.addWidget(file_group)
        
        # Settings group
        settings_group = QGroupBox("設定")
        settings_layout = QVBoxLayout(settings_group)
        
        # Language selection
        lang_layout = QHBoxLayout()
        lang_layout.addWidget(QLabel("言語:"))
        self.language_combo = QComboBox()
        self.language_combo.addItems([
            "ja-JP 日本語（日本）",
            "en-US 英語（米国）",
            "en-GB 英語（英国）", 
            "es-ES スペイン語（スペイン）",
            "fr-FR フランス語（フランス）",
            "de-DE ドイツ語（ドイツ）",
            "it-IT イタリア語（イタリア）",
            "pt-BR ポルトガル語（ブラジル）",
            "ko-KR 韓国語（韓国）",
            "zh-CN 中国語（簡体字）"
        ])
        lang_layout.addWidget(self.language_combo)
        settings_layout.addLayout(lang_layout)
        
        layout.addWidget(settings_group)
        
        # Document type selection group
        doc_type_group = QGroupBox("出力する帳票を選択")
        doc_type_layout = QVBoxLayout(doc_type_group)
        
        # Create checkboxes for document types
        self.doc_type_checkboxes = {}
        
        # サービス担当者会議記録 (maps to sheets 1-8)
        self.doc_type_checkboxes['service_meeting'] = QCheckBox("サービス担当者会議記録")
        self.doc_type_checkboxes['service_meeting'].setChecked(True)  # Default checked
        doc_type_layout.addWidget(self.doc_type_checkboxes['service_meeting'])
        
        # アセスメント票
        self.doc_type_checkboxes['assessment'] = QCheckBox("アセスメント票")
        self.doc_type_checkboxes['assessment'].setChecked(True)  # Default checked
        doc_type_layout.addWidget(self.doc_type_checkboxes['assessment'])
        
        # サービス利用計画案
        self.doc_type_checkboxes['proposed_plan'] = QCheckBox("サービス利用計画案")
        self.doc_type_checkboxes['proposed_plan'].setChecked(True)  # Default checked
        doc_type_layout.addWidget(self.doc_type_checkboxes['proposed_plan'])
        
        # サービス利用計画
        self.doc_type_checkboxes['plan'] = QCheckBox("サービス利用計画")
        self.doc_type_checkboxes['plan'].setChecked(True)  # Default checked
        doc_type_layout.addWidget(self.doc_type_checkboxes['plan'])
        
        # モニタリング表
        self.doc_type_checkboxes['monitoring'] = QCheckBox("モニタリング表")
        self.doc_type_checkboxes['monitoring'].setChecked(True)  # Default checked
        doc_type_layout.addWidget(self.doc_type_checkboxes['monitoring'])
        
        layout.addWidget(doc_type_group)
        
        # Control buttons
        control_group = QGroupBox("操作")
        control_layout = QVBoxLayout(control_group)
        
        self.convert_button = QPushButton("文字起こし開始")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")
        control_layout.addWidget(self.convert_button)
        
        self.summarize_button = QPushButton("回答抽出")
        self.summarize_button.clicked.connect(self.summarize_text)
        self.summarize_button.setEnabled(False)
        self.summarize_button.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 10px; }")
        control_layout.addWidget(self.summarize_button)
        
        self.classification_button = QPushButton("入力")
        self.classification_button.clicked.connect(self.classify_text)
        self.classification_button.setEnabled(False)
        self.classification_button.setStyleSheet("QPushButton { background-color: #FF9800; color: white; font-weight: bold; padding: 10px; }")
        control_layout.addWidget(self.classification_button)
        
        self.stop_button = QPushButton("停止")
        self.stop_button.clicked.connect(self.stop_conversion)
        self.stop_button.setEnabled(False)
        self.stop_button.setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 10px; }")
        control_layout.addWidget(self.stop_button)
        
        layout.addWidget(control_group)
        
        # Add stretch to push everything to top
        layout.addStretch()
        
        return panel
    
    def create_right_panel(self):
        """Create the right panel with progress and results."""
        panel = QFrame()
        panel.setFrameStyle(QFrame.Shape.StyledPanel)
        layout = QVBoxLayout(panel)
        
        # Progress group
        progress_group = QGroupBox("処理状況")
        progress_layout = QVBoxLayout(progress_group)
        
        # Status label
        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet("QLabel { font-weight: bold; color: #333; }")
        progress_layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        
        # Loading indicator (spinning icon)
        self.loading_label = QLabel("⏳")
        self.loading_label.setVisible(False)
        self.loading_label.setStyleSheet("QLabel { font-size: 24px; color: #4CAF50; }")
        progress_layout.addWidget(self.loading_label)
        
        layout.addWidget(progress_group)
        
        # Results group
        results_group = QGroupBox("文字起こし・AI 抽出結果")
        results_layout = QVBoxLayout(results_group)
        
        # Results text area
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)  # Default to read-only mode
        self.results_text.setPlaceholderText("リアルタイムの文字起こし結果と抽出内容がここに表示されます。\n『編集モード切替』で内容を修正できます。")
        # Set a clean font for better readability
        font = QFont("Segoe UI", 10)
        self.results_text.setFont(font)
        # Set text color to make it more readable (read-only styling)
        self.results_text.setStyleSheet("QTextEdit { color: #333; background-color: #f9f9f9; border: 1px solid #ccc; }")
        # Connect text change signal to enable/disable summarize button
        self.results_text.textChanged.connect(self.on_text_changed)
        results_layout.addWidget(self.results_text)
        
        # Results buttons
        results_button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("ファイルへ保存")
        self.save_button.clicked.connect(self.save_results)
        self.save_button.setEnabled(False)
        results_button_layout.addWidget(self.save_button)
        
        self.edit_toggle_button = QPushButton("編集モード切替")
        self.edit_toggle_button.clicked.connect(self.toggle_edit_mode)
        self.edit_toggle_button.setStyleSheet("QPushButton { background-color: #9C27B0; color: white; font-weight: bold; padding: 5px; }")
        results_button_layout.addWidget(self.edit_toggle_button)
        
        self.clear_button = QPushButton("結果をクリア")
        self.clear_button.clicked.connect(self.clear_results)
        results_button_layout.addWidget(self.clear_button)
        
        results_layout.addLayout(results_button_layout)
        
        layout.addWidget(results_group)
        
        return panel
    
    def browse_file(self):
        """Browse for a single audio file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "音声ファイルの選択",
            "",
            "音声ファイル (*.wav *.mp3 *.m4a *.flac *.aiff *.ogg);;すべてのファイル (*)"
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_path_label.setText(f"選択済み: {Path(file_path).name}")
    
    
    def start_conversion(self):
        """Start the audio conversion process."""
        if not hasattr(self, 'selected_file') or not self.selected_file:
            QMessageBox.warning(self, "ファイル未選択", "変換する音声ファイルを選択してください。")
            return
        
        if not os.path.exists(self.selected_file):
            QMessageBox.warning(self, "ファイルが見つかりません", "指定されたファイルが存在しません。")
            return
        
        # Get language code
        language_text = self.language_combo.currentText()
        language_code = language_text.split(' ')[0]  # Extract language code
        
        # Disable controls during conversion
        self.convert_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.loading_label.setVisible(True)
        
        # Clear previous results
        self.results_text.clear()
        self.save_button.setEnabled(False)
        self.summarize_button.setEnabled(False)
        self.classification_button.setEnabled(False)
        
        # Check if ConversionWorker is available
        if not ConversionWorker:
            QMessageBox.critical(self, "エラー", "音声変換モジュールを利用できません。依存関係を確認してください。")
            self.reset_ui()
            return
        
        # Start conversion in worker thread
        self.current_worker = ConversionWorker(self.selected_file, language_code, self.model)
        self.current_worker.progress_updated.connect(self.update_progress)
        self.current_worker.status_updated.connect(self.update_status)
        self.current_worker.partial_result_updated.connect(self.update_partial_results)
        self.current_worker.conversion_completed.connect(self.on_conversion_completed)
        self.current_worker.conversion_failed.connect(self.on_conversion_failed)
        self.current_worker.start()
    
    def stop_conversion(self):
        """Stop the current conversion."""
        if self.current_worker and self.current_worker.isRunning():
            self.current_worker.terminate()
            self.current_worker.wait()
        
        self.reset_ui()
        self.status_label.setText("ユーザー操作により変換を停止しました")
    
    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """Update the status label."""
        self.status_label.setText(message)
        self.statusBar().showMessage(message)
    
    def update_partial_results(self, text):
        """Update the results text area with real-time updates."""
        # Check if this is a transcription update (starts with 📝)
        if text.startswith("📝"):
            # Hide loading icon when transcription starts
            self.loading_label.setVisible(False)
            
            # Extract just the transcription text
            if "📝 文字起こし開始: " in text:
                # 初回の「文字起こし開始」メッセージ
                self.results_text.setPlainText("文字起こしを開始しました…")
            else:
                # This is actual transcription text
                transcription_text = text.replace("📝 ", "")
                self.results_text.setPlainText(transcription_text)
        else:
            # This is a status message, append to current text
            current_text = self.results_text.toPlainText()
            if current_text and not current_text.endswith('\n'):
                current_text += '\n'
            
            # Add timestamp for status messages
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            
            # Update the text area
            self.results_text.setPlainText(current_text + f"[{timestamp}] {text}\n")
        
        # Auto-scroll to bottom to show latest updates
        scrollbar = self.results_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
        # Process events to ensure UI updates immediately
        QApplication.processEvents()
    
    def on_conversion_completed(self, file_path, text):
        """Handle successful conversion completion."""
        # The text is already clean from the streaming display
        final_text = text
        
        # Update the results with clean final text
        self.results_text.setPlainText(final_text)
        # Note: save_button and summarize_button will be enabled by on_text_changed()
        # Keep classification button disabled until summarization is complete
        self.classification_button.setEnabled(False)
        self.reset_ui()
        
    def on_conversion_failed(self, file_path, error):
        """Handle conversion failure."""
        QMessageBox.critical(self, "変換に失敗しました", f"{Path(file_path).name} の変換に失敗しました。\n\n{error}")
        self.reset_ui()
    
    def reset_ui(self):
        """Reset UI elements after conversion."""
        self.convert_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.loading_label.setVisible(False)
        self.current_worker = None
    
    def save_results(self):
        """Save the transcription results to a file."""
        text = self.results_text.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "保存できるテキストがありません", "保存対象となる文字起こし結果がありません。")
            return
        
        default_name = self._default_output_filename()
        initial_path = str(Path.cwd() / default_name)
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "文字起こし結果の保存",
            initial_path,
            "テキストファイル (*.txt);;すべてのファイル (*)"
        )

        if not file_path:
            return

        if not file_path.lower().endswith('.txt'):
            file_path += '.txt'

        if self.save_to_file(text, file_path):
            QMessageBox.information(self, "保存完了", f"文字起こし結果を保存しました:\n{file_path}")
    
    def _default_output_filename(self) -> str:
        if getattr(self, 'selected_file', None):
            base_name = Path(self.selected_file).stem
            return f"{base_name}_transcript.txt"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"transcript_{timestamp}.txt"

    def save_to_file(self, text, filename):
        """Save text to file."""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(text)
            return True
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", f"ファイルを保存できませんでした:\n{str(e)}")
            return False
    
    def clear_results(self):
        """Clear the results text area."""
        self.results_text.clear()
        # Note: buttons will be disabled by on_text_changed()
        self.classification_button.setEnabled(False)
        self.summarized_text = ""
    
    def toggle_edit_mode(self):
        """Toggle between read-only and editable mode for results panel."""
        is_readonly = self.results_text.isReadOnly()
        self.results_text.setReadOnly(not is_readonly)
        
        if is_readonly:
            # 読み取り専用から編集可能へ切り替え
            self.results_text.setStyleSheet("QTextEdit { color: #333; background-color: #ffffff; border: 2px solid #4CAF50; }")
            self.edit_toggle_button.setText("読み取り専用に戻す")
            self.edit_toggle_button.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 5px; }")
            self.statusBar().showMessage("結果パネルを編集できます。分類前に内容を調整してください。")
        else:
            # Switch to read-only mode (currently editable, so make it read-only)
            self.results_text.setStyleSheet("QTextEdit { color: #333; background-color: #f9f9f9; border: 1px solid #ccc; }")
            self.edit_toggle_button.setText("編集モード切替")
            self.edit_toggle_button.setStyleSheet("QPushButton { background-color: #9C27B0; color: white; font-weight: bold; padding: 5px; }")
            self.statusBar().showMessage("結果パネルは読み取り専用になりました。")
    
    def on_text_changed(self):
        """Handle text changes in the results panel."""
        text = self.results_text.toPlainText().strip()
        
        # Enable summarize button if there's text and no worker is running
        if text and not (hasattr(self, 'summarization_worker') and self.summarization_worker and self.summarization_worker.isRunning()):
            self.summarize_button.setEnabled(True)
        else:
            self.summarize_button.setEnabled(False)
        
        # Enable save button if there's text
        self.save_button.setEnabled(bool(text))
    
    def summarize_text(self):
        """Extract answers from result panel text based on Excel structure."""
        text = self.results_text.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "テキストがありません", "結果パネルに抽出対象となるテキストがありません。")
            return
        
        # Get selected document types from checkboxes
        if not hasattr(self, 'doc_type_checkboxes') or not self.doc_type_checkboxes:
            QMessageBox.warning(self, "エラー", "帳票選択機能が初期化されていません。")
            return
        
        selected_doc_types = []
        for doc_type_key, checkbox in self.doc_type_checkboxes.items():
            if checkbox.isChecked():
                selected_doc_types.append(doc_type_key)
        
        if not selected_doc_types:
            QMessageBox.warning(self, "帳票が選択されていません", "出力する帳票を少なくとも1つ選択してください。")
            return
        
        # Use embedded API key
        api_key = OPENAI_API_KEY
        
        # Get selected language
        language_text = self.language_combo.currentText()
        language_code = language_text.split(' ')[0]  # Extract language code
        
        # Disable summarize button and show progress
        self.summarize_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.loading_label.setVisible(True)
        
        # Start summarization in worker thread with language setting and selected document types
        self.summarization_worker = SummarizationWorker(text, api_key, language_code, selected_doc_types)
        self.summarization_worker.progress_updated.connect(self.update_progress)
        self.summarization_worker.status_updated.connect(self.update_status)
        self.summarization_worker.summarization_completed.connect(self.on_summarization_completed)
        self.summarization_worker.summarization_failed.connect(self.on_summarization_failed)
        self.summarization_worker.start()
    
    def classify_text(self):
        """Classify the transcribed text."""
        # Get current content from results panel (user may have edited it)
        current_content = self.results_text.toPlainText()
        if not current_content.strip():
            QMessageBox.warning(self, "内容がありません", "まず文字起こしと回答抽出を実行するか、結果パネルに内容があるか確認してください。")
            return
        
        # Get selected document types from checkboxes
        if not hasattr(self, 'doc_type_checkboxes') or not self.doc_type_checkboxes:
            QMessageBox.warning(self, "エラー", "帳票選択機能が初期化されていません。")
            return
        
        selected_doc_types = []
        for doc_type_key, checkbox in self.doc_type_checkboxes.items():
            if checkbox.isChecked():
                selected_doc_types.append(doc_type_key)
        
        if not selected_doc_types:
            QMessageBox.warning(self, "帳票が選択されていません", "出力する帳票を少なくとも1つ選択してください。")
            return
        
        # Use embedded API key
        api_key = OPENAI_API_KEY
        
        # Disable classification button and show progress
        self.classification_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.loading_label.setVisible(True)
        
        # Prepare session output directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_output_dir = Path("./outputs") / timestamp
        session_output_dir.mkdir(parents=True, exist_ok=True)
        self.current_output_dir = session_output_dir

        # Start classification in worker thread using current content
        self.classification_worker = ClassificationWorker(current_content, api_key, session_output_dir, selected_doc_types)
        self.classification_worker.progress_updated.connect(self.update_progress)
        self.classification_worker.status_updated.connect(self.update_status)
        self.classification_worker.classification_completed.connect(self.on_classification_completed)
        self.classification_worker.classification_failed.connect(self.on_classification_failed)
        self.classification_worker.start()
    
    def on_summarization_completed(self, summary):
        """Handle successful summarization completion."""
        self.summarized_text = summary
        
        # Update the results panel with the summary
        self.results_text.setPlainText(summary)
        
        # Enable classification button
        self.classification_button.setEnabled(True)
        
        # Reset UI
        self.progress_bar.setVisible(False)
        self.loading_label.setVisible(False)
        # Note: summarize_button will be enabled by on_text_changed()
        
        # Show message about editing capability
        QMessageBox.information(self, "抽出完了", 
                               "Excel の構成に基づき回答を抽出しました。\n\n"
                               "分類の前に必要に応じて結果パネルの内容を編集できます。")
    
    def on_summarization_failed(self, error):
        """Handle summarization failure."""
        QMessageBox.critical(self, "要約に失敗しました", f"テキストの要約に失敗しました:\n\n{error}")
        
        # Reset UI
        self.progress_bar.setVisible(False)
        self.loading_label.setVisible(False)
        self.summarize_button.setEnabled(True)
    
    def on_classification_completed(self, results):
        """Handle successful classification completion."""
        classification_results = results.get('classification', {})
        insertion_results = results.get('insertion', {})
        output_dir = results.get('output_dir')
        
        # Show results
        result_text = "分類結果:\n\n"
        for summary_key in SUMMARY_KEYS:
            content = classification_results.get(summary_key)
            if not content:
                continue

            result_text += f"{summary_key}:\n"
            if isinstance(content, dict):
                for key, value in content.items():
                    if isinstance(value, str) and value.strip():
                        result_text += f"  {key}: {value}\n"
            else:
                result_text += f"  {content}\n"
            result_text += "\n"

        result_text += "転記結果:\n"
        for summary_key in SUMMARY_KEYS:
            info = insertion_results.get(summary_key)
            if not info:
                status = "データなし"
            else:
                if info.get("success"):
                    path = info.get("path")
                    status = "✓ 成功"
                    if not info.get("has_data", False):
                        status += " (テンプレートのみ)"
                    if path:
                        status += f" ({Path(path).name})"
                else:
                    status = "✗ 失敗"
                    error = info.get("error")
                    if error:
                        status += f" - {error}"
            result_text += f"  {summary_key}: {status}\n"

        if output_dir:
            result_text += f"\n出力フォルダ: {output_dir}\n"
        
        # Update results panel (preserve edit mode)
        current_edit_mode = not self.results_text.isReadOnly()
        self.results_text.setPlainText(result_text)
        if current_edit_mode:
            self.results_text.setReadOnly(False)
        
        # Reset UI
        self.progress_bar.setVisible(False)
        self.loading_label.setVisible(False)
        self.classification_button.setEnabled(True)
        
        QMessageBox.information(self, "分類が完了しました", "テキストを分類し、Excel ファイルへ転記しました。")
    
    def on_classification_failed(self, error):
        """Handle classification failure."""
        QMessageBox.critical(self, "分類に失敗しました", f"テキストの分類に失敗しました:\n\n{error}")
        
        # Reset UI
        self.progress_bar.setVisible(False)
        self.loading_label.setVisible(False)
        self.classification_button.setEnabled(True)


def main():
    """Main application entry point with error handling."""
    import sys
    import os
    
    # Add current directory to Python path
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    
    try:
        app = QApplication(sys.argv)
        app.setApplicationName("音声テキスト変換ツール")
        
        # Set application style
        app.setStyle('Fusion')
        
        window = AudioToTextGUI()
        window.show()
        
        sys.exit(app.exec())
        
    except ImportError as e:
        print(f"Error importing required modules: {e}")
        print("\nPlease install the required dependencies:")
        print("pip install -r requirements.txt")
        sys.exit(1)
    except Exception as e:
        print(f"Error starting the application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
