#!/usr/bin/env python3
"""
Audio to Text Conversion Module
Handles audio file conversion and speech recognition using Whisper.
"""

import os
import time
import math
import uuid
from pydub import AudioSegment
from PyQt6.QtCore import QThread, pyqtSignal

class ConversionWorker(QThread):
    """Worker thread for audio conversion to prevent GUI freezing using Whisper."""
    
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    partial_result_updated = pyqtSignal(str)  # partial text result
    conversion_completed = pyqtSignal(str, str)  # file_path, text
    conversion_failed = pyqtSignal(str, str)  # file_path, error
    
    def __init__(self, file_path, language='en-US', model=None):
        super().__init__()
        self.file_path = file_path
        self.language = self._convert_language_code(language)
        self.model = model
        self.supported_formats = ['.wav', '.mp3', '.m4a', '.flac', '.aiff', '.ogg']
    
    def _convert_language_code(self, lang_code):
        """Convert language code from format like 'ja-JP' to 'ja' for Whisper."""
        # Extract the base language code (before the hyphen)
        if '-' in lang_code:
            return lang_code.split('-')[0]
        return lang_code
    
    def run(self):
        try:
            self.status_updated.emit("Starting conversion...")
            self.progress_updated.emit(0)
            
            # Check if file exists
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"Audio file not found: {self.file_path}")
            
            # Check if model is provided
            if self.model is None:
                raise Exception("Whisper model not provided. Please ensure model is loaded before conversion.")
           
            # Load audio file
            self.status_updated.emit("🎧 Preparing audio...")
            self.progress_updated.emit(0)
            
            audio_path = os.path.abspath(self.file_path)
            
            # Load and get length using pydub
            audio = AudioSegment.from_file(audio_path)
            duration_ms = len(audio)
            # Get chunk length from .env, defaulting to 60 seconds if not set
            chunk_length_sec = int(os.getenv("AUDIO_CHUNK_LENGTH_SECONDS", "60"))
            chunk_length_ms = chunk_length_sec * 1000  # Convert seconds to milliseconds
            num_chunks = math.ceil(duration_ms / chunk_length_ms)
            
            duration_minutes = duration_ms / 60000
            self.status_updated.emit(f"🕒 Total duration: {duration_minutes:.1f} minutes")
            self.partial_result_updated.emit(f"📦 Splitting into {num_chunks} chunks...")
            self.progress_updated.emit(0)
            
            # Process audio in chunks
            paragraphs = []
            start_time = time.time()
            
            # Use unique temp file prefix to avoid conflicts
            temp_prefix = f"temp_chunk_{uuid.uuid4().hex[:8]}_"
            
            for i in range(num_chunks):
                start_ms = i * chunk_length_ms
                end_ms = min((i + 1) * chunk_length_ms, duration_ms)
                
                # Extract chunk
                chunk = audio[start_ms:end_ms]
                chunk_path = f"{temp_prefix}{i}.wav"
                
                # Export chunk as WAV
                chunk.export(chunk_path, format="wav")
                
                # Transcribe chunk
                self.status_updated.emit(f"📝 Transcribing chunk {i + 1}/{num_chunks}...")
                result = self.model.transcribe(chunk_path, language=self.language, verbose=False)
                text = result["text"].strip()
                paragraphs.append(text)
                
                # Update partial results progressively
                self.partial_result_updated.emit("📝 " + "\n\n".join(paragraphs))
                
                # Calculate and update progress
                percent = int(((i + 1) / num_chunks) * 100)
                elapsed = time.time() - start_time
                
                self.status_updated.emit(f"📈 Extracted {percent}% text from audio... ⏱️ {elapsed:.1f}s")
                
                progress_value = int((percent / 100) * 100)
                self.progress_updated.emit(progress_value)
                
                # Small delay to ensure UI updates
                time.sleep(0.05)
                
                # Remove temp file
                if os.path.exists(chunk_path):
                    try:
                        os.remove(chunk_path)
                    except Exception:
                        pass
            
            self.status_updated.emit("📈 Extracted 100% text from audio ✅")
            self.progress_updated.emit(100)
            
            # Merge paragraphs
            formatted_text = "\n\n".join(paragraphs)
            
            self.conversion_completed.emit(self.file_path, formatted_text)
            
        except Exception as e:
            self.conversion_failed.emit(self.file_path, str(e))
    

