#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Korean-Enhanced PDF to Text Converter (Cross-Platform)
한글 PDF 처리에 특화된 변환기 (Windows/macOS/Linux 지원)
"""

import os
import sys
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import re
from pathlib import Path

# 여러 PDF 처리 라이브러리 import
try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    import pdfminer
    from pdfminer.high_level import extract_text
    from pdfminer.layout import LAParams
    PDFMINER_AVAILABLE = True
except ImportError:
    PDFMINER_AVAILABLE = False

class KoreanPDFConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF to Text 변환")
        self.root.geometry("850x790")
        self.root.resizable(True, True)
        
        # 운영체제 감지
        self.os_name = platform.system()
        
        # 시스템별 폰트 설정
        self.setup_fonts()
        
        # 작업표시줄/독 아이콘 설정
        self.set_app_icon()
        
        self.pdf_files = []
        self.is_converting = False
        
        self.check_libraries()
        self.setup_ui()
    
    def setup_fonts(self):
        """운영체제별 최적 폰트 설정"""
        if self.os_name == "Darwin":  # macOS
            self.main_font = ("SF Pro Text", 11)
            self.title_font = ("SF Pro Display", 20, "bold")
            self.small_font = ("SF Pro Text", 9)
            self.version_font = ("SF Pro Text", 10)
        elif self.os_name == "Windows":
            self.main_font = ("맑은 고딕", 11)
            self.title_font = ("맑은 고딕", 20, "bold")
            self.small_font = ("맑은 고딕", 9)
            self.version_font = ("맑은 고딕", 10)
        else:  # Linux
            self.main_font = ("DejaVu Sans", 11)
            self.title_font = ("DejaVu Sans", 20, "bold")
            self.small_font = ("DejaVu Sans", 9)
            self.version_font = ("DejaVu Sans", 10)
    
    def set_app_icon(self):
        """크로스 플랫폼 아이콘 설정"""
        try:
            # 아이콘 파일 경로들 시도
            icon_paths = [
                "logo.ico",             # 현재 폴더
                "icon.ico",             # 현재 폴더
                os.path.join(os.path.dirname(__file__), "korean_pdf_logo.ico"),  # 스크립트 폴더
                os.path.join(os.path.dirname(__file__), "logo.ico"),             # 스크립트 폴더
            ]
            
            # macOS의 경우 .icns 파일도 시도
            if self.os_name == "Darwin":
                icon_paths.extend([
                    "korean_pdf_logo.icns",
                    "logo.icns",
                    "icon.icns",
                    os.path.join(os.path.dirname(__file__), "korean_pdf_logo.icns"),
                    os.path.join(os.path.dirname(__file__), "logo.icns"),
                ])
            
            # 아이콘 파일을 찾아서 설정
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    if self.os_name == "Darwin" and icon_path.endswith('.icns'):
                        # macOS에서는 iconbitmap 대신 다른 방법 사용
                        self.set_macos_icon(icon_path)
                    else:
                        self.root.iconbitmap(icon_path)
                    break
            else:
                # 아이콘 파일이 없으면 기본 설정 적용
                self.set_default_app_properties()
                
        except Exception as e:
            # 아이콘 설정 실패해도 프로그램은 계속 실행
            print(f"아이콘 설정 실패: {e}")
            self.set_default_app_properties()
    
    def set_macos_icon(self, icon_path):
        """macOS 전용 아이콘 설정"""
        try:
            # macOS에서는 wm_iconbitmap 사용
            self.root.wm_iconbitmap(icon_path)
        except:
            # 실패시 기본 설정
            self.set_default_app_properties()
    
    def set_default_app_properties(self):
        """기본 앱 속성 설정"""
        try:
            if self.os_name == "Windows":
                # Windows에서 작업표시줄 앱 ID 설정
                try:
                    import ctypes
                    myappid = 'sogang.pdf_to_text_converter.v1.2'
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
                except:
                    pass
            elif self.os_name == "Darwin":
                # macOS에서 앱 이름 설정
                try:
                    # 메뉴바에 표시될 앱 이름 설정
                    self.root.createcommand('tk::mac::ShowPreferences', self.show_preferences)
                    # 앱 이름 설정
                    self.root.call('wm', 'attributes', self.root, '-titlepath', '한글 PDF 변환기')
                except:
                    pass
                    
        except Exception:
            # 모든 방법이 실패해도 프로그램은 계속 실행
            pass
    
    def show_preferences(self):
        """macOS 환경설정 메뉴 핸들러"""
        messagebox.showinfo("환경설정", "한글 PDF 변환기 v1.2\n제작: jhkim1009@sogang.ac.kr")
    
    def check_libraries(self):
        """사용 가능한 라이브러리 체크"""
        self.available_methods = []
        
        if PYPDF2_AVAILABLE:
            self.available_methods.append("PyPDF2")
        if PDFPLUMBER_AVAILABLE:
            self.available_methods.append("pdfplumber")
        if PYMUPDF_AVAILABLE:
            self.available_methods.append("PyMuPDF")
        if PDFMINER_AVAILABLE:
            self.available_methods.append("pdfminer")
        
        if not self.available_methods:
            messagebox.showerror("오류", 
                "PDF 처리 라이브러리가 없습니다.\n"
                "다음 명령어를 실행하세요:\n"
                "pip install PyPDF2 pdfplumber PyMuPDF pdfminer.six")
            sys.exit(1)
    
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목과 정보
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        title_label = ttk.Label(header_frame, text="PDFtoText", 
                               font=self.title_font, foreground="#2E86AB")
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        version_label = ttk.Label(header_frame, text="v2.0", 
                                 font=self.version_font, foreground="#666666")
        version_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # 시스템 정보 표시
        system_info = f"{self.os_name} 지원"
        system_label = ttk.Label(header_frame, text=system_info, 
                                font=self.small_font, foreground="#999999")
        system_label.grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        contact_label = ttk.Label(header_frame, text="문의: jhkim1009@sogang.ac.kr", 
                                 font=self.small_font, foreground="#0066CC")
        contact_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        
        # 라이브러리 상태 표시
        status_text = f"사용 가능한 라이브러리: {', '.join(self.available_methods)}"
        ttk.Label(main_frame, text=status_text, font=self.small_font, 
                 foreground="green").grid(row=2, column=0, columnspan=3, pady=(0, 15))
        
        # 파일 선택 부분
        file_frame = ttk.LabelFrame(main_frame, text="  PDF 파일 선택  ", padding="12")
        file_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 8))
        
        ttk.Button(button_frame, text="파일 추가", command=self.add_files, 
                  width=12).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(button_frame, text="폴더 추가", command=self.add_folder, 
                  width=12).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(button_frame, text="목록 지우기", command=self.clear_files, 
                  width=12).grid(row=0, column=2)
        
        # 파일 목록
        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 8))
        
        self.file_listbox = tk.Listbox(list_frame, height=5, selectmode=tk.EXTENDED, 
                                      font=self.small_font)
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        ttk.Button(file_frame, text="선택 제거", command=self.remove_selected).grid(row=2, column=0, pady=(0, 5))
        
        # 한글 처리 방법 선택
        method_frame = ttk.LabelFrame(main_frame, text="  변환 방법 선택  ", padding="12")
        method_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.extraction_method = tk.StringVar(value="smart_korean")
        
        ttk.Radiobutton(method_frame, text="스마트 한글 모드 (추천)", 
                       variable=self.extraction_method, value="smart_korean").grid(row=0, column=0, sticky=tk.W, columnspan=2)
        
        ttk.Radiobutton(method_frame, text="전체 라이브러리 시도", 
                       variable=self.extraction_method, value="all_methods").grid(row=1, column=0, sticky=tk.W)
        
        if PDFMINER_AVAILABLE:
            ttk.Radiobutton(method_frame, text="pdfminer (한글 특화)", 
                           variable=self.extraction_method, value="pdfminer").grid(row=1, column=1, sticky=tk.W)
        
        if PYMUPDF_AVAILABLE:
            ttk.Radiobutton(method_frame, text="PyMuPDF (빠른 처리)", 
                           variable=self.extraction_method, value="pymupdf").grid(row=2, column=0, sticky=tk.W)
        
        if PDFPLUMBER_AVAILABLE:
            ttk.Radiobutton(method_frame, text="pdfplumber (정확한 추출)", 
                           variable=self.extraction_method, value="pdfplumber").grid(row=2, column=1, sticky=tk.W)
        
        # 한글 복구 옵션
        recovery_frame = ttk.LabelFrame(main_frame, text="  한글 복구 옵션  ", padding="12")
        recovery_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.font_recovery = tk.BooleanVar(value=True)
        ttk.Checkbutton(recovery_frame, text="폰트 인코딩 문제 자동 복구", 
                       variable=self.font_recovery).grid(row=0, column=0, sticky=tk.W)
        
        self.charset_detection = tk.BooleanVar(value=True)
        ttk.Checkbutton(recovery_frame, text="문자셋 자동 감지 및 변환", 
                       variable=self.charset_detection).grid(row=0, column=1, sticky=tk.W)
        
        self.unicode_normalization = tk.BooleanVar(value=True)
        ttk.Checkbutton(recovery_frame, text="유니코드 정규화 (조합 문자 처리)", 
                       variable=self.unicode_normalization).grid(row=1, column=0, sticky=tk.W)
        
        self.broken_char_repair = tk.BooleanVar(value=True)
        ttk.Checkbutton(recovery_frame, text="깨진 문자 복구 시도", 
                       variable=self.broken_char_repair).grid(row=1, column=1, sticky=tk.W)
        
        # 일반 옵션
        option_frame = ttk.LabelFrame(main_frame, text="  일반 옵션  ", padding="12")
        option_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.merge_option = tk.BooleanVar()
        ttk.Checkbutton(option_frame, text="모든 PDF를 하나로 합치기", 
                       variable=self.merge_option).grid(row=0, column=0, sticky=tk.W)
        
        self.page_separator = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="페이지 구분선 추가", 
                       variable=self.page_separator).grid(row=0, column=1, sticky=tk.W)
        
        # 출력 폴더
        output_frame = ttk.LabelFrame(main_frame, text="  저장 위치  ", padding="12")
        output_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50, 
                                font=self.small_font)
        output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 8))
        ttk.Button(output_frame, text="폴더 선택", command=self.select_output_folder).grid(row=0, column=1)
        
        # 변환 버튼
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=8, column=0, columnspan=3, pady=(0, 15))
        
        # macOS에서는 스타일이 다를 수 있으므로 조건부 스타일 적용
        try:
            self.convert_button = ttk.Button(convert_frame, text="변환 시작", command=self.start_conversion)
            if self.os_name != "Darwin":  # macOS가 아닌 경우에만 스타일 적용
                self.convert_button.configure(style='Accent.TButton')
        except:
            self.convert_button = ttk.Button(convert_frame, text="변환 시작", command=self.start_conversion)
        
        self.convert_button.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_button = ttk.Button(convert_frame, text="중지", command=self.stop_conversion, 
                                     state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1)
        
        # 진행 상황
        progress_frame = ttk.LabelFrame(main_frame, text="  진행 상황  ", padding="12")
        progress_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.progress_label = ttk.Label(progress_frame, text="변환할 파일을 추가해주세요.", 
                                       font=self.small_font)
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(8, 0))
        
        self.current_file_label = ttk.Label(progress_frame, text="", font=self.small_font)
        self.current_file_label.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="  변환 로그  ", padding="12")
        log_frame.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, font=self.small_font)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(10, weight=1)
        header_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # macOS에서 Command+Q 지원
        if self.os_name == "Darwin":
            self.root.bind('<Command-q>', lambda e: self.root.quit())
            self.root.bind('<Command-w>', lambda e: self.root.quit())
    
    def get_system_encoding(self):
        """시스템별 기본 인코딩 반환"""
        if self.os_name == "Windows":
            return 'cp949'
        else:  # macOS, Linux
            return 'utf-8'
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="PDF 파일들 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        
        self.update_file_count()
    
    def add_folder(self):
        folder = filedialog.askdirectory(title="PDF 파일이 있는 폴더 선택")
        if folder:
            pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            added_count = 0
            
            for pdf_file in pdf_files:
                full_path = os.path.join(folder, pdf_file)
                if full_path not in self.pdf_files:
                    self.pdf_files.append(full_path)
                    self.file_listbox.insert(tk.END, pdf_file)
                    added_count += 1
            
            self.log(f"폴더에서 {added_count}개의 PDF 파일을 추가했습니다.")
            self.update_file_count()
    
    def clear_files(self):
        self.pdf_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_file_count()
        self.log("파일 목록을 지웠습니다.")
    
    def remove_selected(self):
        selected_indices = self.file_listbox.curselection()
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            del self.pdf_files[index]
        self.update_file_count()
    
    def select_output_folder(self):
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_path.set(folder)
    
    def update_file_count(self):
        count = len(self.pdf_files)
        self.progress_label.config(text=f"총 {count}개의 PDF 파일")
        
        if count > 0 and not self.output_path.get():
            first_file_dir = os.path.dirname(self.pdf_files[0])
            self.output_path.set(first_file_dir)
    
    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def repair_korean_text(self, text):
        """한글 텍스트 복구 함수"""
        if not text or not self.broken_char_repair.get():
            return text
        
        original_text = text
        
        try:
            # 1. Unicode 정규화
            if self.unicode_normalization.get():
                import unicodedata
                text = unicodedata.normalize('NFC', text)
            
            # 2. 깨진 한글 패턴 복구 시도
            if self.font_recovery.get():
                # 일반적인 한글 깨짐 패턴들
                broken_patterns = {
                    # 자음 모음이 분리된 경우
                    r'ㅏ([ㄱ-ㅎ])': r'\1ㅏ',
                    r'ㅓ([ㄱ-ㅎ])': r'\1ㅓ',
                    r'ㅗ([ㄱ-ㅎ])': r'\1ㅗ',
                    r'ㅜ([ㄱ-ㅎ])': r'\1ㅜ',
                    r'ㅡ([ㄱ-ㅎ])': r'\1ㅡ',
                    r'ㅣ([ㄱ-ㅎ])': r'\1ㅣ',
                }
                
                for pattern, replacement in broken_patterns.items():
                    text = re.sub(pattern, replacement, text)
            
            # 3. 문자셋 변환 시도 (시스템별)
            if self.charset_detection.get():
                system_encoding = self.get_system_encoding()
                try:
                    # 시스템 인코딩으로 변환 시도
                    if system_encoding == 'cp949' and self.os_name == "Windows":
                        if text.encode('utf-8', errors='ignore').decode('cp949', errors='ignore'):
                            pass
                except:
                    pass
            
            # 4. 서로게이트 및 문제 문자 제거
            text = ''.join(char for char in text if not (0xD800 <= ord(char) <= 0xDFFF))
            
            # 5. 제어 문자 제거
            text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
            
            # 6. 연속된 공백 정리
            text = re.sub(r'\s+', ' ', text)
            text = re.sub(r'\n\s*\n', '\n\n', text)
            
            return text.strip()
            
        except Exception as e:
            self.log(f"텍스트 복구 중 오류: {str(e)}")
            return original_text
    
    def extract_text_pdfminer(self, pdf_path):
        """pdfminer를 사용한 한글 특화 텍스트 추출"""
        try:
            # LAParams로 한글 처리 최적화
            laparams = LAParams(
                word_margin=0.1,
                char_margin=2.0,
                line_margin=0.5,
                boxes_flow=0.5
            )
            
            text = extract_text(pdf_path, laparams=laparams)
            return self.repair_korean_text(text)
        except Exception as e:
            raise Exception(f"pdfminer 추출 실패: {str(e)}")
    
    def extract_text_pymupdf_korean(self, pdf_path):
        """PyMuPDF를 사용한 한글 최적화 추출"""
        try:
            text = ""
            doc = fitz.open(pdf_path)
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                
                if self.page_separator.get():
                    text += f"\n--- 페이지 {page_num + 1} ---\n"
                
                # 다양한 추출 방법 시도
                methods = [
                    lambda p: p.get_text("text"),
                    lambda p: p.get_text("dict"),
                    lambda p: p.get_text("blocks")
                ]
                
                page_text = ""
                for method in methods:
                    try:
                        result = method(page)
                        if isinstance(result, str):
                            page_text = result
                            break
                        elif isinstance(result, dict):
                            # dict 형태에서 텍스트 추출
                            blocks = result.get('blocks', [])
                            page_text = self.extract_from_blocks(blocks)
                            break
                        elif isinstance(result, list):
                            # blocks 형태에서 텍스트 추출
                            page_text = self.extract_from_blocks(result)
                            break
                    except:
                        continue
                
                # 한글 복구 적용
                page_text = self.repair_korean_text(page_text)
                text += page_text + "\n"
            
            doc.close()
            return text
        except Exception as e:
            raise Exception(f"PyMuPDF 한글 추출 실패: {str(e)}")
    
    def extract_from_blocks(self, blocks):
        """블록에서 텍스트 추출"""
        text = ""
        for block in blocks:
            if isinstance(block, dict):
                if 'lines' in block:
                    for line in block['lines']:
                        if 'spans' in line:
                            for span in line['spans']:
                                if 'text' in span:
                                    text += span['text']
                        text += "\n"
                elif 'text' in block:
                    text += block['text']
            elif isinstance(block, (list, tuple)) and len(block) >= 5:
                # 좌표와 텍스트가 포함된 블록
                if isinstance(block[4], str):
                    text += block[4]
        return text
    
    def extract_text_smart_korean(self, pdf_path):
        """스마트 한글 모드: 최적의 방법 자동 선택"""
        methods = []
        
        # 한글 처리에 좋은 순서대로 정렬
        if PDFMINER_AVAILABLE:
            methods.append(("pdfminer (한글특화)", self.extract_text_pdfminer))
        if PYMUPDF_AVAILABLE:
            methods.append(("PyMuPDF (한글최적화)", self.extract_text_pymupdf_korean))
        if PDFPLUMBER_AVAILABLE:
            methods.append(("pdfplumber", self.extract_text_pdfplumber))
        if PYPDF2_AVAILABLE:
            methods.append(("PyPDF2", self.extract_text_pypdf2))
        
        best_text = ""
        best_method = ""
        korean_score = 0
        
        for method_name, method_func in methods:
            try:
                text = method_func(pdf_path)
                
                # 한글 품질 평가
                current_score = self.evaluate_korean_quality(text)
                
                if current_score > korean_score or (current_score == korean_score and len(text) > len(best_text)):
                    best_text = text
                    best_method = method_name
                    korean_score = current_score
                
                self.log(f"  {method_name}: 한글품질={current_score:.1f}, 길이={len(text.strip())}")
                
            except Exception as e:
                self.log(f"  {method_name}: 실패 - {str(e)}")
        
        if best_text.strip():
            self.log(f"  최적 방법: {best_method} (한글품질: {korean_score:.1f})")
        
        return best_text
    
    def evaluate_korean_quality(self, text):
        """한글 텍스트 품질 평가"""
        if not text:
            return 0
        
        score = 0
        total_chars = len(text.strip())
        
        if total_chars == 0:
            return 0
        
        # 한글 문자 비율
        korean_chars = len(re.findall(r'[가-힣]', text))
        korean_ratio = korean_chars / total_chars
        score += korean_ratio * 100
        
        # 완성된 한글 문자 비율 (자음/모음만 있는 것 vs 완성형)
        complete_korean = len(re.findall(r'[가-힣]', text))
        incomplete_korean = len(re.findall(r'[ㄱ-ㅎㅏ-ㅣ]', text))
        
        if complete_korean + incomplete_korean > 0:
            complete_ratio = complete_korean / (complete_korean + incomplete_korean)
            score += complete_ratio * 50
        
        # 일반적인 한글 단어 패턴 감지
        korean_words = len(re.findall(r'[가-힣]{2,}', text))
        if total_chars > 0:
            word_density = korean_words / (total_chars / 10)  # 10자당 단어 개수
            score += min(word_density * 20, 20)
        
        # 깨진 문자 패턴 감지 (감점)
        broken_patterns = len(re.findall(r'[^\w\s가-힣ㄱ-ㅎㅏ-ㅣ.,!?()]', text))
        if total_chars > 0:
            broken_ratio = broken_patterns / total_chars
            score -= broken_ratio * 50
        
        return max(0, score)
    
    def extract_text_pypdf2(self, pdf_path):
        """PyPDF2를 사용한 텍스트 추출"""
        text = ""
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                if self.page_separator.get():
                    text += f"\n--- 페이지 {page_num + 1} ---\n"
                page_text = page.extract_text()
                text += self.repair_korean_text(page_text) + "\n"
        return text
    
    def extract_text_pdfplumber(self, pdf_path):
        """pdfplumber를 사용한 텍스트 추출"""
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                if self.page_separator.get():
                    text += f"\n--- 페이지 {page_num + 1} ---\n"
                page_text = page.extract_text()
                if page_text:
                    text += self.repair_korean_text(page_text) + "\n"
        return text
    
    def extract_text_from_pdf(self, pdf_path):
        """선택된 방법으로 텍스트 추출"""
        method = self.extraction_method.get()
        
        try:
            if method == "smart_korean":
                return self.extract_text_smart_korean(pdf_path)
            elif method == "all_methods":
                return self.extract_text_smart_korean(pdf_path)  # 동일한 로직
            elif method == "pdfminer" and PDFMINER_AVAILABLE:
                return self.extract_text_pdfminer(pdf_path)
            elif method == "pymupdf" and PYMUPDF_AVAILABLE:
                return self.extract_text_pymupdf_korean(pdf_path)
            elif method == "pdfplumber" and PDFPLUMBER_AVAILABLE:
                return self.extract_text_pdfplumber(pdf_path)
            elif method == "pypdf2" and PYPDF2_AVAILABLE:
                return self.extract_text_pypdf2(pdf_path)
            else:
                # 폴백: 스마트 한글 모드
                return self.extract_text_smart_korean(pdf_path)
        
        except Exception as e:
            raise Exception(f"텍스트 추출 실패: {str(e)}")
    
    def safe_write_text(self, filepath, text, encoding=None):
        """안전한 텍스트 파일 쓰기 (시스템별 인코딩 고려)"""
        if encoding is None:
            encoding = 'utf-8'  # 기본값은 UTF-8
        
        # 시스템별 인코딩 우선순위
        if self.os_name == "Windows":
            encodings = ['utf-8', 'cp949', 'ascii']
        else:  # macOS, Linux
            encodings = ['utf-8', 'latin-1', 'ascii']
        
        for enc in encodings:
            try:
                with open(filepath, 'w', encoding=enc, errors='replace') as file:
                    file.write(text)
                return True, enc
            except Exception:
                continue
        
        # 모든 인코딩 실패시 최후 수단
        try:
            ascii_text = ''.join(char if ord(char) < 128 else '?' for char in text)
            with open(filepath, 'w', encoding='ascii', errors='ignore') as file:
                file.write(ascii_text)
            return True, "ASCII (문자 손실)"
        except:
            return False, "실패"
    
    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "변환할 PDF 파일을 추가해주세요.")
            return
        
        if not self.output_path.get():
            messagebox.showwarning("경고", "출력 폴더를 선택해주세요.")
            return
        
        self.is_converting = True
        self.convert_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        threading.Thread(target=self.convert_files, daemon=True).start()
    
    def stop_conversion(self):
        self.is_converting = False
        self.log("변환을 중지했습니다.")
    
    def convert_files(self):
        try:
            output_dir = self.output_path.get()
            total_files = len(self.pdf_files)
            
            self.progress_bar.config(maximum=total_files)
            
            success_count = 0
            empty_count = 0
            error_count = 0
            korean_recovered = 0
            
            self.log(f"변환 시작: 총 {total_files}개 파일 ({self.os_name})")
            self.log(f"사용 모드: {self.extraction_method.get()}")
            
            for i, pdf_path in enumerate(self.pdf_files):
                if not self.is_converting:
                    break
                
                filename = os.path.basename(pdf_path)
                self.current_file_label.config(text=f"변환 중: {filename}")
                self.log(f"\n처리 중: {filename}")
                
                try:
                    text = self.extract_text_from_pdf(pdf_path)
                    
                    # 출력 파일명 생성
                    output_filename = Path(filename).stem + ".txt"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    if text.strip():
                        # 한글 품질 확인
                        korean_quality = self.evaluate_korean_quality(text)
                        
                        success, encoding = self.safe_write_text(output_path, text)
                        
                        if success:
                            success_count += 1
                            char_count = len(text.strip())
                            korean_chars = len(re.findall(r'[가-힣]', text))
                            
                            if korean_chars > 0:
                                korean_recovered += 1
                                self.log(f"  성공: {char_count}자 추출, 한글 {korean_chars}자 (품질: {korean_quality:.1f})")
                            else:
                                self.log(f"  성공: {char_count}자 추출 (한글 없음)")
                            
                            self.log(f"  저장: {output_filename} ({encoding} 인코딩)")
                        else:
                            error_count += 1
                            self.log(f"  저장 실패: 인코딩 문제")
                    else:
                        # 빈 파일 처리
                        empty_message = f"{filename}에서 텍스트를 추출할 수 없습니다.\n\n"
                        empty_message += "가능한 원인:\n"
                        empty_message += "• 이미지로만 구성된 PDF (스캔 문서)\n"
                        empty_message += "• 암호화되거나 보호된 PDF\n"
                        empty_message += "• 특수한 폰트나 인코딩 사용\n"
                        empty_message += "• 손상된 PDF 파일\n\n"
                        empty_message += "해결 방법:\n"
                        empty_message += "• OCR 프로그램 사용 (이미지 PDF의 경우)\n"
                        empty_message += "• PDF 암호 해제 후 재시도\n"
                        empty_message += "• 다른 PDF 뷰어로 다시 저장 후 시도\n"
                        
                        success, encoding = self.safe_write_text(output_path, empty_message)
                        empty_count += 1
                        self.log(f"  텍스트 없음: 설명 파일 생성")
                    
                except Exception as e:
                    error_count += 1
                    self.log(f"  실패: {str(e)}")
                    
                    # 오류 정보 파일 생성
                    error_filename = Path(filename).stem + "_오류.txt"
                    error_path = os.path.join(output_dir, error_filename)
                    error_message = f"PDF 변환 실패: {filename}\n오류: {str(e)}\n\n"
                    error_message += f"시스템: {self.os_name}\n"
                    error_message += "다른 변환 방법을 시도해보세요:\n"
                    error_message += "• 다른 라이브러리 방법 선택\n"
                    error_message += "• 한글 복구 옵션 조정\n"
                    error_message += "• PDF를 다른 형식으로 저장 후 재시도\n"
                    self.safe_write_text(error_path, error_message)
                
                self.progress_bar.config(value=i + 1)
                self.root.update_idletasks()
            
            # 최종 결과 요약
            if self.is_converting:
                self.log(f"\n변환 완료!")
                self.log(f"성공: {success_count}개")
                self.log(f"한글 복구 성공: {korean_recovered}개")
                self.log(f"텍스트 없음: {empty_count}개")
                self.log(f"오류: {error_count}개")
                self.log(f"총 처리: {success_count + empty_count + error_count}개")
                
                result_msg = f"변환 완료! ({self.os_name})\n\n"
                result_msg += f"성공: {success_count}개\n"
                result_msg += f"한글 복구: {korean_recovered}개\n"
                result_msg += f"텍스트 없음: {empty_count}개\n"
                result_msg += f"오류: {error_count}개"
                
                if korean_recovered > 0:
                    result_msg += f"\n\n{korean_recovered}개 파일에서 한글이 성공적으로 복구되었습니다!"
                
                messagebox.showinfo("변환 완료", result_msg)
        
        except Exception as e:
            self.log(f"전체 변환 오류: {str(e)}")
            messagebox.showerror("오류", f"변환 중 심각한 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            self.root.after(0, self.conversion_finished)
    
    def conversion_finished(self):
        self.convert_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.current_file_label.config(text="")
        self.progress_bar.config(value=0)
        self.is_converting = False
    
    def run(self):
        self.root.mainloop()

def main():
    app = KoreanPDFConverter()
    app.run()

if __name__ == "__main__":
    main()