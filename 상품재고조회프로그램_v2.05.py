"""
상품 재고 조회 프로그램 v2.05
- CustomTkinter UI
- 데이터 관리 클래스 분리
- 비동기 데이터 로딩 + 로딩 오버레이
- 속도 최적화
"""
import os
import warnings
import threading
import tkinter as tk
import customtkinter as ctk
import tkinter.messagebox as msgbox
import pandas as pd
from tksheet import Sheet

# 경고 무시 설정
warnings.simplefilter("ignore")

# CustomTkinter 기본 설정
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

# -----------------------------------------------------------------------------
# 상수 및 설정
# -----------------------------------------------------------------------------
FILE_NAME = "1-재고장.xlsm"
DIR_PATHS = [
    r"\\NAS451\team451\01-재고 리스트 남 여",
    r"D:\nas451\DB"
]
SHEET_STOCK = "재고리스트"
SHEET_INFO = "상품정보"

COL_LIST = [
    "상품코드", "옵션번호", "컬러", "사이즈", "OPT", "옵션코드", "현재고", "출고", "입고",
    "수정재고", "상품분류", "제조사", "제조사\n상품코드", "진마니아\n로켓등록", "진마니아\n일반등록",
    "아크시\n로켓등록", "아크시\n일반등록", "종합몰\n등록코드"
]
SEARCH_OPT_LIST = [
    "상품코드로 조회",
    "진마니아 일반등록",
    "아크시 로켓등록",
    "아크시 일반등록",
    "종합몰 등록코드"
]

COL_MAP = {
    "진마니아 일반등록": "진마니아\n일반등록",
    "아크시 로켓등록": "아크시\n로켓등록",
    "아크시 일반등록": "아크시\n일반등록",
    "종합몰 등록코드": "종합몰\n등록코드"
}


# -----------------------------------------------------------------------------
# 유틸리티 함수
# -----------------------------------------------------------------------------
def get_file_path() -> str | None:
    """
    재고 파일 경로 검색 및 반환 함수
    
    설명:
        - DIR_PATHS에 정의된 경로들을 순회하며 FILE_NAME 파일을 찾습니다
        - 파일이 실제로 존재하면 해당 경로를 반환합니다
        - 파일이 없으면 디렉토리만 존재하는 경로를 반환합니다
        - 모든 경로에서 파일을 찾지 못하면 에러 메시지를 표시합니다
    
    반환값:
        str | None: 파일 전체 경로 또는 None
    """
    # 1단계: 실제 파일이 존재하는지 확인
    for dir_path in DIR_PATHS:
        full_path = os.path.join(dir_path, FILE_NAME)
        if os.path.isfile(full_path):
            return full_path
    
    # 2단계: 디렉토리만 존재하는 경우 해당 경로 반환
    for dir_path in DIR_PATHS:
        if os.path.isdir(dir_path):
            return os.path.join(dir_path, FILE_NAME)

    # 3단계: 파일을 찾지 못한 경우 에러 메시지 표시
    msgbox.showerror(
        '에러', 
        '상품정보 파일을 불러올 수 없습니다!\n'
        'NAS451서버에 /team451/01-재고 리스트 남 여 폴더에 1-재고장.xlsm 파일이 존재하는지 확인해 주세요!'
    )
    return None


def convert_to_standard_code(in_code: str) -> str:
    """
    상품 코드를 표준 품번 형식으로 변환하는 함수
    
    설명:
        - 다양한 형식의 상품 코드를 표준화된 품번 형식으로 변환합니다
        - 입력 코드는 대문자로 변환되고 공백이 제거됩니다
        - W, M, T, F, BHP, GS, DS, CS, PAC로 시작하는 코드는 그대로 반환
        - C, D, E, H, R, Z로 시작하는 코드는 패턴에 따라 변환
    
    매개변수:
        in_code (str): 변환할 원본 상품 코드
    
    반환값:
        str: 변환된 표준 품번 또는 "code error"
    """
    # 입력 코드 전처리: 대문자 변환 및 공백 제거
    in_code = in_code.upper().strip()
    if not in_code:
        return "code error"

    # 이미 표준 형식인 코드는 그대로 반환
    if in_code.startswith(("W", "M", "T", "F", "BHP", "GS", "DS", "CS", "PAC")):
        return in_code

    # C, D, E, H, R, Z로 시작하는 코드 변환 로직
    if in_code.startswith(("C", "D", "E", "H", "R", "Z")) and len(in_code) >= 6:
        try:
            # 코드의 3번째와 4번째 문자 추출
            prefix_char = in_code[3] if len(in_code) > 3 else ""
            target_char = in_code[4] if len(in_code) > 4 else ""
            
            # 패턴 1: 3번째 문자가 'T'인 경우
            if prefix_char == "T":
                good_num = in_code[2] + in_code[4] + in_code[5]
                return f"{in_code[3]}{in_code[1]}-{good_num}"
            # 패턴 2: 3번째 문자가 'W' 또는 'M'인 경우
            elif prefix_char in ("W", "M"):
                good_num = in_code[2] + in_code[4] + in_code[5]
                # 1번째 문자를 매핑 테이블에 따라 변환
                prefix_map = {
                    "S": "JS", "U": "JU", "A": "BA", "C": "BC",
                    "L": "JL", "M": "JM", "N": "NA", "E": "EJ"
                }
                mapped_prefix = prefix_map.get(in_code[1], 'XX')
                return f"{in_code[3]}{mapped_prefix}-{good_num}"
            # 패턴 3: 4번째 문자가 'T'인 경우
            elif target_char == "T":
                good_num = in_code[3] + in_code[5] + in_code[6]
                return f"{in_code[4]}{in_code[1]}{in_code[2]}-{good_num}"
            # 패턴 4: 4번째 문자가 'W' 또는 'M'인 경우
            elif target_char in ("W", "M"):
                good_num = in_code[3] + in_code[5] + in_code[6]
                return f"{in_code[4]}BP-{good_num}"
        except IndexError:
            return "code error"

    return "code error"


# -----------------------------------------------------------------------------
# 데이터 관리 클래스
# -----------------------------------------------------------------------------
class InventoryDataManager:
    """
    재고 데이터 관리 클래스
    
    설명:
        - 엑셀 파일에서 재고 데이터를 로딩하고 처리합니다
        - 데이터 검색 기능을 제공합니다
        - 비동기 로딩을 지원합니다
    
    속성:
        df (pd.DataFrame): 재고 데이터를 저장하는 DataFrame
        is_loaded (bool): 데이터 로딩 완료 여부
    """
    
    def __init__(self):
        """데이터 관리자 초기화"""
        # 빈 DataFrame 생성 (컬럼 구조만 정의)
        self.df = pd.DataFrame(columns=COL_LIST)
        # 데이터 로딩 완료 플래그
        self.is_loaded = False
    
    def load_data(self, file_path: str, callback=None, progress_callback=None):
        """
        엑셀 파일에서 재고 데이터를 로드하는 함수
        
        설명:
            - 엑셀 파일에서 재고 시트와 상품정보 시트를 읽습니다
            - 두 시트의 데이터를 병합하여 최종 DataFrame을 생성합니다
            - 진행 상황을 콜백으로 알립니다
        
        매개변수:
            file_path (str): 엑셀 파일 경로
            callback (function): 완료 시 호출할 콜백 함수 (success, message)
            progress_callback (function): 진행 상황 업데이트 콜백 함수
        """
        try:
            # 진행 상황 업데이트: 엑셀 로딩 시작
            if progress_callback:
                progress_callback("엑셀 데이터 로딩 중...")
            
            # 재고 시트와 상품정보 시트 읽기 (헤더는 4번째 행)
            stock_raw_df = pd.read_excel(file_path, sheet_name=SHEET_STOCK, header=3).fillna("")
            info_raw_df = pd.read_excel(file_path, sheet_name=SHEET_INFO, header=3).fillna("")
            
            # 진행 상황 업데이트: 데이터 변환 시작
            if progress_callback:
                progress_callback("데이터 변환 및 병합 중...")
            
            # 원시 데이터를 가공하여 최종 DataFrame 생성
            self.df = self._process_data(stock_raw_df, info_raw_df)
            self.is_loaded = True
            
            # 완료 콜백 호출
            if callback:
                callback(True, f"데이터 로드 완료: 총 {len(self.df)}건")
                
        except Exception as e:
            # 에러 발생 시 처리
            self.is_loaded = False
            if callback:
                callback(False, f"데이터 로드 실패: {e}")
    
    def _process_data(self, stock_df: pd.DataFrame, info_df: pd.DataFrame) -> pd.DataFrame:
        """
        원시 데이터를 가공하여 최종 DataFrame을 생성하는 내부 함수
        
        설명:
            - 재고 시트와 상품정보 시트의 데이터를 행 단위로 처리합니다
            - 각 상품의 옵션별(사이즈별) 재고 정보를 개별 행으로 분리합니다
            - 최대 8개의 옵션을 지원합니다
        
        매개변수:
            stock_df (pd.DataFrame): 재고 시트 원시 데이터
            info_df (pd.DataFrame): 상품정보 시트 원시 데이터
        
        반환값:
            pd.DataFrame: 가공된 최종 데이터
        """
        processed_data = []
        row_count = stock_df.shape[0]
        
        # 각 행을 순회하며 데이터 추출
        for r in range(row_count):
            # 품번 추출 (8번째 컬럼)
            pumbun = stock_df.iloc[r, 8]
            # 헤더 행은 건너뛰기
            if pumbun == "품번":
                continue
            
            # 재고 시트에서 필요한 정보 추출
            color = stock_df.iloc[r, 9]          # 컬러
            part = stock_df.iloc[r, 25]          # 상품분류
            jm_rocket = stock_df.iloc[r, 3]      # 진마니아 로켓등록
            jm_normal = stock_df.iloc[r, 4]      # 진마니아 일반등록
            ac_rocket = stock_df.iloc[r, 5]      # 아크시 로켓등록
            ac_normal = stock_df.iloc[r, 6]      # 아크시 일반등록
            op_code = stock_df.iloc[r, 7]        # 종합몰 등록코드
            
            # 상품정보 시트에서 제조사 정보 추출
            maker = info_df.iloc[r, 21]          # 제조사
            maker_code = info_df.iloc[r, 22]     # 제조사 상품코드

            # 각 옵션(사이즈)별로 재고 정보 처리 (최대 8개 옵션)
            for offset in range(8):
                col_idx = 10 + offset  # 재고 값이 있는 컬럼 인덱스
                stock_val = stock_df.iloc[r, col_idx]
                
                # 재고 값이 있는 경우에만 처리
                if stock_val != "":
                    size_val = info_df.iloc[r, col_idx - 6]  # 해당 사이즈 정보
                    opt_num = offset + 1  # 옵션 번호 (1~8)
                    
                    # 옵션번호와 옵션코드 문자열 생성
                    if color != "":
                        # 컬러가 있는 경우: 품번(컬러)_옵션번호
                        opt_no_str = f"{pumbun}({color})_{opt_num}"
                        opt_code_str = f"{pumbun}({color}) : {size_val}"
                    else:
                        # 컬러가 없는 경우: 품번_옵션번호
                        opt_no_str = f"{pumbun}_{opt_num}"
                        opt_code_str = f"{pumbun} : {size_val}"

                    # 현재고 값을 정수로 변환 (변환 실패 시 0)
                    try:
                        current_stock = int(stock_val)
                    except ValueError:
                        current_stock = 0

                    # 행 데이터 생성 (COL_LIST 순서에 맞춰)
                    row_data = [
                        pumbun, opt_no_str, color, size_val, opt_num, opt_code_str,
                        current_stock, "", "", "", part, maker, maker_code,
                        jm_rocket, jm_normal, ac_rocket, ac_normal, op_code
                    ]
                    processed_data.append(row_data)
        
        # 가공된 데이터로 DataFrame 생성하여 반환
        return pd.DataFrame(processed_data, columns=COL_LIST)
    
    def search(self, search_type: str, codes: list[str]) -> pd.DataFrame:
        """
        검색 조건에 따라 데이터를 필터링하는 함수
        
        설명:
            - 검색어 리스트를 받아 해당하는 데이터를 필터링합니다
            - 검색 타입에 따라 다른 컬럼을 검색합니다
            - 정규식 특수문자를 이스케이프 처리합니다
        
        매개변수:
            search_type (str): 검색 타입 (상품코드, 진마니아, 아크시 등)
            codes (list[str]): 검색할 코드 리스트
        
        반환값:
            pd.DataFrame: 필터링된 데이터
        """
        # 검색어가 없으면 전체 데이터 반환
        if not codes:
            return self.df
        
        # 검색 패턴 생성: 정규식 특수문자 이스케이프 처리 후 OR 조건으로 결합
        pattern = '|'.join([
            pd.Series([c]).str.replace(r'([.\\+*?^$[\]{}()|])', r'\\\1', regex=True).iloc[0] 
            for c in codes
        ])
        
        # 검색 타입에 따라 해당 컬럼에서 검색
        if search_type == "상품코드로 조회":
            return self.df[self.df['상품코드'].astype(str).str.contains(pattern, case=False, na=False)].fillna("")
        else:
            # COL_MAP에서 실제 컬럼명 찾기
            target_col = COL_MAP.get(search_type)
            if target_col:
                return self.df[self.df[target_col].astype(str).str.contains(pattern, case=False, na=False)].fillna("")
        
        # 매칭되는 검색 타입이 없으면 전체 데이터 반환
        return self.df


# -----------------------------------------------------------------------------
# 로딩 오버레이
# -----------------------------------------------------------------------------
class LoadingOverlay(ctk.CTkFrame):
    """
    로딩 오버레이 클래스
    
    설명:
        - 데이터 로딩 중 화면 전체를 덮는 오버레이를 표시합니다
        - 진행 상황 메시지를 업데이트할 수 있습니다
        - 무한 진행 바(spinner)를 표시합니다
    """
    
    def __init__(self, parent):
        """
        로딩 오버레이 초기화
        
        매개변수:
            parent: 부모 위젯
        """
        super().__init__(parent, fg_color=("gray90", "gray20"))
        
        # 중앙 프레임 생성
        self.inner_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.inner_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # 무한 진행 바 (spinner) 생성
        self.spinner = ctk.CTkProgressBar(self.inner_frame, mode="indeterminate", width=200)
        self.spinner.pack(pady=(0, 10))
        self.spinner.start()
        
        # 로딩 메시지 라벨 생성
        self.label = ctk.CTkLabel(
            self.inner_frame, 
            text="데이터를 불러오는 중...",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.label.pack()
    
    def set_message(self, message: str):
        """로딩 메시지 업데이트"""
        self.label.configure(text=message)
    
    def show(self):
        """오버레이 표시"""
        # 화면 전체를 덮도록 배치
        self.place(relx=0, rely=0, relwidth=1, relheight=1)
        # 최상위로 올리기
        self.lift()
    
    def hide(self):
        """오버레이 숨기기"""
        # 진행 바 중지
        self.spinner.stop()
        # 화면에서 제거
        self.place_forget()


# -----------------------------------------------------------------------------
# 메인 애플리케이션 클래스
# -----------------------------------------------------------------------------
class InventoryApp(ctk.CTk):
    """
    재고 조회 프로그램 메인 애플리케이션 클래스
    
    설명:
        - CustomTkinter 기반의 메인 윈도우입니다
        - UI 구성 요소와 데이터 관리자를 초기화합니다
        - 프로그램 시작 시 자동으로 데이터를 로드합니다
    """
    
    def __init__(self):
        """애플리케이션 초기화"""
        super().__init__()
        # 윈도우 기본 설정
        self._init_window()
        # UI 컴포넌트 생성
        self._init_ui()
        
        # 데이터 관리자 및 로딩 오버레이 초기화
        self.data_manager = InventoryDataManager()
        self.loading_overlay = LoadingOverlay(self)
        
        # 빈 시트 렌더링
        self.sheet = None
        self._render_sheet(pd.DataFrame(columns=COL_LIST))
        
        # 100ms 후 데이터 로딩 시작 (UI 초기화 후)
        self.after(100, self._load_data_async)
    
    def _init_window(self):
        """
        윈도우 초기 설정 함수
        
        설명:
            - 윈도우 제목, 크기, 위치를 설정합니다
            - 최소 크기를 설정하고 최대화 상태로 시작합니다
        """
        self.title("상품 재고 조회 프로그램 ver2.04")
        self.geometry("1600x800+10+10")  # 너비x높이+X위치+Y위치
        self.state("zoomed")  # 최대화 상태로 시작
        self.minsize(1600, 800)  # 최소 크기 설정
        self.resizable(True, True)  # 크기 조절 비활성화

    def _init_ui(self):
        """
        UI 컴포넌트 초기화 함수
        
        설명:
            - 상단 툴바, 좌측 입력 영역, 우측 결과 영역, 하단 상태바를 생성합니다
            - 각 영역의 레이아웃을 설정합니다
        """
        # 상단 툴바
        self.toolbar = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.toolbar.pack(fill="x", padx=10, pady=(10, 5))
        self.toolbar.pack_propagate(False)
        
        self._create_toolbar_widgets()
        
        # 메인 컨테이너
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 좌측: 코드 입력 영역
        self.input_frame = ctk.CTkFrame(self.main_container, width=200)
        self.input_frame.pack(side="left", fill="y", padx=(0, 5))
        self.input_frame.pack_propagate(False)
        
        ctk.CTkLabel(
            self.input_frame, 
            text="상품코드 입력",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(pady=(10, 5))
        
        self.txt_input = ctk.CTkTextbox(self.input_frame, width=190, font=ctk.CTkFont(size=12))
        self.txt_input.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        # 우측: 결과 영역
        self.result_frame = ctk.CTkFrame(self.main_container)
        self.result_frame.pack(side="right", fill="both", expand=True)
        
        # 하단 상태바
        self.statusbar = ctk.CTkFrame(self, height=30, corner_radius=0)
        self.statusbar.pack(fill="x", padx=10, pady=(5, 10))
        self.statusbar.pack_propagate(False)
        
        self.lbl_status = ctk.CTkLabel(
            self.statusbar, 
            text="대기 중...",
            font=ctk.CTkFont(size=11),
            anchor="w"
        )
        self.lbl_status.pack(side="left", fill="x", expand=True, padx=10)
    
    def _create_toolbar_widgets(self):
        """
        툴바 위젯 생성 함수
        
        설명:
            - 검색 옵션 콤보박스, 검색 버튼, 설명 라벨, 초기화/닫기 버튼을 생성합니다
            - 각 위젯의 이벤트 핸들러를 연결합니다
        """
        # 검색 옵션 콤보박스
        self.combo_search_type = ctk.CTkComboBox(
            self.toolbar,
            values=SEARCH_OPT_LIST,
            width=180,
            state="readonly",
            font=ctk.CTkFont(size=12)
        )
        self.combo_search_type.set(SEARCH_OPT_LIST[0])
        self.combo_search_type.pack(side="left", padx=(10, 5), pady=10)
        
        # 검색 버튼
        self.btn_search = ctk.CTkButton(
            self.toolbar, 
            text="🔍 정보조회",
            width=100,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._on_search
        )
        self.btn_search.pack(side="left", padx=5, pady=10)
        
        # 설명 라벨
        desc_text = "1. 검색할 데이터를 선택하고 상품코드 또는 옵션코드를 입력해 주세요.\n2. 입력된 코드는 품번으로 자동변환 됩니다"
        ctk.CTkLabel(
            self.toolbar, 
            text=desc_text,
            font=ctk.CTkFont(size=11),
            justify="left"
        ).pack(side="left", padx=15, pady=10)
        
        # 닫기 버튼
        ctk.CTkButton(
            self.toolbar, 
            text="✕ 닫기",
            width=80,
            fg_color="gray50",
            hover_color="gray40",
            command=self.quit
        ).pack(side="right", padx=5, pady=10)
        
        # 초기화 버튼
        ctk.CTkButton(
            self.toolbar, 
            text="↺ 초기화",
            width=80,
            fg_color="gray60",
            hover_color="gray50",
            command=self._on_reset
        ).pack(side="right", padx=5, pady=10)
    
    def _update_status(self, message: str):
        """
        상태바 메시지 업데이트 함수
        
        매개변수:
            message (str): 표시할 상태 메시지
        """
        self.lbl_status.configure(text=message)
    
    def _load_data_async(self):
        """
        비동기 데이터 로딩 함수
        
        설명:
            - 별도의 스레드에서 데이터를 로드하여 UI 블록킹을 방지합니다
            - 로딩 중 오버레이를 표시하고 진행 상황을 업데이트합니다
            - 로딩 완료 후 결과를 화면에 표시합니다
        """
        file_path = get_file_path()
        if not file_path or not os.path.exists(file_path):
            if file_path and not os.path.exists(file_path):
                msgbox.showerror('에러', f'파일을 찾을 수 없습니다: {file_path}')
            return
        
        self.loading_overlay.show()
        
        def progress_update(msg):
            self.after(0, lambda: self.loading_overlay.set_message(msg))
        
        def on_complete(success, message):
            def update_ui():
                self.loading_overlay.hide()
                self._update_status(message)
                if success:
                    self._render_sheet(self.data_manager.df)
            self.after(0, update_ui)
        
        thread = threading.Thread(
            target=self.data_manager.load_data,
            args=(file_path, on_complete, progress_update),
            daemon=True
        )
        thread.start()
    
    def _render_sheet(self, dataframe: pd.DataFrame):
        """
        tksheet를 사용하여 데이터를 표시하는 함수
        
        설명:
            - 기존 시트를 제거하고 새로운 시트를 생성합니다
            - DataFrame 데이터를 Sheet 위젯에 표시합니다
            - 컬럼 너비, 폰트, 정렬 등을 설정합니다
        
        매개변수:
            dataframe (pd.DataFrame): 표시할 데이터
        """
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        
        if dataframe is None or dataframe.empty:
            self.sheet = Sheet(
                self.result_frame,
                data=[],
                headers=list(COL_LIST),
                header_height=36,
                header_fg="#FFFFFF",
                header_bg="#333333"
            )
        else:
            self.sheet = Sheet(
                self.result_frame,
                data=dataframe.values.tolist(),
                headers=list(dataframe.columns),
                header_height=36,
                header_fg="#FFFFFF",
                header_bg="#333333"
            )
        
        self.sheet.header_font(('NanumGothic', 10, 'normal'))
        self.sheet.font(('NanumGothic', 9, 'normal'))
        self.sheet.table_align(align="center")
        self.sheet["A"].align("c")
        
        self.sheet.set_all_column_widths(width=None, only_set_if_too_small=False, redraw=False, recreate_selection_boxes=False)
        self.sheet.column_width(1, 120)
        self.sheet.column_width(7, 50)
        self.sheet.column_width(8, 50)
        self.sheet.column_width(9, 50)
        
        self.sheet.enable_bindings()
        self.sheet.pack(fill="both", expand=True, padx=5, pady=5)
    
    def _on_search(self):
        """
        검색 실행 함수
        
        설명:
            - 입력된 검색어를 가져와 데이터를 필터링합니다
            - 상품코드 검색 시 자동으로 표준 품번으로 변환합니다
            - 검색 결과를 화면에 표시하고 상태바를 업데이트합니다
        """
        raw_input = self.txt_input.get("1.0", "end").strip()
        if not raw_input:
            msgbox.showinfo("알림", "검색어를 입력해 주세요!")
            return
        
        search_type = self.combo_search_type.get()
        search_codes = []
        
        if search_type == "상품코드로 조회":
            raw_codes = raw_input.split()
            for code in raw_codes:
                conv = convert_to_standard_code(code)
                search_codes.append(conv if conv != "code error" else code)
        else:
            search_codes = raw_input.split()
        
        filtered_df = self.data_manager.search(search_type, search_codes)
        self._render_sheet(filtered_df)
        self._update_status(f"검색 결과: {len(filtered_df)}건 (검색어: {len(search_codes)}개)")
    
    def _on_reset(self):
        """
        초기화 함수
        
        설명:
            - 입력된 검색어를 지웁니다
            - 전체 데이터를 다시 표시합니다
            - 상태바를 업데이트합니다
        """
        # 텍스트 입력 영역 초기화
        self.txt_input.delete("1.0", "end")
        # 전체 데이터 다시 표시
        self._render_sheet(self.data_manager.df)
        # 상태바 업데이트
        self._update_status(f"초기화 완료 (전체 {len(self.data_manager.df)}건)")


if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()
