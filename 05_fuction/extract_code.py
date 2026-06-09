def extract_incode(self, inCode):
    """
    원본 상품코드를 내부 표준 형식으로 변환하는 메서드.

    다양한 쇼핑몰/거래처에서 수신된 외부 상품코드를
    회사 내부 규칙에 맞는 표준 코드로 변환합니다.
    변환에 실패하거나 매칭되는 패턴이 없으면 빈 문자열("")을 반환합니다.

    Args:
        inCode: 변환할 원본 상품코드 (str 또는 str로 변환 가능한 타입)

    Returns:
        str: 변환된 내부 표준 상품코드.
             매칭되는 패턴이 없으면 빈 문자열("") 반환.

    코드 인덱스 구조 참고 (C/D/E/H/R/Z 계열):
        [0] 첫 번째 문자 : 대분류 접두사 (C/D/E/H/R/Z)
        [1] 두 번째 문자 : 상품코드 코드
        [2] 세 번째 문자 : 상품코드 코드
        [3] 네 번째 문자 : 시즌/라인 구분자 (T=티셔츠, U=유니섹스, W=여성, M=남성 등)
        [4] 다섯 번째 문자: 추가 구분 코드
        [5][6]           : 상품 일련번호

    prefix_map 매핑표 (inCode[1] → 2자리 표준 접두사):
        S → JS,  U → JU,  A → BA,  C → BC,
        L → JL,  M → JM,  N → NA,  J → EJ,  F → JF

    Note:
        대소문자 구분 없음 (입력값을 자동으로 대문자로 변환하여 처리).
    """

    # 입력값을 문자열로 강제 변환 후 대문자로 통일
    # → None, 숫자 등 비문자열 입력도 안전하게 처리됨
    inCode = str(inCode).upper()

    # ── prefix_map: 단일 문자 → 2자리 표준 접두사 변환 테이블 ──────────────────
    # 주로 C/D/E/H/R/Z 계열 코드의 inCode[1]을 표준 접두사로 변환할 때 사용.
    # 매핑 없는 문자는 get() 호출 시 기본값 'XX' 반환으로 처리.
    prefix_map = {
        "S": "JS", "U": "JU", "A": "BA", "C": "BC",
        "L": "JL", "M": "JM", "N": "NA", "J": "EJ",
        "F": "JF"
    }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [그룹 1] W / M / T / U / BHP / GS / DS / CS / PAC / SPH 접두사 처리
    #   이 그룹에 속하는 코드는 4번째 문자(인덱스 3)의 값에 따라 분기하거나,
    #   더 구체적인 접두사(TS, TPNS 등)를 추가로 확인합니다.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("W", "M", "T", "U", "BHP", "GS", "DS", "CS", "PAC", "SPH")):

        # ── (1-A) 4번째 문자(인덱스 3)가 "A"인 경우 ─────────────────────────
        #   입력 예시: "WBAA12345" → "WBA-12345"
        #   구조: 앞 3자 + "-" + 5번째 이후(goodNum)
        if len(inCode) > 3 and inCode[3] == "A":
            goodNum = inCode[4:]              # 5번째 문자부터 끝까지 = 상품번호
            return f"{inCode[:3]}-{goodNum}"  # "WBA-12345"

        # ── (1-B) 4번째 문자(인덱스 3)가 "F"인 경우 ─────────────────────────
        #   입력 예시: "WBAF12345" → "WJF-12345"
        #   구조: 첫 번째 문자 + "JF-" + 5번째 이후(goodNum)
        #   ※ (1-A) 조건이 먼저 체크되므로 "A"가 아닌 "F"만 여기에 해당
        if len(inCode) > 3 and inCode[3] == "F":
            goodNum = inCode[4:]
            return f"{inCode[0]}JF-{goodNum}"  # "WJF-12345"

        # ── (1-C) "TS"로 시작하는 경우 ───────────────────────────────────────
        #   입력 예시: "TS12345" → "TPS-2345"
        #   구조: "TPS-" + 4번째 이후(goodNum)
        #   ※ T 계열이므로 그룹 1의 T 조건에 걸린 후 여기서 세부 처리
        elif inCode.startswith("TS"):
            goodNum = inCode[3:]      # 4번째 문자(인덱스 2)부터 끝까지
            return f"TPS-{goodNum}"  # "TPS-12345"

        # ── (1-D) "TPNS"로 시작하는 경우 ─────────────────────────────────────
        #   입력 예시: "TPNS12345" → "TPN-12345"
        #   구조: 앞 3자("TPN") + "-" + 5번째 이후(goodNum)
        elif inCode.startswith("TPNS"):
            goodNum = inCode[4:]              # 5번째 문자부터 끝까지
            return f"{inCode[:3]}-{goodNum}"  # "TPN-12345"

        # ── (1-E) "TGJS" 또는 "TLJS"로 시작하는 경우 ────────────────────────
        #   입력 예시: "TGJS12345" → "TG-12345"
        #             "TLJS12345" → "TL-12345"
        #   구조: 앞 2자("TG" 또는 "TL") + "-" + 5번째 이후(goodNum)
        elif inCode.startswith(("TGJS", "TLJS")):
            goodNum = inCode[4:]              # 5번째 문자부터 끝까지
            return f"{inCode[:2]}-{goodNum}"  # "TG-12345"

        # ── (1-F) "UAA-" ~ "UAR-" 계열로 시작하는 경우 ──────────────────────
        #   입력 예시: "UAA-12345" → "UAAO12345"
        #   구조: 앞 3자("UAA") + "O" + 5번째 이후(goodNum)
        #   ※ 대시(-) 다음부터가 실제 상품번호이며, "O"(알파벳)를 삽입
        elif inCode.startswith(("UAA-", "UAB-", "UAC-", "UAD-", "UAE-", "UAF-", "UAR-")):
            goodNum = inCode[4:]              # "-" 이후의 숫자/문자 부분
            return f"{inCode[:3]}O{goodNum}"  # "UAAO12345"

        # ── (1-G) 위 세부 패턴에 모두 해당하지 않는 나머지 ──────────────────
        #   그룹 1 접두사이지만 더 구체적인 패턴이 없으면 원본 코드 그대로 반환
        return inCode

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [그룹 2] C / D / E / H / R / Z 접두사 처리 (복합 패턴)
    #   코드 위치(인덱스)별 의미는 위 docstring의 "코드 인덱스 구조 참고" 참조.
    #   조건 순서가 중요: 더 구체적인 패턴(P7, RS, PN7T13 등)을 먼저 확인.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    elif inCode.startswith(("C", "D", "E", "H", "R", "Z")):

        # ── (2-A) 2~3번째 문자(인덱스 1~2)가 "P7"인 경우 ───────────────────
        #   입력 예시: "_P7_XX"  (언더바=임의 문자)
        #   goodNum = inCode[2]("7") + inCode[4] + inCode[5]
        #   반환: "BHP-7XX"
        if len(inCode) > 3 and inCode[1:3] == "P7":
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"BHP-{goodNum}"

        # ── (2-B) "RS"로 시작하고 4번째 문자(인덱스 3)가 "T"인 경우 ─────────
        #   입력 예시: "RS_T__"  → TPS 계열
        #   goodNum = inCode[2] + inCode[4] + inCode[5]
        #   반환: "TPS-XXX"
        elif len(inCode) > 3 and inCode[3] == "T" and inCode[:2] == "RS":
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"TPS-{goodNum}"

        # ── (2-C) 2번째 문자부터 끝이 "PN7T13"인 경우 ───────────────────────
        #   입력 예시: "_PN7T13"  (총 7자짜리 코드)
        #   goodNum = inCode[3]("T") + inCode[5]("1") + inCode[6]("3")
        #   반환: "BHP-T13"
        elif inCode[1:] in ("PN7T13"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"BHP-{goodNum}"

        # ── (2-D) 2~3번째 문자(인덱스 1~2)가 "PP" 또는 "PK"인 경우 ─────────
        #   입력 예시: "_PP_X__" 또는 "_PK_X__"
        #   goodNum = inCode[3] + inCode[5] + inCode[6]
        #   반환: "BHP-XXX"
        elif inCode[1:3] in ("PP", "PK"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"BHP-{goodNum}"

        # ── (2-E) 2~3번째 문자(인덱스 1~2)가 "PL"/"PG"/"PM"/"PH"인 경우 ─────
        #   입력 예시: "_PL_X__"
        #   goodNum = inCode[3] + inCode[5] + inCode[6]
        #   반환: "[4][1][2]-XXX"
        #   예) inCode[4]='T', inCode[1]='P', inCode[2]='L' → "TPL-XXX"
        elif inCode[1:3] in ("PL", "PG", "PM", "PH"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"

        # ── (2-F) 4번째 문자(인덱스 3)가 "T" 또는 "U"인 경우 ────────────────
        #   goodNum = inCode[2] + inCode[4] + inCode[5]
        #   반환: "[3][1]-XXX"
        #   예) inCode[3]="T", inCode[1]="S" → "TS-XXX"
        #       inCode[3]="U", inCode[1]="J" → "UJ-XXX"
        elif len(inCode) > 3 and inCode[3] in ("T", "U"):
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"{inCode[3]}{inCode[1]}-{goodNum}"

        # ── (2-G) 4번째 문자(인덱스 3)가 "W" 또는 "M"인 경우 ────────────────
        #   goodNum = inCode[2] + inCode[4] + inCode[5]
        #   prefix_map으로 inCode[1]을 2자리 표준 접두사로 변환
        #   반환: "[3]{prefix}-XXX"
        #   예) inCode[3]="W", inCode[1]="S" → prefix="JS" → "WJS-XXX"
        elif len(inCode) > 3 and inCode[3] in ("W", "M"):
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"{inCode[3]}{prefix_map.get(inCode[1], 'XX')}-{goodNum}"

        # ── (2-H) 5번째 문자(인덱스 4)가 "T"인 경우 ─────────────────────────
        #   goodNum = inCode[3] + inCode[5] + inCode[6]
        #   반환: "[4][1][2]-XXX"
        #   예) inCode[4]='T', inCode[1]='P', inCode[2]='L' → "TPL-XXX"
        elif len(inCode) > 4 and inCode[4] in ("T"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"

        # ── (2-I) 5번째 문자(인덱스 4)가 "U"인 경우 ─────────────────────────
        #   goodNum = inCode[3] + inCode[5] + inCode[6]
        #   반환: "[4][1][2]O{goodNum}"  (U 계열은 "O" 삽입)
        #   예) inCode[4]='U', [1]='P', [2]='L' → "UPLOgoodNum"
        elif len(inCode) > 4 and inCode[4] in ("U"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}O{goodNum}"

        # ── (2-J) 5번째 문자(인덱스 4)가 "W" 또는 "M"인 경우 ────────────────
        #   goodNum = inCode[3] + inCode[5] + inCode[6]
        #   반환: "[4]BP-XXX"  ("W"/"M" 계열은 "BP" 고정 접두사)
        #   예) inCode[4]="W" → "WBP-XXX"
        elif len(inCode) > 4 and inCode[4] in ("W", "M"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}BP-{goodNum}"

    # ── 모든 패턴에 해당하지 않으면 빈 문자열 반환 ────────────────────────────
    return ""