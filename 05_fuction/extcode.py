def extcode(inCode):
    """
    상품 코드를 내부 표준 형식으로 변환하는 함수.

    다양한 쇼핑몰/거래처에서 수신된 외부 상품코드를
    회사 내부 규칙에 맞는 표준 코드로 변환합니다.
    변환에 실패하거나 패턴이 일치하지 않으면 빈 문자열("")을 반환합니다.

    Args:
        inCode (str): 변환할 원본 상품코드 (대소문자 무관)

    Returns:
        str: 변환된 내부 표준 상품코드.
             매칭되는 패턴이 없으면 빈 문자열("") 반환.

    변환 우선순위 (위에서 아래 순서로 처리):
        1. TPNS 접두사        → TPN-XXXXX
        2. TGJS / TLJS 접두사 → TG-XXXXX / TL-XXXXX
        3. TS 접두사          → TPS-XXXXX
        4. UAA- ~ UAR- 접두사 → UAAO / UABО / ... 형식
        5. W/M/T/U/BHP/GS/DS/CS/PAC/SPH 접두사
        6. C/D/E/H/R/Z 접두사 (복합 패턴 매칭)
    """

    # ── 입력값 유효성 검사 ──────────────────────────────────────────────────────
    if not inCode:
        # 빈 문자열, None, 0 등 falsy 값이면 즉시 빈 문자열 반환
        return ""

    # 대소문자 통일: 모든 처리는 대문자 기준으로 수행
    inCode = inCode.upper()

    # ── prefix_map: 단일 문자 → 2자리 표준 접두사 변환 테이블 ──────────────────
    # inCode[1] (2번째 문자)를 표준 접두사로 변환할 때 사용.
    # 예) 'S' → 'JS', 'A' → 'BA', 'J' → 'EJ'
    prefix_map = {
        "S": "JS", "U": "JU", "A": "BA", "C": "BC",
        "L": "JL", "M": "JM", "N": "NA", "J": "EJ",
        "F": "JF"
    }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 1] TPNS 접두사 처리
    #   입력 예시: "TPNS12345"
    #   변환 결과: "TPN-12345"  (앞 3자 + "-" + 5자 이후 나머지)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("TPNS",)):
        if len(inCode) > 4:  # "TPNS" + 최소 1자 이상 필요
            goodNum = inCode[4:]          # 5번째 문자부터 끝까지 (상품 번호)
            return f"{inCode[:3]}-{goodNum}"  # "TPN-XXXXX"

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 2] TGJS / TLJS 접두사 처리
    #   입력 예시: "TGJS12345" → "TG-12345"
    #             "TLJS12345" → "TL-12345"
    #   (앞 2자 + "-" + 5번째 이후 나머지)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("TGJS", "TLJS")):
        if len(inCode) > 4:  # "TGJS" + 최소 1자 이상 필요
            goodNum = inCode[4:]          # 5번째 문자부터 끝까지
            return f"{inCode[:2]}-{goodNum}"  # "TG-XXXXX" 또는 "TL-XXXXX"

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 3] TS 접두사 처리  (단, TGJS/TLJS/TPNS는 위에서 먼저 처리됨)
    #   입력 예시: "TS12345" → "TPS-2345"
    #   (4번째 문자부터 끝 = 상품번호; 반환: "TPS-XXXXX")
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("TS",)):
        if len(inCode) > 3:  # "TS" + 최소 2자 이상 필요 (인덱스 3 접근)
            goodNum = inCode[3:]      # 4번째 문자부터 끝까지
            return f"TPS-{goodNum}"  # "TPS-XXXXX"

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 4] UAA- / UAB- / ... / UAR- 접두사 처리
    #   입력 예시: "UAA-12345" → "UAAO12345"
    #   (앞 3자 + "O" + 5번째 이후 나머지; "O"는 문자 O, 숫자 0 아님)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("UAA-", "UAB-", "UAC-", "UAD-", "UAE-", "UAF-", "UAR-")):
        if len(inCode) > 4:  # "UAX-" 4자 + 최소 1자 이상 필요
            goodNum = inCode[4:]              # "-" 이후의 숫자/문자 부분
            return f"{inCode[:3]}O{goodNum}"  # "UAA" + "O" + "12345"

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 5] W / M / T / U / BHP / GS / DS / CS / PAC / SPH 접두사 처리
    #   이 접두사로 시작하는 코드는 4번째 문자(인덱스 3)의 값에 따라 두 갈래로 분기:
    #
    #   (5-A) 4번째 문자 == "A":
    #       입력 예시: "WBAA12345" → "WBA-12345"
    #       (앞 3자 + "-" + 5번째 이후)
    #
    #   (5-B) 4번째 문자 == "F":
    #       입력 예시: "WBAF12345" → "WJF-12345"
    #       (1번째 문자 + "JF-" + 5번째 이후)
    #
    #   위 두 경우 외에는 inCode를 그대로 반환 (변환 없음)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("W", "M", "T", "U", "BHP", "GS", "DS", "CS", "PAC", "SPH")):
        if len(inCode) > 3:
            if inCode[3] == "A":
                # (5-A) 4번째 문자가 "A"인 경우 → "앞3자-나머지" 형식
                if len(inCode) > 4:  # 5번째 문자 이상 존재해야 함
                    goodNum = inCode[4:]
                    return f"{inCode[:3]}-{goodNum}"
            elif inCode[3] == "F":
                # (5-B) 4번째 문자가 "F"인 경우 → "첫글자JF-나머지" 형식
                if len(inCode) > 4:
                    goodNum = inCode[4:]
                    return f"{inCode[0]}JF-{goodNum}"
        # 위 두 조건 모두 해당 없으면 원본 코드 그대로 반환
        return inCode

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # [패턴 6] C / D / E / H / R / Z 접두사 처리 (복합 패턴)
    #   코드 구조 예시 (인덱스 기준):
    #       [0] = 첫 번째 문자 (C/D/E/H/R/Z)
    #       [1] = 브랜드/시리즈 코드
    #       [2] = 추가 분류 코드
    #       [3] = 시즌/라인 구분자 (T=봄여름, U=가을겨울, W=여성, M=남성 등)
    #       [4] = 추가 구분 코드
    #       [5][6] = 상품 일련번호
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if inCode.startswith(("C", "D", "E", "H", "R", "Z")):
        if len(inCode) < 2:
            # 접두사 1자만 있는 경우: 변환 불가
            return ""

        # ── (6-1) "RS"로 시작하고 4번째 문자(인덱스 3)가 "T"인 경우 ──────────
        #   입력 예시: "RS_T__"  (언더바는 임의 문자)
        #   goodNum = [2] + [4] + [5]
        #   반환: "TPS-XXX"
        if inCode.startswith("RS") and len(inCode) > 3 and inCode[3] == "T":
            if len(inCode) > 5:  # 인덱스 5까지 존재해야 함
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"TPS-{goodNum}"

        # ── (6-2) 2~3번째 문자(인덱스 1~2)가 "P7"인 경우 ───────────────────
        #   입력 예시: "_P7_XX"
        #   goodNum = [2] + [4] + [5]  →  "7" + [4] + [5]
        #   반환: "BHP-XXX"
        if len(inCode) > 3 and inCode[1:3] == "P7":
            if len(inCode) > 5:
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"BHP-{goodNum}"

        # ── (6-3) 2번째 문자부터 끝이 "PN7T13"인 경우 ─────────────────────
        #   입력 예시: "_PN7T13"  (총 7자)
        #   goodNum = [3] + [5] + [6]  →  "T" + "1" + "3"
        #   반환: "BHP-T13"
        if len(inCode) > 1 and inCode[1:] == "PN7T13":
            if len(inCode) > 6:  # 인덱스 6까지 존재해야 함
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"BHP-{goodNum}"

        # ── (6-4) 2~3번째 문자(인덱스 1~2)가 "PP" 또는 "PK"인 경우 ─────────
        #   입력 예시: "_PP_X__" 또는 "_PK_X__"
        #   goodNum = [3] + [5] + [6]
        #   반환: "BHP-XXX"
        if len(inCode) > 3 and inCode[1:3] in ("PP", "PK"):
            if len(inCode) > 6:
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"BHP-{goodNum}"

        # ── (6-5) 2~3번째 문자(인덱스 1~2)가 "PL"/"PG"/"PM"/"PH"인 경우 ────
        #   입력 예시: "_PL_X__"
        #   goodNum = [3] + [5] + [6]
        #   반환: "[4][1][2]-XXX"  예) 인덱스[4]가 'T', [1]이 'P', [2]가 'L' 이면 "TPL-XXX"
        if len(inCode) > 3 and inCode[1:3] in ("PL", "PG", "PM", "PH"):
            if len(inCode) > 6:
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"

        # ── (6-6) 4번째 문자(인덱스 3)가 "T" 또는 "U"인 경우 ────────────────
        #   goodNum = [2] + [4] + [5]
        #   반환: "[3][1]-XXX"
        #   예) inCode[3]="T", inCode[1]="S" →  "TS-XXX"
        #       inCode[3]="U", inCode[1]="J" →  "UJ-XXX"
        if len(inCode) > 3 and inCode[3] in ("T", "U"):
            if len(inCode) > 5:
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"{inCode[3]}{inCode[1]}-{goodNum}"

        # ── (6-7) 4번째 문자(인덱스 3)가 "W" 또는 "M"인 경우 ────────────────
        #   goodNum = [2] + [4] + [5]
        #   prefix_map으로 inCode[1]을 2자리 접두사로 변환
        #   반환: "[3]{prefix_map[1]}-XXX"
        #   예) inCode[3]="W", inCode[1]="S" → prefix="JS" → "WJS-XXX"
        if len(inCode) > 3 and inCode[3] in ("W", "M"):
            if len(inCode) > 5:
                goodNum = inCode[2] + inCode[4] + inCode[5]
                prefix = prefix_map.get(inCode[1], "XX")  # 매핑 없으면 "XX" 사용
                return f"{inCode[3]}{prefix}-{goodNum}"

        # ── (6-8) 5번째 문자(인덱스 4)가 "T"인 경우 ─────────────────────────
        #   goodNum = [3] + [5] + [6]
        #   반환: "[4][1][2]-XXX"  예) "T" + "P" + "L" + "-" + goodNum → "TPL-XXX"
        if len(inCode) > 4 and inCode[4] == "T":
            if len(inCode) > 6:
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"

        # ── (6-9) 5번째 문자(인덱스 4)가 "U"인 경우 ─────────────────────────
        #   goodNum = [3] + [5] + [6]
        #   반환: "[4][1][2]O{goodNum}"  (U 계열은 "O" 삽입)
        #   예) "[4]='U', [1]='P', [2]='L'" → "UPLOgoodNum"
        if len(inCode) > 4 and inCode[4] == "U":
            if len(inCode) > 6:
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}{inCode[1]}{inCode[2]}O{goodNum}"

        # ── (6-10) 5번째 문자(인덱스 4)가 "W" 또는 "M"인 경우 ────────────────
        #   goodNum = [3] + [5] + [6]
        #   반환: "[4]BP-XXX"  (W/M 계열은 "BP" 고정 접두사)
        #   예) inCode[4]="W" → "WBP-XXX"
        if len(inCode) > 4 and inCode[4] in ("W", "M"):
            if len(inCode) > 6:
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}BP-{goodNum}"

    # ── 모든 패턴에 해당하지 않으면 빈 문자열 반환 ────────────────────────────
    return ""