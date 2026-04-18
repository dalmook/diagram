#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mermaid Report Designer (Tkinter + WebView2 preview + Excel editable export)

What this app does
- Natural language -> Mermaid flowchart generation via OpenAI-compatible / GPT-OSS-compatible endpoint.
- Tkinter editor for prompt and Mermaid code.
- Preview in a WebView2-backed window through pywebview on Windows.
- Export SVG and PNG without mmdc (SVG from Mermaid JS render, PNG via cairosvg if available).
- Export Excel workbook with editable shapes/connectors (Windows + Excel + pywin32 required).

Important notes
- Preview uses pywebview. On Windows, pywebview typically uses EdgeChromium/WebView2 backend when available.
- Excel editable export currently targets FLOWCHART diagrams and a practical subset of Mermaid flowchart syntax.
- For robust report usage, the LLM prompt is intentionally constrained to generate flowchart-compatible Mermaid.
"""
from __future__ import annotations

import json
import logging
import math
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import requests
except Exception:  # pragma: no cover
    requests = None

try:
    import webview  # pywebview
except Exception:  # pragma: no cover
    webview = None

try:
    import cairosvg
except Exception:  # pragma: no cover
    cairosvg = None

try:
    import win32com.client as win32
except Exception:  # pragma: no cover
    win32 = None

APP_TITLE = "Mermaid Report Designer"
SETTINGS_PATH = Path.home() / ".mermaid_report_designer_settings.json"
APP_STATE_DIR = Path.home() / ".mermaid_report_designer"
LOGS_DIR = APP_STATE_DIR / "logs"
APP_LOG_PATH = LOGS_DIR / "app.log"
DEFAULT_TIMEOUT = 90

OFFICE_CONST_FALLBACKS = {
    "msoShapeRectangle": 1,
    "msoShapeParallelogram": 2,
    "msoShapeDiamond": 4,
    "msoShapeRoundedRectangle": 5,
    "msoShapeHexagon": 10,
    "msoShapeOval": 9,
    "msoShapeFlowchartDocument": 67,
    "msoShapeFlowchartTerminator": 69,
    "msoShapeFlowchartPreparation": 70,
    "msoShapeFlowchartPredefinedProcess": 65,
    "msoShapeFlowchartStoredData": 83,
    "msoConnectorStraight": 1,
    "msoConnectorElbow": 2,
    "msoArrowheadTriangle": 2,
    "msoAnchorMiddle": 3,
    "msoAlignCenter": 2,
    "msoTextOrientationHorizontal": 1,
}

HARD_CODED_LLM_CONFIG = {
    "base_url": "",
    "model": "openai/gpt-oss-120b",
    "api_key": "",
    "gpt_oss_credential": "",
    "gpt_oss_user_id": "",
    "gpt_oss_user_type": "AD_ID",
    "gpt_oss_system_name": "MERMAID_REPORT_DESIGNER",
    "timeout_sec": DEFAULT_TIMEOUT,
}

MERMAID_JS_CDN = "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"

THEMES: Dict[str, Dict[str, str]] = {
    "Executive Blue": {
        "primaryColor": "#EAF2FF",
        "primaryBorderColor": "#2F5AA8",
        "primaryTextColor": "#14305E",
        "lineColor": "#2F5AA8",
        "secondaryColor": "#F2F6FC",
        "tertiaryColor": "#DDEBFF",
        "fontFamily": "Malgun Gothic, Segoe UI, Arial",
    },
    "Calm Green": {
        "primaryColor": "#EAF7F0",
        "primaryBorderColor": "#2F7D5A",
        "primaryTextColor": "#174734",
        "lineColor": "#2F7D5A",
        "secondaryColor": "#F1FBF6",
        "tertiaryColor": "#D9F2E3",
        "fontFamily": "Malgun Gothic, Segoe UI, Arial",
    },
    "Warm Gray": {
        "primaryColor": "#F5F5F5",
        "primaryBorderColor": "#5F6368",
        "primaryTextColor": "#2B2F33",
        "lineColor": "#5F6368",
        "secondaryColor": "#FAFAFA",
        "tertiaryColor": "#EDEDED",
        "fontFamily": "Malgun Gothic, Segoe UI, Arial",
    },
}

THEMES.update(
    {
        "Modern Dark": {
            "primaryColor": "#1F2937",
            "primaryBorderColor": "#60A5FA",
            "primaryTextColor": "#F9FAFB",
            "lineColor": "#93C5FD",
            "secondaryColor": "#111827",
            "tertiaryColor": "#374151",
            "fontFamily": "Segoe UI, Malgun Gothic, Arial",
        },
        "Clean Gray": {
            "primaryColor": "#F3F4F6",
            "primaryBorderColor": "#6B7280",
            "primaryTextColor": "#1F2937",
            "lineColor": "#9CA3AF",
            "secondaryColor": "#FAFAFA",
            "tertiaryColor": "#E5E7EB",
            "fontFamily": "Segoe UI, Malgun Gothic, Arial",
        },
        "Corporate Blue Plus": {
            "primaryColor": "#E9F2FF",
            "primaryBorderColor": "#1565C0",
            "primaryTextColor": "#0F2F57",
            "lineColor": "#1976D2",
            "secondaryColor": "#F4F8FF",
            "tertiaryColor": "#D6E6FF",
            "fontFamily": "Segoe UI, Malgun Gothic, Arial",
        },
    }
)

DEFAULT_PROMPT = """예: 영업팀이 고객 요청을 검토하고, 기술 검증 후 승인/반려로 나뉘며 승인 시 구축/운영으로 이어지는 프로세스를 보고용으로 깔끔한 Mermaid flowchart로 만들어줘."""

DEFAULT_TEMPLATE = """%%{init: {'theme': 'base', 'themeVariables': { 'fontFamily': 'Malgun Gothic, Segoe UI, Arial', 'primaryColor': '#EAF2FF', 'primaryBorderColor': '#2F5AA8', 'primaryTextColor': '#14305E', 'lineColor': '#355FA8', 'secondaryColor': '#F4F8FF', 'tertiaryColor': '#DDEBFF' }}}%%
title 분기 사업 제안 검토 및 실행 체계
flowchart TD
    A((사업 요청 접수)) --> B[사전 분류 및 담당 지정]
    B --> C[/요청서·ROI·리스크 자료 정리/]
    C --> D{추가 검토 필요?}

    D -->|전략 영향 큼| E[전략위원회 상세 검토]
    D -->|표준 안건| F[부문장 신속 심의]

    E --> G{예산 확보 가능?}
    F --> G

    G -->|아니오| H[보완 요청 및 재제출]
    H --> C

    G -->|예| I[(우선순위 포트폴리오 반영)]
    I --> J[구축 계획 수립]
    J --> K[개발/구축 실행]
    K --> L[통합 테스트]
    L --> M{오픈 승인?}

    M -->|보완 필요| N[개선 조치]
    N --> L

    M -->|승인| O((운영 전환))
    O --> P[운영 KPI 모니터링]
    P --> Q{성과 목표 달성?}
    Q -->|미달| R[원인 분석 및 개선 과제]
    R --> J
    Q -->|달성| S((성과 공유 및 종료))

    classDef start fill:#DFF5E8,stroke:#2F7D5A,color:#174734,stroke-width:2px;
    classDef process fill:#EAF2FF,stroke:#2F5AA8,color:#14305E,stroke-width:1.8px;
    classDef decision fill:#FFF4DD,stroke:#C98A1A,color:#6F4A00,stroke-width:1.8px;
    classDef data fill:#F4ECFF,stroke:#7A4BB7,color:#4C2D73,stroke-width:1.8px;
    classDef risk fill:#FFE7E7,stroke:#C23B3B,color:#7A1F1F,stroke-width:1.8px;

    class A,O,S start;
    class B,C,E,F,J,K,L,N,P,R process;
    class D,G,M,Q decision;
    class I data;
    class H risk;
"""

DIAGRAM_TYPES = [
    "자동 추천",
    "flowchart",
    "gantt",
    "sequenceDiagram",
    "org chart",
    "swimlane",
    "journey",
]

FLOW_DIRECTIONS = ["LR", "RL", "TB", "BT"]


def build_flowchart_template(direction: str = "TB") -> str:
    return f"""title 영업 요청 검토 프로세스
flowchart {direction}
    A([요청 접수]) --> B[요건 검토]
    B --> C{{추가 검증 필요}}
    C -->|예| D[기술 검증]
    C -->|아니오| E[승인 검토]
    D --> E
    E --> F{{승인 여부}}
    F -->|승인| G[구축 진행]
    F -->|반려| H[반려 사유 전달]
    G --> I([운영 전환])
"""


def build_flowchart_template_sample2(direction: str = "LR") -> str:
    return f"""title 장애 대응 프로세스
flowchart {direction}
    A([모니터링 감지]) --> B[장애 분류]
    B --> C{{치명도 높음}}
    C -->|예| D[비상 대응 체계 가동]
    C -->|아니오| E[일반 처리 큐 등록]
    D --> F[원인 분석]
    E --> F
    F --> G[조치 실행]
    G --> H[복구 확인]
    H --> I([사후 보고])
"""


def build_gantt_template() -> str:
    return """title 구축 일정 계획
gantt
    dateFormat  YYYY-MM-DD
    title 구축 일정 계획
    excludes    weekends
    section 기획
    요구사항 정리 / PM        :done, req, 2026-04-21, 5d
    범위 확정 / 기획          :active, scope, after req, 4d
    section 구현
    화면 개발 / FE            :dev1, 2026-05-02, 8d
    연동 개발 / BE            :dev2, after dev1, 6d
    section 검증
    통합 테스트 / QA          :test, after dev2, 5d
    오픈 마일스톤            :milestone, go, after test, 0d
"""


def build_gantt_template_sample2() -> str:
    return """title 보고 체계 개선 일정
gantt
    dateFormat  YYYY-MM-DD
    title 보고 체계 개선 일정
    excludes weekends
    section 분석
    현행 인터뷰 / PM          :a1, 2026-05-01, 3d
    데이터 점검 / 분석가      :a2, after a1, 4d
    section 설계
    템플릿 설계 / 기획        :b1, after a2, 5d
    승인 리뷰 / 임원          :milestone, b2, after b1, 0d
    section 적용
    시범 적용 / 운영          :c1, after b2, 7d
    결과 회고 / 전원          :c2, after c1, 2d
"""


def build_sequence_template() -> str:
    return """title 고객 요청 승인 시나리오
sequenceDiagram
    participant User as 요청자
    participant Sales as 영업
    participant Tech as 기술검토
    participant Lead as 팀장

    User->>Sales: 요청서 제출
    Sales->>Tech: 기술 검토 요청
    Tech-->>Sales: 검토 결과 회신
    alt 승인 가능
        Sales->>Lead: 승인 요청
        Lead-->>Sales: 승인
        Sales-->>User: 진행 일정 안내
    else 보완 필요
        Tech-->>User: 보완 요청
    end
    Note over Sales,Lead: 보고용 요약 기록
"""


def build_sequence_template_sample2() -> str:
    return """title 장애 전파 및 복구 커뮤니케이션
sequenceDiagram
    participant Mon as 모니터링
    participant Ops as 운영팀
    participant Dev as 개발팀
    participant Biz as 현업

    Mon->>Ops: 장애 알림 전송
    Ops->>Dev: 원인 분석 요청
    opt 추가 로그 필요
        Dev->>Ops: 로그 수집 요청
        Ops-->>Dev: 로그 전달
    end
    alt 긴급 장애
        Ops-->>Biz: 긴급 공지
    else 일반 장애
        Ops-->>Biz: 예정 복구 시간 공유
    end
    Dev-->>Ops: 조치 완료
    Ops-->>Biz: 복구 완료 안내
"""


def build_org_chart_template(direction: str = "TB") -> str:
    return f"""title 사업 추진 조직도
flowchart {direction}
    CEO[사업 총괄]
    PMO[PMO]
    SALES[영업 리드]
    TECH[기술 리드]
    OPS[운영 리드]
    CEO --> PMO
    PMO --> SALES
    PMO --> TECH
    PMO --> OPS
    SALES --> S1[고객 커뮤니케이션]
    TECH --> T1[솔루션 설계]
    OPS --> O1[운영 안정화]
"""


def build_org_chart_template_sample2(direction: str = "TB") -> str:
    return f"""title 보고 체계 조직도
flowchart {direction}
    HQ[본부장]
    STR[전략실]
    DEV[개발팀]
    DATA[데이터팀]
    QA[품질팀]
    HQ --> STR
    HQ --> DEV
    HQ --> DATA
    HQ --> QA
    DEV --> DEV1[프론트엔드]
    DEV --> DEV2[백엔드]
    DATA --> DATA1[지표 관리]
"""


def build_swimlane_template(direction: str = "LR", lanes: Optional[Sequence[str]] = None) -> str:
    lane_names = list(lanes or ["요청부서", "검토부서", "승인부서", "운영부서"])
    lane_names = [name.strip() for name in lane_names if name.strip()] or ["요청부서", "검토부서", "승인부서", "운영부서"]
    blocks: List[str] = []
    node_ids: List[Tuple[str, str]] = []
    for index, lane in enumerate(lane_names, start=1):
        start = f"L{index}A"
        end = f"L{index}B"
        blocks.append(
            f"""    subgraph Lane{index}[{lane}]
        direction TB
        {start}[{lane} 작업 시작]
        {end}[{lane} 작업 완료]
        {start} --> {end}
    end"""
        )
        node_ids.append((start, end))
    connectors = [f"    {node_ids[i][1]} --> {node_ids[i + 1][0]}" for i in range(len(node_ids) - 1)]
    return "title 부서별 Swimlane 프로세스\n" + f"flowchart {direction}\n" + "\n".join(blocks + connectors)


def build_swimlane_template_sample2(direction: str = "LR") -> str:
    return f"""title 채용 프로세스 Swimlane
flowchart {direction}
    subgraph Lane1[지원자]
        direction TB
        A1[지원서 제출] --> A2[면접 참석]
    end
    subgraph Lane2[인사팀]
        direction TB
        B1[서류 검토] --> B2[일정 안내]
    end
    subgraph Lane3[실무팀]
        direction TB
        C1[기술 면접] --> C2[평가 공유]
    end
    subgraph Lane4[경영진]
        direction TB
        D1[최종 승인] --> D2[처우 확정]
    end
    A1 --> B1
    B2 --> A2
    A2 --> C1
    C2 --> D1
"""


def build_journey_template() -> str:
    return """title 고객 온보딩 여정
journey
    title 고객 온보딩 여정
    section 인지
      서비스 소개 확인: 4: 고객
      초기 문의 접수: 3: 고객, 영업
    section 검토
      요구사항 조율: 3: 고객, 영업, 기술
      제안서 확인: 4: 고객, 영업
    section 전환
      계약 체결: 5: 고객, 영업, 법무
      운영 시작: 4: 고객, 운영
"""


def build_journey_template_sample2() -> str:
    return """title 내부 보고 자동화 도입 여정
journey
    title 내부 보고 자동화 도입 여정
    section 준비
      현행 자료 수집: 2: 실무자
      개선 요구 정리: 3: 실무자, 팀장
    section 구축
      템플릿 설계: 4: 기획, 개발
      시범 운영: 3: 운영, 실무자
    section 정착
      교육 진행: 4: 운영, 실무자
      정기 활용: 5: 전사
"""


def make_swimlane_from_input(lane_text: str, direction: str = "LR") -> str:
    lanes = [item.strip() for item in re.split(r"[,/>\n]+", lane_text or "") if item.strip()]
    return build_swimlane_template(direction=direction, lanes=lanes)


TYPE_SAMPLES: Dict[str, List[str]] = {
    "flowchart": [build_flowchart_template("TB"), build_flowchart_template_sample2("LR")],
    "gantt": [build_gantt_template(), build_gantt_template_sample2()],
    "sequenceDiagram": [build_sequence_template(), build_sequence_template_sample2()],
    "org chart": [build_org_chart_template("TB"), build_org_chart_template_sample2("TB")],
    "swimlane": [build_swimlane_template("LR"), build_swimlane_template_sample2("LR")],
    "journey": [build_journey_template(), build_journey_template_sample2()],
}

GALLERY_SAMPLES: Dict[str, Dict[str, str]] = {
    "임원보고형": {
        "diagram_type": "flowchart",
        "theme": "Executive Blue",
        "direction": "TB",
        "title": "Executive Review Board",
        "description": "Before: 단순 승인 플로우. After: KPI, 리스크, 예외 흐름, 재검토 루프가 포함된 임원 보고형 프로세스.",
        "code": """title Executive Review Board
flowchart TB
    S((Initiative Intake)) --> A[1. Executive Brief Review]
    A --> B[2. Data Pack Validation]
    B --> C{Strategic Fit}
    C -->|High| D[Board Review]
    C -->|Low| R1[Risk Note Registration]
    R1 --> F[Backlog or Re-scope]
    D --> E{Budget Approved}
    E -->|Yes| G[Program Launch]
    E -->|No| H[Revision Request]
    H --> A
    G --> I[KPI Baseline Setup]
    I --> J{KPI On Track}
    J -->|Yes| K[Monthly Executive Update]
    J -->|No| L[Recovery Taskforce]
    L --> M[Root Cause Analysis]
    M --> N[Mitigation Plan]
    N --> I
    K --> O((Quarterly Share-out))
    F --> P((Closed))
    class S,O,P start;
    class A,B,D,G,K process;
    class C,E,J decision;
    class I data;
    class R1,L,M,N risk;
""",
    },
    "업무 프로세스형": {
        "diagram_type": "swimlane",
        "theme": "Corporate Blue Plus",
        "direction": "LR",
        "title": "Operations Excellence Loop",
        "description": "Before: 네모 박스 나열. After: 역할별 lane, 승인/반려, 피드백 루프, 산출물/KPI 구분이 들어간 업무 프로세스형.",
        "code": """title Operations Excellence Loop
flowchart LR
    subgraph Lane1[사업부]
        direction TB
        A1[요청 접수] --> A2[요구사항 정리]
        A2 --> A3[우선순위 확인]
    end
    subgraph Lane2[PMO]
        direction TB
        B1[검토안 작성] --> B2{투입 가능}
        B2 -->|No| B3[보완 요청]
        B2 -->|Yes| B4[실행 승인 상신]
    end
    subgraph Lane3[기술팀]
        direction TB
        C1[기술 타당성 분석] --> C2[(산출물 패키지)]
        C2 --> C3[KPI 영향 예측]
    end
    subgraph Lane4[운영팀]
        direction TB
        D1[운영 준비] --> D2[릴리즈]
        D2 --> D3{지표 정상}
        D3 -->|No| D4[개선 루프]
        D3 -->|Yes| D5[정착 공유]
    end
    A3 --> B1
    B3 --> A2
    B4 --> C1
    C3 --> D1
    D4 --> C1
""",
    },
    "시스템 연동형": {
        "diagram_type": "sequenceDiagram",
        "theme": "Modern Dark",
        "direction": "TB",
        "title": "System Integration Control Tower",
        "description": "Before: 단순 호출 순서. After: user/system/external/DB 역할과 예외, note, 복구 메시지가 드러나는 시스템 연동형.",
        "code": """title System Integration Control Tower
sequenceDiagram
    participant User as User Portal
    participant App as Core System
    participant Ext as External Gateway
    participant DB as KPI DB

    User->>App: 승인 요청 제출
    App->>DB: 정책/권한 조회
    DB-->>App: 정책 결과
    App->>Ext: 외부 검증 요청
    alt 외부 검증 통과
        Ext-->>App: 정상 응답
        App->>DB: KPI 이벤트 저장
        App-->>User: 승인 완료 알림
    else 외부 오류
        Ext-->>App: Timeout / Error
        Note over App,Ext: 예외 로그와 재시도 정책 적용
        App->>DB: 장애 이벤트 적재
        opt 재시도 가능
            App->>Ext: 재검증 요청
            Ext-->>App: 최종 응답
        end
        App-->>User: 수동 확인 안내
    end
""",
    },
    "일정관리형": {
        "diagram_type": "gantt",
        "theme": "Clean Gray",
        "direction": "TB",
        "title": "Executive Delivery Calendar",
        "description": "Before: 기본 일정표. After: plan/in-progress/done/delayed/milestone이 구분되는 일정관리형 샘플.",
        "code": """title Executive Delivery Calendar
gantt
    dateFormat  YYYY-MM-DD
    title Executive Delivery Calendar
    excludes weekends
    section Plan
    Kickoff / PM                 :done, kick, 2026-05-01, 2d
    Scope Freeze / Strategy      :done, scope, after kick, 4d
    section In-Progress
    Platform Build / Core Team   :active, build, 2026-05-09, 8d
    Security Review / Infra      :active, sec, after scope, 5d
    section Delayed
    Vendor API Contract / Ext    :delay1, 2026-05-12, 6d
    section Done
    Reporting Template / PMO     :done, rpt, 2026-05-04, 3d
    section Milestone
    Exec Steering Committee      :milestone, m1, after build, 0d
    Go Live                      :milestone, m2, after sec, 0d
""",
    },
}


# ----------------------------- Utilities -----------------------------

def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def slugify_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name).strip()
    return name or "diagram"


def compact_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def configure_file_logging() -> None:
    ensure_parent(APP_LOG_PATH)
    logging.basicConfig(
        filename=str(APP_LOG_PATH),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )


def append_exception_log(context: str, exc: Exception) -> None:
    logging.error("%s: %s\n%s", context, exc, traceback.format_exc())


def parse_color(value: str, default: str = "#DDEBFF") -> str:
    value = (value or "").strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", value):
        return value.upper()
    return default


def hex_to_bgr_int(hex_color: str) -> int:
    hex_color = parse_color(hex_color)
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    return b << 16 | g << 8 | r


def detect_mermaid_diagram_type(code: str) -> str:
    for raw in code.splitlines():
        line = raw.strip()
        if not line or INIT_COMMENT_RE.match(line):
            continue
        low = line.lower()
        if low.startswith("flowchart ") or low.startswith("graph "):
            if "subgraph" in code:
                return "swimlane"
            return "flowchart"
        if low.startswith("gantt"):
            return "gantt"
        if low.startswith("sequencediagram"):
            return "sequenceDiagram"
        if low.startswith("journey"):
            return "journey"
        if low.startswith("statediagram-v2"):
            return "stateDiagram-v2"
    return "flowchart"


def recommend_diagram_type(user_request: str) -> str:
    text = compact_text(user_request).lower()
    if any(keyword in text for keyword in ["일정", "스케줄", "기간", "마일스톤", "착수", "완료일"]):
        return "gantt"
    if any(keyword in text for keyword in ["대화", "메시지", "호출", "응답", "시퀀스", "api", "인터랙션"]):
        return "sequenceDiagram"
    if any(keyword in text for keyword in ["조직도", "조직", "본부", "팀 구조", "보고 라인"]):
        return "org chart"
    if any(keyword in text for keyword in ["swimlane", "lane", "부서별", "역할별", "담당별", "레인"]):
        return "swimlane"
    if any(keyword in text for keyword in ["여정", "경험", "onboarding", "온보딩", "journey"]):
        return "journey"
    return "flowchart"


def get_template_for_type(diagram_type: str, flow_direction: str = "TB", swimlane_lanes: str = "") -> str:
    if diagram_type == "flowchart":
        return build_flowchart_template(flow_direction)
    if diagram_type == "gantt":
        return build_gantt_template()
    if diagram_type == "sequenceDiagram":
        return build_sequence_template()
    if diagram_type == "org chart":
        return build_org_chart_template(flow_direction)
    if diagram_type == "swimlane":
        return make_swimlane_from_input(swimlane_lanes, direction=flow_direction if flow_direction in {"LR", "RL", "TB", "BT"} else "LR")
    if diagram_type == "journey":
        return build_journey_template()
    return build_flowchart_template(flow_direction)


# ----------------------------- LLM client -----------------------------
class LLMClient:
    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def is_ready(self) -> bool:
        return bool((self.config.get("base_url") or "").strip()) and requests is not None

    def headers(self) -> Dict[str, str]:
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        api_key = (self.config.get("api_key") or "").strip()
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"
        credential = (self.config.get("gpt_oss_credential") or "").strip()
        if credential:
            headers["x-dep-ticket"] = credential
        system_name = (self.config.get("gpt_oss_system_name") or "").strip()
        if system_name:
            headers["Send-System-Name"] = system_name
        user_id = (self.config.get("gpt_oss_user_id") or "").strip()
        if user_id:
            headers["User-Id"] = user_id
        user_type = (self.config.get("gpt_oss_user_type") or "").strip()
        if user_type:
            headers["User-Type"] = user_type
        return headers

    def build_messages(self, user_request: str, theme_name: str) -> List[Dict[str, str]]:
        system = (
            "You are a precise Mermaid flowchart architect for executive reporting. "
            "Always output two XML-like tags only: <title>...</title> and <mermaid>...</mermaid>. "
            "Inside <mermaid>, produce ONLY Mermaid flowchart syntax. Do not use markdown code fences. "
            "Prefer flowchart TD or LR. Use stable node ids like A, B, C, D1. "
            "Use concise Korean labels suitable for business reporting. "
            "Include classDef lines and class assignments for readable color styling. "
            "Do not output unsupported or exotic syntax. Keep the diagram editable and structured."
        )
        user = f"""
다음 자연어 요구를 보고용 Mermaid flowchart로 만들어줘.

요구사항:
{user_request.strip()}

추가 규칙:
- 다이어그램 유형은 flowchart 로만 작성
- 테마 이름 참고: {theme_name}
- 보고서용으로 깔끔하게 정리
- 시작/종료, 의사결정, 처리 단계가 보이면 구분
- classDef 를 포함해서 색상 스타일을 붙여줘
- 결과는 반드시 다음 형식만 출력
<title>다이어그램 제목</title>
<mermaid>...mermaid code...</mermaid>
""".strip()
        return [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]

    def generate_mermaid(self, user_request: str, theme_name: str) -> Dict[str, str]:
        if not self.is_ready():
            raise RuntimeError("LLM 설정이 비어 있거나 requests가 설치되지 않았습니다.")
        payload = {
            "model": self.config.get("model") or "openai/gpt-oss-120b",
            "messages": self.build_messages(user_request, theme_name),
            "temperature": 0.2,
            "max_tokens": 1400,
            "stream": False,
        }
        timeout = safe_int(self.config.get("timeout_sec"), DEFAULT_TIMEOUT)
        response = requests.post(
            (self.config.get("base_url") or "").strip(),
            headers=self.headers(),
            json=payload,
            timeout=timeout,
        )
        response.raise_for_status()
        data = response.json()
        content = self._extract_content(data)
        title = self._extract_tag(content, "title") or "Generated Mermaid"
        mermaid = self._extract_tag(content, "mermaid")
        if not mermaid:
            raise RuntimeError("LLM 응답에서 <mermaid>...</mermaid> 를 찾지 못했습니다.")
        return {"title": title.strip(), "mermaid": mermaid.strip()}

    @staticmethod
    def _extract_content(data: Dict[str, Any]) -> str:
        choices = data.get("choices") or []
        if not choices:
            return ""
        message = (choices[0] or {}).get("message") or {}
        content = message.get("content")
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            chunks: List[str] = []
            for item in content:
                if isinstance(item, dict):
                    text = item.get("text")
                    if isinstance(text, str):
                        chunks.append(text)
                elif isinstance(item, str):
                    chunks.append(item)
            return "\n".join(chunks)
        reasoning = message.get("reasoning")
        if isinstance(reasoning, str):
            return reasoning
        return ""

    @staticmethod
    def _extract_tag(text: str, tag: str) -> str:
        match = re.search(fr"<{tag}>(.*?)</{tag}>", text, flags=re.DOTALL | re.IGNORECASE)
        return match.group(1).strip() if match else ""


def _llm_build_messages(self, user_request: str, theme_name: str, diagram_type: str, flow_direction: str) -> List[Dict[str, str]]:
    target_type = recommend_diagram_type(user_request) if diagram_type == "자동 추천" else diagram_type
    system = (
        "You are a precise Mermaid diagram architect for executive reporting. "
        "Always output exactly three XML-like tags only: <title>...</title>, <description>...</description>, <mermaid>...</mermaid>. "
        "Inside <mermaid>, produce only valid Mermaid syntax for the requested diagram type. "
        "Do not use markdown code fences. Use concise Korean labels and keep syntax conservative."
    )
    user = f"""
다음 자연어 요구를 Mermaid 다이어그램으로 만들어줘.

요구사항:
{user_request.strip()}

추가 규칙:
- 다이어그램 타입: {target_type}
- 사용자가 자동 추천을 선택한 경우 가장 적합한 타입을 먼저 판단
- flowchart, org chart, swimlane 방향: {flow_direction}
- 테마 참고: {theme_name}
- gantt는 시작일, 기간, 담당 또는 마일스톤이 보이게 작성
- sequenceDiagram은 participant, message, alt/opt/note를 적절히 활용
- swimlane는 subgraph 기반 lane 구조 사용
- 설명은 2문장 이내
- 아래 형식만 출력
<title>다이어그램 제목</title>
<description>간단한 설명</description>
<mermaid>...mermaid code...</mermaid>
""".strip()
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": user},
    ]


def _llm_generate_mermaid(self, user_request: str, theme_name: str, diagram_type: str, flow_direction: str) -> Dict[str, str]:
    if not self.is_ready():
        raise RuntimeError("LLM configuration is not ready.")
    payload = {
        "model": self.config.get("model") or "openai/gpt-oss-120b",
        "messages": self.build_messages(user_request, theme_name, diagram_type, flow_direction),
        "temperature": 0.2,
        "max_tokens": 1800,
        "stream": False,
    }
    timeout = safe_int(self.config.get("timeout_sec"), DEFAULT_TIMEOUT)
    response = requests.post(
        (self.config.get("base_url") or "").strip(),
        headers=self.headers(),
        json=payload,
        timeout=timeout,
    )
    response.raise_for_status()
    data = response.json()
    content = self._extract_content(data)
    title = self._extract_tag(content, "title") or "Generated Mermaid"
    description = self._extract_tag(content, "description") or ""
    mermaid = self._extract_tag(content, "mermaid")
    if not mermaid:
        raise RuntimeError("LLM response does not contain a <mermaid> block.")
    return {"title": title.strip(), "description": description.strip(), "mermaid": mermaid.strip()}


LLMClient.build_messages = _llm_build_messages
LLMClient.generate_mermaid = _llm_generate_mermaid


# ----------------------------- Mermaid preview -----------------------------
PREVIEW_HTML_TEMPLATE = r"""
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Mermaid Preview</title>
<style>
  html, body { margin:0; padding:0; background:#f5f7fb; font-family: Segoe UI, Malgun Gothic, Arial, sans-serif; }
  #toolbar { padding:10px 14px; background:#ffffff; border-bottom:1px solid #d9deea; font-size:13px; color:#334155; }
  #status { color:#64748b; }
  #wrap { padding:18px; }
  #diagram-card { background:#ffffff; border:1px solid #d9deea; border-radius:14px; box-shadow:0 4px 18px rgba(15,23,42,.06); padding:18px; overflow:auto; }
  #diagram { min-height:420px; display:flex; align-items:flex-start; justify-content:center; }
  #diagram svg { max-width:100%; height:auto; }
  .error { white-space:pre-wrap; color:#b91c1c; font-family:Consolas, monospace; }
</style>
</head>
<body>
  <div id="toolbar">
    <strong>WebView2 Mermaid Preview</strong>
    <span id="status">준비됨</span>
  </div>
  <div id="wrap">
    <div id="diagram-card"><div id="diagram"></div></div>
  </div>
<script src="__MERMAID_JS__"></script>
<script>
let currentSvg = '';
let currentError = '';

function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = ' · ' + msg;
}

function escapeHtml(str) {
  return (str || '').replace(/[&<>\"]/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[s]));
}

async function renderMermaid(code) {
  const target = document.getElementById('diagram');
  target.innerHTML = '';
  currentSvg = '';
  currentError = '';
  setStatus('렌더링 중...');
  try {
    mermaid.initialize({ startOnLoad: false, securityLevel: 'loose' });
    const id = 'm' + Date.now();
    const result = await mermaid.render(id, code);
    currentSvg = result.svg || '';
    target.innerHTML = currentSvg;
    setStatus('렌더 완료');
    return { ok: true, svg: currentSvg };
  } catch (err) {
    currentError = String(err && err.message ? err.message : err);
    target.innerHTML = '<div class="error">' + escapeHtml(currentError) + '</div>';
    setStatus('렌더 오류');
    return { ok: false, error: currentError };
  }
}

function getSvg() {
  return currentSvg || '';
}

function getError() {
  return currentError || '';
}
</script>
</body>
</html>
"""


class PreviewWindow:
    def __init__(self, log_func):
        self.log = log_func
        self.window = None
        self.last_code = ""
        self.process: Optional[subprocess.Popen] = None
        self.seq = 0
        self.bridge_dir: Optional[Path] = None
        self.request_path: Optional[Path] = None
        self.response_path: Optional[Path] = None
        self.ready_path: Optional[Path] = None
        self.error_path: Optional[Path] = None
        self._allocate_bridge_dir()

    def available(self) -> bool:
        return webview is not None

    def _allocate_bridge_dir(self) -> None:
        self.bridge_dir = Path(tempfile.mkdtemp(prefix="mermaid_preview_bridge_"))
        self.request_path = self.bridge_dir / "request.json"
        self.response_path = self.bridge_dir / "response.json"
        self.ready_path = self.bridge_dir / "ready.flag"
        self.error_path = self.bridge_dir / "error.txt"

    def _cleanup_bridge_files(self) -> None:
        for path in [self.request_path, self.response_path, self.ready_path, self.error_path]:
            if path is not None:
                path.unlink(missing_ok=True)

    def _helper_command(self, html_url: str) -> List[str]:
        script_path = Path(__file__).resolve()
        if getattr(sys, 'frozen', False):
            return [sys.executable, '--preview-helper', str(self.bridge_dir), html_url]
        return [sys.executable, str(script_path), '--preview-helper', str(self.bridge_dir), html_url]

    def _process_alive(self) -> bool:
        return self.process is not None and self.process.poll() is None

    def is_open(self) -> bool:
        return self._process_alive()

    def reset_runtime_state(self) -> None:
        self.window = None
        self.process = None
        self.seq = 0

    def close(self) -> None:
        if self._process_alive():
            self.process.terminate()
            try:
                self.process.wait(timeout=3)
            except Exception:
                self.process.kill()
        self.reset_runtime_state()
        if self.bridge_dir is not None and self.bridge_dir.exists():
            shutil.rmtree(self.bridge_dir, ignore_errors=True)
        self.bridge_dir = None
        self.request_path = None
        self.response_path = None
        self.ready_path = None
        self.error_path = None

    def show(self, html_url: str = MERMAID_JS_CDN) -> None:
        if webview is None:
            raise RuntimeError('pywebview 가 설치되지 않았습니다.')
        if self._process_alive():
            self.window = True
            return

        if self.bridge_dir is None:
            self._allocate_bridge_dir()
        self._cleanup_bridge_files()
        cmd = self._helper_command(html_url)
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        self.process = subprocess.Popen(cmd, creationflags=creationflags)

        deadline = time.time() + 15
        while time.time() < deadline:
            if self.ready_path is not None and self.ready_path.exists():
                self.window = True
                self.log('WebView2 미리보기 창을 열었습니다.')
                return
            if self.error_path is not None and self.error_path.exists():
                error = self.error_path.read_text(encoding='utf-8', errors='ignore')
                self.reset_runtime_state()
                raise RuntimeError(error or '미리보기 보조 프로세스 시작 실패')
            if self.process is not None and self.process.poll() is not None:
                raise RuntimeError('미리보기 보조 프로세스가 바로 종료되었습니다.')
            if self.process is not None and self.process.poll() is not None:
                self.reset_runtime_state()
                self.reset_runtime_state()
                self.reset_runtime_state()
                raise RuntimeError('Preview helper exited before the window became ready.')
            time.sleep(0.2)
        raise RuntimeError('미리보기 창 준비 시간이 초과되었습니다.')

    def render(self, code: str) -> Dict[str, Any]:
        self.last_code = code
        if not self._process_alive():
            self.reset_runtime_state()
            raise RuntimeError('미리보기 창이 아직 열리지 않았습니다.')
        self.seq += 1
        payload = {
            'seq': self.seq,
            'code': code,
            'sent_at': time.time(),
        }
        if self.request_path is None or self.response_path is None:
            raise RuntimeError('誘몃━蹂닿린 蹂댁“ ?뚯씪 寃쎈줈媛 珥덇린??섏? ?딆븯?듬땲??')
        self.request_path.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')

        deadline = time.time() + 20
        while time.time() < deadline:
            if self.response_path is not None and self.response_path.exists():
                try:
                    data = json.loads(self.response_path.read_text(encoding='utf-8'))
                except Exception:
                    data = {}
                if data.get('seq') == self.seq:
                    return data
            if not self._process_alive():
                self.reset_runtime_state()
                raise RuntimeError('미리보기 보조 프로세스가 종료되었습니다.')
            time.sleep(0.25)
        raise RuntimeError('미리보기 렌더 응답 대기 시간이 초과되었습니다.')

    def get_svg(self) -> str:
        if self.response_path is None or not self.response_path.exists():
            return ''
        try:
            data = json.loads(self.response_path.read_text(encoding='utf-8'))
        except Exception:
            return ''
        svg = data.get('svg')
        return svg if isinstance(svg, str) else ''


def _run_preview_helper(bridge_dir: Path, html_url: str) -> int:
    if webview is None:
        bridge_dir.mkdir(parents=True, exist_ok=True)
        (bridge_dir / 'error.txt').write_text('pywebview 가 설치되지 않았습니다.', encoding='utf-8')
        return 1

    bridge_dir.mkdir(parents=True, exist_ok=True)
    request_path = bridge_dir / 'request.json'
    response_path = bridge_dir / 'response.json'
    ready_path = bridge_dir / 'ready.flag'
    error_path = bridge_dir / 'error.txt'
    html = PREVIEW_HTML_TEMPLATE.replace('__MERMAID_JS__', html_url)

    holder: Dict[str, Any] = {'window': None}

    def watcher() -> None:
        ready_path.write_text('ready', encoding='utf-8')
        last_seq = -1
        while True:
            try:
                if request_path.exists():
                    payload = json.loads(request_path.read_text(encoding='utf-8'))
                    seq = int(payload.get('seq', -1))
                    if seq != last_seq:
                        last_seq = seq
                        code = str(payload.get('code') or '')
                        window = holder.get('window')
                        if window is not None:
                            result = window.evaluate_js(f'renderMermaid({json.dumps(code, ensure_ascii=False)})')
                            svg = window.evaluate_js('getSvg()')
                        else:
                            result = {'ok': False, 'error': 'window not ready'}
                            svg = ''
                        response = {'seq': seq, 'ok': True, 'svg': svg if isinstance(svg, str) else ''}
                        if isinstance(result, dict):
                            response.update(result)
                            if 'svg' not in response or not isinstance(response.get('svg'), str):
                                response['svg'] = svg if isinstance(svg, str) else ''
                        elif isinstance(result, str):
                            response['raw'] = result
                        response_path.write_text(json.dumps(response, ensure_ascii=False), encoding='utf-8')
                time.sleep(0.25)
            except Exception as exc:
                error_path.write_text(str(exc), encoding='utf-8')
                time.sleep(0.5)

    try:
        holder['window'] = webview.create_window(
            title='Mermaid Preview',
            html=html,
            width=1280,
            height=860,
            text_select=True,
        )
        try:
            webview.start(watcher, gui='edgechromium', debug=False)
        except TypeError:
            webview.start(watcher, debug=False)
        return 0
    except Exception as exc:
        error_path.write_text(str(exc), encoding='utf-8')
        return 1


# ----------------------------- Mermaid parsing for Excel export -----------------------------
PREVIEW_HTML_TEMPLATE = r"""
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Mermaid Preview</title>
<style>
  html, body { margin:0; padding:0; background:#f5f7fb; font-family: Segoe UI, Malgun Gothic, Arial, sans-serif; }
  #toolbar { padding:10px 14px; background:#ffffff; border-bottom:1px solid #d9deea; font-size:13px; color:#334155; }
  #status { color:#64748b; margin-left:8px; }
  #wrap { padding:18px; }
  #diagram-card { background:#ffffff; border:1px solid #d9deea; border-radius:14px; box-shadow:0 4px 18px rgba(15,23,42,.06); padding:18px; overflow:auto; }
  #diagram { min-height:420px; display:flex; align-items:flex-start; justify-content:center; }
  #diagram svg { max-width:100%; height:auto; }
  .error { white-space:pre-wrap; color:#b91c1c; font-family:Consolas, monospace; width:100%; }
</style>
</head>
<body>
  <div id="toolbar">
    <strong>WebView2 Mermaid Preview</strong>
    <span id="status">준비 중</span>
  </div>
  <div id="wrap">
    <div id="diagram-card"><div id="diagram"></div></div>
  </div>
<script src="__MERMAID_JS__"></script>
<script>
let currentSvg = '';
let currentError = '';

function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = msg || '';
}

function escapeHtml(str) {
  return (str || '').replace(/[&<>\"]/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[s]));
}

function previewReady() {
  return typeof mermaid !== 'undefined' && typeof renderMermaid === 'function';
}

async function renderMermaid(code) {
  const target = document.getElementById('diagram');
  target.innerHTML = '';
  currentSvg = '';
  currentError = '';
  setStatus('렌더링 중');
  try {
    mermaid.initialize({ startOnLoad: false, securityLevel: 'loose' });
    const id = 'm' + Date.now();
    const result = await mermaid.render(id, code || '');
    currentSvg = result.svg || '';
    target.innerHTML = currentSvg;
    setStatus('렌더 완료');
    return { ok: true, svg: currentSvg, error: '' };
  } catch (err) {
    currentError = String(err && err.message ? err.message : err);
    target.innerHTML = '<div class="error">' + escapeHtml(currentError) + '</div>';
    setStatus('렌더 오류');
    return { ok: false, error: currentError, svg: '' };
  }
}

function getSvg() {
  return currentSvg || '';
}

function getError() {
  return currentError || '';
}
</script>
</body>
</html>
"""


class PreviewManager:
    def __init__(self, log_func):
        self.log = log_func
        self.process: Optional[subprocess.Popen] = None
        self.seq = 0
        self.last_code = ""
        self.last_svg = ""
        self.last_state = "closed"
        self.last_html_url = MERMAID_JS_CDN
        self.bridge_dir: Optional[Path] = None
        self.request_path: Optional[Path] = None
        self.response_path: Optional[Path] = None
        self.ready_path: Optional[Path] = None
        self.error_path: Optional[Path] = None
        self._recreate_bridge()

    def available(self) -> bool:
        return webview is not None

    def _recreate_bridge(self) -> None:
        if self.bridge_dir is not None and self.bridge_dir.exists():
            shutil.rmtree(self.bridge_dir, ignore_errors=True)
        self.bridge_dir = Path(tempfile.mkdtemp(prefix="mermaid_preview_bridge_"))
        self.request_path = self.bridge_dir / "request.json"
        self.response_path = self.bridge_dir / "response.json"
        self.ready_path = self.bridge_dir / "ready.flag"
        self.error_path = self.bridge_dir / "error.txt"

    def _helper_command(self, html_url: str) -> List[str]:
        script_path = Path(__file__).resolve()
        if getattr(sys, "frozen", False):
            return [sys.executable, "--preview-helper", str(self.bridge_dir), html_url]
        return [sys.executable, str(script_path), "--preview-helper", str(self.bridge_dir), html_url]

    def _sync_state(self) -> None:
        if self.process is not None and self.process.poll() is not None:
            self.process = None
            self.last_state = "closed"

    def is_open(self) -> bool:
        self._sync_state()
        return self.process is not None

    def _start_helper(self, html_url: str) -> str:
        if webview is None:
            raise RuntimeError("pywebview is not installed.")

        self._sync_state()
        self.last_html_url = html_url
        if self.is_open():
            return "reuse"

        self._recreate_bridge()
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self.process = subprocess.Popen(self._helper_command(html_url), creationflags=creationflags)
        self.last_state = "opening"

        deadline = time.time() + 20
        while time.time() < deadline:
            self._sync_state()
            if self.ready_path is not None and self.ready_path.exists():
                self.last_state = "open"
                return "created"
            if self.error_path is not None and self.error_path.exists():
                error = self.error_path.read_text(encoding="utf-8", errors="ignore").strip()
                self.close()
                raise RuntimeError(error or "Preview helper failed to start.")
            if not self.is_open():
                raise RuntimeError("Preview helper exited before the preview window became ready.")
            time.sleep(0.2)
        self.close()
        raise RuntimeError("Preview window did not become ready within the timeout.")

    def show_or_render(self, code: str, html_url: str) -> Dict[str, Any]:
        code = code or ""
        was_open = self.is_open()
        open_mode = self._start_helper(html_url)
        if open_mode == "created" and was_open:
            self.log("미리보기 창 재생성")
        elif open_mode == "created":
            self.log("미리보기 창 열림")

        self.last_code = code
        self.seq += 1
        if self.request_path is None or self.response_path is None:
            raise RuntimeError("Preview bridge files are not initialized.")

        payload = {"seq": self.seq, "code": code, "sent_at": time.time()}
        self.response_path.unlink(missing_ok=True)
        self.request_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        deadline = time.time() + 25
        while time.time() < deadline:
            self._sync_state()
            if self.response_path.exists():
                try:
                    response = json.loads(self.response_path.read_text(encoding="utf-8"))
                except Exception:
                    response = {}
                if response.get("seq") == self.seq:
                    svg = response.get("svg")
                    self.last_svg = svg if isinstance(svg, str) else ""
                    self.last_state = "open"
                    self.log("미리보기 창 갱신")
                    return response
            if self.error_path is not None and self.error_path.exists():
                error = self.error_path.read_text(encoding="utf-8", errors="ignore").strip()
                raise RuntimeError(error or "Preview render failed.")
            if not self.is_open():
                raise RuntimeError("Preview window closed while rendering.")
            time.sleep(0.2)
        raise RuntimeError("Preview render timed out.")

    def get_svg(self) -> str:
        return self.last_svg or ""

    def close(self) -> None:
        if self.process is not None and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=3)
            except Exception:
                self.process.kill()
        self.process = None
        self.last_state = "closed"
        self.last_svg = ""
        if self.bridge_dir is not None and self.bridge_dir.exists():
            shutil.rmtree(self.bridge_dir, ignore_errors=True)
        self.bridge_dir = None
        self.request_path = None
        self.response_path = None
        self.ready_path = None
        self.error_path = None


def _run_preview_helper(bridge_dir: Path, html_url: str) -> int:
    if webview is None:
        bridge_dir.mkdir(parents=True, exist_ok=True)
        (bridge_dir / "error.txt").write_text("pywebview is not installed.", encoding="utf-8")
        return 1

    bridge_dir.mkdir(parents=True, exist_ok=True)
    request_path = bridge_dir / "request.json"
    response_path = bridge_dir / "response.json"
    ready_path = bridge_dir / "ready.flag"
    error_path = bridge_dir / "error.txt"
    html = PREVIEW_HTML_TEMPLATE.replace("__MERMAID_JS__", html_url)
    holder: Dict[str, Any] = {"window": None}

    def watcher() -> None:
        window = None
        while window is None:
            window = holder.get("window")
            time.sleep(0.1)

        while True:
            try:
                if window.evaluate_js("previewReady()") is True:
                    break
            except Exception:
                pass
            time.sleep(0.2)

        ready_path.write_text("ready", encoding="utf-8")
        last_seq = -1

        while True:
            try:
                if request_path.exists():
                    payload = json.loads(request_path.read_text(encoding="utf-8"))
                    seq = int(payload.get("seq", -1))
                    if seq != last_seq:
                        last_seq = seq
                        code = str(payload.get("code") or "")
                        result = window.evaluate_js(f"renderMermaid({json.dumps(code, ensure_ascii=False)})")
                        svg = window.evaluate_js("getSvg()")
                        error = window.evaluate_js("getError()")
                        response: Dict[str, Any] = {
                            "seq": seq,
                            "ok": True,
                            "svg": svg if isinstance(svg, str) else "",
                            "error": error if isinstance(error, str) else "",
                        }
                        if isinstance(result, dict):
                            response.update(result)
                        response_path.write_text(json.dumps(response, ensure_ascii=False), encoding="utf-8")
                time.sleep(0.15)
            except Exception as exc:
                error_path.write_text(str(exc), encoding="utf-8")
                time.sleep(0.3)

    try:
        holder["window"] = webview.create_window(
            title="Mermaid Preview",
            html=html,
            width=1280,
            height=860,
            text_select=True,
        )
        try:
            webview.start(watcher, gui="edgechromium", debug=False)
        except TypeError:
            webview.start(watcher, debug=False)
        return 0
    except Exception as exc:
        error_path.write_text(str(exc), encoding="utf-8")
        return 1

@dataclass
class MermaidNode:
    node_id: str
    label: str
    shape_kind: str = "process"  # process, decision, terminator, database, subroutine, manual
    class_names: List[str] = field(default_factory=list)
    fill_color: str = "#EAF2FF"
    line_color: str = "#2F5AA8"
    text_color: str = "#14305E"


@dataclass
class MermaidEdge:
    source: str
    target: str
    label: str = ""


@dataclass
class ParsedDiagram:
    direction: str
    nodes: Dict[str, MermaidNode]
    edges: List[MermaidEdge]
    title: str = "Diagram"


SHAPE_PATTERN_ORDER: List[Tuple[str, str]] = [
    (r"^([A-Za-z][\w]*)\(\((.+)\)\)$", "terminator"),
    (r"^([A-Za-z][\w]*)\[\((.+)\)\]$", "database"),
    (r"^([A-Za-z][\w]*)\[\[(.+)\]\]$", "subprocess"),
    (r"^([A-Za-z][\w]*)\[\/(.+)\/\]$", "manual"),
    (r"^([A-Za-z][\w]*)\[\\(.+)\\\]$", "manual"),
    (r"^([A-Za-z][\w]*)\{\{(.+)\}\}$", "decision"),
    (r"^([A-Za-z][\w]*)\{(.+)\}$", "decision"),
    (r"^([A-Za-z][\w]*)\[(.+)\]$", "process"),
    (r"^([A-Za-z][\w]*)\((.+)\)$", "terminator"),
]

EDGE_RE = re.compile(
    r"^([A-Za-z][\w]*)\s*([-]+(?:\.|x|o)?[-]*>)\s*(?:\|([^|]+)\|\s*)?([A-Za-z][\w]*)\s*$"
)
CLASS_ASSIGN_RE = re.compile(r"^class\s+([^;]+?)\s+([A-Za-z_][\w-]*)\s*;?$")
CLASS_DEF_RE = re.compile(r"^classDef\s+([A-Za-z_][\w-]*)\s+(.+?)\s*;?$")
INIT_COMMENT_RE = re.compile(r"^%%\{.*?\}%%\s*$")


def infer_shape_kind(expr: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    expr = expr.strip()
    for pattern, kind in SHAPE_PATTERN_ORDER:
        m = re.match(pattern, expr)
        if m:
            node_id = m.group(1)
            label = m.group(2).strip()
            return node_id, label, kind
    bare = re.match(r"^([A-Za-z][\w]*)$", expr)
    if bare:
        return bare.group(1), bare.group(1), "process"
    return None, None, None


def parse_style_map(style_blob: str) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for piece in style_blob.split(","):
        if ":" in piece:
            key, value = piece.split(":", 1)
            mapping[key.strip()] = value.strip()
    return mapping


def parse_mermaid_flowchart(code: str, theme_name: str) -> ParsedDiagram:
    lines = [line.rstrip() for line in code.splitlines() if line.strip()]
    direction = "TD"
    nodes: Dict[str, MermaidNode] = {}
    edges: List[MermaidEdge] = []
    class_defs: Dict[str, Dict[str, str]] = {}
    class_assignments: Dict[str, List[str]] = {}
    title = "Diagram"
    theme = THEMES.get(theme_name, next(iter(THEMES.values())))

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("%%") and not INIT_COMMENT_RE.match(line):
            continue
        if INIT_COMMENT_RE.match(line):
            continue
        if line.lower().startswith("title "):
            title = line[6:].strip().strip('"')
            continue
        if line.lower().startswith("flowchart "):
            parts = line.split()
            if len(parts) >= 2:
                direction = parts[1].upper()
            continue
        class_def_match = CLASS_DEF_RE.match(line)
        if class_def_match:
            class_defs[class_def_match.group(1)] = parse_style_map(class_def_match.group(2))
            continue
        class_assign_match = CLASS_ASSIGN_RE.match(line)
        if class_assign_match:
            ids = [x.strip() for x in class_assign_match.group(1).split(",") if x.strip()]
            class_name = class_assign_match.group(2).strip()
            for node_id in ids:
                class_assignments.setdefault(node_id, []).append(class_name)
            continue

        edge_match = EDGE_RE.match(line)
        if edge_match:
            src, _arrow, edge_label, dst = edge_match.groups()
            if src not in nodes:
                nodes[src] = MermaidNode(node_id=src, label=src)
            if dst not in nodes:
                nodes[dst] = MermaidNode(node_id=dst, label=dst)
            edges.append(MermaidEdge(source=src, target=dst, label=(edge_label or "").strip()))
            continue

        if "-->" in line or "-.->" in line:
            line = re.sub(r"\|[^|]+\|", "", line)
            parts = re.split(r"[-.ox]+>", line)
            if len(parts) >= 2:
                left, right = parts[0].strip(), parts[-1].strip()
                src_id, src_label, src_kind = infer_shape_kind(left)
                dst_id, dst_label, dst_kind = infer_shape_kind(right)
                if src_id:
                    nodes[src_id] = nodes.get(src_id, MermaidNode(src_id, src_label or src_id, src_kind or "process"))
                    nodes[src_id].label = src_label or nodes[src_id].label
                    nodes[src_id].shape_kind = src_kind or nodes[src_id].shape_kind
                if dst_id:
                    nodes[dst_id] = nodes.get(dst_id, MermaidNode(dst_id, dst_label or dst_id, dst_kind or "process"))
                    nodes[dst_id].label = dst_label or nodes[dst_id].label
                    nodes[dst_id].shape_kind = dst_kind or nodes[dst_id].shape_kind
                if src_id and dst_id:
                    edges.append(MermaidEdge(src_id, dst_id, ""))
                continue

        node_id, label, kind = infer_shape_kind(line)
        if node_id:
            nodes[node_id] = MermaidNode(node_id=node_id, label=label or node_id, shape_kind=kind or "process")

    for node_id, node in nodes.items():
        assigned = class_assignments.get(node_id, [])
        node.class_names = assigned
        style = {}
        for cls in assigned:
            style.update(class_defs.get(cls, {}))
        node.fill_color = parse_color(style.get("fill", theme["primaryColor"]), theme["primaryColor"])
        node.line_color = parse_color(style.get("stroke", theme["primaryBorderColor"]), theme["primaryBorderColor"])
        node.text_color = parse_color(style.get("color", theme["primaryTextColor"]), theme["primaryTextColor"])
        if node.shape_kind == "decision" and not assigned:
            node.fill_color = "#FFF4DD"
            node.line_color = "#C98A1A"
            node.text_color = "#6F4A00"
        elif node.shape_kind == "terminator" and not assigned:
            node.fill_color = theme["secondaryColor"]

    return ParsedDiagram(direction=direction, nodes=nodes, edges=edges, title=title)


def compute_layout(diagram: ParsedDiagram) -> Dict[str, Tuple[float, float, float, float]]:
    node_ids = list(diagram.nodes.keys())
    if not node_ids:
        return {}

    incoming: Dict[str, int] = {nid: 0 for nid in node_ids}
    outgoing: Dict[str, List[str]] = {nid: [] for nid in node_ids}
    for edge in diagram.edges:
        if edge.target in incoming:
            incoming[edge.target] += 1
        outgoing.setdefault(edge.source, []).append(edge.target)

    roots = [nid for nid, cnt in incoming.items() if cnt == 0] or [node_ids[0]]
    depth: Dict[str, int] = {nid: 0 for nid in roots}
    bfs = list(roots)
    seen = set(roots)
    while bfs:
        current = bfs.pop(0)
        for nxt in outgoing.get(current, []):
            new_depth = depth[current] + 1
            if nxt not in depth or new_depth > depth[nxt]:
                depth[nxt] = new_depth
            if nxt not in seen:
                seen.add(nxt)
                bfs.append(nxt)
    for nid in node_ids:
        depth.setdefault(nid, 0)

    levels: Dict[int, List[str]] = {}
    for nid, d in depth.items():
        levels.setdefault(d, []).append(nid)

    for ids in levels.values():
        ids.sort()

    left_margin, top_margin = 30.0, 30.0
    box_w, box_h = 150.0, 60.0
    gap_x, gap_y = 80.0, 45.0
    positions: Dict[str, Tuple[float, float, float, float]] = {}
    direction = diagram.direction.upper()
    horizontal = direction in {"LR", "RL"}

    for level_index in sorted(levels):
        ids = levels[level_index]
        for offset, nid in enumerate(ids):
            if horizontal:
                x = left_margin + level_index * (box_w + gap_x)
                y = top_margin + offset * (box_h + gap_y)
            else:
                x = left_margin + offset * (box_w + gap_x)
                y = top_margin + level_index * (box_h + gap_y)
            positions[nid] = (x, y, box_w, box_h)
    return positions


# ----------------------------- Excel export -----------------------------
class ExcelExporter:
    def __init__(self, log_func):
        self.log = log_func

    def export_editable(self, mermaid_code: str, theme_name: str, path: Path) -> None:
        if win32 is None:
            raise RuntimeError("pywin32 가 설치되지 않았습니다.")
        diagram = parse_mermaid_flowchart(mermaid_code, theme_name)
        if not diagram.nodes:
            raise RuntimeError("Excel 내보내기를 위한 노드를 찾지 못했습니다.")
        layout = compute_layout(diagram)
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "Diagram"
            data_ws = wb.Worksheets.Add(After=ws)
            data_ws.Name = "Data"
            self._write_data_sheet(data_ws, diagram)
            self._draw_shapes(ws, diagram, layout)
            ensure_parent(path)
            wb.SaveAs(str(path))
            self.log(f"엑셀 도형 내보내기 완료: {path}")
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            excel.Quit()

    def _write_data_sheet(self, ws, diagram: ParsedDiagram) -> None:
        ws.Cells(1, 1).Value = "Type"
        ws.Cells(1, 2).Value = "Id"
        ws.Cells(1, 3).Value = "Label"
        ws.Cells(1, 4).Value = "From"
        ws.Cells(1, 5).Value = "To"
        ws.Cells(1, 6).Value = "EdgeLabel"
        row = 2
        for node in diagram.nodes.values():
            ws.Cells(row, 1).Value = "NODE"
            ws.Cells(row, 2).Value = node.node_id
            ws.Cells(row, 3).Value = node.label
            row += 1
        for edge in diagram.edges:
            ws.Cells(row, 1).Value = "EDGE"
            ws.Cells(row, 4).Value = edge.source
            ws.Cells(row, 5).Value = edge.target
            ws.Cells(row, 6).Value = edge.label
            row += 1
        ws.Columns("A:F").AutoFit()

    def _draw_shapes(self, ws, diagram: ParsedDiagram, layout: Dict[str, Tuple[float, float, float, float]]) -> None:
        ws.Cells.ClearFormats()
        ws.Range("A1").Value = diagram.title or "Diagram"
        ws.Range("A1").Font.Size = 20
        ws.Range("A1").Font.Bold = True
        shape_map = {}

        anchor_middle = self._office_const("msoAnchorMiddle")
        align_center = self._office_const("msoAlignCenter")
        connector_elbow = self._office_const("msoConnectorElbow")
        arrow_triangle = self._office_const("msoArrowheadTriangle")

        for node_id, node in diagram.nodes.items():
            left, top, width, height = layout[node_id]
            shape_type = self._shape_type_from_node(node.shape_kind)
            shp = ws.Shapes.AddShape(shape_type, left, top + 40, width, height)
            shp.Name = f"node_{node_id}"
            shp.TextFrame2.TextRange.Text = node.label
            shp.TextFrame2.TextRange.Font.Size = 11
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = hex_to_bgr_int(node.text_color)
            shp.Fill.ForeColor.RGB = hex_to_bgr_int(node.fill_color)
            shp.Line.ForeColor.RGB = hex_to_bgr_int(node.line_color)
            shp.Line.Weight = 1.5
            try:
                shp.TextFrame2.VerticalAnchor = anchor_middle
                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = align_center
            except Exception:
                pass
            shape_map[node_id] = shp

        for idx, edge in enumerate(diagram.edges, start=1):
            if edge.source not in shape_map or edge.target not in shape_map:
                continue
            connector = ws.Shapes.AddConnector(connector_elbow, 0, 0, 40, 40)
            connector.Name = f"edge_{idx}"
            connector.Line.ForeColor.RGB = hex_to_bgr_int("#5B6472")
            connector.Line.EndArrowheadStyle = arrow_triangle
            connector.ConnectorFormat.BeginConnect(shape_map[edge.source], 3)
            connector.ConnectorFormat.EndConnect(shape_map[edge.target], 1)
            connector.RerouteConnections()
            if edge.label:
                self._add_edge_label(ws, connector, edge.label)

        ws.Range("A:Z").ColumnWidth = 3
        ws.Range("1:120").RowHeight = 18

    def _add_edge_label(self, ws, connector, text: str) -> None:
        try:
            x = (connector.Left + connector.Width / 2)
            y = (connector.Top + connector.Height / 2)
            label = ws.Shapes.AddTextbox(self._office_const("msoTextOrientationHorizontal"), x - 30, y - 10, 60, 20)
            label.TextFrame2.TextRange.Text = text
            label.TextFrame2.TextRange.Font.Size = 9
            label.Fill.Visible = False
            label.Line.Visible = False
        except Exception:
            pass

    def _office_const(self, name: str) -> int:
        try:
            return int(getattr(win32.constants, name))
        except Exception:
            return OFFICE_CONST_FALLBACKS[name]

    def _shape_type_from_node(self, shape_kind: str) -> int:
        mapping = {
            "process": self._office_const("msoShapeRoundedRectangle"),
            "decision": self._office_const("msoShapeDiamond"),
            "terminator": self._office_const("msoShapeFlowchartTerminator"),
            "database": self._office_const("msoShapeFlowchartStoredData"),
            "manual": self._office_const("msoShapeParallelogram"),
        }
        return mapping.get(shape_kind, self._office_const("msoShapeRoundedRectangle"))


# ----------------------------- Main App -----------------------------
class MermaidDesignerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1450x920")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.preview_manager = PreviewManager(self.log)
        self.excel_exporter = ExcelExporter(self.log)
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.title_var = tk.StringVar(value="Generated Mermaid")
        self.theme_var = tk.StringVar(value=list(THEMES.keys())[0])
        self.diagram_type_var = tk.StringVar(value="flowchart")
        self.flow_direction_var = tk.StringVar(value="TB")
        self.swimlane_lanes_var = tk.StringVar(value="요청부서, 검토부서, 승인부서, 운영부서")
        self.generation_note_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="준비")

        self.base_url_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["base_url"])
        self.model_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["model"])
        self.api_key_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["api_key"])
        self.credential_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["gpt_oss_credential"])
        self.user_id_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["gpt_oss_user_id"])
        self.user_type_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["gpt_oss_user_type"])
        self.system_name_var = tk.StringVar(value=HARD_CODED_LLM_CONFIG["gpt_oss_system_name"])
        self.timeout_var = tk.StringVar(value=str(HARD_CODED_LLM_CONFIG["timeout_sec"]))
        self.mermaid_js_var = tk.StringVar(value=MERMAID_JS_CDN)
        self.sample_index_by_type: Dict[str, int] = {key: 0 for key in TYPE_SAMPLES}

        self._build_ui()
        self._load_saved_settings(silent=True)
        self._drain_log_queue()
        self._show_startup_hints()

    def _build_ui(self) -> None:
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)

        header = ttk.Frame(self.root, padding=10)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.status_var, foreground="#334155").grid(row=0, column=1, sticky="e")

        main = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        main.grid(row=1, column=0, sticky="nsew")

        left = ttk.Frame(main, padding=(10, 0, 8, 8))
        right = ttk.Frame(main, padding=(8, 0, 10, 8))
        main.add(left, weight=2)
        main.add(right, weight=3)

        left.rowconfigure(3, weight=1)
        left.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # LLM settings
        llm_box = ttk.LabelFrame(left, text="LLM / GPT-OSS 연결 설정", padding=10)
        llm_box.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        for i in range(4):
            llm_box.columnconfigure(i, weight=1)

        self._entry_row(llm_box, 0, "URL", self.base_url_var, 52)
        self._entry_row(llm_box, 1, "Model", self.model_var, 24)
        self._entry_row(llm_box, 2, "API Key", self.api_key_var, 24, show="*")
        self._entry_row(llm_box, 3, "Credential", self.credential_var, 24, show="*")
        self._entry_row(llm_box, 4, "User ID", self.user_id_var, 20)
        self._entry_row(llm_box, 5, "User Type", self.user_type_var, 16)
        self._entry_row(llm_box, 6, "System Name", self.system_name_var, 24)
        self._entry_row(llm_box, 7, "Timeout(sec)", self.timeout_var, 10)
        self._entry_row(llm_box, 8, "Mermaid JS URL", self.mermaid_js_var, 52)

        btns = ttk.Frame(llm_box)
        btns.grid(row=9, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        for i in range(5):
            btns.columnconfigure(i, weight=1)
        ttk.Button(btns, text="하드코딩값 로드", command=self.load_hardcoded_settings).grid(row=0, column=0, padx=3, sticky="ew")
        ttk.Button(btns, text="환경변수 로드", command=self.load_env_settings).grid(row=0, column=1, padx=3, sticky="ew")
        ttk.Button(btns, text="설정 저장", command=self.save_settings).grid(row=0, column=2, padx=3, sticky="ew")
        ttk.Button(btns, text="설정 불러오기", command=self._load_saved_settings).grid(row=0, column=3, padx=3, sticky="ew")
        ttk.Button(btns, text="현재값 적용", command=self.apply_settings).grid(row=0, column=4, padx=3, sticky="ew")

        req_box = ttk.LabelFrame(left, text="자연어 입력", padding=10)
        req_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        req_box.columnconfigure(1, weight=1)
        ttk.Label(req_box, text="제목").grid(row=0, column=0, sticky="w")
        ttk.Entry(req_box, textvariable=self.title_var).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(req_box, text="테마").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(req_box, textvariable=self.theme_var, values=list(THEMES.keys()), state="readonly").grid(row=1, column=1, sticky="w", padx=(6, 0), pady=(6, 0))
        ttk.Label(req_box, text="Diagram Type").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(req_box, textvariable=self.diagram_type_var, values=DIAGRAM_TYPES, state="readonly").grid(row=2, column=1, sticky="w", padx=(6, 0), pady=(6, 0))
        ttk.Label(req_box, text="Flow Direction").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(req_box, textvariable=self.flow_direction_var, values=FLOW_DIRECTIONS, state="readonly").grid(row=3, column=1, sticky="w", padx=(6, 0), pady=(6, 0))
        ttk.Label(req_box, text="Swimlane Lanes").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(req_box, textvariable=self.swimlane_lanes_var).grid(row=4, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))
        self.prompt_text = tk.Text(req_box, height=10, wrap="word")
        self.prompt_text.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.prompt_text.insert("1.0", DEFAULT_PROMPT)

        action_box = ttk.Frame(left)
        action_box.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        for i in range(5):
            action_box.columnconfigure(i, weight=1)
        ttk.Button(action_box, text="생성", command=self.generate_from_prompt).grid(row=0, column=0, padx=3, sticky="ew")
        ttk.Button(action_box, text="코드 개선", command=self.improve_current_code).grid(row=0, column=1, padx=3, sticky="ew")
        ttk.Button(action_box, text="미리보기 창 열기", command=self.open_preview).grid(row=0, column=2, padx=3, sticky="ew")
        ttk.Button(action_box, text="미리보기 갱신", command=self.refresh_preview).grid(row=0, column=3, padx=3, sticky="ew")
        ttk.Button(action_box, text="샘플 로드", command=self.load_template).grid(row=0, column=4, padx=3, sticky="ew")

        editor_box = ttk.LabelFrame(left, text="Mermaid 코드", padding=10)
        editor_box.grid(row=3, column=0, sticky="nsew")
        editor_box.rowconfigure(0, weight=1)
        editor_box.columnconfigure(0, weight=1)
        self.code_text = tk.Text(editor_box, wrap="none", undo=True)
        self.code_text.grid(row=0, column=0, sticky="nsew")
        self.code_text.insert("1.0", DEFAULT_TEMPLATE)
        xscroll = ttk.Scrollbar(editor_box, orient="horizontal", command=self.code_text.xview)
        yscroll = ttk.Scrollbar(editor_box, orient="vertical", command=self.code_text.yview)
        self.code_text.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        xscroll.grid(row=1, column=0, sticky="ew")
        yscroll.grid(row=0, column=1, sticky="ns")

        export_box = ttk.LabelFrame(right, text="내보내기 / 결과", padding=10)
        export_box.grid(row=0, column=0, sticky="ew")
        for i in range(5):
            export_box.columnconfigure(i, weight=1)
        ttk.Button(export_box, text="SVG 저장", command=self.save_svg).grid(row=0, column=0, padx=3, sticky="ew")
        ttk.Button(export_box, text="PNG 저장", command=self.save_png).grid(row=0, column=1, padx=3, sticky="ew")
        ttk.Button(export_box, text="Excel 도형 내보내기", command=self.export_excel_shapes).grid(row=0, column=2, padx=3, sticky="ew")
        ttk.Button(export_box, text="코드 저장", command=self.save_mermaid_code).grid(row=0, column=3, padx=3, sticky="ew")
        ttk.Button(export_box, text="코드 불러오기", command=self.load_mermaid_code).grid(row=0, column=4, padx=3, sticky="ew")

        tabs = ttk.Notebook(right)
        tabs.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        summary_tab = ttk.Frame(tabs)
        log_tab = ttk.Frame(tabs)
        help_tab = ttk.Frame(tabs)
        tabs.add(summary_tab, text="요약")
        tabs.add(log_tab, text="로그")
        tabs.add(help_tab, text="안내")

        summary_tab.rowconfigure(0, weight=1)
        summary_tab.columnconfigure(0, weight=1)
        self.summary_text = tk.Text(summary_tab, wrap="word")
        self.summary_text.grid(row=0, column=0, sticky="nsew")
        self.summary_text.insert(
            "1.0",
            "- 생성된 Mermaid를 보고용으로 다듬고\n"
            "- WebView2 미리보기로 검토한 다음\n"
            "- SVG/PNG 또는 Excel 편집형 도형으로 내보낼 수 있습니다.\n\n"
            "Excel 내보내기는 flowchart 위주로 설계되어 있습니다.\n"
        )

        log_tab.rowconfigure(0, weight=1)
        log_tab.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_tab, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        help_tab.rowconfigure(0, weight=1)
        help_tab.columnconfigure(0, weight=1)
        help_text = tk.Text(help_tab, wrap="word")
        help_text.grid(row=0, column=0, sticky="nsew")
        help_text.insert(
            "1.0",
            "1) 자연어 입력 후 [생성]\n"
            "2) Mermaid 코드 수정\n"
            "3) [미리보기 창 열기] -> [미리보기 갱신]\n"
            "4) SVG / PNG / Excel 도형 내보내기\n\n"
            "주의:\n"
            "- 미리보기는 pywebview + WebView2 계열 환경을 사용합니다.\n"
            "- Excel 도형 내보내기는 Windows + Excel + pywin32가 필요합니다.\n"
            "- Excel 편집형 내보내기는 flowchart 문법을 우선 지원합니다.\n"
        )
        help_text.configure(state="disabled")

    def _entry_row(self, parent, row: int, label: str, variable: tk.StringVar, width: int, show: Optional[str] = None) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        entry = ttk.Entry(parent, textvariable=variable, width=width, show=show)
        entry.grid(row=row, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=2)

    def _show_startup_hints(self) -> None:
        self.log("Mermaid Report Designer 시작")
        self.log("자연어 → Mermaid 생성 → Tkinter 수정 → WebView2 미리보기 → SVG/PNG/Excel 도형 내보내기")
        if webview is None:
            self.log("pywebview 미설치: WebView2 미리보기를 쓰려면 pip install pywebview")
        else:
            self.log("pywebview 준비됨")
        if cairosvg is None:
            self.log("cairosvg 미설치: PNG 저장은 pip install cairosvg 필요")
        else:
            self.log("cairosvg 준비됨")
        if win32 is None:
            self.log("pywin32 미설치: Excel 도형 내보내기는 pip install pywin32 필요")
        else:
            self.log("pywin32 준비됨")

    def current_config(self) -> Dict[str, Any]:
        return {
            "base_url": self.base_url_var.get().strip(),
            "model": self.model_var.get().strip(),
            "api_key": self.api_key_var.get().strip(),
            "gpt_oss_credential": self.credential_var.get().strip(),
            "gpt_oss_user_id": self.user_id_var.get().strip(),
            "gpt_oss_user_type": self.user_type_var.get().strip(),
            "gpt_oss_system_name": self.system_name_var.get().strip(),
            "timeout_sec": safe_int(self.timeout_var.get(), DEFAULT_TIMEOUT),
        }

    def apply_settings(self) -> None:
        self.status_var.set("설정 적용됨")
        self.log("현재 설정을 적용했습니다.")

    def load_hardcoded_settings(self) -> None:
        self.base_url_var.set(HARD_CODED_LLM_CONFIG["base_url"])
        self.model_var.set(HARD_CODED_LLM_CONFIG["model"])
        self.api_key_var.set(HARD_CODED_LLM_CONFIG["api_key"])
        self.credential_var.set(HARD_CODED_LLM_CONFIG["gpt_oss_credential"])
        self.user_id_var.set(HARD_CODED_LLM_CONFIG["gpt_oss_user_id"])
        self.user_type_var.set(HARD_CODED_LLM_CONFIG["gpt_oss_user_type"])
        self.system_name_var.set(HARD_CODED_LLM_CONFIG["gpt_oss_system_name"])
        self.timeout_var.set(str(HARD_CODED_LLM_CONFIG["timeout_sec"]))
        self.log("하드코딩 설정을 로드했습니다.")

    def load_env_settings(self) -> None:
        mapping = {
            self.base_url_var: ["MERMAID_LLM_BASE_URL", "GPT_OSS_BASE_URL"],
            self.model_var: ["MERMAID_LLM_MODEL", "GPT_OSS_MODEL"],
            self.api_key_var: ["MERMAID_LLM_API_KEY", "GPT_OSS_API_KEY"],
            self.credential_var: ["GPT_OSS_CREDENTIAL"],
            self.user_id_var: ["GPT_OSS_USER_ID"],
            self.user_type_var: ["GPT_OSS_USER_TYPE"],
            self.system_name_var: ["GPT_OSS_SYSTEM_NAME"],
        }
        for var, names in mapping.items():
            for name in names:
                value = os.environ.get(name)
                if value:
                    var.set(value)
                    break
        timeout = os.environ.get("MERMAID_LLM_TIMEOUT") or os.environ.get("GPT_OSS_TIMEOUT")
        if timeout:
            self.timeout_var.set(timeout)
        js_url = os.environ.get("MERMAID_JS_URL")
        if js_url:
            self.mermaid_js_var.set(js_url)
        self.log("환경변수에서 설정을 불러왔습니다.")

    def save_settings(self) -> None:
        data = {
            "config": self.current_config(),
            "mermaid_js_url": self.mermaid_js_var.get().strip(),
            "theme": self.theme_var.get().strip(),
            "diagram_type": self.diagram_type_var.get().strip(),
            "flow_direction": self.flow_direction_var.get().strip(),
            "swimlane_lanes": self.swimlane_lanes_var.get().strip(),
        }
        SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log(f"설정 저장: {SETTINGS_PATH}")

    def _load_saved_settings(self, silent: bool = False) -> None:
        if not SETTINGS_PATH.exists():
            if not silent:
                self.log("저장된 설정 파일이 없습니다.")
            return
        try:
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            config = data.get("config") or {}
            self.base_url_var.set(config.get("base_url", self.base_url_var.get()))
            self.model_var.set(config.get("model", self.model_var.get()))
            self.api_key_var.set(config.get("api_key", self.api_key_var.get()))
            self.credential_var.set(config.get("gpt_oss_credential", self.credential_var.get()))
            self.user_id_var.set(config.get("gpt_oss_user_id", self.user_id_var.get()))
            self.user_type_var.set(config.get("gpt_oss_user_type", self.user_type_var.get()))
            self.system_name_var.set(config.get("gpt_oss_system_name", self.system_name_var.get()))
            self.timeout_var.set(str(config.get("timeout_sec", self.timeout_var.get())))
            self.mermaid_js_var.set(data.get("mermaid_js_url", self.mermaid_js_var.get()))
            theme = data.get("theme")
            if theme in THEMES:
                self.theme_var.set(theme)
            diagram_type = data.get("diagram_type")
            if diagram_type in DIAGRAM_TYPES:
                self.diagram_type_var.set(diagram_type)
            flow_direction = data.get("flow_direction")
            if flow_direction in FLOW_DIRECTIONS:
                self.flow_direction_var.set(flow_direction)
            self.swimlane_lanes_var.set(data.get("swimlane_lanes", self.swimlane_lanes_var.get()))
            if not silent:
                self.log("저장된 설정을 불러왔습니다.")
        except Exception as exc:
            if not silent:
                self.log(f"설정 불러오기 실패: {exc}")

    def log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        logging.info(message)
        self.log_queue.put(f"[{timestamp}] {message}\n")

    def _drain_log_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg)
                self.log_text.see("end")
        except queue.Empty:
            pass
        self.root.after(150, self._drain_log_queue)

    def get_current_code(self) -> str:
        return self.code_text.get("1.0", "end").strip()

    def set_current_code(self, code: str) -> None:
        self.code_text.delete("1.0", "end")
        self.code_text.insert("1.0", code.strip() + "\n")

    def generate_from_prompt(self) -> None:
        prompt = self.prompt_text.get("1.0", "end").strip()
        if not prompt:
            messagebox.showwarning(APP_TITLE, "자연어 요구사항을 입력해 주세요.")
            return

        def task():
            try:
                self.status_var.set("생성 중...")
                client = LLMClient(self.current_config())
                result = client.generate_mermaid(prompt, self.theme_var.get())
                self.root.after(0, lambda: self.title_var.set(result["title"]))
                themed = self.apply_theme_to_code(result["mermaid"], self.theme_var.get())
                self.root.after(0, lambda: self.set_current_code(themed))
                self.root.after(0, lambda: self.summary_text.delete("1.0", "end"))
                self.root.after(0, lambda: self.summary_text.insert("1.0", self.build_summary(themed)))
                self.log("LLM Mermaid 생성 완료")
                self.status_var.set("생성 완료")
            except Exception as exc:
                self.log(f"생성 실패: {exc}")
                self.status_var.set("생성 실패")
                self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"생성 실패\n{exc}"))

        threading.Thread(target=task, daemon=True).start()

    def improve_current_code(self) -> None:
        code = self.get_current_code()
        if not code:
            messagebox.showwarning(APP_TITLE, "수정할 Mermaid 코드가 없습니다.")
            return
        prompt = (
            "아래 Mermaid flowchart 코드를 보고 보고용으로 더 정돈해줘. "
            "노드 이름을 간결하게 다듬고, classDef와 class를 적절히 넣고, 흐름이 더 읽기 좋게 정리해줘.\n\n"
            f"현재 코드:\n{code}"
        )
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", prompt)
        self.generate_from_prompt()

    def apply_theme_to_code(self, code: str, theme_name: str) -> str:
        theme = THEMES.get(theme_name, next(iter(THEMES.values())))
        init_line = "%%{init: {'theme': 'base'}}%%"
        lines = [line.rstrip() for line in code.splitlines() if line.strip()]
        lines = [line for line in lines if not INIT_COMMENT_RE.match(line)]
        base_class = (
            f"classDef default fill:{theme['primaryColor']},stroke:{theme['primaryBorderColor']},"
            f"color:{theme['primaryTextColor']},stroke-width:1.6px;"
        )
        if not any(line.startswith("classDef default") for line in lines):
            lines.append(base_class)
        if not any(line.startswith("%%{init:") for line in lines):
            lines.insert(0, init_line)
        return "\n".join(lines)

    def build_summary(self, code: str) -> str:
        diagram = parse_mermaid_flowchart(code, self.theme_var.get())
        node_count = len(diagram.nodes)
        edge_count = len(diagram.edges)
        decisions = sum(1 for n in diagram.nodes.values() if n.shape_kind == "decision")
        starts = [n.label for n in diagram.nodes.values() if n.shape_kind == "terminator"]
        lines = [
            f"제목: {self.title_var.get().strip() or diagram.title}",
            f"방향: {diagram.direction}",
            f"노드 수: {node_count}",
            f"연결 수: {edge_count}",
            f"의사결정 노드 수: {decisions}",
        ]
        if starts:
            lines.append("주요 시작/종료 노드: " + ", ".join(starts[:5]))
        if edge_count:
            preview = [f"{e.source} → {e.target}" + (f" ({e.label})" if e.label else "") for e in diagram.edges[:8]]
            lines.append("")
            lines.append("핵심 흐름:")
            lines.extend(f"- {item}" for item in preview)
        lines.append("")
        lines.append("엑셀 편집형 내보내기는 flowchart 위주로 지원합니다.")
        return "\n".join(lines)

    def open_preview(self) -> None:
        try:
            code = self.get_current_code()
            if code:
                result = self.preview_manager.show_or_render(code, self.mermaid_js_var.get().strip() or MERMAID_JS_CDN)
                if not result.get("ok"):
                    error = result.get("error") or "Preview render failed."
                    messagebox.showwarning(APP_TITLE, f"미리보기 렌더 오류\n{error}")
                self.log("誘몃━蹂닿린 珥덇린 ?뚮뜑 ?꾨즺")
            self.status_var.set("미리보기 창 열림")
        except Exception as exc:
            self.log(f"미리보기 창 오류: {exc}")
            messagebox.showerror(APP_TITLE, f"미리보기 창 열기 실패\n{exc}")

    def refresh_preview(self) -> None:
        code = self.get_current_code()
        if not code:
            messagebox.showwarning(APP_TITLE, "미리볼 Mermaid 코드가 없습니다.")
            return
        try:
            result = self.preview_manager.show_or_render(code, self.mermaid_js_var.get().strip() or MERMAID_JS_CDN)
            if result.get("ok"):
                self.log("미리보기 갱신 완료")
                self.status_var.set("미리보기 완료")
            else:
                error = result.get("error") or "알 수 없는 렌더 오류"
                self.log(f"미리보기 렌더 오류: {error}")
                self.status_var.set("렌더 오류")
                messagebox.showwarning(APP_TITLE, f"렌더 오류\n{error}")
        except Exception as exc:
            self.log(f"미리보기 실패: {exc}")
            self.status_var.set("렌더 실패")
            messagebox.showerror(APP_TITLE, f"미리보기 실패\n{exc}")

    def save_svg(self) -> None:
        code = self.get_current_code()
        if not code:
            messagebox.showwarning(APP_TITLE, "저장할 Mermaid 코드가 없습니다.")
            return
        try:
            self.preview_manager.show_or_render(code, self.mermaid_js_var.get().strip() or MERMAID_JS_CDN)
            svg = self.preview_manager.get_svg()
            if not svg:
                raise RuntimeError("SVG 렌더 결과를 가져오지 못했습니다.")
            path = filedialog.asksaveasfilename(
                title="SVG 저장",
                defaultextension=".svg",
                initialfile=slugify_filename(self.title_var.get()) + ".svg",
                filetypes=[("SVG file", "*.svg")],
            )
            if not path:
                return
            Path(path).write_text(svg, encoding="utf-8")
            self.log(f"SVG 저장 완료: {path}")
        except Exception as exc:
            self.log(f"SVG 저장 실패: {exc}")
            messagebox.showerror(APP_TITLE, f"SVG 저장 실패\n{exc}")

    def save_png(self) -> None:
        if cairosvg is None:
            messagebox.showerror(APP_TITLE, "PNG 저장은 cairosvg 설치가 필요합니다.\npip install cairosvg")
            return
        code = self.get_current_code()
        if not code:
            messagebox.showwarning(APP_TITLE, "저장할 Mermaid 코드가 없습니다.")
            return
        try:
            self.preview_manager.show_or_render(code, self.mermaid_js_var.get().strip() or MERMAID_JS_CDN)
            svg = self.preview_manager.get_svg()
            if not svg:
                raise RuntimeError("PNG 변환용 SVG를 가져오지 못했습니다.")
            path = filedialog.asksaveasfilename(
                title="PNG 저장",
                defaultextension=".png",
                initialfile=slugify_filename(self.title_var.get()) + ".png",
                filetypes=[("PNG file", "*.png")],
            )
            if not path:
                return
            cairosvg.svg2png(bytestring=svg.encode("utf-8"), write_to=path)
            self.log(f"PNG 저장 완료: {path}")
        except Exception as exc:
            self.log(f"PNG 저장 실패: {exc}")
            messagebox.showerror(APP_TITLE, f"PNG 저장 실패\n{exc}")

    def export_excel_shapes(self) -> None:
        code = self.get_current_code()
        if not code:
            messagebox.showwarning(APP_TITLE, "엑셀로 내보낼 Mermaid 코드가 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            title="Excel 도형 내보내기",
            defaultextension=".xlsx",
            initialfile=slugify_filename(self.title_var.get()) + ".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return

        def task():
            try:
                self.status_var.set("Excel 내보내기 중...")
                self.excel_exporter.export_editable(code, self.theme_var.get(), Path(path))
                self.status_var.set("Excel 내보내기 완료")
            except Exception as exc:
                self.log(f"Excel 내보내기 실패: {exc}")
                self.status_var.set("Excel 내보내기 실패")
                self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"Excel 내보내기 실패\n{exc}"))

        threading.Thread(target=task, daemon=True).start()

    def save_mermaid_code(self) -> None:
        code = self.get_current_code()
        path = filedialog.asksaveasfilename(
            title="Mermaid 코드 저장",
            defaultextension=".mmd",
            initialfile=slugify_filename(self.title_var.get()) + ".mmd",
            filetypes=[("Mermaid file", "*.mmd"), ("Text file", "*.txt")],
        )
        if not path:
            return
        Path(path).write_text(code, encoding="utf-8")
        self.log(f"코드 저장 완료: {path}")

    def load_mermaid_code(self) -> None:
        path = filedialog.askopenfilename(
            title="Mermaid 코드 불러오기",
            filetypes=[("Mermaid file", "*.mmd;*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        code = Path(path).read_text(encoding="utf-8")
        self.set_current_code(code)
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", self.build_summary(code))
        self.log(f"코드 불러오기 완료: {path}")

    def load_template(self) -> None:
        self.title_var.set("샘플 프로세스")
        self.set_current_code(self.apply_theme_to_code(DEFAULT_TEMPLATE, self.theme_var.get()))
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", self.build_summary(self.get_current_code()))
        self.log("샘플 Mermaid 코드를 불러왔습니다.")
    def on_close(self) -> None:
        try:
            self.preview_manager.close()
        except Exception as exc:
            append_exception_log("on_close", exc)
        self.root.destroy()


def _app_selected_diagram_type(app: "MermaidDesignerApp") -> str:
    selected = app.diagram_type_var.get().strip() or "flowchart"
    if selected == "자동 추천":
        return recommend_diagram_type(app.prompt_text.get("1.0", "end").strip())
    return selected


def _app_apply_theme_to_code(self, code: str, theme_name: str) -> str:
    theme = THEMES.get(theme_name, next(iter(THEMES.values())))
    lines = [line.rstrip() for line in code.splitlines() if line.strip()]
    lines = [line for line in lines if not INIT_COMMENT_RE.match(line)]
    detected_type = detect_mermaid_diagram_type("\n".join(lines))
    init_line = (
        "%%{init: {'theme': 'base', 'themeVariables': {"
        f" 'fontFamily': '{theme['fontFamily']}',"
        f" 'primaryColor': '{theme['primaryColor']}',"
        f" 'primaryBorderColor': '{theme['primaryBorderColor']}',"
        f" 'primaryTextColor': '{theme['primaryTextColor']}',"
        f" 'lineColor': '{theme['lineColor']}',"
        f" 'secondaryColor': '{theme['secondaryColor']}',"
        f" 'tertiaryColor': '{theme['tertiaryColor']}'"
        " }}}%%"
    )
    if not any(line.startswith("%%{init:") for line in lines):
        lines.insert(0, init_line)
    if detected_type in {"flowchart", "swimlane"}:
        presets = [
            f"classDef default fill:{theme['primaryColor']},stroke:{theme['primaryBorderColor']},color:{theme['primaryTextColor']},stroke-width:1.6px;",
            "classDef start fill:#DFF5E8,stroke:#2F7D5A,color:#174734,stroke-width:2px;",
            "classDef process fill:#EAF2FF,stroke:#2F5AA8,color:#14305E,stroke-width:1.6px;",
            "classDef decision fill:#FFF4DD,stroke:#C98A1A,color:#6F4A00,stroke-width:1.8px;",
            "classDef data fill:#EEF7FF,stroke:#4682B4,color:#1E4C74,stroke-width:1.6px;",
            "classDef risk fill:#FFE7E7,stroke:#C23B3B,color:#7A1F1F,stroke-width:1.8px;",
            "classDef kpi fill:#EAFBF6,stroke:#1D8F6A,color:#14513E,stroke-width:1.8px;",
        ]
        existing = {line.split()[1] for line in lines if line.startswith("classDef ")}
        for preset in presets:
            name = preset.split()[1]
            if name not in existing:
                lines.append(preset)
    return "\n".join(lines)


def _app_build_summary(self, code: str) -> str:
    detected_type = detect_mermaid_diagram_type(code)
    selected_type = self.diagram_type_var.get().strip()
    display_type = detected_type if selected_type == "자동 추천" else selected_type
    title = self.title_var.get().strip() or "Generated Mermaid"
    description = self.generation_note_var.get().strip()
    lines = [
        f"제목: {title}",
        f"유형: {display_type}",
        f"라인 수: {len([line for line in code.splitlines() if line.strip()])}",
    ]
    if detected_type in {"flowchart", "swimlane"}:
        diagram = parse_mermaid_flowchart(code, self.theme_var.get())
        lines.append(f"방향: {diagram.direction}")
        lines.append(f"노드 수: {len(diagram.nodes)}")
        lines.append(f"연결 수: {len(diagram.edges)}")
    if detected_type == "gantt":
        milestone_count = sum(1 for line in code.splitlines() if "milestone" in line.lower())
        lines.append(f"마일스톤 수: {milestone_count}")
    if detected_type == "sequenceDiagram":
        participant_count = sum(1 for line in code.splitlines() if line.strip().startswith("participant "))
        lines.append(f"participant 수: {participant_count}")
    if detected_type == "journey":
        section_count = sum(1 for line in code.splitlines() if line.strip().startswith("section "))
        lines.append(f"section 수: {section_count}")
    if description:
        lines.append("")
        lines.append("설명:")
        lines.append(description)
    lines.append("")
    lines.append("Before / After:")
    lines.append("- Before: 기본 다이어그램 생성 및 단순 요약")
    lines.append("- After: 테마, 범례, 상태 배지, 풍부한 샘플, 보고서형 구조 반영")
    lines.append("")
    lines.append("Excel 편집형 내보내기는 flowchart 위주로 지원합니다.")
    return "\n".join(lines)


def _app_generate_from_prompt(self) -> None:
    prompt = self.prompt_text.get("1.0", "end").strip()
    if not prompt:
        messagebox.showwarning(APP_TITLE, "자연어 요구사항을 입력하세요.")
        return

    selected_type = self.diagram_type_var.get().strip() or "flowchart"
    recommended_type = recommend_diagram_type(prompt) if selected_type == "자동 추천" else selected_type

    def task():
        try:
            self.status_var.set("생성 중...")
            client = LLMClient(self.current_config())
            result = client.generate_mermaid(prompt, self.theme_var.get(), selected_type, self.flow_direction_var.get())
            themed = self.apply_theme_to_code(result["mermaid"], self.theme_var.get())
            description = result.get("description", "").strip()
            self.root.after(0, lambda: self.title_var.set(result["title"]))
            self.root.after(0, lambda: self.generation_note_var.set(description or f"{recommended_type} 유형으로 생성했습니다."))
            self.root.after(0, lambda: self.set_current_code(themed))
            self.root.after(0, lambda: self.summary_text.delete("1.0", "end"))
            self.root.after(0, lambda: self.summary_text.insert("1.0", self.build_summary(themed)))
            self.log(f"LLM Mermaid 생성 완료: {recommended_type}")
            self.status_var.set("생성 완료")
        except Exception as exc:
            self.log(f"생성 실패: {exc}")
            self.status_var.set("생성 실패")
            self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"생성 실패\n{exc}"))

    threading.Thread(target=task, daemon=True).start()


def _app_improve_current_code(self) -> None:
    code = self.get_current_code()
    if not code:
        messagebox.showwarning(APP_TITLE, "수정할 Mermaid 코드가 없습니다.")
        return
    prompt = (
        f"아래 Mermaid {detect_mermaid_diagram_type(code)} 코드를 더 정돈해줘. "
        f"현재 선택 타입은 {self.diagram_type_var.get()} 이고 방향은 {self.flow_direction_var.get()}야. "
        "문법 오류 가능성을 줄이고 보고용으로 정리해줘.\n\n"
        f"현재 코드:\n{code}"
    )
    self.prompt_text.delete("1.0", "end")
    self.prompt_text.insert("1.0", prompt)
    self.generate_from_prompt()


def _app_load_template(self) -> None:
    diagram_type = _app_selected_diagram_type(self)
    if diagram_type == "swimlane":
        current = make_swimlane_from_input(self.swimlane_lanes_var.get(), self.flow_direction_var.get())
        samples = [current] + TYPE_SAMPLES["swimlane"][1:]
    elif diagram_type == "flowchart":
        samples = [build_flowchart_template(self.flow_direction_var.get()), build_flowchart_template_sample2(self.flow_direction_var.get())]
    elif diagram_type == "org chart":
        samples = [build_org_chart_template(self.flow_direction_var.get()), build_org_chart_template_sample2(self.flow_direction_var.get())]
    else:
        samples = TYPE_SAMPLES.get(diagram_type, TYPE_SAMPLES["flowchart"])

    index = self.sample_index_by_type.get(diagram_type, 0) % len(samples)
    self.sample_index_by_type[diagram_type] = index + 1
    code = self.apply_theme_to_code(samples[index], self.theme_var.get())
    self.generation_note_var.set(f"{diagram_type} 샘플 {index + 1} / {len(samples)}")
    self.title_var.set(f"{diagram_type} sample {index + 1}")
    self.set_current_code(code)
    self.summary_text.delete("1.0", "end")
    self.summary_text.insert("1.0", self.build_summary(code))
    self.log(f"샘플 Mermaid 코드 로드: {diagram_type} #{index + 1}")


def _app_apply_current_theme(self) -> None:
    code = self.get_current_code()
    if not code:
        code = get_template_for_type(_app_selected_diagram_type(self), self.flow_direction_var.get(), self.swimlane_lanes_var.get())
    themed = self.apply_theme_to_code(code, self.theme_var.get())
    self.set_current_code(themed)
    self.summary_text.delete("1.0", "end")
    self.summary_text.insert("1.0", self.build_summary(themed))
    self.log(f"테마 적용: {self.theme_var.get()}")


def _app_load_gallery_sample(self, gallery_name: str) -> None:
    sample = GALLERY_SAMPLES[gallery_name]
    if sample["theme"] in THEMES:
        self.theme_var.set(sample["theme"])
    self.diagram_type_var.set(sample["diagram_type"])
    self.flow_direction_var.set(sample["direction"])
    self.title_var.set(sample["title"])
    self.generation_note_var.set(sample["description"])
    code = self.apply_theme_to_code(sample["code"], self.theme_var.get())
    self.set_current_code(code)
    self.summary_text.delete("1.0", "end")
    self.summary_text.insert("1.0", self.build_summary(code))
    self.log(f"샘플 템플릿 갤러리 로드: {gallery_name}")


def _app_build_ui_enhancement(self) -> None:
    try:
        left_parent = self.prompt_text.master.master
    except Exception:
        return
    gallery_box = ttk.LabelFrame(left_parent, text="샘플 템플릿 갤러리", padding=10)
    gallery_box.grid(row=4, column=0, sticky="ew", pady=(8, 8))
    for i in range(4):
        gallery_box.columnconfigure(i, weight=1)
    ttk.Button(gallery_box, text="임원보고형", command=lambda: self.load_gallery_sample("임원보고형")).grid(row=0, column=0, padx=3, sticky="ew")
    ttk.Button(gallery_box, text="업무 프로세스형", command=lambda: self.load_gallery_sample("업무 프로세스형")).grid(row=0, column=1, padx=3, sticky="ew")
    ttk.Button(gallery_box, text="시스템 연동형", command=lambda: self.load_gallery_sample("시스템 연동형")).grid(row=0, column=2, padx=3, sticky="ew")
    ttk.Button(gallery_box, text="일정관리형", command=lambda: self.load_gallery_sample("일정관리형")).grid(row=0, column=3, padx=3, sticky="ew")
    ttk.Button(gallery_box, text="코드에 테마 적용", command=self.apply_current_theme).grid(row=1, column=0, columnspan=4, pady=(8, 0), sticky="ew")


_original_build_ui = MermaidDesignerApp._build_ui


def _app_build_ui_with_gallery(self) -> None:
    _original_build_ui(self)
    _app_build_ui_enhancement(self)


MermaidDesignerApp.apply_theme_to_code = _app_apply_theme_to_code
MermaidDesignerApp.build_summary = _app_build_summary
MermaidDesignerApp.generate_from_prompt = _app_generate_from_prompt
MermaidDesignerApp.improve_current_code = _app_improve_current_code
MermaidDesignerApp.load_template = _app_load_template
MermaidDesignerApp.apply_current_theme = _app_apply_current_theme
MermaidDesignerApp.load_gallery_sample = _app_load_gallery_sample
MermaidDesignerApp._build_ui = _app_build_ui_with_gallery


@dataclass
class ExportTheme:
    name: str
    lane_fill: str
    lane_line: str
    connector: str
    title_fill: str
    title_line: str
    semantic_colors: Dict[str, Tuple[str, str, str]]


@dataclass
class ExportNode:
    node_id: str
    label: str
    shape_kind: str
    semantic_type: str
    lane: str = ""
    role: str = ""
    width: float = 150.0
    height: float = 60.0
    class_names: List[str] = field(default_factory=list)
    is_junction: bool = False
    junction_role: str = ""


@dataclass
class ExportEdge:
    source: str
    target: str
    label: str = ""


@dataclass
class ExportFlowchart:
    title: str
    direction: str
    nodes: Dict[str, ExportNode]
    edges: List[ExportEdge]
    lanes: List[str]


@dataclass
class GanttTask:
    section: str
    name: str
    owner: str
    start: Any
    end: Any
    duration_days: int
    progress: float
    milestone: bool
    dependency: str = ""


@dataclass
class GanttDiagram:
    title: str
    tasks: List[GanttTask]


@dataclass
class SequenceParticipant:
    key: str
    label: str


@dataclass
class SequenceMessage:
    source: str
    target: str
    label: str
    kind: str = "message"
    block_type: str = ""


@dataclass
class SequenceDiagramData:
    title: str
    participants: List[SequenceParticipant]
    messages: List[SequenceMessage]


EXPORT_THEMES: Dict[str, ExportTheme] = {
    "corporate blue": ExportTheme(
        name="corporate blue",
        lane_fill="#F5F9FF",
        lane_line="#C8D8F0",
        connector="#58739B",
        title_fill="#E8F0FF",
        title_line="#9FB7E9",
        semantic_colors={
            "start_end": ("#E3F4EA", "#2F7D5A", "#174734"),
            "process": ("#EAF2FF", "#2F5AA8", "#14305E"),
            "decision": ("#FFF4DD", "#C98A1A", "#6F4A00"),
            "review": ("#F4F0FF", "#7A4BB7", "#4C2D73"),
            "data": ("#EEF7FF", "#4B8AC9", "#1E4C74"),
            "risk": ("#FFE7E7", "#C23B3B", "#7A1F1F"),
            "done": ("#E6F7F2", "#1E8E73", "#125244"),
        },
    ),
    "clean gray": ExportTheme(
        name="clean gray",
        lane_fill="#FAFAFA",
        lane_line="#D7DADF",
        connector="#70757D",
        title_fill="#F1F3F4",
        title_line="#C7CCD1",
        semantic_colors={
            "start_end": ("#EEF3F2", "#62756D", "#31403A"),
            "process": ("#F5F6F7", "#7B848E", "#2F343A"),
            "decision": ("#FAF3E8", "#B2874E", "#5F4428"),
            "review": ("#F3F0F7", "#847097", "#43384E"),
            "data": ("#EFF4F7", "#62859A", "#234353"),
            "risk": ("#FAEDEE", "#AF6466", "#652C2D"),
            "done": ("#EAF4EE", "#5D8C6E", "#2F4A38"),
        },
    ),
    "executive mixed pastel": ExportTheme(
        name="executive mixed pastel",
        lane_fill="#FFFDF8",
        lane_line="#E4DCCF",
        connector="#7C7C8A",
        title_fill="#FFF6E9",
        title_line="#E1CDAA",
        semantic_colors={
            "start_end": ("#E4F6EC", "#4E9D7B", "#1F5140"),
            "process": ("#EAF0FF", "#6C86D9", "#223A7A"),
            "decision": ("#FFF0D9", "#D59C42", "#6D4A13"),
            "review": ("#F5ECFF", "#9670D6", "#56348D"),
            "data": ("#EAF8F5", "#56A6A0", "#1E5A5A"),
            "risk": ("#FFECEC", "#D06C78", "#722B34"),
            "done": ("#EEF7E8", "#7BA857", "#36551F"),
        },
    ),
}


def get_export_theme(theme_name: str) -> ExportTheme:
    mapping = {
        "Executive Blue": "corporate blue",
        "Warm Gray": "clean gray",
        "Calm Green": "executive mixed pastel",
        "Clean Gray": "clean gray",
        "Corporate Blue Plus": "corporate blue",
        "Modern Dark": "executive mixed pastel",
    }
    return EXPORT_THEMES[mapping.get(theme_name, "corporate blue")]


def normalize_label_for_shape(text: str, max_chars: int = 18) -> Tuple[str, float, float]:
    clean = compact_text(text)
    if not clean:
        return "", 120.0, 52.0
    words = clean.split()
    lines: List[str] = []
    current = ""
    for word in words:
        if len(current + " " + word) <= max_chars:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    if len(words) <= 1 and len(clean) > max_chars:
        chunk_size = max(8, max_chars - 4)
        lines = [clean[i:i + chunk_size] for i in range(0, len(clean), chunk_size)]
    elif len(lines) == 1 and len(lines[0]) > max_chars:
        lines = [clean[i:i + max_chars] for i in range(0, len(clean), max_chars)]
    wrapped = "\n".join(lines[:4])
    width = min(260.0, max(150.0, max(len(line) for line in lines[:4]) * 10.5 + 34.0))
    height = max(58.0, 24.0 + len(lines[:4]) * 20.0)
    return wrapped, width, height


def infer_semantic_type(label: str, shape_kind: str, class_names: Sequence[str]) -> str:
    joined = f"{label} {' '.join(class_names)}".lower()
    if "junction" in joined:
        return "junction"
    if shape_kind == "decision" or any(word in joined for word in ["승인 여부", "결정", "판단", "여부", "분기"]):
        return "decision"
    if shape_kind == "terminator" or any(word in joined for word in ["시작", "종료", "완료", "오픈", "전환"]):
        return "start_end"
    if any(word in joined for word in ["문서", "보고서", "report", "document"]):
        return "document"
    if shape_kind in {"database"} or any(word in joined for word in ["db", "데이터", "저장", "storage", "database"]):
        return "database"
    if shape_kind in {"manual"} or any(word in joined for word in ["입력", "출력", "산출물", "작성", "전달"]):
        return "manual_io"
    if shape_kind in {"subprocess"} or any(word in joined for word in ["서브", "subprocess", "재검토", "재처리", "재수행"]):
        return "subprocess"
    if any(word in joined for word in ["리스크", "위험", "이슈", "예외", "반려"]):
        return "risk"
    if any(word in joined for word in ["검토", "승인", "리뷰", "분석", "확인"]):
        return "review"
    if any(word in joined for word in ["공유", "배포", "완료"]):
        return "done"
    return "process"


def infer_excel_shape_name(node: ExportNode) -> str:
    if node.is_junction or node.semantic_type == "junction":
        return "msoShapeOval"
    if node.semantic_type == "start_end":
        return "msoShapeFlowchartTerminator"
    if node.semantic_type == "decision":
        return "msoShapeDiamond"
    if node.semantic_type == "document":
        return "msoShapeFlowchartDocument"
    if node.semantic_type == "database":
        return "msoShapeFlowchartStoredData"
    if node.semantic_type == "manual_io":
        return "msoShapeParallelogram"
    if node.semantic_type == "data":
        if any(word in node.label.lower() for word in ["문서", "보고", "report", "document"]):
            return "msoShapeFlowchartDocument"
        if any(word in node.label.lower() for word in ["db", "데이터", "저장"]):
            return "msoShapeFlowchartStoredData"
        return "msoShapeParallelogram"
    if node.semantic_type == "review":
        return "msoShapeFlowchartPreparation"
    if node.semantic_type == "risk":
        return "msoShapeHexagon"
    if node.semantic_type == "done":
        return "msoShapeRoundedRectangle"
    if node.semantic_type == "subprocess" or any(word in node.label.lower() for word in ["하위", "subprocess", "재검토"]):
        return "msoShapeFlowchartPredefinedProcess"
    return "msoShapeRoundedRectangle"


def parse_flowchart_for_export(code: str) -> ExportFlowchart:
    direction = "TB"
    title = "Diagram"
    nodes: Dict[str, ExportNode] = {}
    edges: List[ExportEdge] = []
    class_assignments: Dict[str, List[str]] = {}
    current_lane = ""
    lanes: List[str] = []
    for raw in code.splitlines():
        line = raw.strip()
        if not line or INIT_COMMENT_RE.match(line):
            continue
        low = line.lower()
        if low.startswith("title "):
            title = line[6:].strip().strip('"')
            continue
        if low.startswith("flowchart ") or low.startswith("graph "):
            parts = line.split()
            if len(parts) >= 2:
                direction = parts[1].upper()
            continue
        if low.startswith("subgraph "):
            lane_match = re.search(r"\[(.+?)\]", line)
            current_lane = lane_match.group(1).strip() if lane_match else line.split(None, 1)[1].strip()
            if current_lane not in lanes:
                lanes.append(current_lane)
            continue
        if low == "end":
            current_lane = ""
            continue
        class_assign_match = CLASS_ASSIGN_RE.match(line)
        if class_assign_match:
            ids = [x.strip() for x in class_assign_match.group(1).split(",") if x.strip()]
            class_name = class_assign_match.group(2).strip()
            for node_id in ids:
                class_assignments.setdefault(node_id, []).append(class_name)
            continue
        edge_match = EDGE_RE.match(line)
        if edge_match:
            src, _arrow, edge_label, dst = edge_match.groups()
            nodes.setdefault(src, ExportNode(src, src, "process", "process", lane=current_lane))
            nodes.setdefault(dst, ExportNode(dst, dst, "process", "process", lane=current_lane))
            edges.append(ExportEdge(src, dst, (edge_label or "").strip()))
            continue
        if "-->" in line or "-.->" in line:
            label_match = re.search(r"\|([^|]+)\|", line)
            edge_label = label_match.group(1).strip() if label_match else ""
            cleaned = re.sub(r"\|[^|]+\|", "", line)
            left, right = re.split(r"[-.ox]+>", cleaned, maxsplit=1)
            src_id, src_label, src_kind = infer_shape_kind(left.strip())
            dst_id, dst_label, dst_kind = infer_shape_kind(right.strip())
            if src_id:
                node = nodes.get(src_id) or ExportNode(src_id, src_label or src_id, src_kind or "process", "process", lane=current_lane)
                if src_label and not (src_label == src_id and src_kind == "process" and node.label != node.node_id):
                    node.label = src_label
                if src_kind and not (src_kind == "process" and node.shape_kind != "process"):
                    node.shape_kind = src_kind
                if current_lane and not node.lane:
                    node.lane = current_lane
                nodes[src_id] = node
            if dst_id:
                node = nodes.get(dst_id) or ExportNode(dst_id, dst_label or dst_id, dst_kind or "process", "process", lane=current_lane)
                if dst_label and not (dst_label == dst_id and dst_kind == "process" and node.label != node.node_id):
                    node.label = dst_label
                if dst_kind and not (dst_kind == "process" and node.shape_kind != "process"):
                    node.shape_kind = dst_kind
                if current_lane and not node.lane:
                    node.lane = current_lane
                nodes[dst_id] = node
            if src_id and dst_id:
                edges.append(ExportEdge(src_id, dst_id, edge_label))
            continue
        node_id, label, kind = infer_shape_kind(line)
        if node_id:
            nodes[node_id] = ExportNode(node_id, label or node_id, kind or "process", "process", lane=current_lane)
    for node in nodes.values():
        node.class_names = class_assignments.get(node.node_id, [])
        node.semantic_type = infer_semantic_type(node.label, node.shape_kind, node.class_names)
        wrapped, width, height = normalize_label_for_shape(node.label)
        node.label = wrapped
        node.width = width
        node.height = height
        if node.shape_kind == "decision":
            node.width = max(node.width, 168.0)
            node.height = max(node.height, 88.0)
        elif node.shape_kind == "terminator":
            node.width = max(node.width, 165.0)
            node.height = max(node.height, 64.0)
        elif node.shape_kind == "database":
            node.width = max(node.width, 168.0)
            node.height = max(node.height, 68.0)
        elif node.shape_kind == "manual":
            node.width = max(node.width, 162.0)
            node.height = max(node.height, 64.0)
        elif node.semantic_type in {"review", "subprocess"}:
            node.width = max(node.width, 170.0)
            node.height = max(node.height, 64.0)
        else:
            node.width = max(node.width, 160.0)
            node.height = max(node.height, 62.0)
        if node.lane and node.lane not in lanes:
            lanes.append(node.lane)
    return ExportFlowchart(title=title, direction=direction, nodes=nodes, edges=edges, lanes=lanes)


def augment_flowchart_with_junctions(diagram: ExportFlowchart) -> ExportFlowchart:
    nodes = {
        node_id: ExportNode(
            node_id=node.node_id,
            label=node.label,
            shape_kind=node.shape_kind,
            semantic_type=node.semantic_type,
            lane=node.lane,
            role=node.role,
            width=node.width,
            height=node.height,
            class_names=list(node.class_names),
            is_junction=node.is_junction,
            junction_role=node.junction_role,
        )
        for node_id, node in diagram.nodes.items()
    }
    edges = [ExportEdge(edge.source, edge.target, edge.label) for edge in diagram.edges]
    lanes = list(diagram.lanes)

    outgoing: Dict[str, List[ExportEdge]] = {}
    incoming: Dict[str, List[ExportEdge]] = {}
    for edge in edges:
        outgoing.setdefault(edge.source, []).append(edge)
        incoming.setdefault(edge.target, []).append(edge)

    split_nodes = [
        node_id
        for node_id, node in nodes.items()
        if len(outgoing.get(node_id, [])) > 1 and (node.semantic_type == "decision" or len(outgoing.get(node_id, [])) >= 3)
    ]
    merge_nodes = [node_id for node_id in nodes if len(incoming.get(node_id, [])) > 1]

    next_index = 1
    for node_id in split_nodes:
        junction_id = f"J_SPLIT_{next_index}"
        next_index += 1
        base_node = nodes[node_id]
        nodes[junction_id] = ExportNode(
            node_id=junction_id,
            label="",
            shape_kind="junction",
            semantic_type="junction",
            lane=base_node.lane,
            width=12.0,
            height=12.0,
            is_junction=True,
            junction_role="split",
        )
        original_edges = list(outgoing.get(node_id, []))
        edges = [edge for edge in edges if edge.source != node_id]
        edges.append(ExportEdge(node_id, junction_id, ""))
        for edge in original_edges:
            edges.append(ExportEdge(junction_id, edge.target, edge.label))

    outgoing = {}
    incoming = {}
    for edge in edges:
        outgoing.setdefault(edge.source, []).append(edge)
        incoming.setdefault(edge.target, []).append(edge)

    for node_id in merge_nodes:
        if node_id not in nodes:
            continue
        node_incoming = list(incoming.get(node_id, []))
        if len(node_incoming) <= 1:
            continue
        junction_id = f"J_MERGE_{next_index}"
        next_index += 1
        base_node = nodes[node_id]
        nodes[junction_id] = ExportNode(
            node_id=junction_id,
            label="",
            shape_kind="junction",
            semantic_type="junction",
            lane=base_node.lane,
            width=12.0,
            height=12.0,
            is_junction=True,
            junction_role="merge",
        )
        edges = [edge for edge in edges if edge.target != node_id]
        for edge in node_incoming:
            edges.append(ExportEdge(edge.source, junction_id, edge.label))
        edges.append(ExportEdge(junction_id, node_id, ""))

    return ExportFlowchart(title=diagram.title, direction=diagram.direction, nodes=nodes, edges=edges, lanes=lanes)


def parse_gantt_for_export(code: str) -> GanttDiagram:
    import datetime as dt

    title = "Gantt"
    section = "Tasks"
    tasks: List[GanttTask] = []
    ref_map: Dict[str, Any] = {}
    for raw in code.splitlines():
        line = raw.strip()
        if not line or line.startswith("%%"):
            continue
        low = line.lower()
        if low.startswith("title "):
            title = line[6:].strip()
            continue
        if low == "gantt" or low.startswith("dateformat") or low.startswith("excludes"):
            continue
        if low.startswith("section "):
            section = line[8:].strip()
            continue
        if ":" not in line:
            continue
        name_part, meta_part = line.split(":", 1)
        name_part = compact_text(name_part)
        meta_items = [compact_text(x) for x in meta_part.split(",") if compact_text(x)]
        owner = ""
        if "/" in name_part:
            name, owner = [compact_text(x) for x in name_part.split("/", 1)]
        else:
            name = name_part
        milestone = any("milestone" in item.lower() for item in meta_items)
        progress = 1.0 if any("done" in item.lower() for item in meta_items) else 0.55 if any("active" in item.lower() for item in meta_items) else 0.15
        ref_id = ""
        start_date = None
        duration_days = 1
        dependency = ""
        for item in meta_items:
            low_item = item.lower()
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", item):
                start_date = dt.datetime.strptime(item, "%Y-%m-%d").date()
            elif re.fullmatch(r"\d+d", low_item):
                duration_days = max(1, int(low_item[:-1]))
            elif low_item.startswith("after "):
                dependency = compact_text(item[6:])
            elif re.fullmatch(r"[A-Za-z_]\w*", item) and item.lower() not in {"done", "active", "milestone"}:
                ref_id = item
        if start_date is None and dependency and dependency in ref_map:
            dep_end = ref_map[dependency]
            start_date = dep_end + dt.timedelta(days=1)
        if start_date is None:
            start_date = dt.date.today()
        end_date = start_date if milestone else start_date + dt.timedelta(days=max(duration_days - 1, 0))
        task = GanttTask(section, name, owner, start_date, end_date, duration_days, progress, milestone, dependency=dependency)
        tasks.append(task)
        if ref_id:
            ref_map[ref_id] = end_date
    return GanttDiagram(title=title, tasks=tasks)


def parse_sequence_for_export(code: str) -> SequenceDiagramData:
    title = "Sequence"
    participants: List[SequenceParticipant] = []
    participant_map: Dict[str, SequenceParticipant] = {}
    messages: List[SequenceMessage] = []
    block_stack: List[str] = []
    msg_re = re.compile(r"^([A-Za-z_]\w*)\s*[-.]+>>?\s*([A-Za-z_]\w*)\s*:\s*(.+)$")
    for raw in code.splitlines():
        line = raw.strip()
        if not line or line.startswith("%%"):
            continue
        low = line.lower()
        if low.startswith("title "):
            title = line[6:].strip()
            continue
        if low == "sequencediagram":
            continue
        if low.startswith("participant "):
            body = line[len("participant "):].strip()
            if " as " in body:
                key, label = [compact_text(x) for x in body.split(" as ", 1)]
            else:
                key, label = body, body
            part = SequenceParticipant(key, label)
            participants.append(part)
            participant_map[key] = part
            continue
        if low.startswith(("alt ", "opt ", "loop ")):
            block_stack.append(line.split()[0].lower())
            messages.append(SequenceMessage("", "", line, kind="block_start", block_type=block_stack[-1]))
            continue
        if low == "end":
            if block_stack:
                messages.append(SequenceMessage("", "", "end", kind="block_end", block_type=block_stack.pop()))
            continue
        if low.startswith("note "):
            messages.append(SequenceMessage("", "", line, kind="note"))
            continue
        m = msg_re.match(line)
        if m:
            src, dst, label = m.groups()
            if src not in participant_map:
                participant_map[src] = SequenceParticipant(src, src)
                participants.append(participant_map[src])
            if dst not in participant_map:
                participant_map[dst] = SequenceParticipant(dst, dst)
                participants.append(participant_map[dst])
            messages.append(SequenceMessage(src, dst, compact_text(label), kind="message"))
    return SequenceDiagramData(title=title, participants=participants, messages=messages)


class ExcelExportStrategy:
    def __init__(self, manager: "ExcelExportManager"):
        self.manager = manager

    def export(self, workbook, code: str, theme_name: str, title: str) -> None:
        raise NotImplementedError


class ExcelExportHelper:
    def __init__(self, manager: "ExcelExportManager", workbook, worksheet, theme: ExportTheme):
        self.manager = manager
        self.workbook = workbook
        self.ws = worksheet
        self.theme = theme
        self.shape_map: Dict[str, Any] = {}

    def office_const(self, name: str, fallback: str = "msoShapeRoundedRectangle") -> int:
        try:
            return int(getattr(win32.constants, name))
        except Exception:
            return OFFICE_CONST_FALLBACKS.get(name, OFFICE_CONST_FALLBACKS.get(fallback, 1))

    def add_title(self, title: str, subtitle: str = "") -> None:
        self.ws.Range("A1").Value = title
        self.ws.Range("A1").Font.Size = 20
        self.ws.Range("A1").Font.Bold = True
        if subtitle:
            self.ws.Range("A2").Value = subtitle
            self.ws.Range("A2").Font.Size = 10
            self.ws.Range("A2").Font.Color = hex_to_bgr_int("#5F6368")

    def add_lane(self, name: str, left: float, top: float, width: float, height: float) -> Any:
        shape = self.ws.Shapes.AddShape(self.office_const("msoShapeRectangle", "msoShapeRectangle"), left, top, width, height)
        shape.TextFrame2.TextRange.Text = name
        shape.TextFrame2.TextRange.Font.Size = 11
        shape.TextFrame2.TextRange.Font.Bold = True
        shape.Fill.ForeColor.RGB = hex_to_bgr_int(self.theme.lane_fill)
        shape.Line.ForeColor.RGB = hex_to_bgr_int(self.theme.lane_line)
        try:
            shape.ZOrder(self.office_const("msoSendToBack", "msoShapeRectangle"))
        except Exception:
            pass
        return shape

    def add_node(self, node: ExportNode, left: float, top: float, width: float, height: float) -> Any:
        shape_name = infer_excel_shape_name(node)
        shape = self.ws.Shapes.AddShape(self.office_const(shape_name), left, top, width, height)
        if node.is_junction:
            fill, line, text = ("#FFFFFF", self.theme.connector, self.theme.connector)
        else:
            fill, line, text = self.theme.semantic_colors.get(node.semantic_type, self.theme.semantic_colors["process"])
        shape.Name = f"node_{node.node_id}"
        shape.TextFrame2.TextRange.Text = node.label
        shape.TextFrame2.TextRange.Font.Size = 10.5 if not node.is_junction else 1
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = hex_to_bgr_int(text)
        shape.Fill.ForeColor.RGB = hex_to_bgr_int(fill)
        shape.Line.ForeColor.RGB = hex_to_bgr_int(line)
        shape.Line.Weight = 1.6 if not node.is_junction else 1.2
        try:
            shape.TextFrame2.WordWrap = True
            shape.TextFrame2.AutoSize = False
            shape.TextFrame2.VerticalAnchor = self.office_const("msoAnchorMiddle", "msoShapeRectangle")
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = self.office_const("msoAlignCenter", "msoShapeRectangle")
        except Exception:
            pass
        self.shape_map[node.node_id] = shape
        return shape

    def connect_shapes(self, source_id: str, target_id: str, label: str = "", direction_hint: str = "TB") -> None:
        if source_id not in self.shape_map or target_id not in self.shape_map:
            return
        src = self.shape_map[source_id]
        dst = self.shape_map[target_id]
        src_mid_x = src.Left + src.Width / 2
        src_mid_y = src.Top + src.Height / 2
        dst_mid_x = dst.Left + dst.Width / 2
        dst_mid_y = dst.Top + dst.Height / 2
        dx = dst_mid_x - src_mid_x
        dy = dst_mid_y - src_mid_y
        if abs(dx) >= abs(dy):
            begin_site = 4 if dx >= 0 else 2
            end_site = 2 if dx >= 0 else 4
        else:
            begin_site = 3 if dy >= 0 else 1
            end_site = 1 if dy >= 0 else 3
        connector = self.ws.Shapes.AddConnector(self.office_const("msoConnectorElbow"), 0, 0, 10, 10)
        connector.Line.ForeColor.RGB = hex_to_bgr_int(self.theme.connector)
        connector.Line.Weight = 1.4
        connector.Line.EndArrowheadStyle = self.office_const("msoArrowheadTriangle", "msoShapeRectangle")
        connector.ConnectorFormat.BeginConnect(src, begin_site)
        connector.ConnectorFormat.EndConnect(dst, end_site)
        try:
            connector.RerouteConnections()
        except Exception:
            pass
        if label:
            self.add_label(label, (src_mid_x + dst_mid_x) / 2, (src_mid_y + dst_mid_y) / 2)

    def add_label(self, text: str, x: float, y: float) -> None:
        box = self.ws.Shapes.AddTextbox(self.office_const("msoTextOrientationHorizontal", "msoShapeRectangle"), x - 42, y - 10, 84, 20)
        box.TextFrame2.TextRange.Text = compact_text(text)
        box.TextFrame2.TextRange.Font.Size = 9
        box.Fill.Visible = False
        box.Line.Visible = False


def compute_professional_layout(diagram: ExportFlowchart) -> Tuple[Dict[str, Tuple[float, float, float, float]], Dict[str, Tuple[float, float, float, float]]]:
    nodes = diagram.nodes
    incoming = {nid: 0 for nid in nodes}
    outgoing: Dict[str, List[str]] = {nid: [] for nid in nodes}
    for edge in diagram.edges:
        if edge.target in incoming:
            incoming[edge.target] += 1
        outgoing.setdefault(edge.source, []).append(edge.target)
    roots = [nid for nid, cnt in incoming.items() if cnt == 0] or list(nodes.keys())[:1]
    levels = {nid: 0 for nid in roots}
    queue_nodes = list(roots)
    while queue_nodes:
        current = queue_nodes.pop(0)
        for nxt in outgoing.get(current, []):
            level = levels[current] + 1
            if nxt not in levels or level > levels[nxt]:
                levels[nxt] = level
            if nxt not in queue_nodes:
                queue_nodes.append(nxt)
    for nid in nodes:
        levels.setdefault(nid, 0)

    lane_names = diagram.lanes or ["Main"]
    lane_index = {lane: idx for idx, lane in enumerate(lane_names)}
    for node in nodes.values():
        if not node.lane:
            node.lane = lane_names[0]

    dirn = diagram.direction.upper()
    horizontal = dirn in {"LR", "RL"}
    margin_x, margin_y = 40.0, 80.0
    gap_x, gap_y = 70.0, 42.0
    lane_gap = 36.0
    lane_width = 280.0
    lane_height = 240.0
    positions: Dict[str, Tuple[float, float, float, float]] = {}
    lane_bounds: Dict[str, Tuple[float, float, float, float]] = {}

    grouped: Dict[Tuple[int, str], List[str]] = {}
    for nid, node in nodes.items():
        grouped.setdefault((levels[nid], node.lane), []).append(nid)
    for ids in grouped.values():
        ids.sort()

    if horizontal:
        max_level = max(levels.values()) if levels else 0
        for lane, idx in lane_index.items():
            lane_top = margin_y + idx * (lane_height + lane_gap)
            lane_left = margin_x
            lane_bounds[lane] = (lane_left, lane_top, (max_level + 1) * (lane_width + gap_x), lane_height)
        for (level, lane), ids in grouped.items():
            x_base = margin_x + level * (lane_width + gap_x) + 20
            y_base = lane_bounds[lane][1] + 54
            for idx, nid in enumerate(ids):
                node = nodes[nid]
                positions[nid] = (x_base, y_base + idx * (node.height + gap_y), node.width, node.height)
    else:
        max_level = max(levels.values()) if levels else 0
        for lane, idx in lane_index.items():
            lane_left = margin_x + idx * (lane_width + lane_gap)
            lane_top = margin_y
            lane_bounds[lane] = (lane_left, lane_top, lane_width, (max_level + 1) * (lane_height + gap_y))
        for (level, lane), ids in grouped.items():
            x_base = lane_bounds[lane][0] + 20
            y_base = margin_y + level * (lane_height + gap_y) + 40
            for idx, nid in enumerate(ids):
                node = nodes[nid]
                positions[nid] = (x_base + idx * (node.width + gap_x), y_base, node.width, node.height)
    return positions, lane_bounds


class FlowchartExportStrategy(ExcelExportStrategy):
    def export(self, workbook, code: str, theme_name: str, title: str) -> None:
        ws = workbook.Worksheets(1)
        ws.Name = "Flowchart"
        theme = get_export_theme(theme_name)
        diagram = parse_flowchart_for_export(code)
        helper = ExcelExportHelper(self.manager, workbook, ws, theme)
        helper.add_title(diagram.title or title, f"Theme: {theme.name}")
        positions, lane_bounds = compute_professional_layout(diagram)
        for lane, bounds in lane_bounds.items():
            helper.add_lane(lane, bounds[0], bounds[1], bounds[2], bounds[3])
        for node_id, node in diagram.nodes.items():
            left, top, width, height = positions[node_id]
            helper.add_node(node, left, top, width, height)
        for edge in diagram.edges:
            helper.connect_shapes(edge.source, edge.target, edge.label, diagram.direction)
        self._add_legend(helper, ws, theme)
        ws.Range("A:AZ").ColumnWidth = 2.4
        ws.Range("1:220").RowHeight = 18

    def _add_legend(self, helper: ExcelExportHelper, ws, theme: ExportTheme) -> None:
        left, top = 20.0, 12.0
        for idx, (semantic, colors) in enumerate(theme.semantic_colors.items()):
            box = ws.Shapes.AddShape(helper.office_const("msoShapeRoundedRectangle"), 760, top + idx * 26, 90, 18)
            box.Fill.ForeColor.RGB = hex_to_bgr_int(colors[0])
            box.Line.ForeColor.RGB = hex_to_bgr_int(colors[1])
            box.TextFrame2.TextRange.Text = semantic
            box.TextFrame2.TextRange.Font.Size = 8


class OrgChartExportStrategy(FlowchartExportStrategy):
    pass


class SwimlaneExportStrategy(FlowchartExportStrategy):
    pass


@dataclass
class FlowExportLayout:
    positions: Dict[str, Tuple[float, float, float, float]]
    lane_bounds: Dict[str, Tuple[float, float, float, float]]
    levels: Dict[str, int]
    orders: Dict[str, int]
    graph_bounds: Tuple[float, float, float, float]
    horizontal: bool


def _semantic_rank(node: ExportNode) -> int:
    ranking = {
        "start_end": 0,
        "process": 1,
        "review": 2,
        "data": 3,
        "decision": 4,
        "risk": 5,
        "done": 6,
    }
    return ranking.get(node.semantic_type, 9)


def _build_flow_metrics(diagram: ExportFlowchart) -> Tuple[Dict[str, List[str]], Dict[str, List[str]], Dict[str, int], Dict[str, int]]:
    incoming_counts = {nid: 0 for nid in diagram.nodes}
    outgoing_counts = {nid: 0 for nid in diagram.nodes}
    parents: Dict[str, List[str]] = {nid: [] for nid in diagram.nodes}
    children: Dict[str, List[str]] = {nid: [] for nid in diagram.nodes}
    for edge in diagram.edges:
        if edge.source in outgoing_counts:
            outgoing_counts[edge.source] += 1
            children[edge.source].append(edge.target)
        if edge.target in incoming_counts:
            incoming_counts[edge.target] += 1
            parents[edge.target].append(edge.source)
    return parents, children, incoming_counts, outgoing_counts


def compute_flowchart_export_layout(diagram: ExportFlowchart) -> FlowExportLayout:
    nodes = diagram.nodes
    if not nodes:
        return FlowExportLayout({}, {}, {}, {}, (0.0, 0.0, 0.0, 0.0), False)

    parents, children, incoming_counts, outgoing_counts = _build_flow_metrics(diagram)
    roots = [nid for nid, count in incoming_counts.items() if count == 0] or [next(iter(nodes))]
    levels = {nid: 0 for nid in roots}
    queue_nodes = list(roots)
    visited = set(roots)
    while queue_nodes:
        current = queue_nodes.pop(0)
        for nxt in children.get(current, []):
            if nxt not in visited:
                levels[nxt] = levels[current] + 1
                visited.add(nxt)
                queue_nodes.append(nxt)
    for nid in nodes:
        levels.setdefault(nid, 0)

    lane_names = list(diagram.lanes or [])
    if not lane_names:
        lane_names = ["Main"]
    lane_index = {lane: idx for idx, lane in enumerate(lane_names)}
    for node in nodes.values():
        if not node.lane:
            node.lane = lane_names[0]

    level_groups: Dict[int, List[str]] = {}
    for nid, level in levels.items():
        level_groups.setdefault(level, []).append(nid)

    orders: Dict[str, int] = {}
    previous_orders: Dict[str, float] = {}
    for level in sorted(level_groups):
        ids = level_groups[level]

        def sort_key(node_id: str) -> Tuple[float, int, int, str]:
            parent_order = (
                sum(previous_orders.get(parent, 0.0) for parent in parents[node_id]) / max(len(parents[node_id]), 1)
                if parents[node_id]
                else float(len(previous_orders) + lane_index.get(nodes[node_id].lane, 0))
            )
            return (
                parent_order,
                lane_index.get(nodes[node_id].lane, 0),
                _semantic_rank(nodes[node_id]),
                node_id,
            )

        ids.sort(key=sort_key)
        for order, node_id in enumerate(ids):
            orders[node_id] = order
            previous_orders[node_id] = float(order)

    horizontal = diagram.direction.upper() in {"LR", "RL"}
    positions: Dict[str, Tuple[float, float, float, float]] = {}
    lane_bounds: Dict[str, Tuple[float, float, float, float]] = {}
    margin_x, margin_y = 56.0, 92.0
    gap_x, gap_y = 64.0, 56.0
    lane_gap = 34.0

    if len(lane_names) > 1:
        lane_width = 320.0
        lane_height = 220.0
        max_level = max(levels.values()) if levels else 0
        if horizontal:
            max_stack = max(
                (sum(1 for nid in nodes if nodes[nid].lane == lane and levels[nid] == level) for lane in lane_names for level in range(max_level + 1)),
                default=1,
            )
            dynamic_lane_height = lane_height + max(0, max_stack - 1) * 88.0
            for lane, idx in lane_index.items():
                lane_top = margin_y + idx * (dynamic_lane_height + lane_gap)
                lane_bounds[lane] = (
                    margin_x,
                    lane_top,
                    (max_level + 1) * (lane_width + gap_x),
                    dynamic_lane_height,
                )
            grouped: Dict[Tuple[int, str], List[str]] = {}
            for nid, node in nodes.items():
                grouped.setdefault((levels[nid], node.lane), []).append(nid)
            for ids in grouped.values():
                ids.sort(key=lambda nid: (orders[nid], nid))
            for (level, lane), ids in grouped.items():
                x = margin_x + level * (lane_width + gap_x) + 30.0
                y_start = lane_bounds[lane][1] + 54.0
                for index, nid in enumerate(ids):
                    node = nodes[nid]
                    positions[nid] = (x, y_start + index * (node.height + gap_y), node.width, node.height)
            total_width = margin_x + (max_level + 1) * (lane_width + gap_x)
            total_height = max((bounds[1] + bounds[3] for bounds in lane_bounds.values()), default=0.0) + 48.0
        else:
            max_stack = max(
                (sum(1 for nid in nodes if nodes[nid].lane == lane and levels[nid] == level) for lane in lane_names for level in range(max_level + 1)),
                default=1,
            )
            dynamic_lane_width = lane_width + max(0, max_stack - 1) * 92.0
            for lane, idx in lane_index.items():
                lane_left = margin_x + idx * (dynamic_lane_width + lane_gap)
                lane_bounds[lane] = (
                    lane_left,
                    margin_y,
                    dynamic_lane_width,
                    (max_level + 1) * (lane_height + gap_y),
                )
            grouped = {}
            for nid, node in nodes.items():
                grouped.setdefault((levels[nid], node.lane), []).append(nid)
            for ids in grouped.values():
                ids.sort(key=lambda nid: (orders[nid], nid))
            for (level, lane), ids in grouped.items():
                y = margin_y + level * (lane_height + gap_y) + 36.0
                x_start = lane_bounds[lane][0] + 24.0
                for index, nid in enumerate(ids):
                    node = nodes[nid]
                    positions[nid] = (x_start + index * (node.width + gap_x), y, node.width, node.height)
            total_width = max((bounds[0] + bounds[2] for bounds in lane_bounds.values()), default=0.0) + 48.0
            total_height = margin_y + (max_level + 1) * (lane_height + gap_y)
    else:
        max_level = max(levels.values()) if levels else 0
        graph_left = margin_x
        graph_top = margin_y
        if horizontal:
            graph_height = 0.0
            for level in range(max_level + 1):
                ids = sorted(level_groups.get(level, []), key=lambda nid: (orders[nid], nid))
                top = graph_top + 36.0
                max_right = 0.0
                for idx, nid in enumerate(ids):
                    node = nodes[nid]
                    split_boost = 20.0 if outgoing_counts.get(nid, 0) > 1 else 0.0
                    x = graph_left + level * (240.0 + gap_x) + split_boost
                    y = top + idx * (node.height + gap_y)
                    positions[nid] = (x, y, node.width, node.height)
                    max_right = max(max_right, x + node.width)
                    graph_height = max(graph_height, y + node.height + 32.0)
            lane_bounds["Main"] = (graph_left - 18.0, graph_top - 22.0, max_right - graph_left + 52.0, graph_height - graph_top + 12.0)
            total_width = max_right + 80.0
            total_height = graph_height + 42.0
        else:
            level_widths: Dict[int, float] = {}
            max_row_width = 0.0
            min_canvas_width = 860.0
            for level in range(max_level + 1):
                ids = sorted(level_groups.get(level, []), key=lambda nid: (orders[nid], nid))
                width = sum(nodes[nid].width for nid in ids) + max(0, len(ids) - 1) * gap_x
                level_widths[level] = width
                max_row_width = max(max_row_width, width)
            max_row_width = max(max_row_width, min_canvas_width)
            current_bottom = graph_top
            max_right = graph_left
            for level in range(max_level + 1):
                ids = sorted(level_groups.get(level, []), key=lambda nid: (orders[nid], nid))
                if not ids:
                    current_bottom += 100.0
                    continue
                row_height = max(nodes[nid].height for nid in ids)
                row_width = max(level_widths[level], 260.0)
                x = graph_left + (max_row_width - row_width) / 2.0
                row_extra = 26.0 if any(outgoing_counts.get(nid, 0) > 1 for nid in ids) else 0.0
                y = current_bottom + row_extra
                for nid in ids:
                    node = nodes[nid]
                    positions[nid] = (x, y, node.width, node.height)
                    x += node.width + gap_x
                    max_right = max(max_right, x)
                current_bottom = y + row_height + gap_y
            lane_bounds["Main"] = (graph_left - 18.0, graph_top - 22.0, max_row_width + 44.0, current_bottom - graph_top + 14.0)
            total_width = graph_left + max_row_width + 80.0
            total_height = current_bottom + 36.0

    graph_bounds = (
        min((bounds[0] for bounds in lane_bounds.values()), default=0.0),
        min((bounds[1] for bounds in lane_bounds.values()), default=0.0),
        total_width,
        total_height,
    )
    return FlowExportLayout(positions, lane_bounds, levels, orders, graph_bounds, horizontal)


def compute_augmented_flowchart_layout(base_diagram: ExportFlowchart, augmented_diagram: ExportFlowchart) -> FlowExportLayout:
    base_layout = compute_flowchart_export_layout(base_diagram)
    positions = dict(base_layout.positions)
    lane_bounds = dict(base_layout.lane_bounds)

    incoming: Dict[str, List[str]] = {}
    outgoing: Dict[str, List[str]] = {}
    for edge in augmented_diagram.edges:
        incoming.setdefault(edge.target, []).append(edge.source)
        outgoing.setdefault(edge.source, []).append(edge.target)

    for node_id, node in augmented_diagram.nodes.items():
        if not node.is_junction:
            continue
        parents = [pid for pid in incoming.get(node_id, []) if pid in positions]
        children = [cid for cid in outgoing.get(node_id, []) if cid in positions]
        if not parents and not children:
            continue
        ref_parent = positions[parents[0]] if parents else None
        ref_child = positions[children[0]] if children else None
        if base_layout.horizontal:
            if node.junction_role == "split" and ref_parent is not None:
                left = ref_parent[0] + ref_parent[2] + 24.0
                top = ref_parent[1] + ref_parent[3] / 2.0 - node.height / 2.0
            elif node.junction_role == "merge" and ref_child is not None:
                left = ref_child[0] - 24.0 - node.width
                top = ref_child[1] + ref_child[3] / 2.0 - node.height / 2.0
            else:
                base_ref = ref_parent or ref_child
                left = base_ref[0] + 24.0
                top = base_ref[1] + base_ref[3] / 2.0 - node.height / 2.0
        else:
            if node.junction_role == "split" and ref_parent is not None:
                left = ref_parent[0] + ref_parent[2] / 2.0 - node.width / 2.0
                top = ref_parent[1] + ref_parent[3] + 24.0
            elif node.junction_role == "merge" and ref_child is not None:
                left = ref_child[0] + ref_child[2] / 2.0 - node.width / 2.0
                top = ref_child[1] - 24.0 - node.height
            else:
                base_ref = ref_parent or ref_child
                left = base_ref[0] + base_ref[2] / 2.0 - node.width / 2.0
                top = base_ref[1] + 24.0
        positions[node_id] = (left, top, node.width, node.height)

    all_left = [x for x, _, _, _ in positions.values()]
    all_top = [y for _, y, _, _ in positions.values()]
    all_right = [x + w for x, _, w, _ in positions.values()]
    all_bottom = [y + h for _, y, _, h in positions.values()]
    graph_bounds = (
        min(all_left, default=0.0),
        min(all_top, default=0.0),
        max(all_right, default=0.0) - min(all_left, default=0.0) + 72.0,
        max(all_bottom, default=0.0) - min(all_top, default=0.0) + 72.0,
    )
    return FlowExportLayout(positions, lane_bounds, dict(base_layout.levels), dict(base_layout.orders), graph_bounds, base_layout.horizontal)


def _route_flowchart_edge_tb(
    edge_index: int,
    edge: ExportEdge,
    layout: FlowExportLayout,
    incoming_slot: int,
    same_level_slot: int,
    loop_slot: int,
) -> List[Tuple[float, float]]:
    sx, sy, sw, sh = layout.positions[edge.source]
    tx, ty, tw, th = layout.positions[edge.target]
    src_center_x = sx + sw / 2.0
    src_center_y = sy + sh / 2.0
    dst_center_x = tx + tw / 2.0
    dst_center_y = ty + th / 2.0
    src_level = layout.levels[edge.source]
    dst_level = layout.levels[edge.target]
    graph_left, _graph_top, graph_width, _graph_height = layout.graph_bounds
    graph_right = graph_left + graph_width

    if dst_level > src_level:
        corridor_y = max(sy + sh + 22.0, ty - 26.0 - incoming_slot * 10.0)
        if corridor_y >= ty - 12.0:
            corridor_y = (sy + sh + ty) / 2.0
        return [(src_center_x, corridor_y), (dst_center_x, corridor_y)]

    if dst_level == src_level:
        route_y = min(sy, ty) - 34.0 - same_level_slot * 18.0
        return [(src_center_x, route_y), (dst_center_x, route_y)]

    corridor_side_right = layout.orders[edge.target] > layout.orders[edge.source]
    corridor_x = graph_right + 32.0 + loop_slot * 28.0 if corridor_side_right else graph_left - 32.0 - loop_slot * 28.0
    entry_y = ty - 28.0 - incoming_slot * 8.0
    return [(corridor_x, src_center_y), (corridor_x, entry_y), (dst_center_x, entry_y)]


def _route_flowchart_edge_lr(
    edge_index: int,
    edge: ExportEdge,
    layout: FlowExportLayout,
    incoming_slot: int,
    same_level_slot: int,
    loop_slot: int,
) -> List[Tuple[float, float]]:
    sx, sy, sw, sh = layout.positions[edge.source]
    tx, ty, tw, th = layout.positions[edge.target]
    src_center_x = sx + sw / 2.0
    src_center_y = sy + sh / 2.0
    dst_center_x = tx + tw / 2.0
    dst_center_y = ty + th / 2.0
    src_level = layout.levels[edge.source]
    dst_level = layout.levels[edge.target]
    _graph_left, graph_top, _graph_width, graph_height = layout.graph_bounds
    graph_bottom = graph_top + graph_height

    if dst_level > src_level:
        corridor_x = max(sx + sw + 22.0, tx - 30.0 - incoming_slot * 10.0)
        if corridor_x >= tx - 14.0:
            corridor_x = (sx + sw + tx) / 2.0
        return [(corridor_x, src_center_y), (corridor_x, dst_center_y)]

    if dst_level == src_level:
        route_x = min(sx, tx) - 38.0 - same_level_slot * 18.0
        return [(route_x, src_center_y), (route_x, dst_center_y)]

    corridor_low = layout.orders[edge.target] >= layout.orders[edge.source]
    corridor_y = graph_bottom + 28.0 + loop_slot * 26.0 if corridor_low else graph_top - 28.0 - loop_slot * 26.0
    entry_x = tx - 28.0 - incoming_slot * 8.0
    return [(src_center_x, corridor_y), (entry_x, corridor_y), (entry_x, dst_center_y)]


def build_flowchart_routes(diagram: ExportFlowchart, layout: FlowExportLayout) -> List[List[Tuple[float, float]]]:
    target_incoming_slots: Dict[str, int] = {}
    same_level_slots: Dict[Tuple[str, str], int] = {}
    loop_slots: Dict[Tuple[str, str], int] = {}
    routes: List[List[Tuple[float, float]]] = []
    for index, edge in enumerate(diagram.edges):
        incoming_slot = target_incoming_slots.get(edge.target, 0)
        target_incoming_slots[edge.target] = incoming_slot + 1
        same_level_key = tuple(sorted((edge.source, edge.target)))
        same_level_slot = same_level_slots.get(same_level_key, 0)
        if layout.levels[edge.target] == layout.levels[edge.source]:
            same_level_slots[same_level_key] = same_level_slot + 1
        loop_key = (edge.source, edge.target)
        loop_slot = loop_slots.get(loop_key, 0)
        if layout.levels[edge.target] < layout.levels[edge.source]:
            loop_slots[loop_key] = loop_slot + 1
        if layout.horizontal:
            routes.append(_route_flowchart_edge_lr(index, edge, layout, incoming_slot, same_level_slot, loop_slot))
        else:
            routes.append(_route_flowchart_edge_tb(index, edge, layout, incoming_slot, same_level_slot, loop_slot))
    return routes


def _helper_side_to_site(self, side: str) -> int:
    return {
        "top": 1,
        "left": 2,
        "bottom": 3,
        "right": 4,
    }.get(side, 3)


def _helper_choose_side(shape: Any, point: Tuple[float, float]) -> str:
    center_x = shape.Left + shape.Width / 2.0
    center_y = shape.Top + shape.Height / 2.0
    dx = point[0] - center_x
    dy = point[1] - center_y
    if abs(dx) >= abs(dy):
        return "right" if dx >= 0 else "left"
    return "bottom" if dy >= 0 else "top"


def _helper_add_route_anchor(self, name: str, x: float, y: float, size: float = 3.6) -> Any:
    anchor = self.ws.Shapes.AddShape(self.office_const("msoShapeRectangle"), x - size / 2.0, y - size / 2.0, size, size)
    anchor.Name = name
    anchor.Fill.Visible = False
    anchor.Line.Visible = False
    self.shape_map[name] = anchor
    return anchor


def _helper_connect_segment(self, start_obj: Any, start_side: str, end_obj: Any, end_side: str, name: str, arrow: bool = False) -> Any:
    connector = self.ws.Shapes.AddConnector(self.office_const("msoConnectorStraight"), 0, 0, 10, 10)
    connector.Name = name
    connector.Line.ForeColor.RGB = hex_to_bgr_int(self.theme.connector)
    connector.Line.Weight = 1.35
    if arrow:
        connector.Line.EndArrowheadStyle = self.office_const("msoArrowheadTriangle", "msoShapeRoundedRectangle")
    connector.ConnectorFormat.BeginConnect(start_obj, self.side_to_site(start_side))
    connector.ConnectorFormat.EndConnect(end_obj, self.side_to_site(end_side))
    return connector


def _helper_connect_shapes_routed(
    self,
    edge_key: str,
    source_id: str,
    target_id: str,
    points: Sequence[Tuple[float, float]],
    label: str = "",
) -> None:
    if source_id not in self.shape_map or target_id not in self.shape_map:
        return
    source_shape = self.shape_map[source_id]
    target_shape = self.shape_map[target_id]
    anchors: List[Any] = []
    for index, (x, y) in enumerate(points):
        anchors.append(self.add_route_anchor(f"{edge_key}_a{index}", x, y))

    if anchors:
        first_point = points[0]
        self.connect_segment(
            source_shape,
            _helper_choose_side(source_shape, first_point),
            anchors[0],
            _helper_choose_side(anchors[0], (source_shape.Left + source_shape.Width / 2.0, source_shape.Top + source_shape.Height / 2.0)),
            f"{edge_key}_seg0",
        )
        for index in range(len(anchors) - 1):
            self.connect_segment(
                anchors[index],
                _helper_choose_side(anchors[index], points[index + 1]),
                anchors[index + 1],
                _helper_choose_side(anchors[index + 1], points[index]),
                f"{edge_key}_seg{index + 1}",
            )
        last_anchor = anchors[-1]
        self.connect_segment(
            last_anchor,
            _helper_choose_side(last_anchor, (target_shape.Left + target_shape.Width / 2.0, target_shape.Top + target_shape.Height / 2.0)),
            target_shape,
            _helper_choose_side(target_shape, points[-1]),
            f"{edge_key}_seg_last",
            arrow=True,
        )
        if label:
            mid_x, mid_y = points[len(points) // 2]
            self.add_label(label, mid_x, mid_y - 10.0)
        return

    src_center = (source_shape.Left + source_shape.Width / 2.0, source_shape.Top + source_shape.Height / 2.0)
    dst_center = (target_shape.Left + target_shape.Width / 2.0, target_shape.Top + target_shape.Height / 2.0)
    self.connect_segment(
        source_shape,
        _helper_choose_side(source_shape, dst_center),
        target_shape,
        _helper_choose_side(target_shape, src_center),
        f"{edge_key}_direct",
        arrow=True,
    )
    if label:
        self.add_label(label, (src_center[0] + dst_center[0]) / 2.0, (src_center[1] + dst_center[1]) / 2.0 - 10.0)


def _flowchart_strategy_export(self, workbook, code: str, theme_name: str, title: str) -> None:
    ws = workbook.Worksheets(1)
    ws.Name = "Flowchart"
    theme = get_export_theme(theme_name)
    diagram = parse_flowchart_for_export(code)
    augmented = augment_flowchart_with_junctions(diagram)
    helper = ExcelExportHelper(self.manager, workbook, ws, theme)
    helper.add_title(diagram.title or title, f"Theme: {theme.name} / Junction layout")
    layout = compute_augmented_flowchart_layout(diagram, augmented)
    for lane, bounds in layout.lane_bounds.items():
        helper.add_lane(lane, bounds[0], bounds[1], bounds[2], bounds[3])
    for node_id, node in augmented.nodes.items():
        left, top, width, height = layout.positions[node_id]
        helper.add_node(node, left, top, width, height)
    for edge in augmented.edges:
        helper.connect_shapes(edge.source, edge.target, edge.label, augmented.direction)
    self._add_legend(helper, ws, theme)
    ws.Range("A:AZ").ColumnWidth = 2.4
    ws.Range("1:240").RowHeight = 18


ExcelExportHelper.side_to_site = _helper_side_to_site
ExcelExportHelper.add_route_anchor = _helper_add_route_anchor
ExcelExportHelper.connect_segment = _helper_connect_segment
ExcelExportHelper.connect_shapes_routed = _helper_connect_shapes_routed
FlowchartExportStrategy.export = _flowchart_strategy_export


class GanttExportStrategy(ExcelExportStrategy):
    def export(self, workbook, code: str, theme_name: str, title: str) -> None:
        import datetime as dt

        ws = workbook.Worksheets(1)
        ws.Name = "Gantt"
        theme = get_export_theme(theme_name)
        helper = ExcelExportHelper(self.manager, workbook, ws, theme)
        diagram = parse_gantt_for_export(code)
        helper.add_title(diagram.title or title, "Editable table + bar overlay")
        tasks = diagram.tasks
        if not tasks:
            raise RuntimeError("No Gantt tasks found.")
        min_date = min(task.start for task in tasks)
        max_date = max(task.end for task in tasks)
        dates: List[Any] = []
        cur = min_date
        while cur <= max_date:
            dates.append(cur)
            cur += dt.timedelta(days=1)
        headers = ["Section", "Task", "Owner", "Start", "End", "Duration", "Progress", "Dependency"]
        for col, header in enumerate(headers, start=1):
            ws.Cells(4, col).Value = header
            ws.Cells(4, col).Font.Bold = True
            ws.Cells(4, col).Interior.Color = hex_to_bgr_int(theme.title_fill)
        start_col = len(headers) + 1
        for idx, day in enumerate(dates, start=start_col):
            ws.Cells(4, idx).Value = day.strftime("%m-%d")
            ws.Cells(4, idx).Font.Size = 8
            ws.Cells(4, idx).Interior.Color = hex_to_bgr_int(theme.lane_fill)
            ws.Columns(idx).ColumnWidth = 3.2
        section_colors: Dict[str, str] = {}
        palette = ["#F5F9FF", "#F8FBF2", "#FFF8EE", "#F7F1FF"]
        for row_idx, task in enumerate(tasks, start=5):
            if task.section not in section_colors:
                section_colors[task.section] = palette[len(section_colors) % len(palette)]
            ws.Cells(row_idx, 1).Value = task.section
            ws.Cells(row_idx, 2).Value = task.name
            ws.Cells(row_idx, 3).Value = task.owner
            ws.Cells(row_idx, 4).Value = task.start.strftime("%Y-%m-%d")
            ws.Cells(row_idx, 5).Value = task.end.strftime("%Y-%m-%d")
            ws.Cells(row_idx, 6).Value = task.duration_days
            ws.Cells(row_idx, 7).Value = int(task.progress * 100)
            ws.Cells(row_idx, 8).Value = task.dependency
            ws.Range(ws.Cells(row_idx, 1), ws.Cells(row_idx, 8)).Interior.Color = hex_to_bgr_int(section_colors[task.section])
            bar_start = start_col + (task.start - min_date).days
            bar_end = start_col + (task.end - min_date).days
            left = ws.Cells(row_idx, bar_start).Left + 1
            top = ws.Cells(row_idx, bar_start).Top + 3
            width = max(10.0, ws.Cells(row_idx, bar_end).Left + ws.Cells(row_idx, bar_end).Width - left - 1)
            height = ws.Cells(row_idx, bar_start).Height - 6
            if task.milestone:
                shp = ws.Shapes.AddShape(helper.office_const("msoShapeDiamond"), left, top, height, height)
            else:
                shp = ws.Shapes.AddShape(helper.office_const("msoShapeRoundedRectangle"), left, top, width, height)
            color_key = "done" if task.progress >= 0.99 else "process" if task.progress < 0.6 else "review"
            fill, line, _text = theme.semantic_colors[color_key]
            shp.Fill.ForeColor.RGB = hex_to_bgr_int(fill)
            shp.Line.ForeColor.RGB = hex_to_bgr_int(line)
            shp.Line.Weight = 1.2
        ws.Columns("A:H").AutoFit()
        ws.Rows("4:120").RowHeight = 20


class SequenceExportStrategy(ExcelExportStrategy):
    def export(self, workbook, code: str, theme_name: str, title: str) -> None:
        ws = workbook.Worksheets(1)
        ws.Name = "Sequence"
        theme = get_export_theme(theme_name)
        helper = ExcelExportHelper(self.manager, workbook, ws, theme)
        data = parse_sequence_for_export(code)
        helper.add_title(data.title or title, "Editable participant headers + lifelines + arrows")
        if not data.participants:
            raise RuntimeError("No sequence participants found.")
        base_left, top = 80.0, 70.0
        col_gap = 150.0
        x_map: Dict[str, float] = {}
        for idx, participant in enumerate(data.participants):
            x = base_left + idx * col_gap
            x_map[participant.key] = x + 50
            header = ws.Shapes.AddShape(helper.office_const("msoShapeRoundedRectangle"), x, top, 100, 30)
            header.TextFrame2.TextRange.Text = participant.label
            header.Fill.ForeColor.RGB = hex_to_bgr_int(theme.title_fill)
            header.Line.ForeColor.RGB = hex_to_bgr_int(theme.title_line)
            line = ws.Shapes.AddLine(x + 50, top + 30, x + 50, top + 560)
            line.Line.ForeColor.RGB = hex_to_bgr_int(theme.connector)
            line.Line.DashStyle = 4
        y = top + 60
        block_start_y = None
        for item in data.messages:
            if item.kind == "block_start":
                block_start_y = y
                ws.Cells(1, 20).Value = item.label
                y += 20
                continue
            if item.kind == "block_end" and block_start_y is not None:
                box = ws.Shapes.AddShape(helper.office_const("msoShapeRectangle"), base_left - 20, block_start_y - 8, col_gap * max(len(data.participants) - 1, 1) + 130, y - block_start_y + 8)
                box.Fill.Transparency = 0.75
                box.Fill.ForeColor.RGB = hex_to_bgr_int(theme.lane_fill)
                box.Line.ForeColor.RGB = hex_to_bgr_int(theme.lane_line)
                try:
                    box.ZOrder(helper.office_const("msoSendToBack", "msoShapeRectangle"))
                except Exception:
                    pass
                block_start_y = None
                y += 10
                continue
            if item.kind == "note":
                note = ws.Shapes.AddShape(helper.office_const("msoShapeRoundedRectangle"), base_left + 10, y - 5, 180, 28)
                note.TextFrame2.TextRange.Text = item.label
                note.Fill.ForeColor.RGB = hex_to_bgr_int(theme.semantic_colors["data"][0])
                note.Line.ForeColor.RGB = hex_to_bgr_int(theme.semantic_colors["data"][1])
                y += 34
                continue
            if item.kind == "message":
                sx = x_map.get(item.source, base_left)
                tx = x_map.get(item.target, base_left + col_gap)
                anchor1 = ws.Shapes.AddShape(helper.office_const("msoShapeRectangle"), sx - 2, y, 4, 4)
                anchor2 = ws.Shapes.AddShape(helper.office_const("msoShapeRectangle"), tx - 2, y, 4, 4)
                anchor1.Fill.Transparency = 1.0
                anchor1.Line.Visible = False
                anchor2.Fill.Transparency = 1.0
                anchor2.Line.Visible = False
                helper.shape_map[f"a_{y}_{item.source}"] = anchor1
                helper.shape_map[f"b_{y}_{item.target}"] = anchor2
                helper.connect_shapes(f"a_{y}_{item.source}", f"b_{y}_{item.target}", item.label, "LR")
                y += 32


class ExcelExportManager:
    def __init__(self, log_func):
        self.log = log_func
        self.strategies = {
            "flowchart": FlowchartExportStrategy(self),
            "org chart": OrgChartExportStrategy(self),
            "swimlane": SwimlaneExportStrategy(self),
            "gantt": GanttExportStrategy(self),
            "sequenceDiagram": SequenceExportStrategy(self),
        }

    def export_editable(self, mermaid_code: str, theme_name: str, path: Path, preferred_type: str = "") -> None:
        if win32 is None:
            raise RuntimeError("pywin32 is required for Excel export.")
        detected = detect_mermaid_diagram_type(mermaid_code)
        export_type = preferred_type or detected
        if export_type == "자동 추천":
            export_type = detected
        strategy = self.strategies.get(export_type)
        if strategy is None:
            raise RuntimeError(f"Excel editable export is not implemented for {export_type}.")
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "Diagram"
            data_ws = wb.Worksheets.Add(After=ws)
            data_ws.Name = "Metadata"
            data_ws.Cells(1, 1).Value = "DiagramType"
            data_ws.Cells(1, 2).Value = export_type
            data_ws.Cells(2, 1).Value = "Theme"
            data_ws.Cells(2, 2).Value = get_export_theme(theme_name).name
            strategy.export(wb, mermaid_code, theme_name, path.stem)
            ensure_parent(path)
            wb.SaveAs(str(path))
            self.log(f"Excel 편집형 내보내기 완료: {export_type} -> {path}")
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            excel.Quit()


def run_excel_export_samples(output_dir: Path) -> None:
    manager = ExcelExportManager(print)
    ensure_parent(output_dir / "dummy.txt")
    sample_plan = [
        ("flowchart", TYPE_SAMPLES["flowchart"][0], "Executive Blue"),
        ("gantt", TYPE_SAMPLES["gantt"][0], "Warm Gray"),
        ("sequenceDiagram", TYPE_SAMPLES["sequenceDiagram"][0], "Calm Green"),
    ]
    for diagram_type, code, theme_name in sample_plan:
        target = output_dir / f"sample_{diagram_type}.xlsx"
        manager.export_editable(code, theme_name, target)


def _app_export_excel_shapes(self) -> None:
    code = self.get_current_code()
    if not code:
        messagebox.showwarning(APP_TITLE, "내보낼 Mermaid 코드가 없습니다.")
        return
    path = filedialog.asksaveasfilename(
        title="Excel 편집형 내보내기",
        defaultextension=".xlsx",
        initialfile=slugify_filename(self.title_var.get()) + ".xlsx",
        filetypes=[("Excel workbook", "*.xlsx")],
    )
    if not path:
        return

    def task():
        try:
            self.status_var.set("Excel 내보내기 중...")
            manager = ExcelExportManager(self.log)
            manager.export_editable(code, self.theme_var.get(), Path(path), _app_selected_diagram_type(self))
            self.status_var.set("Excel 내보내기 완료")
        except Exception as exc:
            self.log(f"Excel 내보내기 실패: {exc}")
            self.status_var.set("Excel 내보내기 실패")
            self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"Excel 내보내기 실패\n{exc}"))

    threading.Thread(target=task, daemon=True).start()


MermaidDesignerApp.export_excel_shapes = _app_export_excel_shapes

def main() -> None:
    configure_file_logging()
    root = tk.Tk()
    app = MermaidDesignerApp(root)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == '--preview-helper':
        bridge = Path(sys.argv[2])
        mermaid_url = sys.argv[3] if len(sys.argv) >= 4 else MERMAID_JS_CDN
        raise SystemExit(_run_preview_helper(bridge, mermaid_url))
    main()
