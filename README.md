[README.md](https://github.com/user-attachments/files/33659519/README.md)
# epoxy-crack-check

균열 보수용 에폭시의 비선형 물성과 콘크리트 물성을 넣으면, Case 1 (관통균열 · 양단 구속 · 온도 하강) 모델에서 콘크리트 응력 · 에폭시 힘 · 최대 변형률을 구하고 콘크리트 균열 / 에폭시 파단을 판정해 엑셀로 정리하는 에이전트 스킬.

## 구조
```
epoxy-crack-check-project/
├── README.md
└── .agents/skills/epoxy-crack-check/
    ├── SKILL.md                  # name · description + 절차 · 출력 형식
    ├── references/
    │   └── research.md           # 배경 · 용어 · 판정 기준 · 식 · 검증
    ├── scripts/
    │   ├── solve_case1.ps1       # 식 (1D 평형) 계산 – 기본
    │   ├── run_ansys_case1.ps1   # ANSYS MAPDL 배치 계산 – 요청 시
    │   └── build_excel.ps1       # 결과 엑셀 작성
    └── assets/
        ├── ds_case1.dat          # ANSYS 입력 템플릿 (Case 1 모델)
        └── post_case1.inp        # ANSYS 결과 추출
```

## 설치
- Codex: 이 폴더의 `.agents/skills/epoxy-crack-check/` 를 그대로 둔다.
- Claude Code: `epoxy-crack-check/` 폴더를 프로젝트의 `.claude/skills/` 또는 사용자 폴더 `~/.claude/skills/` 로 복사한다.

## 필요 환경
- Windows, PowerShell 5.1, Microsoft Excel (엑셀 작성)
- ANSYS 2025 R1 (MAPDL) 과 라이선스 – ANSYS 로 확인할 때만

## 사용 예
- "항복 2 인장 5, E 100 이랑 500, 끝점 100 150 200 %, 콘크리트 인장강도 3 4 5 에서 판정해줘"
- "f_ck 24 에 항복 1 인장 2.65 신율 340 % 에폭시 괜찮아? 균열폭 0.2 mm"
- "이 조합 ANSYS 로도 확인해줘"

## 검증
식 계산은 사용자 ANSYS 2D 모델 결과와 0.0002 MPa 이내로 일치한다 (예: 항복 2 · 인장 5 · E 100 · 끝점 150 % → 4.6664 vs 4.6665 MPa). 자세한 값은 `references/research.md`.
