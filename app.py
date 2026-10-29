# -*- coding: utf-8 -*-
"""
웹 화면 — '웹화면_실행.bat'을 더블클릭하면 브라우저에서 열립니다.
(직접 실행: python -m streamlit run app.py)
"""
import os
import re
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import keyphrase_cloud as kc

UPLOAD_DIR = kc.BASE_DIR / "uploads"
SECTION_OPTIONS = {"제목": "title", "초록": "abstract", "저자키워드": "keywords"}
ANGLE_OPTIONS = {"mixed": "다양한 각도", "right": "가로 + 세로", "horizontal": "가로만"}
CONFIG_FILES = {
    "제외할 키프레이즈": ("exclude_phrases.txt", "결과에서 **완전히 뺄** 항목을 한 줄에 하나씩 적습니다. (예: `Korea`)"),
    "같은 뜻으로 합치기": ("synonyms.txt", "형식: `변형표기 => 대표표기` — 여러 개는 `|`로 구분 (예: `SARS-CoV-2 | COVID-19 pandemic => COVID-19`)"),
    "항상 포함할 구문": ("custom_phrases.txt", "하나의 구문으로 인식시키고, 일반어 필터와 관계없이 **항상 결과에 포함**할 항목"),
    "1단어 일반어": ("stopwords.txt", "1단어로는 집계하지 않을 일반 단어 (다른 단어와 결합한 구문의 일부로는 집계됨)"),
}

st.set_page_config(page_title="논문 키프레이즈 워드클라우드", page_icon="☁️", layout="wide")


# ── 도우미 ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="파일 읽는 중…")
def inspect_file(path: str, mtime: float):
    df, _ = kc.load_records([path], None, None)
    dist = df["Year"].value_counts().to_dict() if "Year" in df.columns else {}
    return len(df), {y: c for y, c in sorted(dist.items()) if y}


def save_upload(f) -> Path:
    UPLOAD_DIR.mkdir(exist_ok=True)
    path = UPLOAD_DIR / Path(f.name).name
    data = f.getvalue()
    if not path.exists() or path.stat().st_size != len(data) or path.read_bytes() != data:
        path.write_bytes(data)
    return path


def append_config_lines(fname: str, items: list[str]) -> int:
    path = kc.CONFIG_DIR / fname
    text = path.read_text(encoding="utf-8-sig") if path.exists() else ""
    existing = {line.strip().lower() for line in text.splitlines()}
    new = [i for i in items if i.strip().lower() not in existing]
    if new:
        path.write_text(text.rstrip("\n") + "\n" + "\n".join(new) + "\n", encoding="utf-8")
    return len(new)


def safe_label(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip() or "result"


# ── 분석 실행 ─────────────────────────────────────────────────────────────────
def run_analysis(path: Path, years, label, sections, weight, top, ratio_pct, min_phrase, angles):
    argv = [
        str(path), "--label", label, "--top", str(top), "--angles", angles,
        "--weight", weight, "--min-keyword-ratio", str(ratio_pct / 100),
        "--min-phrase-papers", str(min_phrase), "--sections", *[SECTION_OPTIONS[s] for s in sections],
    ]
    if years:
        argv += ["--year", *years]
    args = kc.parse_args(argv)
    if ratio_pct == 0:
        args.min_keyword_papers = 0

    logs = []
    with st.status("분석 중… 처음 분석하는 파일은 3~5분 정도 걸립니다.", expanded=True) as status:
        bar = st.progress(0.0, text="준비 중")
        log_box = st.empty()

        def on_log(msg):
            logs.append(msg)
            log_box.code("\n".join(logs[-10:]), language=None)

        def on_progress(done, total):
            bar.progress(done / total, text=f"문장 분석(품사 태깅) {done:,} / {total:,}")

        kc.set_logger(on_log)
        try:
            res = kc.run_pipeline(args, progress=on_progress)
        except SystemExit as e:
            status.update(label=f"오류: {e}", state="error")
            return
        except Exception as e:  # noqa: BLE001
            status.update(label="오류가 발생했습니다", state="error")
            st.exception(e)
            return
        finally:
            kc.set_logger(None)
        bar.progress(1.0, text="완료")
        status.update(label=f"분석 완료 ({res['seconds']}초)", state="complete", expanded=False)
    st.session_state["result"] = res


# ── 결과 표시 ─────────────────────────────────────────────────────────────────
def render_result():
    res = st.session_state.get("result")
    if not res:
        return
    table, dropped = res["table"], res["dropped"]
    outputs = [Path(p) for p in res["outputs"]]

    st.divider()
    st.subheader(f"3. 결과 — {res['label']}")
    m1, m2, m3 = st.columns(3)
    m1.metric("분석한 논문", f"{res['settings']['집계 논문 수']:,}편")
    m2.metric("키프레이즈", f"{len(table):,}개")
    m3.metric("그중 2단어 이상 구문", f"{int((table['단어수'] >= 2).sum()):,}개")

    reports = [p for p in outputs if p.suffix == ".html"]
    if reports:
        # HTML 안에서 클릭/검색/정렬하므로 단어를 누를 때 분석을 다시 실행하지 않는다.
        html = reports[0].read_text(encoding="utf-8")
        if hasattr(st, "iframe"):
            st.iframe(html, height="content")
        else:
            components.html(html, height=850, scrolling=True)
    else:
        for p in outputs:
            if p.suffix == ".png":
                st.image(str(p), width="stretch")

    st.markdown("**내려받기**")
    cols = st.columns(len(outputs) + 1)
    for col, p in zip(cols, outputs):
        if p.suffix == ".xlsx":
            name, mime = "📥 엑셀 표", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif p.suffix == ".html":
            name, mime = "🔎 논문 탐색 HTML", "text/html"
        else:
            name = "🖼️ 전체 워드클라우드" if p.stem.endswith("전체") else "🖼️ 구문 워드클라우드"
            mime = "image/png"
        col.download_button(name, data=p.read_bytes(), file_name=p.name, mime=mime, key=f"dl_{p.name}")
    if hasattr(os, "startfile") and cols[-1].button("📂 결과 폴더 열기"):  # 내 컴퓨터에서 실행할 때만
        os.startfile(outputs[0].parent)

    st.markdown("**키프레이즈 표**")
    c1, c2 = st.columns([2, 3])
    query = c1.text_input("검색", placeholder="예: cancer", key="search")
    kind = c2.radio("보기", ["전체", "구문(2단어 이상)", "단어(1단어)"], horizontal=True, key="kind")
    view = table
    if kind.startswith("구문"):
        view = view[view["단어수"] >= 2]
    elif kind.startswith("단어"):
        view = view[view["단어수"] == 1]
    if query:
        hit = view["병합된 표기"].str.contains(query, case=False, regex=False)
        view = view[hit | view["키프레이즈"].str.contains(query, case=False, regex=False)]
    st.dataframe(to_display(view), height=460)

    with st.expander("✂️ 결과에서 빼고 싶은 키프레이즈가 있나요?"):
        picks = st.multiselect("제외할 항목 선택 (상위 300개 중)", table["키프레이즈"].head(300).tolist(), key="picks")
        if st.button("제외 목록에 추가", disabled=not picks):
            n = append_config_lines("exclude_phrases.txt", picks)
            st.success(f"{n}개를 제외 목록에 추가했습니다. 위의 **분석 시작**을 다시 누르면 반영됩니다 (수십 초).")

    with st.expander(f"🔍 일반어 필터로 제외된 항목 보기 ({len(dropped):,}개)"):
        st.caption("초록에는 자주 나오지만 저자키워드로는 거의 쓰이지 않는 일반어입니다. "
                   "살리고 싶은 항목은 아래에서 선택하거나 ‘⚙️ 단어 설정 → 항상 포함할 구문’에 추가하세요.")
        keep = st.multiselect("항상 포함할 항목으로 추가", dropped["키프레이즈"].head(300).tolist(), key="keep")
        if st.button("항상 포함 목록에 추가", disabled=not keep):
            n = append_config_lines("custom_phrases.txt", keep)
            st.success(f"{n}개를 추가했습니다. **분석 시작**을 다시 누르면 반영됩니다.")
        st.dataframe(to_display(dropped.head(500)), height=360)


def to_display(df):
    out = df.drop(columns=["정규화 키"]).reset_index(drop=True)
    out["저자키워드 사용률"] = (out["저자키워드 사용률"] * 100).round(1).astype(str) + "%"
    out.index += 1
    return out


# ── 탭 1: 분석 ────────────────────────────────────────────────────────────────
def render_run_tab():
    st.subheader("1. 파일 올리기")
    uploaded = st.file_uploader("Scopus에서 내보낸 논문 목록 (.xlsx 또는 .csv)", type=["xlsx", "csv"])
    with st.expander("또는 컴퓨터에 있는 파일 경로로 불러오기"):
        typed = st.text_input("파일 경로", key="typed_path",
                              placeholder=r"C:\Users\User\Desktop\YRS 2026_scopus.xlsx").strip().strip('"')

    path = None
    if uploaded:
        path = save_upload(uploaded)
    elif typed:
        if Path(typed).is_file():
            path = Path(typed)
        else:
            st.error("해당 경로에 파일이 없습니다.")
    if path is None:
        st.info("파일을 올리면 옵션과 **분석 시작** 버튼이 나타납니다.")
        render_result()
        return

    try:
        n_papers, year_dist = inspect_file(str(path), path.stat().st_mtime)
    except SystemExit as e:
        st.error(f"파일을 읽을 수 없습니다: {e}")
        return
    years_txt = " · ".join(f"{y}년 {c:,}편" for y, c in year_dist.items())
    st.success(f"**{path.name}** — 논문 {n_papers:,}편" + (f"  ({years_txt})" if years_txt else ""))

    st.subheader("2. 옵션")
    c1, c2 = st.columns(2)
    with c1:
        years = []
        if year_dist:
            years = st.multiselect("발행연도", list(year_dist), default=list(year_dist))
        default_label = path.stem if not years or len(years) == len(year_dist) else f"{path.stem}_{'-'.join(years)}"
        label = safe_label(st.text_input("결과 이름 (파일 이름에 사용)", value=default_label))
    with c2:
        sections = st.multiselect("집계할 부분", list(SECTION_OPTIONS), default=list(SECTION_OPTIONS))
        weight = st.radio("워드클라우드 글자 크기 기준", ["count", "papers"], horizontal=True,
                          format_func={"count": "노출빈도 (총 등장 횟수)", "papers": "논문수 (등장한 논문 수)"}.get)
        angles = st.radio("워드클라우드 글자 각도", list(ANGLE_OPTIONS), horizontal=True, format_func=ANGLE_OPTIONS.get)
    with st.expander("고급 설정"):
        top = st.slider("워드클라우드에 표시할 키프레이즈 수", 50, 300, 150, 10)
        ratio = st.slider("일반어 필터 강도 — 저자키워드 사용률 하한 (%)", 0, 30, 5,
                          help="키프레이즈가 등장한 논문 중 저자키워드로도 쓴 논문의 비율이 이 값 이상인 것만 남깁니다. "
                               "0이면 필터를 끕니다. 높일수록 일반적인 단어가 더 많이 빠집니다.")
        min_phrase = st.number_input("제목·초록에서 구문으로 인정할 최소 논문 수", 2, 20, 3,
                                     help="데이터가 적을 때(예: 몇 백 편)는 2로 낮추세요.")

    if st.button("🚀 분석 시작", type="primary", disabled=not sections):
        run_analysis(path, years if len(years) != len(year_dist) else [], label, sections, weight, top, ratio,
                     int(min_phrase), angles)
    render_result()


# ── 탭 2: 단어 설정 ───────────────────────────────────────────────────────────
def render_config_tab():
    st.caption("수정 후 **저장**하고, ‘📊 분석’ 탭에서 **분석 시작**을 다시 누르면 반영됩니다 (두 번째부터는 수십 초). "
               "단수/복수·대소문자·하이픈 차이는 자동으로 합쳐지므로 한 가지 형태만 적으면 됩니다. `#` 뒤는 설명입니다.")
    for tab, (name, (fname, desc)) in zip(st.tabs(list(CONFIG_FILES)), CONFIG_FILES.items()):
        with tab:
            st.markdown(desc)
            path = kc.CONFIG_DIR / fname
            text = path.read_text(encoding="utf-8-sig") if path.exists() else ""
            mtime = path.stat().st_mtime_ns if path.exists() else 0
            new = st.text_area(name, value=text, height=420, key=f"cfg_{fname}_{mtime}", label_visibility="collapsed")
            if st.button("💾 저장", key=f"save_{fname}"):
                path.write_text(new, encoding="utf-8")
                st.toast(f"‘{name}’ 저장 완료")
                st.rerun()


# ── 탭 3: 도움말 ──────────────────────────────────────────────────────────────
HELP = """
#### 사용 순서
1. Scopus에서 논문 목록을 **Excel 또는 CSV로 내보내기** (Title · Abstract · Author Keywords 열 포함)
2. **📊 분석** 탭에 파일을 끌어다 놓기 → 연도 확인 → **분석 시작**
3. 워드클라우드·엑셀 표 확인 및 내려받기 (내 컴퓨터에서 실행하면 `output` 폴더에도 저장됩니다)
4. 빼고 싶은 단어는 결과 아래 **✂️**에서, 표기 합치기 등은 **⚙️ 단어 설정** 탭에서 고친 뒤 다시 분석

집계·필터 기준은 옆의 **처리 방식** 탭에 자세히 정리되어 있습니다.

#### 엑셀 열 설명
- **총빈도**: 제목+초록+저자키워드에 등장한 총 횟수 (노출빈도)
- **논문수**: 한 번이라도 등장한 논문 수
- **저자키워드 사용률**: 등장한 논문 중 저자키워드로도 쓴 논문의 비율
- **병합된 표기**: 하나로 합쳐진 실제 표기들 (단복수·표기 차이 확인용)

#### 종료
내 컴퓨터에서 실행한 경우, 검은색 실행 창을 닫으면 프로그램이 종료됩니다.
"""

METHOD_HELP = """
결과는 아래 6단계를 거쳐 나옵니다. 숫자는 **기본값** 기준이며, 일부는 ‘2. 옵션 → 고급 설정’에서 바꿀 수 있습니다.

**① 논문 고르기 → ② 텍스트 정리 → ③ 표기 통일 → ④ 구문 사전 만들기 → ⑤ 세기 → ⑥ 거르기·정렬**

##### ① 논문 고르기
- **읽는 열**: `Title`, `Abstract`, `Author Keywords` 세 열만 사용
- **연도**: 발행연도를 고르면 `Year` 열 기준으로 해당 연도만 남김
- **중복**: `EID`가 같으면 한 편으로 셈 (EID가 없으면 제목으로 판단)
- **빈 행**: 제목과 초록이 모두 비어 있으면 제외. `[No abstract available]`은 초록 없음으로 처리

##### ② 텍스트 정리
주제어가 아닌 문구를 먼저 지웁니다.

| 지우는 것 | 예 |
|---|---|
| HTML 태그 | `<sup>`, `<i>` |
| 저작권 문구 (그 뒤 전부) | `© 2025 Elsevier…`, `Copyright 2024…` |
| 구조화 초록 소제목 | `Background:`, `Methods:`, `Main outcome measures:` |
| 통계 괄호 | `(95% CI 1.2–3.4)`, `(p<0.05)`, `(OR = 2.1)` |
| 인용 괄호 | `(Kim et al 2020; …)` |

**약어 풀기** — 한 논문 안에서 `large language models (LLMs)`처럼 정의된 약어를 찾아,
**그 논문의** 제목·초록·키워드에 나오는 `LLMs`·`LLM`을 `large language models`로 바꿉니다. 다른 논문에는 적용하지 않습니다.

##### ③ 표기 통일
- **대소문자 무시**: `Deep Learning` = `deep learning`
- **하이픈·소유격**: `meta-analysis` = `meta analysis`, `Alzheimer's disease` = `Alzheimer disease`
- **단수형** (영어 명사 사전 lemminflect): `models → model`, `analyses → analysis`, `mice → mouse`, `LLMs → LLM`
- **예외**: 대문자 약어(`AIDS`, `SARS`)와 3글자 이하 단어는 그대로 · 사전에 없는 단어는 단수형이 파일 안에 실제로 나올 때만 변환
- 무엇이 합쳐졌는지는 표의 **병합된 표기** 열에서 확인

##### ④ 구문 사전 만들기
품사 태깅(spaCy) 후, 아래 세 출처를 합쳐 ‘한 덩어리로 볼 구문’ 목록을 만듭니다.

| 출처 | 조건 |
|---|---|
| 저자키워드 | 전체 논문의 저자키워드 중 **2–8단어**짜리 전부 (한 편에만 있어도 포함) |
| 자동 추출 명사구 | 형용사·명사가 연달아 나오는 **2–4단어** 구간 · 끝 단어는 명사 · **3편 이상** 논문에 등장 |
| 항상 포함할 구문 | ‘⚙️ 단어 설정 → 항상 포함할 구문’에 적은 항목 |

자동 추출에서는 `novel`, `significant` 같은 수식어로 **시작**하거나 `model`, `approach` 같은 일반 명사로 **끝나는** 구문을 뺍니다.

##### ⑤ 세기
- **제목·초록**: 구문 사전에서 **가장 긴 구문부터** 맞춤 (쉼표·마침표를 넘어서는 매칭하지 않음).
  구문에 걸리지 않은 **명사만** 1단어로 세고, 형용사·동사 단독은 세지 않음
- **저자키워드**: 세미콜론으로 나눈 키워드 하나하나를 그대로 1회씩 셈 (1단어 키워드 포함)
- **합치기·빼기**: ‘같은 뜻으로 합치기’로 대표 표기에 합치고, ‘제외할 키프레이즈’·‘1단어 일반어’,
  영어 기능어(the, of …), 1글자·숫자만 있는 항목은 이때 뺌

> ⚠️ **긴 구문으로 잡힌 부분은 짧은 구문으로 다시 세지 않습니다.**
> 초록의 `non-small cell lung cancer`는 그 구문으로만 1회 세고, `lung cancer`나 `cancer`로는 세지 않습니다.

| 기록하는 값 | 뜻 |
|---|---|
| 제목 · 초록 · 저자키워드 | 각 부분의 **등장 횟수** (한 초록에 5번 나오면 5) |
| 총빈도 | 세 값의 합 (노출빈도) |
| 논문수 | 한 번이라도 등장한 논문 수 |
| 저자키워드 사용률 | 저자키워드로 쓴 논문 수 ÷ 등장한 논문 수 |

##### ⑥ 거르기와 정렬
결과에 남으려면 **세 조건을 모두** 만족해야 합니다.

1. 등장 논문 **2편 이상** (미달이면 표에서 아예 빠짐)
2. 저자키워드로 쓴 논문 **2편 이상**
3. 저자키워드 사용률 **5% 이상** (고급 설정의 ‘일반어 필터 강도’ — 0으로 하면 2·3번 조건을 끔)

- 2·3번에서 탈락한 항목은 결과 아래 **🔍 일반어 필터로 제외된 항목**과 엑셀 ‘제외됨(일반어 필터)’ 시트에서 확인.
  ‘항상 포함할 구문’에 넣은 항목은 조건과 관계없이 남음
- **정렬**: 총빈도 순, 같으면 논문수 순 (글자 크기 기준을 ‘논문수’로 고르면 반대)
- **표시 이름**: 합쳐진 표기 중 가장 많이 쓰인 형태 (예: `COVID-19`)
- **워드클라우드**: 남은 항목 중 상위 150개 (구문 이미지는 2단어 이상만 따로 150개).
  글자 크기는 총빈도(또는 논문수)를 따르지만 정확히 비례하지는 않으므로 정확한 수치는 표에서 확인.
  상위 3개는 가로로, 나머지는 선택한 각도로 가운데부터 배치

##### 예시 한 편으로 따라가 보기
> **초록** — *Large language models (LLMs) are increasingly used. We evaluated LLMs for predicting survival in non-small cell lung cancer patients.*
>
> **저자키워드** — *LLM; Lung cancer; Survival analysis*

| 키프레이즈 | 어떻게 세어지나 |
|---|---|
| large language model | 초록 2 (`LLMs` 약어 풀기) + 키워드 1 (`LLM`도 치환) → 총빈도 3, 논문 1 |
| non-small cell lung cancer | 구문 사전에 있으면 초록 1 (다른 논문의 키워드에 있거나 3편 이상에 반복된 경우) |
| lung cancer | 키워드 1 — 초록 속 lung cancer는 더 긴 구문에 포함돼 따로 세지 않음 |
| survival analysis | 키워드 1 |
| survival · patient | 1단어 명사로 초록 각 1 — 저자키워드로 거의 안 쓰이면 ⑥에서 빠질 수 있음 |

##### 해석할 때 주의할 점
- **저자키워드를 빼면 결과가 거의 사라집니다.** ‘집계할 부분’에서 저자키워드를 해제하면 사용률이 모두 0이 되어
  ‘항상 포함할 구문’만 남습니다. 이때는 고급 설정의 일반어 필터 강도를 **0**으로 바꾸세요.
- **긴 구문이 우선합니다.** 짧은 구문의 초록 빈도가 생각보다 낮다면 더 긴 구문에 흡수됐을 수 있습니다. 필요하면 ‘같은 뜻으로 합치기’를 쓰세요.
- **1단어 형용사는 나오지 않습니다.** `pediatric`처럼 명사가 아닌 단어는 구문의 일부로만 집계됩니다.
- **품사 태깅은 자동이라 가끔 틀립니다.** 꼭 잡아야 하는 표현은 ‘항상 포함할 구문’에 적으세요.
"""


st.title("☁️ 논문 키프레이즈 워드클라우드")
st.caption("Scopus 논문 목록 파일을 올리면 제목·초록·저자키워드의 키프레이즈 노출빈도를 집계해 워드클라우드와 엑셀 표를 만듭니다.")
tab_run, tab_cfg, tab_help = st.tabs(["📊 분석", "⚙️ 단어 설정", "❓ 도움말"])
with tab_run:
    render_run_tab()
with tab_cfg:
    render_config_tab()
with tab_help:
    help_use, help_method = st.tabs(["사용법", "처리 방식 (집계·필터 기준)"])
    with help_use:
        st.markdown(HELP)
    with help_method:
        st.markdown(METHOD_HELP)
