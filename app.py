import io
import os
import zipfile
import xml.etree.ElementTree as ET

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# .env 파일에서 환경변수 불러오기
load_dotenv()

# 모델명은 여기서 쉽게 바꿀 수 있도록 상수로 분리
CHAT_MODEL = "gpt-5.4-nano"          # 일반 챗봇 / 문서 기반 답변에 사용할 모델
WEB_SEARCH_MODEL = "gpt-5.4-nano"    # 웹검색 모드에서 사용할 모델

# 문서가 너무 길 경우 앞부분 일부만 사용 (최대 글자 수 제한)
MAX_DOC_CHARS = 6000

# 보고서가 들어 있는 데이터 폴더 (분야별 하위 폴더)
DATA_DIR = r"C:\Users\jihyun\Desktop\KEITI_AD\ecolab\데이터"
CATEGORIES = {
    "국가별 보고서": "country_report",
    "정책규제 보고서": "policy_report",
}

# 문서에서 답을 찾지 못했을 때 사용할 고정 안내 문구
NOT_FOUND_MESSAGE = "업로드된 문서에서 확인하기 어렵습니다"

# 한글 문서 추출 실패 시 공통 안내 문구
HWP_FAIL_MESSAGE = (
    "해당 한글 문서는 텍스트 추출이 어렵습니다. "
    "PDF 또는 DOCX로 변환 후 다시 업로드해주세요."
)

st.title("환경 정책 문서 Q&A 도우미")

# API Key 확인
# Streamlit Cloud 배포 환경에서는 st.secrets를, 로컬에서는 .env(os.getenv)를 사용한다.
# st.secrets는 secrets.toml이 없으면 예외를 던지므로 try로 감싼다.
def get_api_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY")


api_key = get_api_key()
if not api_key:
    st.error(
        "OPENAI_API_KEY를 찾을 수 없습니다. "
        "로컬 실행 시에는 프로젝트 폴더의 .env 파일에, "
        "Streamlit Cloud 배포 시에는 앱 설정의 Secrets에 "
        "OPENAI_API_KEY를 설정했는지 확인해주세요."
    )
    st.stop()

client = OpenAI(api_key=api_key)


# =====================================================================
# 문서 텍스트 추출 함수들
# =====================================================================
class DocExtractionError(Exception):
    """사용자에게 보여줄 안내 문구를 담은 추출 오류."""


def extract_pdf_text(data):
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    parts = [page.extract_text() or "" for page in reader.pages]
    return "\n".join(parts)


def extract_docx_text(data):
    from docx import Document

    document = Document(io.BytesIO(data))
    parts = [p.text for p in document.paragraphs]
    return "\n".join(parts)


def extract_hwpx_text(data):
    """HWPX는 zip 구조이므로 내부 section XML에서 텍스트(<hp:t>)를 모은다."""
    texts = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = sorted(
            n for n in z.namelist()
            if n.startswith("Contents/section") and n.endswith(".xml")
        )
        for name in names:
            root = ET.fromstring(z.read(name))
            for el in root.iter():
                # 네임스페이스가 붙어도 태그 끝이 't' 이면 텍스트 노드
                if el.tag.endswith("}t") or el.tag == "t":
                    if el.text:
                        texts.append(el.text)
            texts.append("\n")
    return "".join(texts)


def _parse_hwp_records(data):
    """HWP BodyText 레코드 스트림에서 문단 텍스트(tag 67)를 추출한다."""
    import struct

    out = []
    i = 0
    n = len(data)
    while i + 4 <= n:
        header = struct.unpack_from("<I", data, i)[0]
        i += 4
        tag_id = header & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            size = struct.unpack_from("<I", data, i)[0]
            i += 4
        rec = data[i:i + size]
        i += size
        if tag_id == 67:  # HWPTAG_PARA_TEXT
            out.append(_decode_hwp_para(rec))
    return "".join(out)


def _decode_hwp_para(rec):
    """문단 레코드(UTF-16LE + 제어문자)를 사람이 읽을 텍스트로 변환한다."""
    chars = []
    i = 0
    n = len(rec)
    while i + 1 < n:
        code = rec[i] | (rec[i + 1] << 8)
        i += 2
        if code in (10, 13):
            chars.append("\n")
        elif code == 9:
            chars.append("\t")
        elif code in (0, 24, 25, 26, 27, 28, 29, 30, 31):
            # 단일 워드 제어문자 → 무시
            continue
        elif code in (1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23):
            # 확장 제어문자 → 7워드(14바이트) 추가로 건너뜀
            i += 14
        else:
            chars.append(chr(code))
    return "".join(chars)


def extract_hwp_text(data):
    """HWP(OLE) 문서에서 텍스트 추출을 시도한다. 실패하면 빈 문자열."""
    import zlib
    import olefile

    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        # 압축 여부 확인 (FileHeader의 속성 플래그 비트 0)
        is_compressed = False
        if ole.exists("FileHeader"):
            header = ole.openstream("FileHeader").read()
            if len(header) > 36:
                is_compressed = bool(header[36] & 0x01)

        # BodyText/Section* 스트림 수집
        sections = sorted(
            entry for entry in ole.listdir()
            if len(entry) >= 2 and entry[0] == "BodyText"
        )
        parts = []
        for entry in sections:
            raw = ole.openstream(entry).read()
            if is_compressed:
                try:
                    raw = zlib.decompress(raw, -15)
                except Exception:
                    continue
            parts.append(_parse_hwp_records(raw))
        text = "".join(parts)

        # BodyText에서 못 얻으면 미리보기 텍스트(PrvText)로 폴백
        if not text.strip() and ole.exists("PrvText"):
            text = ole.openstream("PrvText").read().decode("utf-16-le", "ignore")
        return text
    finally:
        ole.close()


def extract_text_by_ext(name, data):
    """파일명과 바이트에서 (확장자, 텍스트)를 반환. 실패 시 DocExtractionError."""
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    try:
        if ext == "pdf":
            text = extract_pdf_text(data)
        elif ext == "docx":
            text = extract_docx_text(data)
        elif ext == "hwpx":
            text = extract_hwpx_text(data)
        elif ext == "hwp":
            text = extract_hwp_text(data)
        else:
            raise DocExtractionError(
                f"지원하지 않는 파일 형식입니다: .{ext} "
                "(PDF, DOCX, HWP, HWPX만 가능합니다.)"
            )
    except DocExtractionError:
        raise
    except Exception:
        # HWP/HWPX는 구조에 따라 추출이 실패할 수 있으므로 친절한 안내
        if ext in ("hwp", "hwpx"):
            raise DocExtractionError(HWP_FAIL_MESSAGE)
        raise DocExtractionError(
            f"문서에서 텍스트를 추출하지 못했습니다. 파일이 손상되었는지 확인해주세요. (형식: .{ext})"
        )

    # 한글 문서인데 내용이 비어 있으면 추출 실패로 간주
    if ext in ("hwp", "hwpx") and not text.strip():
        raise DocExtractionError(HWP_FAIL_MESSAGE)

    return ext, text


# =====================================================================
# 데이터 폴더에서 문서 목록 / 내용 불러오기
# =====================================================================
def list_documents(folder):
    """폴더 안의 지원되는 문서 파일명을 정렬해 반환한다."""
    if not os.path.isdir(folder):
        return []
    return sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith((".pdf", ".docx", ".hwp", ".hwpx"))
    )


@st.cache_data(show_spinner=False)
def load_document_text(path):
    """데이터 폴더의 문서 파일을 읽어 (확장자, 텍스트)를 반환한다.

    같은 파일은 캐시되어 질문할 때마다 다시 읽지 않는다.
    """
    with open(path, "rb") as f:
        data = f.read()
    return extract_text_by_ext(os.path.basename(path), data)


# =====================================================================
# 사이드바 UI - 답변 방식 선택
# =====================================================================
st.sidebar.subheader("답변 방식")
answer_mode = st.sidebar.radio(
    "어떤 방식으로 답변할까요?",
    ["문서 기반 답변", "웹 검색 답변"],
    label_visibility="collapsed",
)
st.sidebar.caption(
    "기본은 데이터 폴더의 보고서를 근거로 답변합니다. "
    "필요할 때만 ‘웹 검색 답변’을 사용해주세요."
)
use_web_search = answer_mode == "웹 검색 답변"


# =====================================================================
# 본문 상단 - 문서 선택 (문서 기반 답변일 때만)
# =====================================================================
st.caption(
    "데이터 폴더의 환경 정책 보고서를 선택하고, "
    "그 내용에 대해 질문하면 문서를 근거로 답변합니다."
)

doc_text = None
doc_name = None
if not use_web_search:
    col1, col2 = st.columns([1, 2])
    category_label = col1.selectbox("분야 선택", list(CATEGORIES.keys()))
    folder = os.path.join(DATA_DIR, CATEGORIES[category_label])

    files = list_documents(folder)
    if not files:
        st.warning(
            f"문서 폴더에서 파일을 찾지 못했습니다.\n\n경로: {folder}"
        )
    else:
        selected = col2.selectbox("문서 선택", files)
        path = os.path.join(folder, selected)
        try:
            ext, doc_text = load_document_text(path)
            doc_name = selected
            note = (
                f" (길어서 앞 {MAX_DOC_CHARS:,}자만 사용)"
                if len(doc_text) > MAX_DOC_CHARS else ""
            )
            st.success(f"선택한 문서: {selected}{note}")
        except DocExtractionError as e:
            st.error(str(e))

st.divider()


# =====================================================================
# 대화 기록
# =====================================================================
if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])


# =====================================================================
# 답변 생성 함수들
# =====================================================================
def extract_citations(response):
    """Responses API 응답에서 출처(URL) 목록을 추출한다."""
    citations = []
    seen = set()
    try:
        for item in response.output:
            if getattr(item, "type", None) != "message":
                continue
            for content in getattr(item, "content", []) or []:
                for ann in getattr(content, "annotations", []) or []:
                    if getattr(ann, "type", None) == "url_citation":
                        url = getattr(ann, "url", None)
                        if url and url not in seen:
                            seen.add(url)
                            title = getattr(ann, "title", None) or url
                            citations.append((title, url))
    except Exception:
        pass
    return citations


def generate_web_search_answer(prompt):
    """OpenAI Responses API의 web_search 도구로 답변을 생성한다."""
    response = client.responses.create(
        model=WEB_SEARCH_MODEL,
        tools=[{"type": "web_search"}],
        input=prompt,
    )
    answer = response.output_text
    citations = extract_citations(response)
    if citations:
        sources = "\n".join(f"- [{title}]({url})" for title, url in citations)
        answer = f"{answer}\n\n---\n**참고한 출처**\n{sources}"
    return answer


def generate_doc_answer(prompt, document_text, document_name):
    """선택한 문서 내용만 근거로 정해진 형식의 답변을 생성한다 (단순 방식, RAG 아님).

    문서를 user 메시지 안에 넣는다. system 프롬프트에 거절 문구를 직접 넣으면
    일부 소형 모델이 항상 거절 문구만 반환하는 문제가 있어 이 방식을 사용한다.
    """
    context = document_text[:MAX_DOC_CHARS]
    user_message = (
        "너는 환경 정책 보고서를 읽고 질문에 답하는 도우미야. "
        "아래 [문서]만 근거로 [질문]에 답해줘. 추측하지 말고, "
        "문서에 있는 내용만 사용해. 반드시 아래 형식을 그대로 지켜서 답해줘.\n\n"
        "### 요약 답변\n"
        "질문에 대한 핵심 답을 2~3문장으로 정리.\n\n"
        "### 문서에서 확인한 근거\n"
        "답변의 근거가 되는 문서 속 문장이나 수치를 짧게 인용. "
        "근거를 찾지 못했으면 '해당 근거를 찾지 못했습니다'라고 적어.\n\n"
        "### 추가로 확인할 점\n"
        "문서만으로는 부족해 추가로 확인이 필요한 부분을 적어. 없으면 '특이사항 없음'.\n\n"
        "단, 문서에서 질문과 관련된 내용을 전혀 찾을 수 없으면, "
        "위 형식을 쓰지 말고 다른 말 없이 "
        f"'{NOT_FOUND_MESSAGE}' 라고만 답해줘.\n\n"
        f"[문서 파일명] {document_name}\n"
        f"[문서]\n{context}\n\n"
        f"[질문] {prompt}"
    )
    response = client.chat.completions.create(
        model=CHAT_MODEL,
        messages=[{"role": "user", "content": user_message}],
    )
    answer = response.choices[0].message.content
    # 답변 맨 아래에 출처(파일명)를 표시한다.
    if NOT_FOUND_MESSAGE not in answer:
        answer = f"{answer}\n\n---\n**출처**: {document_name}"
    return answer


# =====================================================================
# 사용자 입력 처리
# =====================================================================
if prompt := st.chat_input("궁금한 점을 입력해주세요"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        try:
            if use_web_search:
                with st.spinner("웹을 검색하는 중입니다..."):
                    answer = generate_web_search_answer(prompt)
            elif not doc_text:
                # 문서 기반 답변인데 선택된 문서가 없을 때 안내
                answer = "먼저 위에서 분야와 문서를 선택해주세요."
            else:
                with st.spinner("문서 내용을 확인하는 중입니다..."):
                    answer = generate_doc_answer(prompt, doc_text, doc_name)
            st.markdown(answer)
            st.session_state.messages.append(
                {"role": "assistant", "content": answer}
            )
        except Exception as e:
            st.error(
                "답변을 생성하는 중 문제가 발생했습니다. "
                "잠시 후 다시 시도하거나, 사이드바 옵션을 바꿔서 시도해주세요.\n\n"
                f"(자세한 내용: {e})"
            )
