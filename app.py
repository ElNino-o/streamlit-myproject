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

# 한글 문서 추출 실패 시 공통 안내 문구
HWP_FAIL_MESSAGE = (
    "해당 한글 문서는 텍스트 추출이 어렵습니다. "
    "PDF 또는 DOCX로 변환 후 다시 업로드해주세요."
)

st.title("나의 AI 챗봇")

# API Key 확인
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    st.error(
        "OPENAI_API_KEY를 찾을 수 없습니다. "
        "프로젝트 폴더의 .env 파일에 OPENAI_API_KEY를 설정했는지 확인해주세요."
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


def extract_document_text(uploaded_file):
    """업로드 파일에서 (확장자, 텍스트)를 반환. 실패 시 DocExtractionError."""
    name = uploaded_file.name
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    data = uploaded_file.getvalue()

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
# 사이드바 UI
# =====================================================================
use_web_search = st.sidebar.checkbox("웹검색 사용하기", value=False)
use_doc_qa = st.sidebar.checkbox("문서 기반 답변 사용하기", value=False)
st.sidebar.caption(
    "‘문서 기반 답변’을 켜고 파일을 업로드하면, 업로드한 문서 내용을 참고해 답변합니다."
)

uploaded_file = st.sidebar.file_uploader(
    "문서 업로드 (PDF, DOCX, HWP, HWPX)",
    type=["pdf", "docx", "hwp", "hwpx"],
)

# 업로드된 문서에서 텍스트 추출 (같은 파일은 한 번만 처리하도록 캐시)
doc_text = None
if uploaded_file is not None:
    cache_key = (uploaded_file.name, uploaded_file.size)
    if st.session_state.get("doc_cache_key") != cache_key:
        try:
            ext, text = extract_document_text(uploaded_file)
            st.session_state["doc_cache_key"] = cache_key
            st.session_state["doc_text"] = text
            st.session_state["doc_ext"] = ext
            st.session_state["doc_name"] = uploaded_file.name
            st.session_state["doc_error"] = None
        except DocExtractionError as e:
            st.session_state["doc_cache_key"] = cache_key
            st.session_state["doc_text"] = None
            st.session_state["doc_error"] = str(e)

    if st.session_state.get("doc_error"):
        st.sidebar.error(st.session_state["doc_error"])
    elif st.session_state.get("doc_text") is not None:
        doc_text = st.session_state["doc_text"]
        st.sidebar.success("문서 업로드 완료")
        st.sidebar.markdown(
            f"- **파일명**: {st.session_state['doc_name']}\n"
            f"- **형식**: {st.session_state['doc_ext'].upper()}\n"
            f"- **추출된 텍스트 길이**: {len(doc_text):,}자"
        )
        if len(doc_text) > MAX_DOC_CHARS:
            st.sidebar.caption(
                f"문서가 길어 앞 {MAX_DOC_CHARS:,}자만 답변에 사용합니다."
            )


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
def generate_chat_answer(messages):
    """기존 챗봇 방식 (Chat Completions API)."""
    response = client.chat.completions.create(
        model=CHAT_MODEL,
        messages=[{"role": m["role"], "content": m["content"]} for m in messages],
    )
    return response.choices[0].message.content


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


def generate_doc_answer(prompt, document_text):
    """업로드한 문서 내용만 근거로 답변을 생성한다 (단순 방식, RAG 아님).

    문서를 user 메시지 안에 넣는다. system 프롬프트에 거절 문구를 직접 넣으면
    일부 소형 모델이 항상 거절 문구만 반환하는 문제가 있어 이 방식을 사용한다.
    """
    context = document_text[:MAX_DOC_CHARS]
    user_message = (
        "아래 [문서]를 읽고 [질문]에 답해줘. "
        "문서에 근거가 있으면 그 내용을 바탕으로 답하고, "
        "문서에서 찾을 수 없는 내용이면 다른 말 없이 "
        "'업로드된 문서에서 확인하기 어렵습니다' 라고만 답해줘.\n\n"
        f"[문서]\n{context}\n\n[질문] {prompt}"
    )
    response = client.chat.completions.create(
        model=CHAT_MODEL,
        messages=[{"role": "user", "content": user_message}],
    )
    return response.choices[0].message.content


# =====================================================================
# 사용자 입력 처리
# =====================================================================
if prompt := st.chat_input("궁금한 점을 입력해주세요"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        try:
            if use_doc_qa:
                # 문서 기반 답변: 문서가 없으면 안내
                if not doc_text:
                    answer = "먼저 문서를 업로드해주세요."
                    st.markdown(answer)
                    st.session_state.messages.append(
                        {"role": "assistant", "content": answer}
                    )
                else:
                    with st.spinner("문서 내용을 확인하는 중입니다..."):
                        answer = generate_doc_answer(prompt, doc_text)
                    st.markdown(answer)
                    st.session_state.messages.append(
                        {"role": "assistant", "content": answer}
                    )
            elif use_web_search:
                with st.spinner("웹을 검색하는 중입니다..."):
                    answer = generate_web_search_answer(prompt)
                st.markdown(answer)
                st.session_state.messages.append(
                    {"role": "assistant", "content": answer}
                )
            else:
                with st.spinner("답변을 생성하는 중입니다..."):
                    answer = generate_chat_answer(st.session_state.messages)
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
