from __future__ import annotations
import base64, json, mimetypes, os
from pathlib import Path

SYSTEM = """You extract blind/curtain manufacturing purchase orders into JSON.
Never invent missing measurements or product codes. Unknown values must be null.
Return an object with keys 거래처, 주문번호, 고객명, 전체기재사항, 배송, items.
Each item may include 품목코드, 종류, 타입, 가로, 세로, 창개수, 좌개수, 우개수,
손잡이방향, 손잡이길이, 연창, 기재사항, 설치장소, 특이.
"""

def extract_with_openai(path: str | Path, forced_client: str | None = None):
    from openai import OpenAI
    p = Path(path)
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    model = os.getenv("FITORDER_MODEL","gpt-5-mini")
    mime = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
    if mime.startswith("image/"):
        data = base64.b64encode(p.read_bytes()).decode()
        user_content = [
            {"type":"input_text","text":f"거래처 힌트: {forced_client or '자동판별'}"},
            {"type":"input_image","image_url":f"data:{mime};base64,{data}"},
        ]
    else:
        raise ValueError("AI fallback은 현재 이미지에만 연결되어 있습니다. PDF는 렌더링 후 전달하도록 확장하세요.")
    r = client.responses.create(
        model=model,
        input=[{"role":"system","content":SYSTEM},{"role":"user","content":user_content}],
    )
    text = r.output_text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.split("\n",1)[-1]
        if text.endswith("```"): text = text[:-3]
    obj = json.loads(text)
    if forced_client:
        obj["거래처"] = forced_client
    obj["_source_path"] = str(p.resolve())
    return [obj]
