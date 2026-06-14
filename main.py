from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import re, zipfile, io

app = FastAPI()

# 允許前端跨域呼叫
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class ContractData(BaseModel):
    房東姓名: str = ""
    房客姓名: str = ""
    完整承租地址: str = ""
    地段: str = ""
    小段: str = ""
    地號: str = ""
    建號: str = ""
    建物平方公尺: str = ""
    純租金: str = ""
    房東金融機構: str = ""
    房東分行: str = ""
    房東帳號: str = ""
    押金月: str = "2"
    押金金額: str = ""
    管理費每月多少: str = ""
    房東ID: str = ""
    房東戶籍地完整: str = ""
    房東手機: str = ""
    房客ID: str = ""
    房客戶籍地: str = ""
    房客手機: str = ""

def replace_ph(xml, name, value):
    pat = r'\{\{[^{]{0,50}?' + re.escape(name) + r'[^}]{0,50}?\}\}'
    def repl(m):
        orig = m.group(0)
        rpr_m = re.search(r'<w:rPr>.*?</w:rPr>', orig, re.DOTALL)
        rpr = rpr_m.group(0) if rpr_m else ''
        safe_value = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f'<w:r>{rpr}<w:t xml:space="preserve">{safe_value}</w:t></w:r>'
    return re.sub(pat, repl, xml, flags=re.DOTALL)

@app.get("/")
def root():
    return {"status": "合約 API 運行中 ✅"}

@app.post("/generate-contract")
def generate_contract(data: ContractData):
    # 讀取範本
    with open("template.docx", "rb") as f:
        raw = f.read()

    zin = zipfile.ZipFile(io.BytesIO(raw))
    doc_xml = zin.read("word/document.xml").decode("utf-8")

    # 填入所有欄位
    fields = data.dict()
    new_xml = doc_xml
    for name, value in fields.items():
        if value:
            new_xml = replace_ph(new_xml, name, str(value))

    # 打包輸出
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_xml.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()

    out_buf.seek(0)
    filename = f"租賃契約書_{data.房東姓名}_{data.房客姓名}.docx"

    return StreamingResponse(
        out_buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"}
    )
