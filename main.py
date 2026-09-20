from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
import re, zipfile, io, os, tempfile, shutil, subprocess

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

CONTRACTS = {
    "代管_委託租賃契約書":                       "_代管_1_第5期-社會住宅代租代管計畫-委託租賃契約書範本-1141001.docx",
    "代管_委託管理契約書":                       "_代管_2_第5期-社會住宅代租代管計畫-委託管理契約書範本-1141001.docx",
    "代管_社會住宅租賃契約書（不公證）":         "_代管_3_0第5期-社會住宅-租賃契約書範本_不公證_-1141001.docx",
    "代管_社會住宅租賃契約書（不公證雙租補）":   "_代管_3_0第5期-社會住宅-租賃契約書範本_不公證_-1141001_-_雙租補.docx",
    "代管_社會住宅租賃契約書（公證）":           "_代管_3_0第5期-社會住宅-租賃契約書範本_公證_-1141001.docx",
    "代管_社會住宅租賃契約書（公證雙租補）":     "_代管_3_0第5期-社會住宅-租賃契約書範本_公證_-1141001_-_雙租補.docx",
    "代管_設備點交清單":                         "_代管_3_1設備點交清單-代管-1130614.docx",
    "代管_補充協議書":                           "_代管_3_2補充協議書-1141118.docx",
    "代管_補充協議書（雙租補）":                 "_代管_3_2補充協議書-1141118_-_雙租補.docx",
    "代管_補充協議書（有車位）":                 "_代管_3_2補充協議書-1141118_有車位_.docx",
    "代管_補充協議書（有寵物）":                 "_代管_3_2補充協議書_有寵物_-1141118.docx",
    "代管_補充協議書（有寵物雙租補）":           "_代管_3_2補充協議書_有寵物_-1141118_-_雙租補.docx",
    "代管_補充協議書（保證人）":                 "_代管_補充協議書-保證人-1120703.docx",
    "代管_補充協議書（代理人）":                 "_代管_補充協議書-代理人-1120703.docx",
    "包租_包租契約書":                           "_包租_1_第5期-社會住宅-包租契約書範本_-1141001.docx",
    "包租_設備點交清單":                         "_包租_2_設備點交清單-包租-1130614.docx",
    "包租_包租止付":                             "_包租_3_包租止付.docx",
    "包租_補充協議書（包租止付）":               "_包租_3_補充協議書-包租止付.docx",
    "轉租_轉租契約書":                           "_轉租_1_第5期社會住宅-轉租契約書範本-1141001.docx",
    "轉租_轉租契約書（雙租補）":                 "_轉租_1_第5期社會住宅-轉租契約書範本-1141001雙租補.docx",
    "轉租_設備點交清單":                         "_轉租_2_設備點交清單-轉租-1130614.docx",
    "轉租_補充協議書":                           "_轉租_3_補充協議書-1141118.docx",
    "轉租_補充協議書（雙租補）":                 "_轉租_3_補充協議書-1141118雙租補.docx",
    "轉租_補充協議書（有寵物）":                 "_轉租_3_補充協議書_有寵物_-1141118.docx",
    "轉租_補充協議書（有寵物雙租補）":           "_轉租_3_補充協議書_有寵物_-1141118雙租補.docx",
    "轉租_補充協議書（有車位）":                 "_轉租_3_補充協議書_有車位_-1141118.docx",
    "轉租_補充協議書（保證人）":                 "_轉租_8_補充協議書-保證人-1120703.docx",
    "申請書_出租人出租住宅申請書（1141001）":    "1_第5期-表單1_出租人出租住宅申請書-1141001.docx",
    "申請書_出租人出租住宅申請書（1150521）":    "第5期-表單1_出租人出租住宅申請書-1150521.docx",
    "申請書_承租住宅申請書（房客）":             "1_第5期-表單5_民眾_房客_承租住宅申請書-1150101.docx",
    "申請書_承租住宅申請書（房客雙租補）":       "1_第5期-表單5_民眾_房客_承租住宅申請書-1150101房2雙租補.docx",
    "聲明書_承租人聲明書":                       "2_承租人聲明書-1141016.docx",
    "聲明書_承租人聲明書（雙租補）":             "2_承租人聲明書-1141016_雙租補.docx",
    "聲明書_出租人聲明書":                       "3_出租人聲明書-1141204.docx",
    "其他_屋況及租屋安全檢核表":                 "第5期-表單2_屋況及租屋安全檢核表_租賃標的現況確認書_-1150521.docx",
    "其他_補充協議書（代收付更改帳戶）":         "補充協議書-代收付更改帳戶.docx",
    "其他_補充協議書（共住人）":                 "補充協議書-共住人-1150109_轉租_.docx",
}

class ContractData(BaseModel):
    房東姓名: Optional[str] = ""
    房東ID: Optional[str] = ""
    房東手機: Optional[str] = ""
    房東縣市: Optional[str] = ""
    房東鄉市鎮區: Optional[str] = ""
    房東街路: Optional[str] = ""
    房東戶籍地完整: Optional[str] = ""
    房東金融機構: Optional[str] = ""
    房東金融機構代碼: Optional[str] = ""
    房東分行: Optional[str] = ""
    房東分行代碼: Optional[str] = ""
    房東帳號: Optional[str] = ""
    房客姓名: Optional[str] = ""
    房客ID: Optional[str] = ""
    房客手機: Optional[str] = ""
    房客戶號: Optional[str] = ""
    房客縣市: Optional[str] = ""
    房客鄉市鎮區: Optional[str] = ""
    房客街路: Optional[str] = ""
    房客生日年: Optional[str] = ""
    房客生日月: Optional[str] = ""
    房客生日日: Optional[str] = ""
    房客年月日: Optional[str] = ""
    房客戶籍地: Optional[str] = ""
    未成年子女: Optional[str] = ""
    房客金融機構: Optional[str] = ""
    房客金融機構代碼: Optional[str] = ""
    房客分行: Optional[str] = ""
    房客分行代碼: Optional[str] = ""
    房客帳號: Optional[str] = ""
    房2姓名: Optional[str] = ""
    房2ID: Optional[str] = ""
    房2手機: Optional[str] = ""
    房2戶號: Optional[str] = ""
    房2戶籍地: Optional[str] = ""
    房2生日年: Optional[str] = ""
    房2生日月: Optional[str] = ""
    房2生日日: Optional[str] = ""
    房2年月日: Optional[str] = ""
    房2金融機構: Optional[str] = ""
    房2金融機構代碼: Optional[str] = ""
    房2分行: Optional[str] = ""
    房2分行代碼: Optional[str] = ""
    房2帳號: Optional[str] = ""
    保ID: Optional[str] = ""
    保手機: Optional[str] = ""
    保戶籍地: Optional[str] = ""
    完整承租地址: Optional[str] = ""
    鄉鎮市區: Optional[str] = ""
    街路: Optional[str] = ""
    段: Optional[str] = ""
    巷: Optional[str] = ""
    弄: Optional[str] = ""
    號: Optional[str] = ""
    樓: Optional[str] = ""
    地段: Optional[str] = ""
    小段: Optional[str] = ""
    地號: Optional[str] = ""
    建號: Optional[str] = ""
    建物平方公尺: Optional[str] = ""
    建物坪數: Optional[str] = ""
    建築年: Optional[str] = ""
    建築月: Optional[str] = ""
    建築日: Optional[str] = ""
    層數: Optional[str] = ""
    層次: Optional[str] = ""
    純租金: Optional[str] = ""
    押金月: Optional[str] = ""
    押金金額: Optional[str] = ""
    管理費每月多少: Optional[str] = ""
    管理費坪元: Optional[str] = ""

class GenerateRequest(BaseModel):
    contracts: list
    data: ContractData

def fill_contract(filename, data_dict):
    """填入合約佔位符，直接操作 ZIP/XML，不依賴外部腳本"""
    if not os.path.exists(filename):
        return None
    try:
        with open(filename, "rb") as f:
            raw = f.read()

        zin = zipfile.ZipFile(io.BytesIO(raw))
        doc_xml = zin.read("word/document.xml").decode("utf-8")

        def safe(v):
            return str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

        # 步驟1：替換 MERGEFIELD 欄位
        def replace_mergefield(x, name, value):
            pat = (
                r'<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?\s*'
                r'<w:fldChar[^>]*w:fldCharType=["\']begin["\'][^/]*/>\s*</w:r>'
                r'.*?MERGEFIELD\s+' + re.escape(name) + r'[\s\\]*.*?'
                r'<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?\s*'
                r'<w:fldChar[^>]*w:fldCharType=["\']end["\'][^/]*/>\s*</w:r>'
            )
            repl = f'<w:r><w:t xml:space="preserve">{safe(value)}</w:t></w:r>'
            return re.sub(pat, repl, x, flags=re.DOTALL)

        # 步驟2：替換跨 run 的 {{ }} — 合併所有 w:t 文字後替換
        def replace_split(x, replacements):
            wt_re = re.compile(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', re.DOTALL)
            result = x
            for name, value in replacements.items():
                if not value:
                    continue
                target = "{{" + name + "}}"
                sv = safe(value)
                # 先試直接替換（佔位符在單一 w:t 內）
                if target in result:
                    result = result.replace(target, sv)
                    continue
                # 再試跨 run 替換
                entries = [(m.start(2), m.end(2), m.group(2)) for m in wt_re.finditer(result)]
                full = "".join(e[2] for e in entries)
                idx = full.find(target)
                if idx < 0:
                    continue
                char_pos = 0
                r_start = r_end = None
                for xs, xe, txt in entries:
                    end_pos = char_pos + len(txt)
                    if r_start is None and end_pos > idx:
                        r_start = xs + (idx - char_pos)
                    if r_start is not None and end_pos >= idx + len(target):
                        r_end = xs + (idx + len(target) - char_pos)
                        break
                    char_pos = end_pos
                if r_start is not None and r_end is not None:
                    result = result[:r_start] + sv + result[r_end:]
            return result

        # 先處理 MERGEFIELD，再處理 {{ }}
        for name, value in data_dict.items():
            if value:
                doc_xml = replace_mergefield(doc_xml, name, value)
        doc_xml = replace_split(doc_xml, data_dict)

        # 重新打包
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, doc_xml.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
        zin.close()
        out.seek(0)
        return out
    except Exception as e:
        print(f"fill_contract error: {e}")
        return None

@app.get("/")
def root():
    return {"status": "合約 API v2 ✅", "contracts": len(CONTRACTS)}

@app.get("/contracts")
def list_contracts():
    return {"contracts": list(CONTRACTS.keys())}

@app.post("/generate")
def generate(req: GenerateRequest):
    data_dict = {k: v for k, v in req.data.dict().items() if v}
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in req.contracts:
            filename = CONTRACTS.get(name)
            if not filename:
                continue
            filled = fill_contract(filename, data_dict)
            if filled:
                safe_name = name.replace("/", "_") + ".docx"
                zout.writestr(safe_name, filled.read())
    zip_buf.seek(0)
    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''contracts.zip"}
    )

@app.post("/generate-contract")
def generate_single(data: ContractData):
    filename = "_代管_3_0第5期-社會住宅-租賃契約書範本_不公證_-1141001.docx"
    data_dict = {k: v for k, v in data.dict().items() if v}
    filled = fill_contract(filename, data_dict)
    if not filled:
        return JSONResponse(status_code=404, content={"error": "找不到範本"})
    return StreamingResponse(
        filled,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''contract.docx"}
    )
