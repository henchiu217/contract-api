from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
import re, zipfile, io, os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 所有合約清單（顯示名稱 → 檔案名稱）
CONTRACTS = {
    # 代管
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
    # 包租
    "包租_包租契約書":                           "_包租_1_第5期-社會住宅-包租契約書範本_-1141001.docx",
    "包租_設備點交清單":                         "_包租_2_設備點交清單-包租-1130614.docx",
    "包租_包租止付":                             "_包租_3_包租止付.docx",
    "包租_補充協議書（包租止付）":               "_包租_3_補充協議書-包租止付.docx",
    # 轉租
    "轉租_轉租契約書":                           "_轉租_1_第5期社會住宅-轉租契約書範本-1141001.docx",
    "轉租_轉租契約書（雙租補）":                 "_轉租_1_第5期社會住宅-轉租契約書範本-1141001雙租補.docx",
    "轉租_設備點交清單":                         "_轉租_2_設備點交清單-轉租-1130614.docx",
    "轉租_補充協議書":                           "_轉租_3_補充協議書-1141118.docx",
    "轉租_補充協議書（雙租補）":                 "_轉租_3_補充協議書-1141118雙租補.docx",
    "轉租_補充協議書（有寵物）":                 "_轉租_3_補充協議書_有寵物_-1141118.docx",
    "轉租_補充協議書（有寵物雙租補）":           "_轉租_3_補充協議書_有寵物_-1141118雙租補.docx",
    "轉租_補充協議書（有車位）":                 "_轉租_3_補充協議書_有車位_-1141118.docx",
    "轉租_補充協議書（保證人）":                 "_轉租_8_補充協議書-保證人-1120703.docx",
    # 申請書
    "申請書_出租人出租住宅申請書（1141001）":    "1_第5期-表單1_出租人出租住宅申請書-1141001.docx",
    "申請書_出租人出租住宅申請書（1150521）":    "第5期-表單1_出租人出租住宅申請書-1150521.docx",
    "申請書_承租住宅申請書（房客）":             "1_第5期-表單5_民眾_房客_承租住宅申請書-1150101.docx",
    "申請書_承租住宅申請書（房客雙租補）":       "1_第5期-表單5_民眾_房客_承租住宅申請書-1150101房2雙租補.docx",
    # 聲明書
    "聲明書_承租人聲明書":                       "2_承租人聲明書-1141016.docx",
    "聲明書_承租人聲明書（雙租補）":             "2_承租人聲明書-1141016_雙租補.docx",
    "聲明書_出租人聲明書":                       "3_出租人聲明書-1141204.docx",
    # 其他
    "其他_屋況及租屋安全檢核表":                 "第5期-表單2_屋況及租屋安全檢核表_租賃標的現況確認書_-1150521.docx",
    "其他_補充協議書（代收付更改帳戶）":         "補充協議書-代收付更改帳戶.docx",
    "其他_補充協議書（共住人）":                 "補充協議書-共住人-1150109_轉租_.docx",
}

class ContractData(BaseModel):
    # 出租人
    房東姓名: Optional[str] = ""
    房東性別: Optional[str] = ""
    房東ID: Optional[str] = ""
    房東生日年: Optional[str] = ""
    房東生日月: Optional[str] = ""
    房東生日日: Optional[str] = ""
    房東手機: Optional[str] = ""
    房東縣市: Optional[str] = ""
    房東鄉市鎮區: Optional[str] = ""
    房東街路: Optional[str] = ""
    房東段: Optional[str] = ""
    房東巷: Optional[str] = ""
    房東弄: Optional[str] = ""
    房東號: Optional[str] = ""
    房東樓: Optional[str] = ""
    房東之: Optional[str] = ""
    房東戶籍地完整: Optional[str] = ""
    房東金融機構: Optional[str] = ""
    房東金融機構代碼: Optional[str] = ""
    房東分行: Optional[str] = ""
    房東分行代碼: Optional[str] = ""
    房東帳號: Optional[str] = ""
    # 承租人
    房客姓名: Optional[str] = ""
    房客性別: Optional[str] = ""
    房客ID: Optional[str] = ""
    房客戶號: Optional[str] = ""
    房客手機: Optional[str] = ""
    房客生日年: Optional[str] = ""
    房客生日月: Optional[str] = ""
    房客生日日: Optional[str] = ""
    房客年月日: Optional[str] = ""
    房客縣市: Optional[str] = ""
    房客鄉市鎮區: Optional[str] = ""
    房客街路: Optional[str] = ""
    房客段: Optional[str] = ""
    房客巷: Optional[str] = ""
    房客弄: Optional[str] = ""
    房客號: Optional[str] = ""
    房客樓: Optional[str] = ""
    房客之: Optional[str] = ""
    房客戶籍地: Optional[str] = ""
    未成年子女: Optional[str] = ""
    房客金融機構: Optional[str] = ""
    房客金融機構代碼: Optional[str] = ""
    房客分行: Optional[str] = ""
    房客分行代碼: Optional[str] = ""
    房客帳號: Optional[str] = ""
    # 雙租補（房2）
    房2姓名: Optional[str] = ""
    房2性別: Optional[str] = ""
    房2ID: Optional[str] = ""
    房2戶號: Optional[str] = ""
    房2手機: Optional[str] = ""
    房2生日年: Optional[str] = ""
    房2生日月: Optional[str] = ""
    房2生日日: Optional[str] = ""
    房2年月日: Optional[str] = ""
    房2縣市: Optional[str] = ""
    房2鄉市鎮區: Optional[str] = ""
    房2街路: Optional[str] = ""
    房2段: Optional[str] = ""
    房2巷: Optional[str] = ""
    房2弄: Optional[str] = ""
    房2號: Optional[str] = ""
    房2樓: Optional[str] = ""
    房2之: Optional[str] = ""
    房2戶籍地: Optional[str] = ""
    房2金融機構: Optional[str] = ""
    房2金融機構代碼: Optional[str] = ""
    房2分行: Optional[str] = ""
    房2分行代碼: Optional[str] = ""
    房2帳號: Optional[str] = ""
    # 保證人
    保ID: Optional[str] = ""
    保手機: Optional[str] = ""
    保戶籍地: Optional[str] = ""
    # 房屋資料
    完整承租地址: Optional[str] = ""
    鄉鎮市區: Optional[str] = ""
    街路: Optional[str] = ""
    段: Optional[str] = ""
    巷: Optional[str] = ""
    弄: Optional[str] = ""
    號: Optional[str] = ""
    樓: Optional[str] = ""
    之: Optional[str] = ""
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
    謄鄉鎮市區: Optional[str] = ""
    謄街路: Optional[str] = ""
    謄段: Optional[str] = ""
    謄巷: Optional[str] = ""
    謄弄: Optional[str] = ""
    謄號: Optional[str] = ""
    謄樓: Optional[str] = ""
    謄之: Optional[str] = ""
    純租金: Optional[str] = ""
    押金月: Optional[str] = ""
    押金金額: Optional[str] = ""
    管理費每月多少: Optional[str] = ""
    管理費坪元: Optional[str] = ""
    房客2ID: Optional[str] = ""

class GenerateRequest(BaseModel):
    contracts: list[str]   # 要產生的合約名稱清單
    data: ContractData

def replace_ph(xml, name, value):
    if not value:
        return xml
    pat = r'\{\{[^{]{0,50}?' + re.escape(name) + r'[^}]{0,50}?\}\}'
    def repl(m):
        orig = m.group(0)
        rpr_m = re.search(r'<w:rPr>.*?</w:rPr>', orig, re.DOTALL)
        rpr = rpr_m.group(0) if rpr_m else ''
        safe = value.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        return f'<w:r>{rpr}<w:t xml:space="preserve">{safe}</w:t></w:r>'
    return re.sub(pat, repl, xml, flags=re.DOTALL)

def fill_contract(filename, data_dict):
    path = filename
    if not os.path.exists(path):
        return None
    with open(path, 'rb') as f:
        raw = f.read()

    zin = zipfile.ZipFile(io.BytesIO(raw))
    doc_xml = zin.read('word/document.xml').decode('utf-8')

    # 策略：提取所有 w:t 文字組合後替換，處理跨 run 的佔位符
    wt_pattern = r'(<w:t[^>]*>)(.*?)(</w:t>)'
    combined = ''
    positions = []
    for m in re.finditer(wt_pattern, doc_xml, re.DOTALL):
        text = m.group(2)
        start_in_xml = m.start(2)
        combined += text
        for i in range(len(text)):
            positions.append(start_in_xml + i)

    result = doc_xml
    offset = 0
    for name, value in data_dict.items():
        if not value:
            continue
        target = '{{' + name + '}}'
        idx = combined.find(target)
        if idx < 0:
            continue
        safe = str(value).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        xml_start = positions[idx]
        xml_end = positions[idx + len(target) - 1] + 1
        result = result[:xml_start + offset] + safe + result[xml_end + offset:]
        offset_diff = len(safe) - len(target)
        offset += offset_diff
        new_combined = combined[:idx] + value + combined[idx+len(target):]
        new_positions = positions[:idx] + [positions[idx]] * len(value) + positions[idx+len(target):]
        combined = new_combined
        positions = new_positions

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, result.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))
    zin.close()
    out.seek(0)
    return out

@app.get("/")
def root():
    return {"status": "合約 API v2 運行中 ✅", "contracts": list(CONTRACTS.keys())}

@app.get("/contracts")
def list_contracts():
    return {"contracts": list(CONTRACTS.keys())}

@app.post("/generate")
def generate(req: GenerateRequest):
    data_dict = {k: v for k, v in req.data.dict().items() if v}

    # 產生多份合約，壓縮成 ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for contract_name in req.contracts:
            filename = CONTRACTS.get(contract_name)
            if not filename:
                continue
            filled = fill_contract(filename, data_dict)
            if filled:
                safe_name = contract_name.replace("/", "_") + ".docx"
                zout.writestr(safe_name, filled.read())

    zip_buf.seek(0)
    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''contracts.zip"}
    )

# 單份合約下載（相容舊版）
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
