"""有效清单原行索引、Word反查及逐字段差异。行号、列号均从1起算。"""
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata
from source_coverage import fresh, resolve

HEADERS = {"item": ("序号","项号","清单序号"), "code": ("项目编码","清单编码"), "name": ("项目名称",),
           "feature": ("项目特征描述","项目特征"), "quantity": ("工程量","数量"), "unit": ("计量单位","单位")}
LABEL = dict(item="项号",code="编码",name="项目名称",feature="项目特征",quantity="数量",unit="单位")
MARK = re.compile(r"(?<!\d)\d{12}(?!\d)|第\s*\d+\s*项|(?:清单|项目)(?:序号|项号|编码)\s*[：:]?\s*[A-Za-z0-9-]+")
QUANT = re.compile(r"(\d+(?:\.\d+)?)\s*(m[²³23]?|㎡|㎥|米|平方(?:米)|立方(?:米)|项|套|樘|个|台|组|块|吨|kg)(?![a-z])", re.I)


def clean(value):
    if value is None:
        return ""
    if isinstance(value,float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC",str(value)))


def worksheets(path):
    if path.suffix.lower()==".xls":
        import xlrd
        book=xlrd.open_workbook(str(path))
        try:
            for ws in book.sheets():
                yield ws.name, [ws.row_values(i) for i in range(ws.nrows)]
        finally:
            book.release_resources()
    else:
        from openpyxl import load_workbook
        book=load_workbook(path,read_only=True,data_only=True)
        try:
            for ws in book:
                yield ws.title,list(ws.values)
        finally:
            book.close()


def index(state, base):
    sources = {s["id"]:s for s in state["startup"]["sources"]}
    specs = state["plan"]["hard_checks"]["bills"]
    result = {}
    for spec in specs:
        s = sources[spec["source"]]
        if s["role"] not in ("bill","boq") or not fresh(s,base) or s["sha256"]!=spec["sha256"]:
            raise ValueError("清单版本不匹配：" + spec["source"])
        source_path = resolve(s["path"],base)
        if source_path.suffix.lower() not in {".xls",".xlsx",".xlsm"}:
            copy=s.get("readable_copy",{})
            if (not fresh(copy,base) or copy.get("original_sha256")!=s["sha256"] or
                    copy.get("correspondence_reviewed") is not True or not copy.get("conclusion")):
                raise ValueError("原件需逐行对应核验的表格转换件："+str(source_path))
            source_path=resolve(copy["path"],base)
        for sheet, rows in worksheets(source_path):
            for tab in spec["tables"]:
                if sheet != tab["sheet"]:
                    continue
                title, trade = tab["title"], tab["trade"]
                if not title.strip() or not trade.strip():
                    raise ValueError("清单缺实际表名或专业")
                def cell(key):
                    r,c=tab[key]
                    if min(r,c)<1:
                        raise ValueError(key+"须为原表坐标")
                    return clean(rows[r-1][c-1])
                if (clean(title) not in clean(sheet) and clean(title) not in cell("title_cell")) or clean(trade) not in cell("trade_cell"):
                    raise ValueError(f"{sheet} 表名或专业与原表不符")
                key=(s["id"],sheet,trade,title)
                if key in result:
                    raise ValueError("有效表重复登记："+sheet)
                records=result[key]={}
                columns={}
                for n,row in enumerate(rows,1):
                    vals=[clean(v) for v in row]
                    hdr={k:next((i for i,v in enumerate(vals) if v in names),None) for k,names in HEADERS.items()}
                    if all(v is not None for v in hdr.values()):
                        columns=hdr
                    elif columns and vals[columns["item"]].isdigit() and vals[columns["code"]]:
                        records[n]={k:vals[col] for k,col in columns.items()}
                if not records:
                    raise ValueError(f"{sheet} 缺可定位表头或原行")
        for tab in spec["tables"]:
            if (s["id"],tab["sheet"],tab["trade"],tab["title"]) not in result:
                raise ValueError("未找到登记的有效清单表："+tab["sheet"])
    return result


def same(field, value, actual):
    value, actual = clean(value),clean(actual)
    if not value or not actual:
        return False
    if field == "quantity":
        try:
            return Decimal(value)==Decimal(actual)
        except InvalidOperation:
            return False
    if field=="feature":
        return value in actual
    return value==actual


def check(locations, claims, tables, versions, exclusions=()):
    errors, covered = [], {}
    names={r["name"] for table in tables.values() for r in table.values() if r["name"]}
    trades={k[2] for k in tables}
    bill_words=re.compile(r"(?:清单|计价表)(?:中|内|所列|列明|载明|列项|项目|摘录|引用|对应|[：:])")
    for c in claims:
        loc, quote = c["loc"], c["quote"]
        key = (c["source"],c["sheet"],c["trade"],c["title"])
        row = tables.get(key,{}).get(c["row"])
        prefix = f"Word {loc}「{quote}」 ↔ {c['source']}[{c['sha256'][:10]}]/{c['trade']}/{c['sheet']}/{c['title']} 第{c['row']}行"
        if not quote or quote not in locations.get(loc,""):
            errors.append(prefix+" 原文定位失败")
            continue
        if versions.get(c["source"])!=c["sha256"] or row is None:
            errors.append(prefix+" 版本、专业、表名或原行定位失败")
            continue
        fields=c["fields"]
        if not fields or fields.keys()-HEADERS.keys():
            errors.append(prefix+" 缺实际明示字段或字段名无效")
            continue
        for name in names:
            if name in quote and name!=row["name"] and name not in row["feature"] and row["name"].find(name)<0:
                errors.append(prefix+f" 项目名称 Word={name} 原行={row['name']}")
        if row["name"] in quote and fields.get("name")!=row["name"]:
            errors.append(prefix+" 项目名称已明示但未逐字段登记："+row["name"])
        for trade in trades:
            if trade != c["trade"] and trade in quote:
                errors.append(prefix+" Word所引专业为"+trade)
        for field,pattern in (("name",r"项目名称[：:]([^，。；|]+)"),("feature",r"项目特征(?:描述)?[：:]([^；。|]+)"),
                              ("unit",r"(?:计量)?单位[：:]([^，。；|\s]+)"),("quantity",r"(?:工程量|数量)[：:]?\s*(\d+(?:\.\d+)?)")):
            for m in re.finditer(pattern,quote):
                if not same(field,m[1],row[field]) or not same(field,m[1],fields.get(field,"")):
                    errors.append(prefix+f" {LABEL[field]} Word={m[1]} 原行={row[field]}")
        found = list(MARK.finditer(quote))
        for m in found:
            token=clean(m.group())
            k="code" if re.search(r"\d{12}",token) or "编码" in token else "item"
            digits=re.search(r"\d+",token)
            if not digits or clean(fields.get(k))!=digits.group():
                errors.append(prefix+f" {LABEL[k]} Word={token} 原行={row[k]}")
        for m in QUANT.finditer(quote):
            if m.start() and quote[m.start()-1]=="第" or any(a.start()<=m.start()<a.end() for a in found):
                continue
            # 数量、单位的明示值一起核查。工艺规格数字放在对应项目特征引文内。
            if m.group() in str(fields.get("feature","")):
                continue
            if not same("quantity",m[1],fields.get("quantity","")) or clean(m[2])!=clean(fields.get("unit","")):
                errors.append(prefix+f" 数量/单位 Word={m.group()} 原行={row['quantity']}{row['unit']}")
        for field, value in fields.items():
            token=clean(value)
            if token not in clean(quote) or not same(field,value,row[field]):
                errors.append(prefix+f" {LABEL[field]} Word={value} 原行={row[field]}")
        covered.setdefault(loc,[]).append(quote)
    for c in exclusions:
        if c["quote"] not in locations.get(c["loc"],"") or not c["reason"].strip() or len(c["reason"])<6 or re.search(r"清单|\d{12}",c["quote"]):
            errors.append("非清单项误排除："+str(c))
        else:
            covered.setdefault(c["loc"],[]).append(c["quote"])
    for loc, original in locations.items():
        rest=original
        for quote in covered.get(loc,[]):
            rest=rest.replace(quote," ",1)
        if MARK.search(rest) or bill_words.search(rest) or ("清单" in rest and any(n in rest for n in names)):
            errors.append(f"Word {loc}「{rest.strip()}」 ↔ 清单原行未定位／漏登明示项号或编码")
    return errors
