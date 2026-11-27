"""Check the bytes of a project-controlled anonymous submission and its reviewed visual objects."""
from hashlib import sha256
from pathlib import Path
from xml.etree import ElementTree as ET
import unicodedata
import zipfile

from bid_requirements import meaningful


def folded(text):
    return "".join(unicodedata.normalize("NFKC",str(text)).casefold().split())


def surfaces(path):
    """Yield location, readable text, and the digest of any object requiring visual confirmation."""
    path=Path(path)
    if path.suffix.lower()==".pdf":
        import fitz
        with fitz.open(path) as doc:
            yield "文件属性"," ".join(str(v or "") for v in doc.metadata.values()),None
            for i,page in enumerate(doc,1):
                yield f"第{i}页",page.get_text()+" "+" ".join(str(a.info) for a in (page.annots() or [])),sha256(page.get_pixmap().samples).hexdigest()
        return
    if path.suffix.lower()!=".docx":
        raise ValueError("匿名核验只接受实际DOCX或PDF")
    with zipfile.ZipFile(path) as z:
        for member in z.namelist():
            if member.startswith("word/media/"):
                yield member,"",sha256(z.read(member)).hexdigest()
            elif member.endswith(".xml") and member.startswith(("word/","docProps/","customXml/")):
                root=ET.fromstring(z.read(member))
                paragraphs=[el for el in root.iter() if el.tag.rsplit("}",1)[-1] in ("p","coreProperties","Properties","comment")]
                if not paragraphs:
                    paragraphs=[root]
                for i,el in enumerate(paragraphs,1):
                    # Join split runs before matching a complete company name.
                    value="".join(el.itertext())+" "+" ".join(v for c in el.iter() for v in c.attrib.values())
                    yield f"{member}#{i}",value,None
                if member.startswith("word/") and any(el.tag.rsplit("}",1)[-1] in ("drawing","pict","object") for el in root.iter()):
                    yield member+"#版面","",sha256(z.read(member)).hexdigest()


def errors(path,fmt,record,logical_name=None):
    if fmt["anonymous"] is False:
        return []
    tokens=fmt["identifiers"]
    checks=record.get("visual",{})
    problems=[]
    if not tokens or record.get("sha256")!=sha256(Path(path).read_bytes()).hexdigest():
        return ["暗标核验缺身份词或未绑定本次提交字节"]
    for loc,value,image in [("文件名",logical_name or Path(path).name,None),*surfaces(path)]:
        permits=fmt.get("permitted",{}).get(loc,[])
        for token in tokens:
            if folded(token) in folded(value) and token not in permits:
                problems.append("暗标身份命中："+loc+" / "+token)
        if image:
            finding=checks.get(loc,{})
            if finding.get("sha256")!=image or finding.get("clean") is not True or not meaningful(finding.get("finding")):
                problems.append("暗标图像/版面缺同稿人工核对："+loc)
    return problems
