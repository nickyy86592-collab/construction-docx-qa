"""私有 Drive 送审、回执和整批发布。发布身份由受控运行器提供。"""
from __future__ import annotations

import json
import mimetypes
import re
import os
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from uuid import uuid4

from release_task import GateError, canonical, file_sha, require, sha256

DRIVE = "https://www.googleapis.com/drive/v3"
UPLOAD = "https://www.googleapis.com/upload/drive/v3"
FOLDER = "application/vnd.google-apps.folder"
META = "id,name,mimeType,parents,sha256Checksum,capabilities,trashed,appProperties,webViewLink"


class DriveError(GateError):
    def __init__(self, status, reasons=()):
        super().__init__(f"Drive 请求失败（HTTP {status}）")
        self.status = status
        self.reasons = reasons


class Drive:
    def __init__(self, credential):
        form = {k: credential[k] for k in ("client_id", "refresh_token")}
        if credential.get("client_secret"):
            form["client_secret"] = credential["client_secret"]
        form["grant_type"] = "refresh_token"
        self.form = urlencode(form).encode()
        self.token = self.refresh()

    @classmethod
    def from_env(cls, name):
        return cls(json.loads(os.environ[name]))

    def refresh(self):
        request = Request("https://oauth2.googleapis.com/token", data=self.form,
                          headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with urlopen(request, timeout=30) as response:
                return json.load(response)["access_token"]
        except HTTPError as exc:
            raise DriveError(exc.code) from exc

    def request(self, path, method="GET", params=None, body=None, kind=None, upload=False, raw=False):
        url = (UPLOAD if upload else DRIVE) + path
        if params:
            url += "?" + urlencode(params)
        for attempt in range(2):
            headers = {"Authorization": "Bearer " + self.token}
            if kind:
                headers["Content-Type"] = kind
            request = Request(url, data=body, headers=headers, method=method)
            try:
                with urlopen(request, timeout=180) as response:
                    data = response.read()
                    return data if raw else json.loads(data or b"{}")
            except HTTPError as exc:
                if exc.code == 401 and attempt == 0:
                    self.token = self.refresh()
                    continue
                try:
                    reasons = tuple(x.get("reason") for x in json.load(exc).get("error", {}).get("errors", ()))
                except (ValueError, AttributeError):
                    reasons = ()
                raise DriveError(exc.code, reasons) from exc

    def uid(self):
        return self.request("/about", params={"fields": "user(permissionId)"})["user"]["permissionId"]

    def meta(self, fid):
        return self.request("/files/" + checked_id(fid), params={"fields": META, "supportsAllDrives": "true"})

    def permissions(self, fid):
        params = {"fields": "nextPageToken,permissions(type,role,id,deleted)",
                  "supportsAllDrives": "true", "pageSize": 100}
        while True:
            result = self.request("/files/" + checked_id(fid) + "/permissions", params=params)
            yield from result.get("permissions", ())
            if not result.get("nextPageToken"):
                break
            params["pageToken"] = result["nextPageToken"]

    def private(self, fid):
        acl = list(self.permissions(fid))
        require(not any(x["type"] in ("anyone", "domain") and not x.get("deleted") for x in acl),
                "项目路径必须采用私有的 Drive 访问范围")
        return acl

    def children(self, fid):
        params = {"q": f"'{checked_id(fid)}' in parents and trashed = false",
                  "fields": "nextPageToken,files(" + META + ")", "pageSize": 100,
                  "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}
        while True:
            result = self.request("/files", params=params)
            yield from result.get("files", ())
            if not result.get("nextPageToken"):
                break
            params["pageToken"] = result["nextPageToken"]

    def download(self, fid):
        return self.request("/files/" + checked_id(fid),
                            params={"alt": "media", "supportsAllDrives": "true"}, raw=True)

    def create(self, name, parent, media=None, props=None):
        metadata = {"name": name, "parents": [checked_id(parent)]}
        if props:
            metadata["appProperties"] = props
        params = {"fields": META, "supportsAllDrives": "true"}
        if media is None:
            metadata["mimeType"] = FOLDER
            return self.request("/files", method="POST", params=params, body=canonical(metadata),
                                kind="application/json; charset=UTF-8")
        boundary = uuid4().hex
        part = b"--" + boundary.encode()
        content_type = (mimetypes.guess_type(name)[0] or "application/octet-stream").encode("ascii")
        body = (part + b"\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + canonical(metadata) +
                b"\r\n" + part + b"\r\nContent-Type: " + content_type + b"\r\n\r\n" +
                media + b"\r\n" + part + b"--\r\n")
        params["uploadType"] = "multipart"
        return self.request("/files", method="POST", params=params, body=body,
                            kind="multipart/related; boundary=" + boundary, upload=True)

    def move(self, fid, source, target):
        return self.request("/files/" + checked_id(fid), method="PATCH",
                            params={"fields": META, "addParents": checked_id(target),
                                    "removeParents": checked_id(source), "supportsAllDrives": "true"},
                            body=b"{}", kind="application/json")


def checked_id(value):
    require(isinstance(value, str) and re.fullmatch(r"[A-Za-z0-9_-]+", value),
            "Drive 配置或送审文件 ID 无效")
    return value


def check_access(author, publisher, settings):
    author_id, publisher_id = author.uid(), publisher.uid()
    require(author_id and publisher_id and author_id != publisher_id, "编制和发布须使用不同的 Drive 身份")
    names = ("inbox", "evidence", "staging", "formal")
    folders = {name: checked_id(settings["drive"][name]) for name in names}
    require(len(set(folders.values())) == len(names), "收稿、证据、暂存及正式目录须分别登记")
    for name, fid in folders.items():
        details = publisher.meta(fid)
        require(details["mimeType"] == FOLDER and not details.get("trashed") and
                (name == "inbox" or details["capabilities"].get("canAddChildren") is True),
                "发布身份无法使用已登记的私有目录")
        acl = publisher.private(fid)
        if name == "staging":
            require(not any(x.get("id") != publisher_id and not x.get("deleted") for x in acl),
                    "暂存目录必须在整批核验前仅对发布身份可见")
    cap = author.meta(folders["formal"])["capabilities"]
    require(all(cap.get(k) is False for k in ("canAddChildren", "canEdit", "canShare")),
            "普通编制账号仍可写入或改动正式目录，正式交付已阻止")
    require(author.meta(folders["evidence"])["mimeType"] == FOLDER, "编制账号无法取回私有 Word 回执")
    require(author.meta(folders["inbox"])["capabilities"].get("canAddChildren") is True,
            "编制账号无法走已登记的私有送审通道")
    return folders


def test_author_upload(author, folder):
    """部署验收时使用已授权的测试目录，确认真实编制身份上传被拒。"""
    try:
        made = author.create("construction-permission-test.txt", folder, b"permission test")
    except DriveError as exc:
        require(exc.status == 403 and set(exc.reasons) & {
            "insufficientFilePermissions", "forbidden"}, "须明确证实拒绝原因是目录写权限")
        return
    author.request("/files/" + checked_id(made["id"]), method="PATCH",
                   params={"supportsAllDrives": "true"}, body=b'{"trashed":true}', kind="application/json")
    raise GateError("编制账号实际上传成功，须先修正正式目录权限")


def receive(publisher, folders, file_id, expected, output):
    meta = publisher.meta(file_id)
    require(folders["inbox"] in meta.get("parents", ()) and not meta.get("trashed"),
            "批次输入不在已登记的私有收稿目录")
    publisher.private(file_id)
    data = publisher.download(file_id)
    require(sha256(data).hexdigest() == expected, "私有传输后的压缩包与送审摘要不一致")
    Path(output).write_bytes(data)


def check_bytes(drive, node, expected):
    actual = node.get("sha256Checksum") or sha256(drive.download(node["id"])).hexdigest()
    require(actual == expected, "Drive 中存放的文件与已审字节不一致")


def ensure_folder(drive, parent, name, key):
    entries = [x for x in drive.children(parent) if x.get("appProperties", {}).get("release_key") == key]
    require(len(entries) <= 1, "发布目录存在重复批次，须定位后重试")
    if entries:
        require(entries[0]["mimeType"] == FOLDER, "批次身份与已登记文件冲突")
        return entries[0]
    return drive.create(name, parent, props={"release_key": key})


def publish(drive, folders, release_key, batch_id, artifacts):
    """Copy verified bytes under restricted staging, then move the complete folder as a unit."""
    require(artifacts and len({name.casefold() for name in artifacts}) == len(artifacts), "正式稿缺项或文件名冲突")
    expected = {name: checksum for name, (_, checksum) in artifacts.items()}
    final = [n for n in drive.children(folders["formal"]) if
             n.get("appProperties", {}).get("release_key") == release_key]
    require(len(final) <= 1, "同批正式目录存在重复")
    folder = final[0] if final else ensure_folder(drive, folders["staging"], "施工组织设计-" + batch_id, release_key)
    existing_list = list(drive.children(folder["id"]))
    existing = {item["name"]: item for item in existing_list}
    require(len(existing) == len(existing_list) and set(existing) <= set(expected), "批次暂存文件清单异常")
    if final:
        require(set(existing) == set(expected), "已发布的批次存在文件缺失")
    for name, (path, checksum) in artifacts.items():
        data = Path(path).read_bytes()
        require(sha256(data).hexdigest() == checksum, "审核后文件已改变")
        entry = existing.get(name) or drive.create(name, folder["id"], data)
        check_bytes(drive, entry, checksum)
    # A retry after an interrupted response verifies the folder already in the formal parent.
    if not final:
        drive.move(folder["id"], folders["staging"], folders["formal"])
    folder = drive.meta(folder["id"])
    require(folders["formal"] in folder.get("parents", ()), "整批文件未进入正式目录")
    return folder.get("webViewLink") or ("https://drive.google.com/drive/folders/" + checked_id(folder["id"]))


def save_evidence(drive, folder, request_id, path):
    name = "回执-" + request_id + ".zip"
    exact = [x for x in drive.children(folder) if x["name"] == name]
    require(len(exact) <= 1, "私有回执目录存在重复记录")
    item = exact[0] if exact else drive.create(name, folder, Path(path).read_bytes())
    check_bytes(drive, item, file_sha(path))
    return item["id"]

