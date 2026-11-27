"""公开入口只把私有文件引用交给已登记的生产工作流。"""
from __future__ import annotations

import json
import os
import re
from urllib.request import HTTPRedirectHandler, Request, build_opener


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def forward(env, send):
    if not (env.get("CONSTRUCTION_REPOSITORY_PRIVATE") == "false"
            and env.get("GITHUB_EVENT_NAME") == "workflow_dispatch"
            and env.get("GITHUB_REF", "").startswith("refs/heads/")
            and env.get("GITHUB_REF") == env.get("CONSTRUCTION_DEFAULT_REF")):
        raise ValueError("GATEWAY=UNAPPROVED_ENTRY")
    repo = env.get("CONSTRUCTION_PRODUCTION_REPO", "")
    branch = env.get("CONSTRUCTION_PRODUCTION_BRANCH", "")
    token = env.get("CONSTRUCTION_FORWARD_TOKEN", "")
    file_id = env.get("CONSTRUCTION_REQUEST_FILE", "")
    checksum = env.get("CONSTRUCTION_REQUEST_HASH", "")
    if not (re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo)
            and repo.casefold() != env.get("GITHUB_REPOSITORY", "").casefold()
            and branch.strip() and token and re.fullmatch(r"[A-Za-z0-9_-]+", file_id)
            and re.fullmatch(r"[0-9a-f]{64}", checksum)):
        raise ValueError("GATEWAY=INVALID_CONFIGURATION_OR_REFERENCE")
    payload = {"ref": branch, "inputs": {"drive_file": file_id, "packet_sha256": checksum}}
    endpoint = f"https://api.github.com/repos/{repo}/actions/workflows/construction-versioned-release.yml/dispatches"
    request = Request(endpoint, data=json.dumps(payload).encode(), method="POST",
                      headers={"Authorization": "Bearer " + token, "Accept": "application/vnd.github+json",
                               "Content-Type": "application/json"})
    with send(request, timeout=30) as response:
        if response.status not in (200, 204):
            raise ValueError("GATEWAY=PRODUCTION_REJECTED")


if __name__ == "__main__":
    try:
        forward(os.environ, build_opener(NoRedirect()).open)
    except (OSError, ValueError):
        print("GATEWAY=FAILED")
        raise SystemExit(2)
    print("GATEWAY=ACCEPTED")
