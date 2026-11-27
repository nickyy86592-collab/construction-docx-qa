"""将本地已解包两版注册到指定的私有生产仓库检出目录。"""
from __future__ import annotations

import argparse
from pathlib import Path
import re
import shutil
import sys

from release_task import GATEWAY_WORKFLOW, PRODUCTION_WORKFLOW, TRUST, dispatch_target, profile, read, require, write


def register(packages, checkout, settings_path, gateway_checkout=None):
    checkout = checkout.resolve()
    require((checkout / ".git").exists(), "须明确指定生产仓库的本地检出目录")
    settings = read(settings_path)
    require(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", settings["repository"]) and settings["branch"]
            and settings.get("public_key") and all(settings["drive"].get(x) for x in ("inbox", "evidence", "staging", "formal")),
            "生产信任配置缺仓库、分支、回执公钥或私有目录")
    dispatch_target(settings)
    if gateway_checkout is not None:
        gateway_checkout = gateway_checkout.resolve()
        require(settings.get("gateway") and (gateway_checkout / ".git").exists() and gateway_checkout != checkout,
                "须在配置填写公开入口并明确指定其独立本地检出目录")
    source = Path(__file__).resolve().parent
    target = checkout / "review-runtime"
    target.mkdir(exist_ok=True)
    (target / ".gitattributes").write_bytes(b"* -text\n")
    registry_path = target / TRUST
    existing = read(registry_path) if registry_path.is_file() else {}
    if existing:
        require(all(existing[key] == settings[key] for key in ("repository", "branch", "public_key", "drive")),
                "保留既有信任配置，变更生产仓、账号或签名须单独迁移")
    programs = dict(existing.get("profiles", {}))
    registered = []
    for pack in packages:
        manifest = profile(pack)
        version = manifest["skill_version"]
        require(re.fullmatch(r"r[0-9]+(?:\.[0-9]+)?", version), "技能版本标识无效")
        destination = target / "versions" / version
        require(not destination.exists() or (profile(destination) == manifest and
                programs.get(version) in (None, manifest["audit_sha256"])), "同名版本已有另一个审核实现")
        registered.append((pack, manifest, destination))
        programs[version] = manifest["audit_sha256"]
    for pack, manifest, destination in registered:
        for name in ["release-profile.json", *manifest["audit_files"]]:
            output = destination / name
            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pack / name, output)
        require(profile(destination) == manifest, "安装后的审核代码不完整")
    for name in ("release_task.py", "runner_release.py", "drive_release.py", "word_native.py", "requirements.txt"):
        shutil.copy2(source / name, target / name)
    settings["profiles"] = programs
    write(registry_path, settings)
    workflow = checkout / ".github" / "workflows" / PRODUCTION_WORKFLOW
    workflow.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source / PRODUCTION_WORKFLOW, workflow)
    if gateway_checkout is not None:
        workflow = gateway_checkout / ".github" / "workflows" / GATEWAY_WORKFLOW
        implementation = gateway_checkout / ".github" / "construction" / "forward_request.py"
        workflow.parent.mkdir(parents=True, exist_ok=True)
        implementation.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source / GATEWAY_WORKFLOW, workflow)
        shutil.copy2(source / "forward_request.py", implementation)
    return programs


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository-checkout", type=Path, required=True)
    parser.add_argument("--settings", type=Path, required=True)
    parser.add_argument("--gateway-checkout", type=Path)
    parser.add_argument("packages", nargs="+", type=Path)
    args = parser.parse_args()
    for version, digest in register(args.packages, args.repository_checkout, args.settings, args.gateway_checkout).items():
        print(version, digest)


if __name__ == "__main__":
    main()

