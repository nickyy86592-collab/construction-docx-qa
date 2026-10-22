#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用、轻量的施工组织设计评分清单校验器。

不读取/解释招标文件，不调用模型，不运行 Word，也不产生主观分数。
AI 在编制前从招标评分原文提取项目清单，按需在 40%、75%、final 节点调用此脚本。
检查来源、评分项完整性、模拟分数范围、响应位置、得分解释及计算一致性。

命令：python 评分响应核验.py 项目评分清单.json --phase 40 --report score-40.json
      python 评分响应核验.py 项目评分清单.json --phase final --report score-final.json

输入格式详见 references/03_优秀区间评分与实景图.md。
返回码：0=PASS，2=FAIL，3=ERROR。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from decimal import Decimal, InvalidOperation
from pathlib import Path


def valid_num(x):
    return isinstance(x, (float, int)) and not isinstance(x, bool) and math.isfinite(x)


def dec(x):
    if not valid_num(x):
        raise InvalidOperation('不是有效数值')
    return Decimal(str(x))


def nonempty(x):
    return isinstance(x, str) and bool(x.strip())


def check(data, phase):
    issues, warnings, summaries = [], [], []
    if not isinstance(data, dict):
        return {"phase": phase, "status": "FAIL", "failures": ["评分清单顶层必须是JSON对象"], "warnings": []}
    if not nonempty(data.get('tender_source')):
        issues.append('缺少评分标准原文来源 tender_source（招标文件版本/页码或条款）')
    items = data.get('items')
    if not isinstance(items, list) or not items:
        issues.append('缺少有效的 items 评分项目列表')
        items = []
    seen = set()
    total_max = Decimal('0')
    completed_score = Decimal('0')
    completed_count = 0
    for n, item in enumerate(items, 1):
        if not isinstance(item, dict):
            issues.append(f'第{n}项不是对象')
            continue
        key = str(item.get('id', '')).strip()
        name = str(item.get('name', '')).strip()
        label = f'{key or n}({name or "未命名"})'
        if not key or key in seen:
            issues.append(f'{label} 评分项ID缺失或重复')
        seen.add(key)
        if not name or not nonempty(item.get('source_ref')):
            issues.append(f'{label} 缺评分名称或招标原文位置 source_ref')
        criteria = item.get('criteria')
        if not isinstance(criteria, list) or not criteria or any(not nonempty(x) for x in criteria):
            issues.append(f'{label} 缺招标原文的优秀/评判条件 criteria')
        max_score = item.get('max_score')
        if not valid_num(max_score) or max_score <= 0:
            issues.append(f'{label} max_score 缺失或无效')
            continue
        max_d = dec(max_score)
        total_max += max_d
        excellent = item.get('excellent_range')
        if excellent is not None:
            if not (isinstance(excellent, list) and len(excellent) == 2 and all(valid_num(x) for x in excellent)
                    and 0 <= dec(excellent[0]) <= dec(excellent[1]) <= max_d):
                issues.append(f'{label} excellent_range 越界或格式错误')
        elif not nonempty(item.get('excellent_note')):
            issues.append(f'{label} 未提供优秀区间，须说明原文未明确或无法确认的原因 excellent_note')

        status = item.get('completion')
        if status not in ('complete', 'partial', 'pending'):
            issues.append(f'{label} completion 只能是 complete/partial/pending')
            continue
        if phase == 'final' and status != 'complete':
            issues.append(f'{label} 最终阶段仍为 {status}，不得形成完整模拟自评')
        if status == 'complete':
            evidence = item.get('evidence')
            if not isinstance(evidence, list) or not evidence:
                issues.append(f'{label} 已完成项缺少正文定位与措施证据 evidence')
            else:
                for i, e in enumerate(evidence, 1):
                    if not isinstance(e, dict) or not nonempty(e.get('location')) or not nonempty(e.get('response')):
                        issues.append(f'{label} evidence第{i}项必须有具体章节/位置location及实际响应response')
            score = item.get('self_score')
            if not valid_num(score) or not (Decimal('0') <= dec(score) <= max_d):
                issues.append(f'{label} self_score 未填写或不在0到满分之间')
            else:
                completed_score += dec(score)
                completed_count += 1
            if not nonempty(item.get('reason')):
                issues.append(f'{label} 完成项缺少模拟得分理由 reason')
            if not nonempty(item.get('deduction_risk')):
                issues.append(f'{label} 完成项缺少扣分风险/保留意见 deduction_risk；无明显风险也须说明判断依据')
            if (isinstance(excellent, list) and len(excellent) == 2
                    and all(valid_num(x) for x in excellent) and valid_num(score)
                    and dec(score) >= dec(excellent[0])):
                if not nonempty(item.get('excellent_justification')):
                    issues.append(f'{label} 自评进入优秀区间但缺对应优秀条件的证据解释 excellent_justification')
        else:
            if phase != 'final' and item.get('self_score') is not None:
                warnings.append(f'{label} 尚未完成，self_score不纳入阶段已完成项得分；不得据此推算全标得分')
        summaries.append({
            'id': key, 'name': name, 'max_score': str(max_d), 'completion': status,
            'self_score': item.get('self_score') if status == 'complete' else None,
        })
    if phase == 'final':
        declared = data.get('declared_self_total')
        if declared is not None:
            if not valid_num(declared) or dec(declared) != completed_score:
                issues.append(f'报告声明总分 {declared} 与逐项模拟得分合计 {completed_score} 不一致')
    else:
        warnings.append('阶段检查只核验已完成评分项；未完成项不计分，不得把阶段得分当成整体预测得分')
    if not items:
        warnings.append('未产生有效评分项目清单')
    return {
        'phase': phase, 'status': 'FAIL' if issues else 'PASS',
        'page_source': None,  # 不参与/替代 Word 页数校核
        'item_count': len(items), 'completed_count': completed_count,
        'maximum_total': str(total_max), 'completed_items_score': str(completed_score),
        'score_nature': 'AI模拟自评；不是评标专家得分保证',
        'items': summaries, 'failures': issues, 'warnings': warnings,
    }


def main():
    p = argparse.ArgumentParser(description='通用评分响应完整性及分数算术核验（不作主观评价）')
    p.add_argument('input', type=Path)
    p.add_argument('--phase', choices=['40', '75', 'final'], required=True)
    p.add_argument('--document',type=Path,required=True,help='本阶段实际DOCX，报告绑定其SHA256')
    p.add_argument('--report', type=Path, default=None)
    args = p.parse_args()
    try:
        data = json.loads(args.input.read_text(encoding='utf-8'))
        report = check(data, args.phase)
        report["document_sha256"]=hashlib.sha256(args.document.read_bytes()).hexdigest()
        report["checks"]=[{"name":w,"status":"WARN","message":w} for w in report.get("warnings",[])]
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({k: v for k, v in report.items() if k != 'items'}, ensure_ascii=False, indent=2))
        return 0 if report['status'] == 'PASS' else 2
    except (OSError, ValueError, TypeError, InvalidOperation) as exc:
        print(json.dumps({'status': 'ERROR', 'reason': str(exc)}, ensure_ascii=False))
        return 3


if __name__ == '__main__':
    raise SystemExit(main())
