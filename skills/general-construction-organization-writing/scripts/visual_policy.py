#!/usr/bin/env python3
"""启动、静态审核和交付共用的项目图件决策校验；不代替技术/真实性判断。"""

BASE_FIGURES = ('施工功能分区示意图', '施工进度横道图', '施工进度网络图')
_UNSET = object()


def visual_policy_template():
    return {
        'version': 1,
        'named_exhibits': [],
        'real_photos': {'decision': '', 'reason': '', 'replacement_figures': []},
        'color_figures': {'mode': 'color'},
        'additional_figures': {'names': [], 'reason': ''},
    }


def base_names(policy):
    items = policy.get('named_exhibits', []) if isinstance(policy, dict) else []
    mapping = {item['replaces']: item['name'] for item in items if isinstance(item, dict)
               and item.get('replaces') in BASE_FIGURES and isinstance(item.get('name'), str)} if isinstance(items, list) else {}
    return [mapping.get(name, name) for name in BASE_FIGURES]


def _names(value):
    return (isinstance(value, list)
            and all(isinstance(x, str) and x.strip() for x in value)
            and len(value) == len(set(value)))


def _override_valid(value):
    return (isinstance(value, dict)
            and value.get('source_type') in ('prompt', 'agents', 'tender')
            and all(isinstance(value.get(k), str) and value[k].strip()
                    for k in ('source_ref', 'quote')))


def color_profile(policy):
    policy = policy if isinstance(policy, dict) else {}
    color = policy.get('color_figures')
    color = color if isinstance(color, dict) else {}
    mode = color.get('mode')
    if mode == 'monochrome':
        return {'color_required': False, 'min_color_ratio': 0,
                'min_foreground_color_ratio': 0, 'min_hue_groups': 0,
                'max_color_ratio': .001}
    if mode == 'custom':
        metrics = color.get('metrics')
        return {**(metrics if isinstance(metrics, dict) else {}), 'color_required': color.get('color_required')}
    return {'color_required': True, 'min_color_ratio': .01,
            'min_foreground_color_ratio': .04, 'min_hue_groups': 2,
            'max_color_ratio': 1}


def visual_policy_errors(policy, slots, figures=_UNSET):
    """省略figures为启动规划；传图件清单时同步验证交付规则。"""
    errors = []
    if not isinstance(policy, dict) or type(policy.get('version')) is not int or policy.get('version') != 1:
        return ['未记录有效visual_policy：须由AI按项目整体判断照片和功能图安排']
    exhibits = policy.get('named_exhibits', [])
    if not isinstance(exhibits, list) or any(not isinstance(x, dict) or x.get('kind') not in ('图','表')
            or not all(isinstance(x.get(k),str) and x[k].strip() for k in ('name','source_ref','quote'))
            or x['name'] not in x['quote'] for x in exhibits):
        return ['招标指定图表须登记原名、图或表、原文与位置']
    replaced = [x['replaces'] for x in exhibits if x.get('replaces')]
    if (not _names([x['name'] for x in exhibits]) or not _names(replaced)
            or any(x.get('replaces') and (x['kind'] != '图' or x['replaces'] not in BASE_FIGURES) for x in exhibits)):
        return ['指定图表重名或默认图替换关系无效']
    base = base_names(policy)
    if len(set(base)) != len(BASE_FIGURES):
        return ['不同必备图不能共用图名']
    if not isinstance(slots, list) or any(not isinstance(x, dict) for x in slots):
        errors.append('照片主题清单必须为对象数组；不需要照片时明确填[]')
        slots = []
    slot_names = [x.get('name') for x in slots]
    if not _names(slot_names):
        errors.append('照片主题须非空且不重复')
    photo = policy.get('real_photos')
    color = policy.get('color_figures')
    extra = policy.get('additional_figures')
    if not all(isinstance(x, dict) for x in (photo, color, extra)):
        return errors + ['visual_policy须明确real_photos、color_figures和additional_figures']
    for label, part in (('照片', photo), ('配色', color), ('增补图', extra)):
        if 'override' in part and not _override_valid(part['override']):
            errors.append(label+'例外须引用临时提示词、AGENTS.md或有效招标条款的原文及位置')
    decision = photo.get('decision')
    replacement = photo.get('replacement_figures')
    if decision not in ('required', 'schematic', 'omit') or not isinstance(photo.get('reason'), str) or not photo['reason'].strip():
        errors.append('须明确照片决策required/schematic/omit及项目整体价值依据')
    if not _names(replacement):
        errors.append('replacement_figures须为不重复的图名数组')
        replacement = []
    if decision == 'required' and not slots:
        errors.append('已判断照片有价值或明确要求照片：须至少登记一个照片/占位主题')
    if decision in ('schematic', 'omit') and slots:
        errors.append('不采用照片的决策与非空照片清单矛盾')
    if decision == 'schematic' and not replacement:
        errors.append('不采用实景图时须策划并实际生成有针对性的替代示意图')
    if decision != 'schematic' and replacement:
        errors.append('只有schematic决策才登记替代示意图')
    if decision == 'omit' and not _override_valid(photo.get('override')):
        errors.append('同时省略照片与替代示意图须有明确的特殊要求依据')
    if any(x in base for x in replacement):
        errors.append('替代实景的示意图须针对该主题，不能拿必备三图充数')
    mode = color.get('mode')
    if mode not in ('color', 'monochrome', 'custom'):
        errors.append('配色模式须为color/monochrome/custom')
    if mode in ('monochrome', 'custom') and not _override_valid(color.get('override')):
        errors.append('改变默认彩图要求须有明确的特殊要求依据')
    if mode == 'custom':
        profile = color_profile(policy)
        keys = ('min_color_ratio', 'min_foreground_color_ratio', 'max_color_ratio')
        if (type(profile.get('color_required')) is not bool
                or any(type(profile.get(k)) not in (int, float) or not 0 <= profile[k] <= 1 for k in keys)
                or type(profile.get('min_hue_groups')) is not int
                or not 0 <= profile.get('min_hue_groups', -1) <= 12
                or profile.get('min_color_ratio', 2) > profile.get('max_color_ratio', -1)):
            errors.append('custom配色须按依据明确color_required及有效metrics阈值')
    names = extra.get('names')
    if not _names(names):
        errors.append('增补功能图names须为不重复的图名数组')
        names = []
    if not isinstance(extra.get('reason'), str) or not extra['reason'].strip():
        errors.append('须说明增补功能图的项目价值及与正文的关系')
    if not 1 <= len(names) <= 3 and not _override_valid(extra.get('override')):
        errors.append('默认须在必备三图之外增加1至3张功能图；数量例外须有明确依据')
    if any(x in base for x in names):
        errors.append('必备三图不能计入增补1至3张功能图')
    if figures is not _UNSET:
        if not isinstance(figures, list) or any(not isinstance(x, dict) for x in figures):
            return errors + ['required_figures须为对象数组']
        all_names = [x.get('name') for x in figures]
        if not _names(all_names):
            errors.append('图件清单图名须非空且不重复')
        profile = color_profile(policy)
        for name in dict.fromkeys([*base, *names, *replacement]):
            item = next((x for x in figures if x.get('name') == name), None)
            if not item or item.get('hard') is not True or item.get('color_required') is not profile.get('color_required'):
                errors.append('图件缺失或hard/color_required与有效决策不一致：'+name)
            if name not in base and item and any(not item.get(k) for k in ('purpose', 'location', 'context_anchor')):
                errors.append('功能图须明确purpose、location及正文衔接原文context_anchor：'+name)
    return errors


def planned_figure_names(policy):
    policy = policy if isinstance(policy, dict) else {}
    extra = policy.get('additional_figures')
    photo = policy.get('real_photos')
    extra = extra if isinstance(extra, dict) else {}
    photo = photo if isinstance(photo, dict) else {}
    extras = extra.get('names') if isinstance(extra.get('names'), list) else []
    replacements = photo.get('replacement_figures') if isinstance(photo.get('replacement_figures'), list) else []
    return list(dict.fromkeys([*base_names(policy), *(x for x in extras + replacements if isinstance(x, str))]))
