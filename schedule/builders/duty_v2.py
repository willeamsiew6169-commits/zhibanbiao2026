"""V2 值班地点规划：纯计算，无数据库写入。已发布的 rows 仅作为不可变覆盖参考。"""
from schedule.builders.time_utils import parse_min, min_to_ampm

PLACES = ("观音堂", "活动中心")
BOUNDARIES = (600, 840)  # 绿/橙/黄班边界


def shift_of(t):
    return "绿" if t < 600 else "橙" if t < 840 else "黄"


def _start(p):
    return parse_min(p.get("start_time") or p.get("开始时间"))


def _end(p):
    return parse_min(p.get("end_time") or p.get("结束时间"))


def _id(p):
    return str(p.get("volunteer_id") or p.get("编号") or "").strip().upper()


def _history_score(p, place, history):
    counts = history.get(_id(p), {}).get("counts", {})
    return counts.get(place, 0)


def _options(person):
    s, e = _start(person), _end(person)
    if s is None or e is None or e <= s:
        return []
    fixed = "活动中心" if _id(person) == "STW-152" else None
    places = (fixed,) if fixed else PLACES
    opts = [[(s, e, place)] for place in places]
    # 非必要不拆；仅有实际换岗意义的切点。跨班界线在输出时单独切开。
    if not fixed and e - s >= 240:
        mid = (s + e) // 2
        cuts = set(BOUNDARIES)
        cuts.update(range(((s + 59) // 60) * 60, e, 60))
        cuts.add(mid)
        for cut in sorted(cuts):
            if cut - s < 60 or e - cut < 60:
                continue
            for place in PLACES:
                other = PLACES[1] if place == PLACES[0] else PLACES[0]
                opts.append([(s, cut, place), (cut, e, other)])
    return opts


def _score(assignments, fixed, history):
    all_segments = list(fixed)
    for _, segments in assignments:
        all_segments.extend(segments)
    edges = sorted({v for s,e,_ in all_segments for v in (s,e)})
    uncovered = 0
    imbalance = 0
    center_priority = 0
    # 10:00–16:00 同时有奇数位义工时，活动中心优先多一位。
    # 只对实际同时在岗的人数评分，不按整日报名人数。
    edges = sorted(set(edges) | {600, 960})
    for s,e in zip(edges, edges[1:]):
        if s >= e:
            continue
        counts = {place: 0 for place in PLACES}
        for a,b,place in all_segments:
            if a <= s and b >= e and place in counts:
                counts[place] += 1
        # 只评估有人报名的时间；无任何人报名时不影响地点决策
        if sum(counts.values()):
            uncovered += (e-s) * sum(n == 0 for n in counts.values())
            imbalance += (e-s) * abs(counts[PLACES[0]] - counts[PLACES[1]])
            if 600 <= s < 960:
                total = sum(counts.values())
                if total >= 3:
                    desired_center = (total + 1) // 2
                    center_priority += (e-s) * max(0, desired_center - counts["活动中心"])
    history_penalty = sum(sum((b-a)*_history_score(person,place,history) for a,b,place in segs)
                          for person,segs in assignments)
    changes = sum(len(segs)-1 for _,segs in assignments)
    midpoint_penalty = sum(abs(segs[0][1] - (_start(person)+_end(person))//2)
                           for person,segs in assignments if len(segs)>1)
    # 两位同一完整时段的义工，在不造成空岗时公平交换地点。
    rotation_penalty = 0
    if not fixed and len(assignments) == 2:
        (p1, seg1), (p2, seg2) = assignments
        if (_start(p1) == _start(p2) and _end(p1) == _end(p2)
                and _end(p1) - _start(p1) >= 240
                and _id(p1) != "STW-152" and _id(p2) != "STW-152"):
            rotation_penalty = sum(len(segs) == 1 for _, segs in assignments)
    # 覆盖优先；10–16点活动中心多一人优先于历史公平；其余保持原有次序。
    return (uncovered, center_priority, rotation_penalty, history_penalty, imbalance, changes, midpoint_penalty)


def _split_shift(segments):
    for a,b,place in segments:
        cuts = [a] + [x for x in BOUNDARIES if a < x < b] + [b]
        for s,e in zip(cuts,cuts[1:]):
            yield s,e,place


def plan(people, existing=(), history=None):
    """people: 本次允许安排的义工；existing: 已发布、不可修改的地点区间。"""
    history = history or {}
    fixed = []
    for row in existing:
        s,e = _start(row),_end(row)
        place = row.get("assigned_place") or row.get("place")
        if s is not None and e is not None and e > s and place in PLACES:
            fixed.append((s,e,place))
    people = [p for p in people if _options(p)]
    # 先安排时间覆盖较长的人，确保少人日的基本覆盖。
    people.sort(key=lambda p: (_start(p), -(_end(p)-_start(p)), _id(p), str(p.get("name") or "")))
    chosen = []
    for p in people:
        options = _options(p)
        best = min(options, key=lambda segs: _score(chosen+[(p,segs)], fixed, history))
        chosen.append((p,best))
    # 未发布完整排班可反复改善，补排时 existing 永远不变。
    for _ in range(5):
        improved = False
        for i,(p,current) in enumerate(chosen):
            options = _options(p)
            best = min(options, key=lambda segs: _score(chosen[:i]+[(p,segs)]+chosen[i+1:], fixed, history))
            if _score(chosen[:i]+[(p,best)]+chosen[i+1:], fixed, history) < _score(chosen,fixed,history):
                chosen[i] = (p,best)
                improved = True
        if not improved:
            break
    # 两位相同时段必须联合选择，单人逐个优化会卡在局部最优。
    if len(chosen) == 2 and not fixed:
        p1, _ = chosen[0]
        p2, _ = chosen[1]
        best_pair = min(((a,b) for a in _options(p1) for b in _options(p2)),
                        key=lambda pair: _score([(p1,pair[0]),(p2,pair[1])], fixed, history))
        chosen = [(p1,best_pair[0]), (p2,best_pair[1])]
    result = []
    for person,segments in chosen:
        for s,e,place in _split_shift(segments):
            item = dict(person)
            item.update(start_time=min_to_ampm(s),end_time=min_to_ampm(e),
                        开始时间=min_to_ampm(s),结束时间=min_to_ampm(e),
                        role="值班",岗位="值班")
            result.append((shift_of(s)+place,item))
    return result
