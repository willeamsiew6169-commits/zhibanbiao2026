"""V2 值班地点规划：纯计算，无数据库写入。已发布的 rows 仅作为不可变覆盖参考。"""
from schedule.builders.time_utils import parse_min, min_to_ampm

PLACES = ("观音堂", "活动中心")
BOUNDARIES = (600, 840)  # 绿/橙/黄班边界

MEAL_CENTER_IDS = frozenset({"STW-158", "CHE-228", "CHE-116", "CHE-184", "CHE-50"})

def is_meal_center_day(target_date):
    """只在初一、十五、观音三大日及恩师涅槃日生效。"""
    if not target_date:
        return False
    from datetime import date
    from lunar_rules import get_special_day_info
    day = date.fromisoformat(str(target_date)[:10])
    info = get_special_day_info(day)
    return (info.get("lunar_day") in (1, 15)
            or (info.get("lunar_month"), info.get("lunar_day"))
            in {(2, 19), (6, 19), (9, 19), (10, 6)})

def is_fixed_center(person, meal_center_day=False):
    return (_id(person) == "STW-152"
            or (meal_center_day and _id(person) in MEAL_CENTER_IDS
                and _start(person) is not None and _start(person) < 780))


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


def _options(person, meal_center_day=False):
    s, e = _start(person), _end(person)
    if s is None or e is None or e <= s:
        return []
    fixed = "活动中心" if is_fixed_center(person, meal_center_day) else None
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


def _score(assignments, fixed, history, existing=(), cleaning=None, meal_center_day=False):
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
    # 方案 C：按班次、编号统计人数。跨地点者按该班服务时间比例分摊，
    # 避免为凑人数而额外拆班或把一人重复算作两人。
    from fractions import Fraction
    shift_people = {}
    def add_shift_person(identity, segments):
        for a, b, place in segments:
            for start, end, location in _split_shift([(a, b, place)]):
                minutes = shift_people.setdefault(shift_of(start), {}).setdefault(identity, dict.fromkeys(PLACES, 0))
                minutes[location] += end - start
    for i, row in enumerate(existing):
        a, b = _start(row), _end(row)
        place = row.get("assigned_place") or row.get("place")
        if a is not None and b is not None and b > a and place in PLACES:
            add_shift_person(_id(row) or ("existing", i), [(a, b, place)])
    for i, (person, segments) in enumerate(assignments):
        add_shift_person(_id(person) or ("new", i), segments)
    shift_imbalance = shift_center_penalty = 0
    for shift, persons in shift_people.items():
        center = sum(Fraction(v[PLACES[1]], sum(v.values())) for v in persons.values())
        total = len(persons)
        # 奇数允许相差一人；偶数力求相等。
        shift_imbalance += max(0, abs(total - 2 * center) - total % 2)
        if shift == "橙" and total >= 2:
            shift_center_penalty += max(0, (total + 1) // 2 - center) if total % 2 else 0
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
                and not is_fixed_center(p1, meal_center_day) and not is_fixed_center(p2, meal_center_day)):
            rotation_penalty = sum(len(segs) == 1 for _, segs in assignments)
    # 同日轮替按编号汇总，已发布记录仅作为参考。
    by_id = {}
    active = {_id(person) for person, _ in assignments if not is_fixed_center(person, meal_center_day)} - {""}
    for row in existing:
        identity = _id(row)
        a,b = _start(row),_end(row)
        place = row.get("assigned_place") or row.get("place")
        if identity in active and a is not None and b is not None and b > a and place in PLACES:
            by_id.setdefault(identity, []).append((a,b,place))
    for person,segs in assignments:
        if _id(person) in active:
            by_id.setdefault(_id(person), []).extend(segs)
    clean_penalty = day_penalty = 0
    for identity,segs in by_id.items():
        before = {place: sum(max(0,min(b,840)-max(a,600)) for a,b,p in segs if p == place) for place in PLACES}
        after = {place: sum(max(0,b-max(a,840)) for a,b,p in segs if p == place) for place in PLACES}
        if sum(before.values()) and sum(after.values()):
            day_penalty += sum(before[p]*after[p] for p in PLACES) / (sum(before.values())*sum(after.values()))
        clean_place = (cleaning or {}).get(identity)
        if clean_place in PLACES:
            clean_penalty += int(min(segs)[2] == clean_place)
    return (uncovered, imbalance, shift_imbalance, shift_center_penalty, center_priority, clean_penalty, day_penalty,
            rotation_penalty, history_penalty, changes, midpoint_penalty)


def _split_shift(segments):
    for a,b,place in segments:
        cuts = [a] + [x for x in BOUNDARIES if a < x < b] + [b]
        for s,e in zip(cuts,cuts[1:]):
            yield s,e,place


def plan(people, existing=(), history=None, cleaning=None, target_date=None):
    """people: 本次允许安排的义工；existing: 已发布、不可修改的地点区间。"""
    history = history or {}
    meal_center_day = is_meal_center_day(target_date)
    existing = list(existing)
    fixed = []
    for row in existing:
        s,e = _start(row),_end(row)
        place = row.get("assigned_place") or row.get("place")
        if s is not None and e is not None and e > s and place in PLACES:
            fixed.append((s,e,place))
    people = [p for p in people if _options(p, meal_center_day)]
    # 先安排时间覆盖较长的人，确保少人日的基本覆盖。
    people.sort(key=lambda p: (_start(p), -(_end(p)-_start(p)), _id(p), str(p.get("name") or "")))
    chosen = []
    for p in people:
        options = _options(p, meal_center_day)
        best = min(options, key=lambda segs: _score(chosen+[(p,segs)], fixed, history, existing, cleaning, meal_center_day))
        chosen.append((p,best))
    # 未发布完整排班可反复改善，补排时 existing 永远不变。
    for _ in range(5):
        improved = False
        for i,(p,current) in enumerate(chosen):
            options = _options(p, meal_center_day)
            best = min(options, key=lambda segs: _score(chosen[:i]+[(p,segs)]+chosen[i+1:], fixed, history, existing, cleaning, meal_center_day))
            if _score(chosen[:i]+[(p,best)]+chosen[i+1:], fixed, history, existing, cleaning, meal_center_day) < _score(chosen,fixed,history,existing,cleaning,meal_center_day):
                chosen[i] = (p,best)
                improved = True
        if not improved:
            break
    # 联合改善：单人调整可能暂时制造空岗，无法完成公平交换。
    # 限制轮数，避免大日子无限搜索。
    for _ in range(3):
        improved = False
        for i in range(len(chosen)):
            for j in range(i+1,len(chosen)):
                p1,s1 = chosen[i]
                p2,s2 = chosen[j]
                current_score = _score(chosen, fixed, history, existing, cleaning, meal_center_day)
                best = None
                for a in _options(p1, meal_center_day):
                    for b in _options(p2, meal_center_day):
                        candidate = list(chosen)
                        candidate[i],candidate[j] = (p1,a),(p2,b)
                        score = _score(candidate, fixed, history, existing, cleaning, meal_center_day)
                        if score < current_score:
                            current_score,best = score,candidate
                if best is not None:
                    chosen = best
                    improved = True
        if not improved:
            break
    # 两位相同时段必须联合选择，单人逐个优化会卡在局部最优。
    if len(chosen) == 2 and not fixed:
        p1, _ = chosen[0]
        p2, _ = chosen[1]
        best_pair = min(((a,b) for a in _options(p1, meal_center_day) for b in _options(p2, meal_center_day)),
                        key=lambda pair: _score([(p1,pair[0]),(p2,pair[1])], fixed, history, existing, cleaning, meal_center_day))
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
