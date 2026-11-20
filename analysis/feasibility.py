"""재건축 사업성분석 자동화.

입력 TOML(단지별) → 블로그 원고 초안(report.md) + 엑셀 표(tables.xlsx) + 민감도 차트(PNG)

    python analysis/feasibility.py analysis/complexes/mokdong13.toml

금액 단위: 천원 / 면적: ㎡ / 단가: 천원/평
"""
from __future__ import annotations

import argparse
import copy
import sys
import tomllib
from dataclasses import dataclass
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm  # noqa: E402

PYEONG = 1 / 0.3025  # ㎡ → 평 (㎡ × 0.3025)
ROOT = Path(__file__).resolve().parent.parent

# ── 표기 도우미 ──────────────────────────────────────────────


def py(m2: float) -> float:
    return m2 / PYEONG


def won(v: float) -> str:
    """천원 금액을 블로그 표기로: 1,150,000"""
    return f"{v:,.0f}"


def eok(v: float, digits: int = 0) -> str:
    """천원 → '3조 9,628억원' / '831억원' / '11.2억원'"""
    e = v / 100_000
    if abs(e) >= 10_000:
        jo, rest = divmod(round(abs(e)), 10_000)
        sign = "-" if e < 0 else ""
        return f"{sign}{jo}조 {rest:,}억원" if rest else f"{sign}{jo}조원"
    return f"{e:,.{digits}f}억원"


def man(v: float) -> str:
    """천원/평 → '5,500만원'"""
    return f"{v / 10:,.0f}만원"


def pct(v: float) -> str:
    return f"{v * 100:.2f}%"


def md_table(df: pd.DataFrame) -> str:
    cols = [str(c) for c in df.columns]
    lines = ["| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)]
    for row in df.itertuples(index=False):
        lines.append("| " + " | ".join("" if pd.isna(x) else str(x) for x in row) + " |")
    return "\n".join(lines)


# ── 계산 ─────────────────────────────────────────────────────


@dataclass
class Result:
    params: dict
    construction: float
    costs: list[tuple[str, float]]
    revenues: list[tuple[str, float]]
    prior_total: float
    member_price: float

    @property
    def cost_total(self) -> float:
        return sum(v for _, v in self.costs)

    @property
    def revenue_total(self) -> float:
        return sum(v for _, v in self.revenues)

    @property
    def ratio(self) -> float:
        """비례율 = (총수입 - 총사업비) / 종전자산평가액"""
        return (self.revenue_total - self.cost_total) / self.prior_total


def prior_value(e: dict, appraisal: float) -> float:
    """종전자산 1건(또는 lump 합계)의 평가액"""
    return e["price"] * appraisal


def prior_total(cfg: dict, p: dict) -> float:
    total = 0.0
    for e in cfg["existing"]:
        n = 1 if e.get("lump") else e.get("units", 0)
        total += prior_value(e, p["appraisal_ratio"]) * n
    return total


def unit_price(n: dict, per_pyeong: float) -> float:
    """신축 1세대 분양가 (천원)"""
    return py(n["supply_m2"]) * per_pyeong * n.get("coef", 1.0)


def compute(cfg: dict, overrides: dict | None = None) -> Result:
    p = {**cfg["base"], **(overrides or {})}
    base_c = cfg["base"]["construction_per_pyeong"]
    construction = p["gfa_pyeong"] * p["construction_per_pyeong"]
    growth = p["construction_per_pyeong"] / base_c - 1

    costs = [("공사비(철거 포함)", construction)]
    for c in cfg["cost"]:
        costs.append((c["name"], c["amount"] * (1 + c.get("elasticity", 0) * growth)))

    gp = p["general_price"]
    mp = gp * p["member_price_ratio"]
    member = sum(unit_price(n, mp) * n.get("member", 0) for n in cfg["new"])
    general = sum(unit_price(n, gp) * n.get("general", 0) for n in cfg["new"])
    if "public_amount" in p and "public_basic_cost" not in (overrides or {}):
        public = p["public_amount"]
    else:
        area = sum(py(n["supply_m2"]) * n.get("public", 0) for n in cfg["new"])
        public = area * p["public_basic_cost"] * p.get("public_ratio", 1.0)
    rv = cfg.get("revenue", {})
    revenues = [
        ("조합원분양", member),
        ("일반분양", general),
        ("상가분양", rv.get("commercial", 0)),
        ("공공주택", public),
        ("매출부가세", rv.get("vat", 0)),
    ]
    return Result(p, construction, costs, revenues, prior_total(cfg, p), mp)


def contribution(cfg: dict, r: Result, ex: dict, nw: dict) -> float:
    """분담금(천원) = 신축 조합원분양가 - 권리가액(종전평가액 × 비례율). 음수 = 환급"""
    return unit_price(nw, r.member_price) - prior_value(ex, r.params["appraisal_ratio"]) * r.ratio


def residential(cfg: dict) -> list[dict]:
    return [e for e in cfg["existing"] if not e.get("lump")]


# ── 표 만들기 ────────────────────────────────────────────────


def existing_table(cfg: dict) -> pd.DataFrame:
    rows = []
    for e in cfg["existing"]:
        lump = e.get("lump")
        rows.append({
            "타 입": e["type"],
            "전용면적": "-" if lump else f"{e['exclusive_m2']:.2f} ㎡",
            "공급면적": "-" if lump else f"{e['supply_m2']:.2f} ㎡",
            "세대수": "-" if lump else f"{e['units']:,}",
            "대지지분(평)": f"{e['land_pyeong']:,.1f} 평",
            "추정시세(천원)": won(e["price"]),
        })
    res = residential(cfg)
    rows.append({
        "타 입": "합 계",
        "전용면적": f"{sum(e['exclusive_m2'] * e['units'] for e in res):,.2f} ㎡",
        "공급면적": f"{sum(e['supply_m2'] * e['units'] for e in res):,.2f} ㎡",
        "세대수": f"{sum(e['units'] for e in res):,}",
        "대지지분(평)": f"{sum(e['land_pyeong'] * (1 if e.get('lump') else e['units']) for e in cfg['existing']):,.1f} 평",
        "추정시세(천원)": won(prior_total(cfg, {"appraisal_ratio": 1.0})),
    })
    return pd.DataFrame(rows)


def new_table(cfg: dict) -> pd.DataFrame:
    rows = []
    for n in cfg["new"]:
        total = sum(n.get(k, 0) for k in ("member", "general", "public", "donation"))
        parts = [f"{label} {n[k]:,}" for k, label in
                 (("member", "조합원"), ("general", "일반"), ("public", "공공"), ("donation", "기부채납"))
                 if n.get(k)]
        rows.append({
            "타 입": n["type"],
            "전용면적": f"{n['exclusive_m2']:.2f} ㎡",
            "공급면적": f"{n['supply_m2']:.2f} ㎡",
            "세대수": f"{total:,}",
            "대지지분(평)": f"{n.get('land_pyeong', 0):.1f} 평",
            "구성": " / ".join(parts),
        })
    tot = lambda k: sum(n.get(k, 0) for n in cfg["new"])  # noqa: E731
    rows.append({
        "타 입": "합 계",
        "전용면적": f"{sum(n['exclusive_m2'] * sum(n.get(k, 0) for k in ('member', 'general', 'public', 'donation')) for n in cfg['new']):,.2f} ㎡",
        "공급면적": f"{sum(n['supply_m2'] * sum(n.get(k, 0) for k in ('member', 'general', 'public', 'donation')) for n in cfg['new']):,.2f} ㎡",
        "세대수": f"{tot('member') + tot('general') + tot('public') + tot('donation'):,}",
        "대지지분(평)": "-",
        "구성": f"조합원 {tot('member'):,} / 일반 {tot('general'):,} / 공공 {tot('public'):,} / 기부채납 {tot('donation'):,}",
    })
    return pd.DataFrame(rows)


def budget_table(r: Result) -> pd.DataFrame:
    n = max(len(r.costs), len(r.revenues))
    c = r.costs + [("", None)] * (n - len(r.costs))
    v = r.revenues + [("", None)] * (n - len(r.revenues))
    rows = [{"사업비 구분": a, "사업비(천원)": won(b) if b is not None else "",
             "수입 구분": x, "수입(천원)": won(y) if y is not None else ""} for (a, b), (x, y) in zip(c, v)]
    rows.append({"사업비 구분": "사업비 합계", "사업비(천원)": won(r.cost_total),
                 "수입 구분": "총수입 합계", "수입(천원)": won(r.revenue_total)})
    return pd.DataFrame(rows)


def ratio_table(r: Result) -> pd.DataFrame:
    return pd.DataFrame([
        {"구 분": "총 수입합계", "금 액(천원)": won(r.revenue_total)},
        {"구 분": "종전감정평가액", "금 액(천원)": won(r.prior_total)},
        {"구 분": "총사업비", "금 액(천원)": won(r.cost_total)},
        {"구 분": "비례율", "금 액(천원)": pct(r.ratio)},
    ])


def compare_table(before: list, after: list, label: str) -> pd.DataFrame:
    rows = []
    for (name, b), (_, a) in zip(before, after):
        diff = "변경없음" if abs(a - b) < 1 else (f"{(a / b - 1) * 100:+.1f}%" if b else "")
        rows.append({"구 분": name, "변경 전(천원)": won(b), "변경 후(천원)": won(a), "비 고": diff})
    tb, ta = sum(v for _, v in before), sum(v for _, v in after)
    rows.append({"구 분": f"{label} 합계", "변경 전(천원)": won(tb), "변경 후(천원)": won(ta),
                 "비 고": f"{(ta / tb - 1) * 100:+.1f}%"})
    return pd.DataFrame(rows)


def contribution_table(cfg: dict, r: Result) -> pd.DataFrame:
    """종전 타입(행) × 신축 타입(열) 분담금, 억원. 음수 = 환급"""
    rows = []
    for e in residential(cfg):
        row = {"종전 \\ 신축": e["type"], "권리가액": f"{e['price'] * r.params['appraisal_ratio'] * r.ratio / 100_000:.2f}"}
        for n in cfg["new"]:
            row[n["type"]] = f"{contribution(cfg, r, e, n) / 100_000:.2f}"
        rows.append(row)
    row = {"종전 \\ 신축": "조합원분양가", "권리가액": ""}
    for n in cfg["new"]:
        row[n["type"]] = f"{unit_price(n, r.member_price) / 100_000:.2f}"
    rows.append(row)
    return pd.DataFrame(rows)


# ── 민감도 차트 ──────────────────────────────────────────────
# 발산형 팔레트: 빨강(조합원 불리) ↔ 회색(기준) ↔ 파랑(조합원 유리)
INK, INK2, SURFACE = "#0b0b0b", "#52514e", "#fcfcfb"
DIVERGING = LinearSegmentedColormap.from_list("div", ["#e34948", "#f0efec", "#2a78d6"])


def set_font() -> None:
    for name in ("NanumGothic", "NanumBarunGothic", "Malgun Gothic", "AppleGothic"):
        if any(f.name == name for f in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = name
            break
    plt.rcParams["axes.unicode_minus"] = False


def heatmap(grid: pd.DataFrame, title: str, subtitle: str, fmt, center: float,
            good_high: bool, base: tuple, path: Path) -> None:
    set_font()
    fig, ax = plt.subplots(figsize=(8, 5.6), dpi=200, facecolor=SURFACE)
    ax.set_facecolor(SURFACE)
    vals = grid.values.astype(float)
    span = max(abs(vals.max() - center), abs(vals.min() - center)) or 1
    cmap = DIVERGING if good_high else DIVERGING.reversed()
    norm = TwoSlopeNorm(vcenter=center, vmin=center - span, vmax=center + span)
    ax.imshow(vals, cmap=cmap, norm=norm, aspect="auto")
    for i in range(vals.shape[0]):
        for j in range(vals.shape[1]):
            rgb = cmap(norm(vals[i, j]))[:3]
            lum = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
            ax.text(j, i, fmt(vals[i, j]), ha="center", va="center", fontsize=11,
                    color="white" if lum < 0.45 else INK,
                    fontweight="semibold" if (i, j) == base else "normal")
    # 셀 사이 2px 간격
    ax.set_xticks([x - 0.5 for x in range(1, vals.shape[1])], minor=True)
    ax.set_yticks([y - 0.5 for y in range(1, vals.shape[0])], minor=True)
    ax.grid(which="minor", color=SURFACE, linewidth=2)
    ax.tick_params(which="minor", length=0)
    if base:
        ax.add_patch(plt.Rectangle((base[1] - 0.46, base[0] - 0.46), 0.92, 0.92, fill=False, ec=INK, lw=1.6, zorder=5))
    ax.set_xticks(range(len(grid.columns)), [f"{c / 10:,.0f}" for c in grid.columns], color=INK2)
    ax.set_yticks(range(len(grid.index)), [f"{i / 10:,.0f}" for i in grid.index], color=INK2)
    ax.set_xlabel("일반분양가 (만원/평)", color=INK2)
    ax.set_ylabel("평당공사비 (만원/평)", color=INK2)
    ax.tick_params(length=0)
    for s in ax.spines.values():
        s.set_visible(False)
    fig.suptitle(title, x=0.06, ha="left", fontsize=14, fontweight="semibold", color=INK)
    ax.set_title(subtitle, loc="left", fontsize=9.5, color=INK2, pad=10)
    fig.tight_layout()
    fig.savefig(path, facecolor=SURFACE)
    plt.close(fig)


def sensitivity(cfg: dict, base_r: Result, out: Path) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    s = cfg["sensitivity"]
    cs, gs = s["construction_per_pyeong"], s["general_price"]
    ex = next(e for e in cfg["existing"] if e["type"] == s["reference_existing"])
    nw = next(n for n in cfg["new"] if n["type"] == s["reference_new"])
    ratio = pd.DataFrame(index=cs, columns=gs, dtype=float)
    contrib = pd.DataFrame(index=cs, columns=gs, dtype=float)
    for c in cs:
        for g in gs:
            r = compute(cfg, {"construction_per_pyeong": c, "general_price": g})
            ratio.loc[c, g] = r.ratio * 100
            contrib.loc[c, g] = contribution(cfg, r, ex, nw) / 100_000
    bp = base_r.params
    base = ((cs.index(bp["construction_per_pyeong"]), gs.index(bp["general_price"]))
            if bp["construction_per_pyeong"] in cs and bp["general_price"] in gs else None)
    name = cfg["info"]["name"]
    note = (f"조합원분양가 = 일반분양가 × {bp['member_price_ratio'] * 100:.1f}% · 굵은 테두리 = 정비계획 기준값 · "
            "파랑 = 조합원 유리, 빨강 = 불리")
    heatmap(ratio, f"{name} 비례율 민감도 (%)", note, lambda v: f"{v:.1f}", 100, True, base,
            out / "sens_ratio.png")
    heatmap(contrib, f"{name} 분담금 민감도: {ex['type']} → {nw['type']} (억원)",
            note + " · 음수 = 환급", lambda v: f"{v:+.2f}", 0, False, base, out / "sens_contrib.png")
    return ratio, contrib, ["sens_ratio.png", "sens_contrib.png"]


# ── 원고 초안 ────────────────────────────────────────────────

TODO = "> ✍️ [작성]"


def report(cfg: dict, base: Result, scen: list[tuple[dict, Result]], sens: tuple, out: Path) -> str:
    i, b = cfg["info"], cfg["base"]
    res = residential(cfg)
    house = sum(e["price"] * e["units"] for e in res) * b["appraisal_ratio"]
    lump = sum(e["price"] for e in cfg["existing"] if e.get("lump")) * b["appraisal_ratio"]
    rv = cfg.get("revenue", {})
    contribs = [contribution(cfg, base, e, n) for e in res for n in cfg["new"]]
    s = cfg["sensitivity"]

    L = [f"# [사업성분석#{i['series_no']}] {i['name']} 재건축 분담금, 분양가, 비례율, 공사비&분양가 민감도", ""]
    L += ["## 1) " + f"{i['name']} 아파트 기본정보", "", f"**{i['name']}아파트** · {i['address']}", "",
          f"- 준공일 : {i['completion']} ({i['age']}년차)", f"- 세대수 : {i['households']:,}세대",
          f"- 용적률 : {i['far']} %", f"- 평형대 : {i['size_range']}", "",
          TODO + " 이 단지를 고른 이유, 최근 이슈(시공사 선정·심의 등)", ""]
    L += ["## 2) 대지지분, 시세", "",
          f"- 대지면적 : {i['site_area_m2']:,.1f}㎡ ({py(i['site_area_m2']):,.1f}평)", "",
          "타입별 전용면적, 공급면적, 세대수, 대지지분, 현재시세 표", "", md_table(existing_table(cfg)), "",
          f"토지등소유자 수 : {i['owners']:,}명 ({i['owners_note']})", "",
          f"총 종전자산가액은 주택 {eok(house)}" + (f", 상가 {eok(lump)}" if lump else "") + " 수준으로 나온다.", ""]
    L += ["## 3) 재건축 예상 계획", "", f"{i['notice']}에 따라 {i['name']}의 정비구역 지정이 고시되었다.", "",
          TODO + " 고시문 PDF 첨부, 조감도·결정도 이미지 삽입", "",
          f"- 계획 용적률 : {i['plan_far']}", "", "[공동주택 계획]", "", md_table(new_table(cfg)), ""]
    L += ["## 4) 사업성 분석 (기존 정비구역 결정고시 추정치)", "", "[사업비, 비례율 분석]", "",
          f"- 일반분양가 : {man(b['general_price'])}/평",
          f"- 조합원분양가 : {man(base.member_price)}/평 (일반분양가 {b['member_price_ratio'] * 100:.0f}%)",
          f"- 상가매각가 : {eok(rv.get('commercial', 0), 0)}" + (f" ({rv['commercial_note']})" if rv.get("commercial_note") else ""),
          f"- 평당공사비 : {man(b['construction_per_pyeong'])}",
          "- 공공주택은 기본형 건축비로 책정하여 매각",
          f"- 공사기간 : {i['construction_months']}개월",
          f"- 총 사업비 : {eok(base.cost_total)}", "",
          md_table(budget_table(base)), "", "비례율 계산", "", md_table(ratio_table(base)), "",
          "[평형별 추정 분담금표] (단위: 억원, 음수 = 환급)", "", md_table(contribution_table(cfg, base)), "",
          f"분담금 범위는 최대 {eok(-min(contribs), 1)} 환급 ~ 최대 {eok(max(contribs), 1)} 부담 구간이다.", "",
          TODO + " 분담금표 해석 (환급 구간, 1+1 가능성 등)", ""]
    for n, (sc, r) in enumerate(scen, start=5):
        L += [f"## {n}) {sc['title']}", "", "[총 사업비 변동사항]", "",
              md_table(compare_table(base.costs, r.costs, "사업비")), "",
              f"- 공사비 : 평당 {man(b['construction_per_pyeong'])} → 평당 {man(r.params['construction_per_pyeong'])}", "",
              "[분양가 변동사항]", "", TODO + " 주변 신축 실거래가 근거 (단지명, 실거래가, 평단가)", "",
              md_table(compare_table(base.revenues, r.revenues, "총수입")), "",
              "[변경 비례율 및 추정 분담금표]", "", md_table(ratio_table(r)), "",
              f"비례율은 {pct(base.ratio)} → {pct(r.ratio)}로 {(r.ratio - base.ratio) * 100:+.2f}%p 변한다.", "",
              md_table(contribution_table(cfg, r)), ""]
    n = 5 + len(scen)
    ratio, contrib, imgs = sens
    L += [f"## {n}) 공사비 & 분양가 민감도", "",
          f"평당공사비({', '.join(man(c) for c in s['construction_per_pyeong'])})와 "
          f"일반분양가({', '.join(man(g) for g in s['general_price'])})를 바꿔가며 비례율과 분담금을 계산했다.", "",
          f"![비례율 민감도]({imgs[0]})", "", f"![분담금 민감도]({imgs[1]})", "",
          f"- 비례율 범위 : {ratio.values.min():.1f}% ~ {ratio.values.max():.1f}%",
          f"- {s['reference_existing']} → {s['reference_new']} 분담금 범위 : "
          f"{contrib.values.min():+.2f}억원 ~ {contrib.values.max():+.2f}억원", "",
          TODO + " 민감도 해석 (분양가 100만원 상승 vs 공사비 100만원 상승 효과 비교 등)", ""]
    L += [f"## {n + 1}) 마무리", "", TODO + " 향후 일정, 모니터링 포인트, 리스크 요인", "",
          "*개인적으로 분석해 본 내용으로서, 여러 가정 요인들이 많기에 참고용으로만 봐주시길 바라며, "
          "혹 틀린부분이나 보완할 부분에 대해 댓글 등으로 의견 주시면 감사드리겠습니다.", ""]
    return "\n".join(L)


def excel(cfg: dict, base: Result, scen: list, sens: tuple, path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        existing_table(cfg).to_excel(w, sheet_name="대지지분_시세", index=False)
        new_table(cfg).to_excel(w, sheet_name="신축계획", index=False)
        budget_table(base).to_excel(w, sheet_name="수지표", index=False)
        ratio_table(base).to_excel(w, sheet_name="비례율", index=False)
        contribution_table(cfg, base).to_excel(w, sheet_name="분담금", index=False)
        for k, (sc, r) in enumerate(scen, start=1):
            compare_table(base.costs, r.costs, "사업비").to_excel(w, sheet_name=f"시나리오{k}_사업비", index=False)
            compare_table(base.revenues, r.revenues, "총수입").to_excel(w, sheet_name=f"시나리오{k}_수입", index=False)
            ratio_table(r).to_excel(w, sheet_name=f"시나리오{k}_비례율", index=False)
            contribution_table(cfg, r).to_excel(w, sheet_name=f"시나리오{k}_분담금", index=False)
        ratio, contrib, _ = sens
        (ratio.round(2).rename(index=man, columns=man)).to_excel(w, sheet_name="민감도_비례율(%)")
        (contrib.round(2).rename(index=man, columns=man)).to_excel(w, sheet_name="민감도_분담금(억원)")
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(
                    10, max(len(str(c.value or "")) for c in col) * 1.6)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("config", type=Path, help="단지 입력 파일 (.toml)")
    ap.add_argument("-o", "--out", type=Path, help="출력 폴더 (기본: output/<파일명>)")
    a = ap.parse_args(argv)

    cfg = tomllib.loads(a.config.read_text(encoding="utf-8"))
    cfg["existing"], cfg["new"], cfg["cost"] = cfg["existing"]["types"], cfg["new"]["types"], cfg["cost"]["items"]
    out = a.out or ROOT / "output" / a.config.stem
    out.mkdir(parents=True, exist_ok=True)

    base = compute(cfg)
    scen = []
    for sc in cfg.get("scenario", []):
        ov = {k: v for k, v in sc.items() if k != "title"}
        scen.append((sc, compute(cfg, ov)))
    sens = sensitivity(cfg, base, out)

    (out / "report.md").write_text(report(cfg, base, scen, sens, out), encoding="utf-8")
    excel(cfg, base, scen, sens, out / "tables.xlsx")

    print(f"[{cfg['info']['name']}] 기본 비례율 {pct(base.ratio)} "
          f"(총수입 {eok(base.revenue_total)}, 총사업비 {eok(base.cost_total)}, 종전자산 {eok(base.prior_total)})")
    for sc, r in scen:
        print(f"  · {sc['title']}: 비례율 {pct(r.ratio)} (총사업비 {eok(r.cost_total)}, 총수입 {eok(r.revenue_total)})")
    print(f"→ {out.relative_to(ROOT) if out.is_relative_to(ROOT) else out}/ : report.md, tables.xlsx, sens_ratio.png, sens_contrib.png")
    return 0


if __name__ == "__main__":
    sys.exit(main())
