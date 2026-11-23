#!/usr/bin/env python3
"""QA hook scanner: find source-grounded hooks (superlatives, hedges, limitations, assumptions)
that are natural question targets, to seed qa_plan.json drafting.

Usage:
  python3 qa_draft.py --source /app/work/paper.txt [--out /app/work/qa_hooks.json] [--max 12]

Advisory only (always exit 0). Output: [{type, page, snippet}].
No-script fallback: search the keyword list (printed below) in paper.txt manually.
"""
import argparse
import json
import re
import sys
from pathlib import Path

HOOKS = [
    ("mechanism", r"\b(first|novel|only|sole|state-of-the-art|SOTA|outperform\w*|superior|assume[sd]?|assumption|isotropic|idealized|theor(?:y|em|etical)|converg\w*)\b|首次|唯一|最优|最先进|假设|理想化|理论|收敛"),
    ("data", r"(?:±|\bp\s*[<≤=]\s*0\.\d|\bvariance\b|\bstd\b|\bstandard\s+deviation\b|\bconfidence\s+interval\b|\bseed[s]?\b|\bnoisy\b|\bsample\s+size\b|\bdataset\b|\boverhead\b|\blatency\b|\bGPU[- ]hours?\b|\bmemory\b)|方差|标准差|噪声|样本量|开销|延迟|显存|代价"),
    ("baseline", r"\b(baseline[s]?|compared\s+(?:to|with)|ablation|without\s+\w+|fair\s+comparison|same\s+budget|hyperparameter|tuning|compute\s+budget)\b|基线|对照|消融|同预算|超参|公平对比"),
    ("limitation", r"\b(limitation[s]?|however|future\s+work|remains?|unclear|untested|unverified|only\s+(?:on|for|in)|restricted\s+to|preliminary|fail(?:s|ure)?|degrad\w*|boundar(?:y|ies))\b|局限|然而|未来工作|尚未|未验证|仅在|限于|初步|失效|边界"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--out")
    ap.add_argument("--max", type=int, default=12)
    ap.add_argument("--suggest", action="store_true", help="emit suggested question skeletons per hook")
    a = ap.parse_args()
    src = Path(a.source)
    if not src.exists():
        print(f"FATAL: source text missing: {src}", file=sys.stderr)
        sys.exit(3)
    text = src.read_text(encoding="utf-8", errors="ignore")
    marker = list(re.finditer(r"--- PAGE \d+(?: \([^)]+\))? ---", text))

    def page_of(pos):
        pg = 1
        for mk in marker:
            if mk.start() <= pos:
                pg = int(re.search(r"\d+", mk.group(0)).group(0))
            else:
                break
        return pg

    def sentence_span(pos):
        n = len(text)
        s = pos
        while s > 0 and pos - s < 220:
            ch = text[s - 1]
            if ch in ".!?。！？\n":
                decimal = ch == "." and s >= 2 and text[s - 2].isdigit() and text[s].isdigit()
                if not decimal:
                    break
            s -= 1
        e = pos
        while e < n and e - pos < 220:
            ch = text[e]
            if ch in ".!?。！？\n":
                decimal = ch == "." and e >= 1 and e + 1 < n and text[e - 1].isdigit() and text[e + 1].isdigit()
                if not decimal:
                    e += 1
                    break
            e += 1
        return text[s:e].replace("\n", " ").strip()

    def _tok_set(s):
        return set(re.findall(r"[a-z0-9]{3,}|[\u4e00-\u9fff]{2}", s.lower()))

    def _near_dup(s1, s2):
        if s1[:36].lower() == s2[:36].lower():
            return True
        t1, t2 = _tok_set(s1), _tok_set(s2)
        if not t1 or not t2:
            return False
        inter = len(t1 & t2)
        return (inter / min(len(t1), len(t2))) >= 0.60

    per_type = {typ: [] for typ, _ in HOOKS}
    all_kept = []
    for typ, pat in HOOKS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            span = sentence_span(m.start())
            if not span or len(span) < 16 or span.count("|") >= 2:
                continue
            if span.startswith(("--- PAGE", "--- EXTRA", "--- STARTING", "--- DECK", "--- SUPPLEMENT")):
                continue
            if len(span) > 130:
                cut = span[:130]
                span = cut[:cut.rfind(" ")] if " " in cut[40:] else cut
            pg = page_of(m.start())
            if any(_near_dup(span, prev["snippet"]) for prev in all_kept):
                continue
            item = {"type": typ, "page": pg, "snippet": span}
            per_type[typ].append(item)
            all_kept.append(item)
            if len(per_type[typ]) >= 4:
                break

    # Round-robin across the 4 rubric dimensions (mechanism/data/baseline/limitation) for balanced coverage
    hooks = []
    for r in range(4):
        for typ, _ in HOOKS:
            if r < len(per_type[typ]):
                hooks.append(per_type[typ][r])
                if len(hooks) >= a.max:
                    break
        if len(hooks) >= a.max:
            break
    if a.suggest:
        src_cjk = sum(1 for c in text[:4000] if "\u4e00" <= c <= "\u9fff") >= 20
        tmpl = ({
            "mechanism": ("关于机制/假设「{snippet}」，核心原理与成立前提是什么？", "mechanism"),
            "data": ("关于数据/开销「{snippet}」，统计显著性与成本口径如何？", "data"),
            "baseline": ("关于对照/消融「{snippet}」，对比预算与设置是否公平？", "baseline"),
            "limitation": ("关于边界/局限「{snippet}」，外推到其他场景会如何失效？", "limitation"),
        } if src_cjk else {
            "mechanism": ("What mechanistic rationale or evidence supports '{snippet}'?", "mechanism"),
            "data": ("What variance, sample-size, or cost evidence backs '{snippet}'?", "data"),
            "baseline": ("How does '{snippet}' ensure fair baseline/ablation budgets?", "baseline"),
            "limitation": ("How does this boundary ('{snippet}') limit external validity?", "limitation"),
        })
        for h in hooks:
            default_q = ("该发现的前提与边界是什么？", "mechanism") if src_cjk else ("What are the boundary conditions of this finding?", "mechanism")
            q_, cat = tmpl.get(h["type"], default_q)
            frag = h["snippet"][:48]
            if " " in frag[20:]:
                frag = frag[:frag.rfind(" ")]
            h["suggested_question"] = q_.format(snippet=frag)
            h["suggested_category"] = cat
    rep = {"n_hooks": len(hooks), "hooks": hooks}
    if a.out:
        Path(a.out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"qa hooks: {len(hooks)} found")
    for h in hooks:
        print(f"  [{h['type']}@p{h['page']}] ...{h['snippet'][:90]}...")
    sys.exit(0)


if __name__ == "__main__":
    main()
