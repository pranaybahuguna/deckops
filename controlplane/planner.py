"""The model step. Produces DeckPlans; never sees XML, coordinates or file bytes.

Two implementations behind one interface:
- StubPlanner: offline and deterministic, so the POC runs here with no model access.
- OpenAICompatPlanner: talks to any OpenAI-compatible /v1/chat/completions endpoint
  (vLLM, an internal gateway, a hosted API). Configured explicitly - never from environment variables.
"""
from __future__ import annotations

import json
import urllib.request

TRANSLATE_SYSTEM = """You translate PowerPoint paragraphs. Each paragraph is given as tagged text:
<rN>...</rN> is a formatted run (keep its formatting by keeping the tag), <br/> is a line break,
<fN/> is a field (slide number/date) that must appear exactly once, unchanged.
Rules: translate only the text inside tags; you MAY reorder, merge (drop a tag) or repeat tags
when the target grammar needs it; never invent tag numbers; keep numbers, tickers and product names.
Return JSON: {"paragraphs":[{"ref":"...","tagged":"..."}]} with the same refs."""

SHORTEN_SYSTEM = """Shorten each tagged paragraph by about {pct}% in {lang} while keeping its meaning and its tags.
Return JSON: {{"paragraphs":[{{"ref":"...","tagged":"..."}}]}}."""

CREATE_SYSTEM = """Plan a slide deck. Output JSON {"slides":[...]}; each slide has
"message" (the one sentence the slide proves), "layout_family" (one of: title, section,
title_and_body, two_content, title_only), "blocks" ([{"role":"title|subtitle|body",
"content":{"kind":"text","text":...} or {"kind":"bullets","items":[...]}}]) and optional "notes".
Never output coordinates, sizes, colours or fonts."""


KIT_SYSTEM = """You plan a slide deck using ONLY the slide patterns listed below, and ONLY facts from the
content pack (never invent numbers). Output JSON:
{"footer": "<short deck footer>", "slides": [{"pattern": "<name>", "fields": {...}, "items": [...],
  "table": {"header": [...], "rows": [[...]]}, "notes": "<optional speaker notes>",
  "sources": ["<fact ids used>"]}]}
Use "items" only for patterns that list items, "table" only for table patterns, and
"charts": {"<slot>": {"type": "...", "dataset": "<id>", "category": "<column>", "series": ["<column>", ...],
"title": "...", "number_format": "#,##0"}} only for chart patterns - pick a dataset from the content pack;
the numbers are taken from it automatically, never write numbers into a chart yourself. Keep text short:
headings under 6 words, bullet points under 15 words. Never output positions, sizes, colours or fonts.
Patterns:
{patterns}"""


class StubPlanner:
    name = "stub (offline, deterministic)"

    FR = {
        "Q3 Results": "<r0>Résultats du troisième trimestre</r0>",
        # word order changes: the bold "12%" run moves to the front
        "Revenue grew 12% year on year": "<r1>+12 %</r1><r0> de croissance du chiffre d'affaires</r0><r2> sur un an</r2>",
        "Cost-to-income ratio improved to 58%": "<r0>Le coefficient d'exploitation s'est amélioré à 58 %</r0>",
        "Client franchise expanded across all regions": "<r0>La franchise clients s'est développée dans toutes les régions</r0>",
        "Revenue by region": "<r0>Chiffre d'affaires par région</r0>",
        "Region": "<r0>Région</r0>", "Change": "<r0>Variation</r0>",
        "Europe": "<r0>Europe</r0>", "Asia Pacific": "<r0>Asie-Pacifique</r0>",
        "€1.2bn": "<r0>1,2 Md€</r0>", "€0.6bn": "<r0>0,6 Md€</r0>",
        "Q3 2026": "<r0>T3 2026</r0>",
        "Strengths and risks": "<r0>Forces et risques</r0>",
        "Strong fee income": "<r0>Solides revenus de commissions</r0>",
        "Rate volatility": "<r0>Volatilité des taux</r0>",
        "Emphasise the fee income story before discussing risks.":
            "<r0>Insister sur les revenus de commissions avant d'aborder les risques.</r0>",
        "Outlook for the fourth quarter": "<r0>Perspectives pour le quatrième trimestre</r0>",
        "Deposit margins expected to stabilise as rate cuts slow":
            "<r0>Les marges sur dépôts devraient se stabiliser à mesure que le rythme des baisses de taux ralentit</r0>",
        "Pipeline in structured financing remains strong into year end":
            "<r0>Le portefeuille d'opérations en financements structurés reste solide jusqu'à la fin de l'exercice</r0>",
        "Cost discipline continues with technology investment protected":
            "<r0>La discipline sur les coûts se poursuit tout en préservant les investissements technologiques</r0>",
        "Capital ratios stay comfortably above regulatory requirements":
            "<r0>Les ratios de fonds propres restent confortablement supérieurs aux exigences réglementaires</r0>",
        "Client activity in Asia expected to remain the main growth driver":
            "<r0>L'activité clients en Asie devrait rester le principal moteur de croissance</r0>",
    }
    FR_SHORT = {
        "Deposit margins expected to stabilise as rate cuts slow":
            "<r0>Marges sur dépôts en voie de stabilisation</r0>",
        "Pipeline in structured financing remains strong into year end":
            "<r0>Financements structurés : portefeuille solide</r0>",
        "Cost discipline continues with technology investment protected":
            "<r0>Coûts maîtrisés, investissements tech préservés</r0>",
    }

    def translate(self, paragraphs: list[dict], lang: str) -> dict:
        out = []
        for p in paragraphs:
            if p["text"].strip() == "":
                continue
            out.append({"ref": p["ref"], "tagged": self.FR.get(p["text"].replace("\n", " "), p["tagged"])})
        return {"mode": "translate", "target_lang": lang, "paragraphs": out}

    def shorten(self, paragraphs: list[dict], lang: str, pct: int) -> dict:
        out = [{"ref": p["ref"], "tagged": self.FR_SHORT[p["text"]]} for p in paragraphs if p["text"] in self.FR_SHORT]
        return {"paragraphs": out}

    def plan_from_kit(self, brief: str, kit_description: str, content: dict) -> dict:
        raise NotImplementedError("the offline stub has no model: pass a ready-made plan instead")

    def plan_create(self, brief: str) -> dict:
        return {"mode": "create", "slides": [
            {"message": "Title", "layout_family": "title", "blocks": [
                {"role": "title", "content": {"kind": "text", "text": "Passive mandates: fee compression"}},
                {"role": "subtitle", "content": {"kind": "text", "text": "Asset Management strategy review · Nov 2026"}}]},
            {"message": "Fee compression is concentrated in passive mandates", "layout_family": "title_and_body",
             "blocks": [
                 {"role": "title", "content": {"kind": "text", "text": "Fee compression is concentrated in passive mandates"}},
                 {"role": "body", "content": {"kind": "bullets", "items": [
                     "Passive fees down 9 bps over three years", "Active fees broadly stable",
                     "Mix shift explains two thirds of revenue decline"]}}],
             "notes": "Source: internal fee database, ev-12."},
            {"message": "Options compared", "layout_family": "two_content", "blocks": [
                {"role": "title", "content": {"kind": "text", "text": "Two responses"}},
                {"role": "body", "content": {"kind": "bullets", "items": ["Scale: consolidate index range", "Lower unit cost"]}},
                {"role": "body", "content": {"kind": "bullets", "items": ["Differentiate: thematic ETFs", "Higher margin, higher risk"]}}]},
            {"message": "Next steps", "layout_family": "section", "blocks": [
                {"role": "title", "content": {"kind": "text", "text": "Decision needed by Q1"}},
                {"role": "body", "content": {"kind": "text", "text": "Steering committee, 15 January"}}]},
        ]}


class OpenAICompatPlanner:
    def __init__(self, base_url: str, model: str, api_key: str = ""):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.model = model
        self.name = f"openai-compatible ({self.model})"

    def _chat(self, system, user) -> dict:
        body = json.dumps({"model": self.model, "temperature": 0,
                           "response_format": {"type": "json_object"},
                           "messages": [{"role": "system", "content": system},
                                        {"role": "user", "content": user}]}).encode()
        req = urllib.request.Request(f"{self.base_url}/v1/chat/completions", data=body, headers={
            "Content-Type": "application/json", "Authorization": f"Bearer {self.api_key}"})
        with urllib.request.urlopen(req, timeout=120) as r:
            return json.loads(json.load(r)["choices"][0]["message"]["content"])

    def translate(self, paragraphs, lang):
        items = [{"ref": p["ref"], "tagged": p["tagged"]} for p in paragraphs if p["text"].strip()]
        res = self._chat(TRANSLATE_SYSTEM, json.dumps({"target_language": lang, "paragraphs": items}, ensure_ascii=False))
        return {"mode": "translate", "target_lang": lang, "paragraphs": res["paragraphs"]}

    def shorten(self, paragraphs, lang, pct):
        items = [{"ref": p["ref"], "tagged": p["tagged"]} for p in paragraphs]
        return self._chat(SHORTEN_SYSTEM.format(pct=pct, lang=lang), json.dumps({"paragraphs": items}, ensure_ascii=False))

    def plan_create(self, brief):
        return {"mode": "create", **self._chat(CREATE_SYSTEM, brief)}

    def plan_from_kit(self, brief, kit_description, content):
        user = json.dumps({"request": brief, "content_pack": content}, ensure_ascii=False)
        return self._chat(KIT_SYSTEM.replace("{patterns}", kit_description), user)

    @classmethod
    def from_file(cls, path: str) -> "OpenAICompatPlanner":
        """JSON file: {"base_url": "...", "model": "...", "api_key": "..."}. Keep it out of git."""
        cfg = json.load(open(path))
        return cls(cfg["base_url"], cfg["model"], cfg.get("api_key", ""))
