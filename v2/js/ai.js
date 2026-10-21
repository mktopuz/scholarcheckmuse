"use strict";
/*
 * Optional AI review.
 *
 * Rule checks never send anything anywhere. This module is the only part
 * of the app that talks to a network, and only when the user has entered
 * their own API key and pressed "Run AI review". It sends TEXT EXCERPTS
 * (headings, front matter, samples) — never the document file — to the
 * provider at the base URL the user configured.
 *
 * The key lives in this browser's localStorage, on the user's device only.
 */

const AI_KEYS = { baseUrl: "tfc_ai_base_url", model: "tfc_ai_model", key: "tfc_ai_key" };

function aiLoadSettings() {
  return {
    baseUrl: localStorage.getItem(AI_KEYS.baseUrl) || "https://api.openai.com/v1",
    model: localStorage.getItem(AI_KEYS.model) || "gpt-4o-mini",
    key: localStorage.getItem(AI_KEYS.key) || ""
  };
}
function aiSaveSettings(s) {
  localStorage.setItem(AI_KEYS.baseUrl, s.baseUrl);
  localStorage.setItem(AI_KEYS.model, s.model);
  if (s.key) localStorage.setItem(AI_KEYS.key, s.key);
  else localStorage.removeItem(AI_KEYS.key);
}

function extractForAI(model) {
  const paras = model.paragraphs;
  const headings = paras
    .filter((p) => /^heading\s*[123]$/i.test(p.styleName || "") || /^Heading[123]$/.test(p.styleId || ""))
    .map((p) => ({ style: p.styleName || p.styleId, text: p.text.trim().slice(0, 160) }))
    .slice(0, 200);
  const frontMatter = paras.slice(0, 40).map((p) => p.text).join("\n").slice(0, 6000);
  const refIdx = paras.findIndex((p) => /^(references|bibliography|works cited)\b/i.test(p.text.trim()));
  const references = refIdx >= 0
    ? paras.slice(refIdx, refIdx + 25).map((p) => p.text).join("\n").slice(0, 5000)
    : "";
  const midStart = Math.max(0, Math.floor(paras.length / 2) - 10);
  const bodySample = paras.slice(midStart, midStart + 20).map((p) => p.text).join("\n").slice(0, 5000);
  return { headings, frontMatter, references, bodySample };
}

async function runAIReview(settings, extracted, ruleResults) {
  const base = settings.baseUrl.replace(/\/+$/, "");
  const ruleSummary = ruleResults
    .map((r) => "- [" + r.status.toUpperCase() + "] " + r.label + ": expected " + r.expected + "; found " + r.found)
    .join("\n");
  const userContent = [
    "You are reviewing the FORMATTING of a thesis/dissertation, not its arguments.",
    "Rule-based checks already ran with these results:",
    ruleSummary,
    "",
    "Headings found (style, text):",
    JSON.stringify(extracted.headings),
    "",
    "Front matter (first pages' text):",
    extracted.frontMatter || "(none)",
    "",
    "References sample:",
    extracted.references || "(no References heading found)",
    "",
    "Body text sample:",
    extracted.bodySample || "(none)",
    "",
    "Flag only formatting/structure issues a rule engine cannot judge: inconsistent heading capitalization or numbering, a missing or misplaced required part, caption or list wording that drifts, inconsistent reference formatting within the sample, anything else visibly irregular.",
    "Answer with a JSON array only, no prose outside it. Each item: {\"issue\": string, \"evidence\": string (short quote), \"suggestion\": string}. At most 8 items. If nothing is irregular, answer []."
  ].join("\n");

  const resp = await fetch(base + "/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + settings.key
    },
    body: JSON.stringify({
      model: settings.model,
      messages: [
        { role: "system", content: "You are a meticulous graduate-school format reviewer. Output JSON only." },
        { role: "user", content: userContent }
      ],
      temperature: 0.2
    })
  });
  if (!resp.ok) {
    let detail = "";
    try { const j = await resp.json(); detail = (j.error && j.error.message) || ""; } catch (e) { /* keep status only */ }
    throw new Error("The AI provider answered HTTP " + resp.status + (detail ? " — " + detail : "") + ". Check the base URL, model name and key.");
  }
  const data = await resp.json();
  const content = data && data.choices && data.choices[0] && data.choices[0].message
    ? data.choices[0].message.content : "";
  const jsonText = String(content).replace(/```json|```/g, "").trim();
  let items;
  try { items = JSON.parse(jsonText); } catch (e) {
    throw new Error("The AI answer was not the expected JSON list. Raw answer starts: " + String(content).slice(0, 120));
  }
  if (!Array.isArray(items)) throw new Error("The AI answer was not a list.");
  return items.slice(0, 8).map((it) => ({
    issue: String(it.issue || ""),
    evidence: String(it.evidence || ""),
    suggestion: String(it.suggestion || "")
  }));
}
