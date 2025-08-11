/* global Office, Word, PowerPoint */
const $ = (id) => document.getElementById(id);
const API = {
  generate: (prompt) => call("/api/generate", { prompt }),
  rewrite: (text, intent) => call("/api/rewrite", { text, intent })
};
let host = "word"; // "word" | "powerpoint"

Office.onReady((info) => {
  host = (info.host === Office.HostType.PowerPoint) ? "powerpoint" : "word";
  bindUI();
});

function bindUI() {
  $("btnGenerate").onclick = onGenerate;
  $("btnEdit").onclick = onEditSelection;
  $("btnApply").onclick = onApply;
  $("btnCopy").onclick = () => navigator.clipboard.writeText($("out").value || "");
}

function note(m) { $("msg").textContent = m; }
async function call(path, payload) {
  const r = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await r.json();
  if (!r.ok) throw new Error(data.error || "Erreur API");
  return data;
}

/** ---------- CRÉATION ---------- **/
async function onGenerate() {
  const prompt = $("prompt").value.trim();
  if (!prompt) return note("Prompt vide.");
  try {
    const { result } = await API.generate(prompt);
    $("out").value = result;
    note("Texte généré.");
  } catch (e) { note(e.message); }
}


/** ---------- PROPOSITION D'ÉDITION ---------- **/
async function onEditSelection() {
  const prompt = $("createPrompt").value.trim();
  if (!prompt) return note("Veuillez décrire ce que vous voulez créer.");
  const intent = $("intent").value;
  try {
    const { result } = await API.generate(`${prompt} (Style: ${intent})`);
    $("out").value = result;
    note("Proposition prête.");
  } catch (e) { note(e.message); }
}

/** ---------- APPLICATION + HISTORIQUE ---------- **/
async function onApply() {
  const suggestion = $("out").value.trim();
  if (!suggestion) return note("Rien à appliquer.");

  if (host === "word") {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();

      if ($("asComment").checked) {
        sel.insertComment("Suggestion GPT:\n\n" + suggestion);
      } else {
        sel.insertText(suggestion, Word.InsertLocation.replace);
      }

      await logHistoryWord(ctx, {
        t: new Date().toISOString(),
        host: "word",
        action: $("asComment").checked ? "comment" : "replace",
        original: sel.text || "",
        suggestion
      });
      await ctx.sync();
      note("Appliqué (Word) + historisé.");
    });
  } else {
    await PowerPoint.run(async (ctx) => {
      const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
      tr.load("text", "isNullObject");
      await ctx.sync();

      if (tr.isNullObject) return note("Sélection vide (PowerPoint).");

      const original = tr.text || "";
      tr.text = suggestion;

      await logHistoryPpt(ctx, {
        t: new Date().toISOString(),
        host: "powerpoint",
        action: "replace",
        original,
        suggestion
      });
      await ctx.sync();
      note("Appliqué (PowerPoint) + historisé.");
    });
  }
}

/** ---- Historique Word : propriétés personnalisées ---- */
async function logHistoryWord(ctx, entry) {
  const props = ctx.document.properties.customProperties;
  const existing = props.getItemOrNullObject("GPT_History");
  existing.load(["value", "isNullObject"]);
  await ctx.sync();

  let list = [];
  if (!existing.isNullObject && existing.value) {
    try { list = JSON.parse(existing.value); } catch {}
  }
  list.push(entry);
  props.add("GPT_History", JSON.stringify(list)); // écrase / remplace la valeur
}

/** ---- Historique PowerPoint : propriétés personnalisées (API 1.7) ---- */
async function logHistoryPpt(ctx, entry) {
  const props = ctx.presentation.properties.customProperties;
  const existing = props.getItemOrNullObject("GPT_History");
  existing.load(["value", "isNullObject"]);
  await ctx.sync();

  let list = [];
  if (!existing.isNullObject && existing.value) {
    try { list = JSON.parse(existing.value); } catch {}
  }
  list.push(entry);
  props.add("GPT_History", JSON.stringify(list));
}
