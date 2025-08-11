import "dotenv/config";
import express from "express";
import cors from "cors";
import https from "https";
import path from "path";
import fs from "fs/promises";
import { fileURLToPath } from "url";
import * as devCerts from "office-addin-dev-certs";
import OpenAI from "openai";
import multer from "multer";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 3001;
const app = express();

// Sécurité & robustesse de base
app.use(express.json({ limit: "200kb" }));
// WSL: Allow both localhost and IP access
const wsIPv4 = process.env.WSL_HOST_IP || 'localhost';
app.use(cors({ 
  origin: [
    `https://localhost:${PORT}`, 
    `https://${wsIPv4}:${PORT}`,
    'https://word-edit.officeapps.live.com',
    'https://outlook.live.com'
  ] 
}));

// Statique: / → taskpane.html
app.use(express.static(path.join(__dirname, "public"), { index: "taskpane.html" }));

// Healthcheck
app.get("/healthz", (_, res) => res.json({ ok: true }));

// OpenAI client (clé via .env, jamais en clair)
const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// Tracking des tokens et coûts par session
let sessionStats = {
  "gpt-5": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 },
  "gpt-5-mini": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 },
  "gpt-5-nano": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 }
};

// Prix officiels par token (USD), convertis depuis USD par 1M tokens
// gpt-5: $1.25 (input), $0.125 (cache), $10.00 (output) per 1M → per token
// gpt-5-mini: $0.25 (input), $0.025 (cache), $2.00 (output) per 1M → per token
// gpt-5-nano: $0.05 (input), $0.005 (cache), $0.40 (output) per 1M → per token
const PRICING_PER_TOKEN = {
  "gpt-5": { input: 1.25 / 1_000_000, output: 10.0 / 1_000_000, cache: 0.125 / 1_000_000 },
  "gpt-5-mini": { input: 0.25 / 1_000_000, output: 2.0 / 1_000_000, cache: 0.025 / 1_000_000 },
  "gpt-5-nano": { input: 0.05 / 1_000_000, output: 0.4 / 1_000_000, cache: 0.005 / 1_000_000 }
};

// Fonction pour extraire les métadonnées de tokens de la réponse OpenAI
function extractTokenUsage(response, model) {
  const usage = response?.usage || {};

  // Support both Chat Completions-style and Responses API-style fields
  const inputTokens =
    (typeof usage.input_tokens === "number" ? usage.input_tokens : 0) ||
    (typeof usage.prompt_tokens === "number" ? usage.prompt_tokens : 0) ||
    0;

  const outputTokens =
    (typeof usage.output_tokens === "number" ? usage.output_tokens : 0) ||
    (typeof usage.completion_tokens === "number" ? usage.completion_tokens : 0) ||
    0;

  // Cached tokens can appear under different keys depending on API/version
  const cacheTokens =
    (usage.input_token_details && typeof usage.input_token_details.cached_tokens === "number"
      ? usage.input_token_details.cached_tokens
      : 0) ||
    (usage.prompt_tokens_details && typeof usage.prompt_tokens_details.cached_tokens === "number"
      ? usage.prompt_tokens_details.cached_tokens
      : 0) ||
    (typeof usage.prompt_tokens_cached === "number" ? usage.prompt_tokens_cached : 0);

  // Mettre à jour les statistiques de session (init si modèle inconnu)
  if (!sessionStats[model]) {
    sessionStats[model] = { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 };
  }
  sessionStats[model].requests += 1;
  sessionStats[model].inputTokens += inputTokens;
  sessionStats[model].outputTokens += outputTokens;
  sessionStats[model].cacheTokens += cacheTokens;

  return { inputTokens, outputTokens, cacheTokens };
}

// Fonction pour générer un rapport de session dans les logs
function generateSessionReport() {
  const pricing = PRICING_PER_TOKEN;

  console.log("\n" + "=".repeat(60));
  console.log("📊 RAPPORT SESSION - USAGE TOKENS & COÛTS");
  console.log("=".repeat(60));
  
  let globalTotalCost = 0;
  let globalTotalTokens = { input: 0, output: 0, cache: 0, total: 0 };
  let totalRequests = 0;

  // Rapport par modèle
  for (const [model, stats] of Object.entries(sessionStats)) {
    if (stats.requests > 0) {
      const modelPricing = pricing[model] || { input: 0, output: 0, cache: 0 };
      
      const inputCost = stats.inputTokens * modelPricing.input;
      const outputCost = stats.outputTokens * modelPricing.output;
      const cacheCost = stats.cacheTokens * modelPricing.cache;
      const modelTotal = inputCost + outputCost + cacheCost;
      const tokenTotal = stats.inputTokens + stats.outputTokens + stats.cacheTokens;
      
      console.log(`\n🤖 ${model.toUpperCase()}`);
      console.log(`   Requêtes: ${stats.requests}`);
      console.log(`   Tokens   - Input: ${stats.inputTokens.toLocaleString()}, Output: ${stats.outputTokens.toLocaleString()}, Cache: ${stats.cacheTokens.toLocaleString()}`);
      console.log(`   Total    - ${tokenTotal.toLocaleString()} tokens`);
      console.log(`   Coûts    - Input: $${inputCost.toFixed(6)}, Output: $${outputCost.toFixed(6)}, Cache: $${cacheCost.toFixed(6)}`);
      console.log(`   Total    - $${modelTotal.toFixed(6)}`);
      
      globalTotalCost += modelTotal;
      globalTotalTokens.input += stats.inputTokens;
      globalTotalTokens.output += stats.outputTokens;
      globalTotalTokens.cache += stats.cacheTokens;
      totalRequests += stats.requests;
    }
  }
  
  globalTotalTokens.total = globalTotalTokens.input + globalTotalTokens.output + globalTotalTokens.cache;
  
  // Résumé global
  console.log(`\n💰 TOTAL SESSION`);
  console.log(`   Requêtes: ${totalRequests}`);
  console.log(`   Tokens   - Input: ${globalTotalTokens.input.toLocaleString()}, Output: ${globalTotalTokens.output.toLocaleString()}, Cache: ${globalTotalTokens.cache.toLocaleString()}`);
  console.log(`   Total    - ${globalTotalTokens.total.toLocaleString()} tokens`);
  console.log(`   Coût     - $${globalTotalCost.toFixed(6)} USD`);
  
  console.log("=".repeat(60) + "\n");
}

// ---- API: Génération (prompt libre) ----
app.post("/api/generate", async (req, res) => {
  try {
    const { prompt, config } = req.body || {};
    if (!prompt || typeof prompt !== "string") throw new Error("Prompt invalide");
    
    // Configuration par défaut si non fournie
    const gptConfig = config || { model: "gpt-5", reasoningEffort: "medium", verbosity: "medium" };
    
    const requestParams = {
      model: gptConfig.model,
      reasoning: {
        effort: gptConfig.reasoningEffort
      },
      input: [
        { role: "system", content: "Tu es un rédacteur francophone précis. Respecte la consigne. Réponds uniquement par le texte demandé." },
        { role: "user", content: `Crée un texte répondant à: ${prompt}` }
      ]
    };
    
    // Ajouter verbosity si disponible (certains modèles le supportent)
    if (gptConfig.verbosity && gptConfig.verbosity !== "medium") {
      requestParams.verbosity = gptConfig.verbosity;
    }
    
    console.log("=== GÉNÉRATION ===");
    console.log("Config GPT:", gptConfig);
    console.log("Prompt utilisateur:", prompt);
    console.log("Paramètres envoyés à GPT:", JSON.stringify(requestParams, null, 2));
    
    const rsp = await client.responses.create(requestParams);
    const result = (rsp.output_text || "").trim();
    
    // Extraire et tracker l'usage des tokens
    const tokenUsage = extractTokenUsage(rsp, gptConfig.model);
    
    console.log("Réponse GPT reçue:", result);
    console.log("=== FIN GÉNÉRATION ===");
    
    // Générer le rapport de session après la réponse
    generateSessionReport();
    
    res.json({ result, tokenUsage });
  } catch (e) {
    console.error("=== ERREUR GÉNÉRATION ===");
    console.error("Message d'erreur:", e.message);
    console.error("Stack trace:", e.stack);
    console.error("=== FIN ERREUR ===\n");
    res.status(500).json({ error: e.message || "OpenAI error" });
  }
});

// ---- API: Réécriture (édition sur sélection) ----
app.post("/api/rewrite", async (req, res) => {
  try {
    const { text, intent, config } = req.body || {};
    if (!text || typeof text !== "string") throw new Error("Texte manquant");
    
    // Configuration par défaut si non fournie
    const gptConfig = config || { model: "gpt-5", reasoningEffort: "medium", verbosity: "medium" };
    
    const system = "Tu es un éditeur francophone rigoureux. Ne change pas le sens. Garde la voix de l'auteur. Réponds uniquement par le texte révisé.";
    const user = `Intention: ${intent || "clarify"}\nTexte:\n${text}`;
    
    const requestParams = {
      model: gptConfig.model,
      reasoning: {
        effort: gptConfig.reasoningEffort
      },
      input: [
        { role: "system", content: system },
        { role: "user", content: user }
      ]
    };
    
    // Ajouter verbosity si disponible
    if (gptConfig.verbosity && gptConfig.verbosity !== "medium") {
      requestParams.verbosity = gptConfig.verbosity;
    }
    
    console.log("=== RÉÉCRITURE ===");
    console.log("Config GPT:", gptConfig);
    console.log("Texte original:", text);
    console.log("Intention:", intent);
    console.log("Paramètres envoyés à GPT:", JSON.stringify(requestParams, null, 2));
    
    const rsp = await client.responses.create(requestParams);
    const result = (rsp.output_text || "").trim();
    
    // Extraire et tracker l'usage des tokens
    const tokenUsage = extractTokenUsage(rsp, gptConfig.model);
    
    console.log("Texte réécrit:", result);
    console.log("=== FIN RÉÉCRITURE ===");
    
    // Générer le rapport de session après la réponse
    generateSessionReport();
    
    res.json({ result, tokenUsage });
  } catch (e) {
    console.error("=== ERREUR RÉÉCRITURE ===");
    console.error("Message d'erreur:", e.message);
    console.error("Stack trace:", e.stack);
    console.error("=== FIN ERREUR ===\n");
    res.status(500).json({ error: e.message || "OpenAI error" });
  }
});

// ---- API: Budget et statistiques de session ----
app.get("/api/budget", (req, res) => {
  try {
    // Tarifs officiels en USD par token (convertis depuis /1M)
    const pricing = PRICING_PER_TOKEN;

    const budget = {};
    let totalCost = 0;
    let totalTokens = { input: 0, output: 0, cache: 0 };

    for (const [model, stats] of Object.entries(sessionStats)) {
      const modelPricing = pricing[model] || { input: 0, output: 0, cache: 0 };
      
      const inputCost = stats.inputTokens * modelPricing.input;
      const outputCost = stats.outputTokens * modelPricing.output;
      const cacheCost = stats.cacheTokens * modelPricing.cache;
      const modelTotal = inputCost + outputCost + cacheCost;
      
      budget[model] = {
        requests: stats.requests,
        tokens: {
          input: stats.inputTokens,
          output: stats.outputTokens,
          cache: stats.cacheTokens,
          total: stats.inputTokens + stats.outputTokens + stats.cacheTokens
        },
        costs: {
          input: inputCost,
          output: outputCost,
          cache: cacheCost,
          total: modelTotal
        }
      };
      
      totalCost += modelTotal;
      totalTokens.input += stats.inputTokens;
      totalTokens.output += stats.outputTokens;
      totalTokens.cache += stats.cacheTokens;
    }

    res.json({
      session: {
        totalRequests: Object.values(sessionStats).reduce((sum, stats) => sum + stats.requests, 0),
        totalTokens: {
          input: totalTokens.input,
          output: totalTokens.output,
          cache: totalTokens.cache,
          total: totalTokens.input + totalTokens.output + totalTokens.cache
        },
        totalCostUSD: totalCost
      },
      models: budget,
      pricing
    });
  } catch (e) {
    console.error("Erreur budget:", e.message);
    res.status(500).json({ error: e.message });
  }
});

// ---- API: Reset des statistiques ----
app.post("/api/budget/reset", (req, res) => {
  try {
    sessionStats = {
      "gpt-5": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 },
      "gpt-5-mini": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 },
      "gpt-5-nano": { requests: 0, inputTokens: 0, outputTokens: 0, cacheTokens: 0 }
    };
    console.log("Statistiques de session réinitialisées");
    res.json({ message: "Statistiques réinitialisées", stats: sessionStats });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ---- Configuration multer pour l'upload de fichiers ----
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      const profileId = req.params.profileId;
      const uploadPath = path.join(__dirname, 'profiles', profileId, 'docs');
      fs.mkdir(uploadPath, { recursive: true }).then(() => {
        cb(null, uploadPath);
      }).catch(cb);
    },
    filename: (req, file, cb) => {
      cb(null, file.originalname);
    }
  }),
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.txt', '.md', '.json'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error(`Type de fichier non supporté: ${ext}`));
    }
  },
  limits: { fileSize: 5 * 1024 * 1024 } // 5MB max
});

// ---- API: Gestion des profils d'instructions ----

// Fonction utilitaire pour lire les profils
async function readProfiles() {
  try {
    const data = await fs.readFile('profiles.json', 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.error('Erreur lecture profiles.json:', error);
    return { profiles: { default: { name: "Par défaut", createInstruction: "", editInstruction: "", docFiles: [] } }, activeProfile: "default" };
  }
}

// Fonction utilitaire pour sauvegarder les profils
async function saveProfiles(profilesData) {
  try {
    await fs.writeFile('profiles.json', JSON.stringify(profilesData, null, 2));
  } catch (error) {
    console.error('Erreur sauvegarde profiles.json:', error);
    throw error;
  }
}

// GET /api/profiles - Récupérer tous les profils
app.get('/api/profiles', async (req, res) => {
  try {
    const profilesData = await readProfiles();
    res.json(profilesData);
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la lecture des profils' });
  }
});

// POST /api/profiles - Créer/Modifier un profil
app.post('/api/profiles', async (req, res) => {
  try {
    const { profileId, profile } = req.body;
    if (!profileId || !profile) {
      return res.status(400).json({ error: 'profileId et profile requis' });
    }

    const profilesData = await readProfiles();
    profilesData.profiles[profileId] = profile;
    
    // Créer le dossier docs si nécessaire
    const docsPath = path.join(__dirname, 'profiles', profileId, 'docs');
    await fs.mkdir(docsPath, { recursive: true });
    
    await saveProfiles(profilesData);
    res.json({ message: 'Profil sauvegardé avec succès' });
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la sauvegarde du profil' });
  }
});

// DELETE /api/profiles/:profileId - Supprimer un profil
app.delete('/api/profiles/:profileId', async (req, res) => {
  try {
    const { profileId } = req.params;
    if (profileId === 'default') {
      return res.status(400).json({ error: 'Impossible de supprimer le profil par défaut' });
    }

    const profilesData = await readProfiles();
    if (!profilesData.profiles[profileId]) {
      return res.status(404).json({ error: 'Profil non trouvé' });
    }

    delete profilesData.profiles[profileId];
    
    // Si le profil actif était celui supprimé, revenir au défaut
    if (profilesData.activeProfile === profileId) {
      profilesData.activeProfile = 'default';
    }

    // Supprimer le dossier du profil
    const profilePath = path.join(__dirname, 'profiles', profileId);
    await fs.rm(profilePath, { recursive: true, force: true });
    
    await saveProfiles(profilesData);
    res.json({ message: 'Profil supprimé avec succès' });
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la suppression du profil' });
  }
});

// POST /api/profiles/:profileId/upload - Upload fichier documentation
app.post('/api/profiles/:profileId/upload', upload.array('docs'), async (req, res) => {
  try {
    const { profileId } = req.params;
    const files = req.files;

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'Aucun fichier fourni' });
    }

    const profilesData = await readProfiles();
    if (!profilesData.profiles[profileId]) {
      return res.status(404).json({ error: 'Profil non trouvé' });
    }

    // Ajouter les fichiers à la liste
    const newFiles = files.map(file => file.filename);
    const existingFiles = profilesData.profiles[profileId].docFiles || [];
    profilesData.profiles[profileId].docFiles = [...existingFiles, ...newFiles];

    await saveProfiles(profilesData);
    res.json({ message: `${files.length} fichier(s) uploadé(s) avec succès`, files: newFiles });
  } catch (error) {
    console.error('Erreur upload:', error);
    res.status(500).json({ error: 'Erreur lors de l\'upload' });
  }
});

// DELETE /api/profiles/:profileId/docs/:filename - Supprimer un fichier documentation
app.delete('/api/profiles/:profileId/docs/:filename', async (req, res) => {
  try {
    const { profileId, filename } = req.params;
    
    const profilesData = await readProfiles();
    if (!profilesData.profiles[profileId]) {
      return res.status(404).json({ error: 'Profil non trouvé' });
    }

    // Supprimer le fichier de la liste
    const docFiles = profilesData.profiles[profileId].docFiles || [];
    profilesData.profiles[profileId].docFiles = docFiles.filter(f => f !== filename);

    // Supprimer le fichier physique
    const filePath = path.join(__dirname, 'profiles', profileId, 'docs', filename);
    await fs.unlink(filePath);

    await saveProfiles(profilesData);
    res.json({ message: 'Fichier supprimé avec succès' });
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la suppression du fichier' });
  }
});

// POST /api/profiles/active - Changer le profil actif
app.post('/api/profiles/active', async (req, res) => {
  try {
    const { profileId } = req.body;
    if (!profileId) {
      return res.status(400).json({ error: 'profileId requis' });
    }

    const profilesData = await readProfiles();
    if (!profilesData.profiles[profileId]) {
      return res.status(404).json({ error: 'Profil non trouvé' });
    }

    profilesData.activeProfile = profileId;
    await saveProfiles(profilesData);
    res.json({ message: 'Profil actif mis à jour' });
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors du changement de profil actif' });
  }
});

// Fonction utilitaire pour lire le contenu des docs d'un profil
async function readProfileDocs(profileId) {
  try {
    const profilesData = await readProfiles();
    const profile = profilesData.profiles[profileId];
    if (!profile || !profile.docFiles) return '';

    const docsPath = path.join(__dirname, 'profiles', profileId, 'docs');
    const docsContent = [];

    for (const filename of profile.docFiles) {
      try {
        const filePath = path.join(docsPath, filename);
        const content = await fs.readFile(filePath, 'utf8');
        docsContent.push(`\n--- Documentation: ${filename} ---\n${content}\n`);
      } catch (error) {
        console.warn(`Erreur lecture fichier ${filename}:`, error.message);
      }
    }

    return docsContent.join('\n');
  } catch (error) {
    console.error('Erreur lecture docs profil:', error);
    return '';
  }
}

// HTTPS local avec certificats dev (auto-install)
const start = async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(PORT, '0.0.0.0', () => {
    console.log(`HTTPS add-in server on https://localhost:${PORT}`);
    if (process.env.WSL_HOST_IP) {
      console.log(`WSL: Also accessible via https://${process.env.WSL_HOST_IP}:${PORT}`);
    }
  });
};
start().catch(err => {
  console.error("Failed to start HTTPS server:", err);
  process.exit(1);
});
