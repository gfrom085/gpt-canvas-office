/* global Office, Word, PowerPoint */

Office.onReady((info) => {
  // Initialize the add-in immediately
  console.log("Office.onReady called, host:", info.host);
  try {
    initializeAddIn();
    console.log("Add-in initialized successfully");
  } catch (error) {
    console.error("Error during initialization:", error);
  }
});

const $ = (id) => document.getElementById(id);

// Configuration GPT par défaut
let gptConfig = {
  model: "gpt-5",
  reasoningEffort: "medium",
  verbosity: "medium"
};

// Gestion des profils d'instructions
let profilesData = null;
let currentProfile = 'default';

const API = {
  generate: (prompt) => call("/api/generate", { prompt, config: gptConfig }),
  rewrite: (text, intent) => call("/api/rewrite", { text, intent, config: gptConfig })
};
let host = "word"; // "word" | "powerpoint"

function initializeAddIn() {
  host = Office.context.host === Office.HostType.PowerPoint ? "powerpoint" : "word";
  bindUI();
}

function bindUI() {
  console.log("Binding UI events...");
  
  $("btnExecute").onclick = onExecute;
  $("btnGrab").onclick = onGrabSelection;
  
  // Gestionnaires pour les paramètres
  $("btnSettings").onclick = openSettings;
  $("btnSaveSettings").onclick = saveSettings;
  $("btnCloseSettings").onclick = closeSettings;
  
  // Gestionnaires pour les profils d'instructions
  $("profileSelect").onchange = onProfileChange;
  $("btnNewProfile").onclick = createNewProfile;
  $("btnSaveProfile").onclick = saveCurrentProfile;
  $("btnDuplicateProfile").onclick = duplicateCurrentProfile;
  $("btnDeleteProfile").onclick = deleteCurrentProfile;
  $("docUpload").onchange = uploadDocuments;
  
  // Auto-chargement de la sélection au focus de l'add-in
  setTimeout(autoGrabSelection, 500);
  
  // Mise à jour dynamique des placeholders
  $("src").oninput = updatePromptPlaceholder;
  updatePromptPlaceholder();
  
  // Charger les profils d'instructions
  loadProfiles();
  
  console.log("All UI events bound");
}

function note(m) { 
  $("msg").textContent = m; 
  console.log("GPT Canvas:", m);
}

/** ---------- LOADING OVERLAY ---------- **/
function showLoading(message = "Traitement en cours...") {
  const overlay = $("loadingOverlay");
  const text = overlay.querySelector('.loading-text');
  text.textContent = message;
  overlay.style.display = "flex";
}

function hideLoading() {
  $("loadingOverlay").style.display = "none";
}

/** ---------- PURGE DES CHAMPS ---------- **/
function clearFields() {
  // Vider les champs
  $("createPrompt").value = "";
  $("src").value = "";
  
  // Remettre les placeholders (ils sont déjà définis dans updatePromptPlaceholder)
  updatePromptPlaceholder();
  
  // Petit effet de surbrillance sur les placeholders
  const createField = $("createPrompt");
  const srcField = $("src");
  
  createField.style.backgroundColor = "#fff3cd";
  srcField.style.backgroundColor = "#fff3cd";
  
  setTimeout(() => {
    createField.style.backgroundColor = "";
    srcField.style.backgroundColor = "";
  }, 1000);
}

async function call(path, payload) {
  try {
    console.log("Current URL:", window.location.href);
    console.log("Making API call to:", path, "with payload:", payload);
    // Ensure absolute URL for API calls
    const fullUrl = window.location.origin + path;
    console.log("Full API URL:", fullUrl);
    const r = await fetch(fullUrl, {
      method: "POST",
      headers: { 
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      body: JSON.stringify(payload)
    });
    console.log("API Response status:", r.status, "OK:", r.ok);
    const data = await r.json();
    console.log("API Response data:", data);
    if (!r.ok) throw new Error(data.error || `HTTP ${r.status}: ${r.statusText}`);
    return data;
  } catch (error) {
    console.error("API Error:", error);
    throw error;
  }
}

/** ---------- EXÉCUTION UNIFIÉE ---------- **/
async function onExecute() {
  const prompt = $("createPrompt").value.trim();
  const selection = $("src").value.trim();
  
  if (!prompt) {
    return note("Veuillez saisir un prompt.");
  }
  
  // RÈGLE SIMPLE : Auto-détection du mode
  if (selection) {
    // Mode EDIT : il y a du texte sélectionné
    await executeEdit(selection, prompt);
  } else {
    // Mode CREATE : pas de sélection - insérer au curseur
    await executeCreate(prompt);
  }
}

async function executeCreate(prompt) {
  showLoading("Création du texte...");
  try {
    const { result } = await API.generate(prompt);
    
    // TOUJOURS insérer à la position du curseur (jamais en commentaire pour la création)
    await insertAtCursor(result);
    
    hideLoading();
    note("Texte créé et inséré au curseur.");
    clearFields();
  } catch (e) { 
    hideLoading();
    note("Erreur: " + e.message); 
  }
}

async function executeEdit(selection, instructions) {
  showLoading("Édition du texte...");
  try {
    // Utiliser les instructions personnalisées de l'utilisateur
    const { result } = await API.rewrite(selection, instructions);
    
    // Remplacer la sélection par le texte édité
    await replaceSelection(result);
    
    hideLoading();
    note("Texte modifié avec succès.");
    clearFields();
  } catch (e) { 
    hideLoading();
    note("Erreur: " + e.message); 
  }
}

/** ---------- SÉLECTION ---------- **/
async function onGrabSelection() {
  if (host === "word") {
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        $("src").value = sel.text || "";
        updatePromptPlaceholder();
        note("Sélection Word chargée.");
      });
    } catch (error) {
      note("Erreur lors de la récupération: " + error.message);
    }
  } else {
    try {
      await PowerPoint.run(async (ctx) => {
        const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
        tr.load("text", "isNullObject");
        await ctx.sync();
        if (tr.isNullObject) {
          $("src").value = "";
          return note("Aucun texte sélectionné dans PowerPoint.");
        }
        $("src").value = tr.text || "";
        updatePromptPlaceholder();
        note("Sélection PowerPoint chargée.");
      });
    } catch (error) {
      note("Erreur lors de la récupération: " + error.message);
    }
  }
}

/** ---------- AUTO-CHARGEMENT DE LA SÉLECTION ---------- **/
async function autoGrabSelection() {
  if (host === "word") {
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        if (sel.text && sel.text.trim()) {
          $("src").value = sel.text;
          // Mettre à jour le placeholder pour le mode edit
          updatePromptPlaceholder();
        }
      });
    } catch (error) {
      console.log("Pas de sélection automatique:", error.message);
    }
  } else {
    try {
      await PowerPoint.run(async (ctx) => {
        const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
        tr.load("text", "isNullObject");
        await ctx.sync();
        if (!tr.isNullObject && tr.text && tr.text.trim()) {
          $("src").value = tr.text;
          // Mettre à jour le placeholder pour le mode edit
          updatePromptPlaceholder();
        }
      });
    } catch (error) {
      console.log("Pas de sélection automatique:", error.message);
    }
  }
}

/** ---------- GESTION DU PLACEHOLDER ---------- **/
function updatePromptPlaceholder() {
  const selection = $("src").value.trim();
  const promptField = $("createPrompt");
  
  if (selection) {
    // Mode EDIT : il y a du texte sélectionné
    promptField.placeholder = "Instructions pour modifier le texte sélectionné...";
  } else {
    // Mode CREATE : pas de sélection
    promptField.placeholder = "Décris ce que tu veux créer...";
  }
}

/** ---- Historique Word : propriétés personnalisées ---- */
async function logHistoryWord(ctx, entry) {
  try {
    const props = ctx.document.properties.customProperties;
    const existing = props.getItemOrNullObject("GPT_History");
    existing.load(["value", "isNullObject"]);
    await ctx.sync();

    let list = [];
    if (!existing.isNullObject && existing.value) {
      try { list = JSON.parse(existing.value); } catch {}
    }
    list.push(entry);
    props.add("GPT_History", JSON.stringify(list));
  } catch (error) {
    console.warn("Could not save history:", error);
  }
}

/** ---- Historique PowerPoint : propriétés personnalisées ---- */
async function logHistoryPpt(ctx, entry) {
  try {
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
  } catch (error) {
    console.warn("Could not save history:", error);
  }
}

/** ---------- FONCTIONS UTILITAIRES D'INSERTION ---------- **/

/** Insérer du texte à la position actuelle du curseur */
async function insertAtCursor(text) {
  if (host === "word") {
    await Word.run(async (ctx) => {
      const selection = ctx.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
    });
  } else {
    await PowerPoint.run(async (ctx) => {
      const textRange = ctx.presentation.getSelectedTextRangeOrNullObject();
      textRange.load("isNullObject");
      await ctx.sync();
      if (!textRange.isNullObject) {
        textRange.text = text;
        await ctx.sync();
      }
    });
  }
}

/** ---------- INSERTION COMME COMMENTAIRE ---------- **/
async function insertAsComment(text) {
  if (host === "word") {
    await Word.run(async (ctx) => {
      const selection = ctx.document.getSelection();
      selection.insertComment("GPT Canvas:\n\n" + text);
      await ctx.sync();
    });
  }
  // PowerPoint ne supporte pas les commentaires de la même manière
}

/** Remplacer la sélection courante par le nouveau texte */
async function replaceSelection(newText) {
  if (host === "word") {
    await Word.run(async (ctx) => {
      const selection = ctx.document.getSelection();
      selection.insertText(newText, Word.InsertLocation.replace);
      await ctx.sync();
    });
  } else {
    await PowerPoint.run(async (ctx) => {
      const textRange = ctx.presentation.getSelectedTextRangeOrNullObject();
      textRange.load("isNullObject");
      await ctx.sync();
      if (!textRange.isNullObject) {
        textRange.text = newText;
        await ctx.sync();
      }
    });
  }
}

/** ---------- GESTION DES PARAMÈTRES ---------- **/

function openSettings() {
  // Charger les valeurs actuelles dans l'interface
  $("gptModel").value = gptConfig.model;
  $("reasoningEffort").value = gptConfig.reasoningEffort;
  $("verbosity").value = gptConfig.verbosity;
  
  // Afficher le panneau
  $("settingsPanel").style.display = "block";
}

function saveSettings() {
  // Sauvegarder les nouveaux paramètres
  gptConfig.model = $("gptModel").value;
  gptConfig.reasoningEffort = $("reasoningEffort").value;
  gptConfig.verbosity = $("verbosity").value;
  
  console.log("Paramètres GPT sauvegardés:", gptConfig);
  note(`Paramètres sauvegardés: ${gptConfig.model}`);
  
  // Fermer le panneau
  closeSettings();
}

function closeSettings() {
  $("settingsPanel").style.display = "none";
}

/** ---------- GESTION DES PROFILS D'INSTRUCTIONS ---------- **/

// Charger tous les profils depuis le serveur
async function loadProfiles() {
  try {
    const response = await fetch('/api/profiles');
    profilesData = await response.json();
    currentProfile = profilesData.activeProfile;
    
    // Mettre à jour le dropdown des profils
    updateProfileSelect();
    
    // Charger le profil actif
    loadCurrentProfile();
    
    console.log("Profils chargés:", profilesData);
  } catch (error) {
    console.error("Erreur chargement profils:", error);
    note("Erreur lors du chargement des profils");
  }
}

// Mettre à jour le dropdown des profils
function updateProfileSelect() {
  const select = $("profileSelect");
  select.innerHTML = "";
  
  for (const [id, profile] of Object.entries(profilesData.profiles)) {
    const option = document.createElement("option");
    option.value = id;
    option.textContent = profile.name;
    option.selected = id === currentProfile;
    select.appendChild(option);
  }
}

// Charger le profil actuel dans l'interface
function loadCurrentProfile() {
  const profile = profilesData.profiles[currentProfile];
  if (profile) {
    $("createInstruction").value = profile.createInstruction || "";
    $("editInstruction").value = profile.editInstruction || "";
    updateDocsList(profile.docFiles || []);
  }
}

// Mettre à jour la liste des documents
function updateDocsList(docFiles) {
  const docsList = $("docsList");
  docsList.innerHTML = "";
  
  if (!docFiles || docFiles.length === 0) {
    docsList.innerHTML = '<div style="color: #666; font-style: italic;">Aucun document ajouté</div>';
    return;
  }
  
  docFiles.forEach(filename => {
    const docItem = document.createElement("div");
    docItem.className = "doc-item";
    docItem.innerHTML = `
      <span class="doc-filename">${filename}</span>
      <button class="doc-delete" onclick="deleteDocument('${filename}')">✕</button>
    `;
    docsList.appendChild(docItem);
  });
}

// Changement de profil
async function onProfileChange() {
  const newProfileId = $("profileSelect").value;
  if (newProfileId !== currentProfile) {
    currentProfile = newProfileId;
    
    // Changer le profil actif côté serveur
    try {
      await fetch('/api/profiles/active', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ profileId: currentProfile })
      });
      
      loadCurrentProfile();
      note(`Profil "${profilesData.profiles[currentProfile].name}" activé`);
    } catch (error) {
      console.error("Erreur changement profil:", error);
      note("Erreur lors du changement de profil");
    }
  }
}

// Créer un nouveau profil
async function createNewProfile() {
  const profileName = prompt("Nom du nouveau profil :");
  if (!profileName) return;
  
  const profileId = profileName.toLowerCase().replace(/[^a-z0-9]/g, '-');
  
  if (profilesData.profiles[profileId]) {
    alert("Un profil avec ce nom existe déjà");
    return;
  }
  
  const newProfile = {
    name: profileName,
    createInstruction: "",
    editInstruction: "",
    docFiles: []
  };
  
  try {
    await fetch('/api/profiles', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ profileId, profile: newProfile })
    });
    
    // Recharger les profils
    await loadProfiles();
    
    // Activer le nouveau profil
    currentProfile = profileId;
    $("profileSelect").value = currentProfile;
    loadCurrentProfile();
    
    note(`Nouveau profil "${profileName}" créé`);
  } catch (error) {
    console.error("Erreur création profil:", error);
    note("Erreur lors de la création du profil");
  }
}

// Sauvegarder le profil actuel
async function saveCurrentProfile() {
  const profile = {
    name: profilesData.profiles[currentProfile].name,
    createInstruction: $("createInstruction").value,
    editInstruction: $("editInstruction").value,
    docFiles: profilesData.profiles[currentProfile].docFiles || []
  };
  
  try {
    await fetch('/api/profiles', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ profileId: currentProfile, profile })
    });
    
    // Mettre à jour les données locales
    profilesData.profiles[currentProfile] = profile;
    
    note("Profil sauvegardé avec succès");
  } catch (error) {
    console.error("Erreur sauvegarde profil:", error);
    note("Erreur lors de la sauvegarde");
  }
}

// Dupliquer le profil actuel
async function duplicateCurrentProfile() {
  const currentProfileData = profilesData.profiles[currentProfile];
  const newName = prompt(`Nom pour la copie de "${currentProfileData.name}" :`);
  if (!newName) return;
  
  const newProfileId = newName.toLowerCase().replace(/[^a-z0-9]/g, '-');
  
  if (profilesData.profiles[newProfileId]) {
    alert("Un profil avec ce nom existe déjà");
    return;
  }
  
  const duplicatedProfile = {
    name: newName,
    createInstruction: currentProfileData.createInstruction,
    editInstruction: currentProfileData.editInstruction,
    docFiles: [] // Ne pas copier les fichiers
  };
  
  try {
    await fetch('/api/profiles', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ profileId: newProfileId, profile: duplicatedProfile })
    });
    
    await loadProfiles();
    note(`Profil dupliqué: "${newName}"`);
  } catch (error) {
    console.error("Erreur duplication profil:", error);
    note("Erreur lors de la duplication");
  }
}

// Supprimer le profil actuel
async function deleteCurrentProfile() {
  if (currentProfile === 'default') {
    alert("Impossible de supprimer le profil par défaut");
    return;
  }
  
  const currentProfileData = profilesData.profiles[currentProfile];
  if (!confirm(`Voulez-vous vraiment supprimer le profil "${currentProfileData.name}" ?`)) {
    return;
  }
  
  try {
    await fetch(`/api/profiles/${currentProfile}`, { method: 'DELETE' });
    
    await loadProfiles();
    note("Profil supprimé");
  } catch (error) {
    console.error("Erreur suppression profil:", error);
    note("Erreur lors de la suppression");
  }
}

// Upload de documents
async function uploadDocuments() {
  const files = $("docUpload").files;
  if (!files || files.length === 0) return;
  
  const formData = new FormData();
  for (const file of files) {
    formData.append('docs', file);
  }
  
  try {
    const response = await fetch(`/api/profiles/${currentProfile}/upload`, {
      method: 'POST',
      body: formData
    });
    
    const result = await response.json();
    if (response.ok) {
      await loadProfiles();
      $("docUpload").value = ""; // Reset l'input
      note(`${result.files.length} fichier(s) ajouté(s)`);
    } else {
      throw new Error(result.error);
    }
  } catch (error) {
    console.error("Erreur upload:", error);
    note("Erreur lors de l'upload: " + error.message);
  }
}

// Supprimer un document
async function deleteDocument(filename) {
  if (!confirm(`Supprimer le fichier "${filename}" ?`)) return;
  
  try {
    await fetch(`/api/profiles/${currentProfile}/docs/${filename}`, {
      method: 'DELETE'
    });
    
    await loadProfiles();
    note("Fichier supprimé");
  } catch (error) {
    console.error("Erreur suppression fichier:", error);
    note("Erreur lors de la suppression du fichier");
  }
}