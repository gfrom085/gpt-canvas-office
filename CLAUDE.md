# GPT Canvas Lite - Instructions de Développement

## 📋 Vue d'ensemble
GPT Canvas Lite est un add-in Office pour Word qui permet de générer et éditer du texte avec GPT-5. Le projet utilise un serveur Node.js HTTPS sur le port 3001 pour servir l'interface et les APIs.

## 🚀 Démarrage Rapide

### 1. Démarrer le serveur API
```cmd
cd D:\projet_gab\gpt-canvas-office
npm run serve
```
**Attendez de voir :** `HTTPS add-in server on https://localhost:3001`

### 2. Ouvrir Word avec l'add-in (Mode Debug)
```cmd
cd D:\projet_gab\gpt-canvas-office-v2
npx office-addin-debugging start manifest.xml --app word
```

## 🧹 Nettoyage des Caches (si problème)
```cmd
# Fermer tous les processus Office
powershell.exe "Get-Process winword,excel,powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force"

# Vider les caches Office
powershell.exe "Remove-Item '$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*' -Recurse -Force -ErrorAction SilentlyContinue"
powershell.exe "Remove-Item '$env:LOCALAPPDATA\Microsoft\EdgeWebView\*' -Recurse -Force -ErrorAction SilentlyContinue"

# Nettoyer les settings d'add-ins
cd D:\projet_gab\gpt-canvas-office-v2
npx office-addin-dev-settings clear manifest.xml
```

## 🛠️ Commandes Utiles

### Gestion des Add-ins
```cmd
# Arrêter le debugging
npx office-addin-debugging stop manifest.xml

# Valider le manifest
npx office-addin-manifest validate manifest.xml

# Mode sideload simple
npx office-addin-dev-settings sideload manifest.xml --app word

# Activer les DevTools
npx office-addin-dev-settings debugging manifest.xml --enable --open-dev-tools

# Activer les logs runtime
npx office-addin-dev-settings runtime-log --enable
```

### Gestion des Certificats
```cmd
# Installer/réinstaller les certificats HTTPS
npm run certs
```

## 📁 Structure du Projet

```
gpt-canvas-office/
├── server.js                    # Serveur Node.js HTTPS (port 3001)
├── public/
│   ├── taskpane-v2.html        # Interface principale
│   ├── taskpane-v2.js          # Logique JavaScript
│   ├── taskpane-v2.css         # Styles Office Fabric UI
│   ├── taskpane.html           # Interface originale (backup)
│   └── taskpane.js             # Logique originale (backup)
├── .env                        # Clé API OpenAI
└── package.json

gpt-canvas-office-v2/
├── manifest.xml                # Manifest Office (pointe vers port 3001)
└── package.json               # Outils de développement Office
```

## ⚙️ Configuration

### Variables d'Environnement
Créer un fichier `.env` dans la racine :
```
OPENAI_API_KEY=sk-proj-...
```

### Paramètres GPT Disponibles
- **Modèles** : `gpt-5`, `gpt-5-mini`
- **Effort de raisonnement** : `minimal`, `low`, `medium`, `high`
- **Verbosité** : `low`, `medium`, `high`

## 🔍 Debugging

### Logs Serveur
Le serveur affiche des logs détaillés :
```
=== GÉNÉRATION ===
Config GPT: { model: "gpt-5-mini", reasoningEffort: "high" }
Prompt utilisateur: [prompt]
Paramètres envoyés à GPT: [JSON complet]
Réponse GPT reçue: [réponse]
=== FIN GÉNÉRATION ===
```

### Console Browser
Dans Word, clic droit sur l'add-in → "Inspecter l'élément" → Console

### Logs Runtime Office
```cmd
npx office-addin-dev-settings runtime-log --enable
# Fichier : C:\Users\[user]\AppData\Local\Temp\OfficeAddins.log.txt
```

## 🎯 Fonctionnalités

### Interface Utilisateur
- **Génération** : Auto-insertion à la position du curseur
- **Édition** : Prompt personnalisé + droplist prédéfinie
- **Auto-remplacement** : La sélection éditée remplace l'original
- **Menu paramètres** : ⚙️ Choix du modèle et paramètres GPT

### API Endpoints
- `POST /api/generate` - Génération de texte
- `POST /api/rewrite` - Réécriture/édition de texte
- `GET /healthz` - Health check

## 🚨 Résolution de Problèmes

### Add-in vide ou ne charge pas
1. Vérifier que le serveur sur 3001 fonctionne
2. Vider les caches (voir section nettoyage)
3. Vérifier les certificats HTTPS : `npm run certs`

### Erreurs API
1. Vérifier la clé API dans `.env`
2. Consulter les logs serveur détaillés
3. Tester les endpoints : `curl -k https://localhost:3001/healthz`

### Word ne détecte pas l'add-in
1. S'assurer que le manifest pointe vers `https://localhost:3001/taskpane-v2.html`
2. Nettoyer les settings : `npx office-addin-dev-settings clear manifest.xml`
3. Redémarrer en mode debug

## 📝 Notes de Développement
- Le serveur sur 3001 sert à la fois l'API et l'interface HTML
- Les paramètres GPT sont configurables via l'interface ⚙️
- L'add-in supporte Word et PowerPoint (manifest configuré pour les deux)
- Utilise Office.js et Office Fabric UI pour l'interface moderne