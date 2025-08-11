README (installation & usage)
1) Prérequis
Node.js LTS + npm

Word et/ou PowerPoint (Bureau ou Web)

Connexion Internet

Une clé OpenAI (nouvelle, pas celle que tu as collée ici)

2) Installation
bash
Copier
Modifier
git clone <ce-projet> gpt-canvas-office
cd gpt-canvas-office
npm install
cp .env.example .env
# Ouvre .env et colle ta clé dans OPENAI_API_KEY
3) Lancer en local (sideload auto)
Word (Desktop)
bash
Copier
Modifier
npm run start:word
Le script :

installe un certificat HTTPS localhost,

lance le serveur HTTPS sur https://localhost:3000,

sideload le complément dans Word (ouvre Word si besoin). 
Npm
+1
Microsoft Learn

PowerPoint (Desktop)
bash
Copier
Modifier
npm run start:ppt
Stopper proprement
bash
Copier
Modifier
npm run stop
Alternative : Word/PowerPoint Web → “Upload My Add-in” et sélectionne manifest.xml (pratique si tu veux éviter la config locale). 
Microsoft Learn

4) Utilisation
Dans Word/PowerPoint, ouvre Mes compléments → GPT Canvas Lite (si le panneau n’est pas déjà là).

Créer : saisis un prompt → Générer → le texte apparaît dans “Proposition”; tu peux le coller ou l’insérer en commentaire (Word).

Éditer : sélectionne du texte dans le doc/slide → Prendre la sélection → Proposer → ajuste si tu veux → Appliquer.

Historique : stocké dans GPT_History (Propriétés personnalisées du fichier).

Word : Fichier → Infos → Propriétés → Propriétés avancées → Personnaliser.

PowerPoint : idem (selon version). Les APIs confirment la présence. 
Microsoft Learn
+1

5) Notes & limites
SourceLocation doit être HTTPS (localhost avec cert dev), sinon Office refuse de charger. 
Microsoft Learn

PowerPoint : assure-toi d’avoir PowerPointApi ≥ 1.7 pour l’historique via custom properties. Si ton client est ancien, l’édition marchera, mais l’historique PPt sera désactivé (tu verras un message dans la console). 
Microsoft Learn

Track Changes Word n’est pas exposé finement : on utilise commentaires/remplacement + historique JSON.

La clé OpenAI reste côté serveur (.env), jamais dans le client (panneau). C’est volontairement plus sûr.

6) Dépannage
Certificat invalide : relance npm run certs. Si Office boude, supprime l’ancien cert et réinstalle. 
Microsoft Learn

Manifest invalide : npm run manifest:validate pour des erreurs courantes. 
Microsoft Learn

Sideload ne démarre pas : lance d’abord npm run serve, puis npx office-addin-debugging start manifest.xml word.