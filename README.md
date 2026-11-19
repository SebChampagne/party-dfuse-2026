# Party annuel DFuse 2026 — landing page RSVP

Page statique (GitHub Pages) + Google Sheet / Apps Script pour enregistrer les réponses et envoyer les confirmations.

```
index.html            la page (tout le texte, le style et le JS)
assets/               logos DFuse, Tartaro, Fiola Import
apps-script/Code.gs   script à coller dans la Google Sheet
```

## 1. Google Sheet + Apps Script (≈ 5 min)

1. Crée une Google Sheet vide, idéalement avec le compte Google de DFuse (les confirmations partiront de ce compte), nomme-la « RSVP Party DFuse 2026 ».
2. Menu **Extensions → Apps Script**. Efface le contenu et colle tout `apps-script/Code.gs`.
3. En haut du fichier, ajuste `CONFIG` :
   - `NOTIFY_EMAIL` : l'adresse qui reçoit un avis à chaque réponse (laisser vide pour aucun avis).
   - `REPLY_TO` : l'adresse où les invités répondent s'ils répondent au courriel de confirmation.
4. Enregistre, choisis la fonction **setup** dans la liste et clique **Exécuter**. Accepte les autorisations (Google affiche « application non validée » : *Paramètres avancés → Accéder au projet*). Les onglets **RSVP** et **Résumé** sont créés.
5. **Déployer → Nouveau déploiement** → type **Application Web** :
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**
6. Copie l'URL qui finit par `/exec` et colle-la dans `index.html`, ligne `const RSVP_ENDPOINT = "";`.

> Si tu modifies le script plus tard : **Déployer → Gérer les déploiements → ✏️ → Nouvelle version** (l'URL reste la même).

Quota : un compte Gmail gratuit peut envoyer environ 100 courriels/jour via le script (1 500 avec Google Workspace).

## 2. Mise en ligne (GitHub Pages)

Le dépôt est publié sur GitHub Pages depuis la branche `main` (dossier racine). Chaque `git push` met la page à jour en environ 1 minute.

Adresse personnalisée (optionnel) : dans le DNS de dfuse.ca, ajoute un CNAME `party` → `sebchampagne.github.io`, puis inscris `party.dfuse.ca` dans *Settings → Pages → Custom domain*.

## 3. Le jour J

L'onglet **Résumé** donne en direct : nombre de personnes attendues (invités + accompagnateurs), absents, préférences de boissons et nombre de réponses avec restrictions alimentaires — à transmettre à Tartaro et Fiola Import.
