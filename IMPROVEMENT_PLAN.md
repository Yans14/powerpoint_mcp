# Plan d'amélioration — PowerPoint MCP

Audit du 2026-11-09 sur le commit `9e635e7`. Environnement d'audit : macOS, Node 20.19, Python 3.11, moteur OOXML.
Le moteur COM (Windows) n'a pas pu être exécuté : les points COM viennent de la lecture du code.

## 1. État des lieux

| Vérification | Résultat |
|---|---|
| `npm run build` | OK |
| `vitest` | 13 tests OK (catalogue + unités seulement) |
| `pytest` | 22 tests OK (checkers, executor, unités — aucun test moteur) |
| `ruff`, `black --check` | OK |
| `tools/list` via stdio réel | 68 outils, **0 avec paramètres** (`inputSchema: {}`) |
| Lancement hors du dossier du repo | **Échec au démarrage** |
| Fermeture de stdin | **Le serveur ne s'arrête pas** |

Taille : ~12 000 lignes, dont `ooxml_engine.py` (2 814) et `com_engine.py` (2 524).

## 2. Phase 0 — Correctifs bloquants (reproduits)

> **Statut (2026-11-11) : corrigée, non commitée.** `npm run check` passe (22 tests TS, 46 tests Python).
> Reste à valider hors de cette machine : le correctif 7 côté COM (écrit sans Windows), et l'ouverture
> des fichiers produits dans PowerPoint réel (vérifiés ici par relecture python-pptx uniquement).
> Limite connue : SmartArt dupliqué partage ses parts de données avec l'original au lieu de les copier.

Ces défauts rendent le serveur inutilisable ou corrompent des fichiers. À traiter avant tout le reste.

1. **Schémas d'outils vides.** `zod-to-json-schema` v3 ne sait pas lire zod v4 : les 68 outils sortent avec `inputSchema: {}` ([src/server.ts:31](src/server.ts:31)). Le client LLM ne voit aucun paramètre.
   Correctif : remplacer par `z.toJSONSchema()` natif de zod 4, ou passer à `McpServer.registerTool`. Supprimer la dépendance `zod-to-json-schema`.
   Test : un test qui vérifie que chaque outil expose `type: "object"` et ses `properties`.
2. **Chemin du bridge relatif au `cwd`.** [src/bridge/client.ts:29](src/bridge/client.ts:29) résout `python/bridge.py` depuis `process.cwd()`. Lancé par Claude Desktop ou tout autre client, le serveur meurt au démarrage.
   Correctif : résoudre depuis `import.meta.url`.
3. **Processus zombie.** [src/index.ts](src/index.ts) n'écoute que SIGINT/SIGTERM. À la fermeture de stdin, Node et le bridge Python restent en vie.
   Correctif : fermer proprement sur `transport.onclose` et `stdin` `end`.
4. **`pptx_duplicate_slide` corrompt les images.** Les formes sont copiées par `deepcopy` sans leurs relations ([python/engines/ooxml_engine.py:315](python/engines/ooxml_engine.py:315)) : l'image de la copie pointe vers un `rId` inexistant (`KeyError: no relationship with key 'rId2'` à la relecture). Même risque pour graphiques, médias, liens.
   Correctif : copier les parts liées et réécrire les `rId`. Vérifier `pptx_clone_shape` et `pptx_copy_shape_between_decks` dans le même passage.
5. **Snapshot envoyé deux fois.** [src/server.ts:92](src/server.ts:92) sérialise le résultat complet (base64 inclus) en texte, puis rajoute l'image. Coût en tokens doublé, voire pire.
   Correctif : retirer `snapshot_base64` du bloc texte.
6. **Lecture qui marque le deck modifié.** `pptx_get_slide_snapshot` appelle `_persist`, qui passe `dirty` à `True` (reproduit : `False` → `True`).
7. **`width`/`height` ignorés avec `template_path`** dans `pptx_create_presentation` (reproduit : 9 144 000 EMU au lieu de 4 572 000).
8. **Aucune validation côté Python.** Un paramètre manquant ou mal typé lève `KeyError`/`ValueError`, renvoyé en `internal_error` avec traceback complet. Touche l'UI web et l'orchestrateur, qui contournent zod.

Critère de sortie : un test d'intégration lance le serveur depuis un autre dossier, liste les outils, crée un deck avec image, duplique la slide, sauvegarde, rouvre sans erreur.

## 3. Phase 1 — Robustesse du bridge

> **Statut (2026-11-11) : faite, non commitée.** `npm run check` passe (34 tests TS, 74 tests Python).
> - Timeouts : 120 s par appel (`PPTX_BRIDGE_TIMEOUT_MS`), 600 s pour les outils `pptx_agent_*`, 90 s pour le rendu (`PPTX_RENDER_TIMEOUT_SEC`). Appel bloqué = bridge arrêté (SIGTERM puis SIGKILL), erreur `timeout`, relance au prochain appel.
> - Sessions perdues : erreur `session_lost` côté client TS pour tout identifiant connu avant le crash. Pas de registre persistant ni de réouverture automatique : les modifications non sauvegardées sont perdues (conséquence assumée de l'écriture différée ci-dessous).
> - Fichiers : un dossier par processus (`<tmp>/pptx-mcp/bridge-<pid>-…`, déplaçable via `PPTX_WORKSPACE_DIR`), un sous-dossier par session (copie de travail, snapshot de rollback, PDF de rendu). Supprimé à la fermeture ; les dossiers de processus morts sont balayés au démarrage.
> - Écriture différée : la copie de travail n'est réécrite que pour `save`, `get_theme`, snapshot et rollback. Mesure sur un deck de 20 Mo : 30 ajouts de zone de texte, 15,1 s avec écriture à chaque mutation, 46 ms sans.
> - Snapshots : PDF mis en cache par révision de session ; `slide_indices` (20 max) rend plusieurs slides en un appel.
> - Logs : `PPTX_LOG_LEVEL`, une ligne par requête (`id`, `method`, `status`, `duration_ms`). Tracebacks sur stderr uniquement.
> - Ajouts hors liste : stdin du bridge décodé en UTF-8 explicite (sinon cp1252 sous Windows), `stdin=DEVNULL` pour `soffice`/`pdftoppm`, 10 dernières lignes stderr jointes à l'erreur quand le bridge meurt.
>
> Non vérifié : chemin Windows (`_windows_process_alive`, arrêt par signal, `slide_indices` côté COM), rendu réel (LibreOffice absent, renderers simulés dans les tests). Moteur COM laissé sur ses fichiers temporaires d'origine.

- **Timeouts.** Aucun sur `bridge.call`, ni sur `soffice`/`pdftoppm` (`subprocess.run` sans `timeout`). Un appel bloqué gèle tout le serveur. Ajouter un timeout par appel et tuer/relancer le bridge.
- **Redémarrage du bridge.** Après un crash, `call()` relance le processus en silence : toutes les sessions sont perdues sans que le client le sache. Renvoyer une erreur explicite `session_lost`, et envisager un registre de sessions persistant (chemin de travail + chemin d'origine) pour réouverture.
- **Erreurs de lancement.** `spawn` n'a pas de handler `error` : un `python3` introuvable produit une exception non gérée. Ajouter le handler et un message d'installation clair.
- **Tracebacks.** Les journaliser sur stderr, ne plus les renvoyer au client ([python/bridge.py:61](python/bridge.py:61)).
- **Protection de stdout.** Rediriger `sys.stdout` vers stderr dans le bridge et garder un descripteur dédié au protocole : un `print` d'une dépendance casse aujourd'hui une réponse.
- **Fichiers temporaires.** Les `.snapshot` de l'orchestrateur ne sont supprimés que sur `cancel`. Regrouper tout dans un dossier de session nettoyé à la fermeture et au démarrage.
- **Performance.** `_persist` réécrit le fichier entier après chaque mutation (28 appels). Sauvegarder seulement quand un consommateur a besoin du fichier (save, snapshot, rollback).
- **Snapshots.** Chaque capture convertit le deck entier en PDF. Mettre le PDF en cache par révision du deck, et permettre plusieurs slides en un appel.
- **Journalisation.** Logs structurés sur stderr avec identifiant de requête, durée, méthode. Niveau réglable par variable d'environnement.

## 4. Phase 2 — Conformité MCP et ergonomie pour le LLM

> **Statut (2026-11-11) : faite en partie, non commitée.** `npm run check` passe (68 tests TS, 102 tests Python).
> - Annotations déclarées par outil ([src/tools/traits.ts](src/tools/traits.ts)) ; version lue dans `package.json`.
> - Réponses : JSON compact, `slide_count` par défaut, `include_state` pour l'état complet. Ajout de slide sur deck de 100 slides : 12 836 → 155 caractères.
> - `pptx_batch` : 50 opérations max sur une présentation, validation de toutes les opérations avant exécution, rollback si échec (OOXML). Pas de références entre opérations (résultat de l'une réutilisé par la suivante) : non fait.
> - Profils : `PPTX_TOOL_PROFILE=core` expose 31 outils. Regroupement des variantes `set_*` : non fait.
> - Sorties structurées : `outputSchema` + `structuredContent` sur 5 outils de lecture seulement, ceux dont la forme est identique dans les deux moteurs. Les 20 autres restent en texte.
> - Ressources : templates, un seul appel bridge pour la liste, snapshot en blob JPEG, erreur protocole -32002 sur URI inconnue, `resources/list_changed` à l'ouverture et à la fermeture.
> - Prompts : `create_deck_from_outline`, `apply_template_to_deck`.
> - Erreurs : champ `hint` (outil à appeler ensuite) ; layout introuvable passe de `validation_error` à `not_found`.
>
> Coût : `tools/list` complet passe d'environ 65 000 à 81 000 caractères (annotations, `include_state`, schémas de sortie, `pptx_batch`). Profil `core` : 41 000.
> Non vérifié : moteur COM (code `not_found` du layout, `hint` de session, batch sans rollback), affichage réel des annotations et prompts dans un client MCP.
> Constat : la validation Python accepte les paramètres inconnus (catalogue Python incomplet), donc un nom de paramètre erroné est ignoré en silence hors couche TS. À traiter avec la source unique de phase 3.

- **Annotations.** `destructiveHint` est déduit du nom de l'outil ([src/server.ts:44](src/server.ts:44)). Le déclarer explicitement par outil, avec `idempotentHint`, `openWorldHint` et `title`. `pptx_save_presentation` écrase un fichier existant sans le signaler.
- **Sorties structurées.** Ajouter `outputSchema` + `structuredContent` aux outils de lecture.
- **Taille des réponses.** Chaque mutation renvoie le `presentation_state` complet. Sur un deck de 100 slides, c'est du bruit. Le rendre optionnel (`include_state`, défaut `false`) et renvoyer un résumé.
- **Surface d'outils.** 68 outils chargés d'office. Regrouper les variantes proches (ex. les `set_*` de forme et de texte) ou proposer des profils (`core`, `full`) par variable d'environnement.
- **Opérations par lot.** Ajouter `pptx_batch` : une liste d'opérations, une seule sauvegarde, rollback si échec. Réduit fortement les allers-retours.
- **Ressources.** Utiliser des `resourceTemplates` au lieu d'énumérer chaque slide (N+1 appels dans `listResources`). Servir le snapshot en `blob` image, pas en JSON base64. Lever une vraie erreur sur URI inconnue au lieu de renvoyer l'erreur comme contenu ([src/server.ts:136](src/server.ts:136)). Émettre `resources/list_changed`.
- **Prompts MCP.** Exposer des prompts guidés (« créer un deck depuis un plan », « mettre un deck à la charte »).
- **Messages d'erreur.** Uniformiser codes et suggestions (« layout introuvable → appeler `pptx_get_layouts` », avec la liste des noms valides).
- **Version.** `0.1.0` est codée en dur dans `server.ts` : la lire depuis `package.json`.

## 5. Phase 3 — Architecture et dette

> **Statut (2026-11-13) : faite, non vérifiée sous Windows.**
> - Source unique : `src/tools/catalog.ts` (avec une `category` par outil) est exporté par `npm run gen:catalog` vers `python/generated/tool_catalog.json`. Python en dérive validation (jsonschema, mêmes schémas que le client MCP), routage `pptx_<nom>` → `engine.<nom>` (dossier `handlers/` supprimé), formulaires de l'UI web et manifeste de l'agent. CI et `npm run check` échouent si le JSON est périmé.
> - Conséquence : un paramètre inconnu est maintenant refusé côté Python (UI web, plans de l'agent) au lieu d'être ignoré ; les valeurs `null` sont retirées avant l'appel moteur.
> - Moteurs découpés en 9 modules chacun (`engines/ooxml/`, `engines/com/`) ; `ooxml_engine.py` et `com_engine.py` restent les points d'entrée. Découpage mécanique : source de chacune des 105 + 151 méthodes vérifiée identique avant/après.
> - Couplage : `BaseEngine.get_pptx()` remplace `extra["prs"]` dans checkers et agent ; en COM, export par `SaveCopyAs`. Rollback de l'agent et `pptx_batch` passent par `copy_working_file` / `restore_working_file`.
> - Checkpoints exposés en outils : `pptx_create_checkpoint`, `pptx_restore_checkpoint` (OOXML, 10 par présentation).
> - Agent : désactivé par défaut, activé par `PPTX_ENABLE_AGENT=1` (avant : actif dès qu'une `ANTHROPIC_API_KEY` traînait dans l'environnement) ; outils `pptx_agent_*` masqués sinon ; modèle par défaut `claude-sonnet-5-5` ; `PPTX_AGENT_TIMEOUT` appliqué entre les étapes ; SDK dans `requirements-agent.txt`.
> - Dépendances : `requirements.txt` (bridge, + `jsonschema`), `requirements-ui.txt` (FastAPI, uvicorn, websockets), `requirements-agent.txt`. Versions alignées : Node ≥ 22, Python ≥ 3.11.
>
> Non vérifié : tout le code COM (découpage, `get_pptx` par `SaveCopyAs`) faute de Windows ; agent jamais exécuté avec un vrai LLM ; modèle OpenAI par défaut (`gpt-4o`) laissé tel quel.

- **Source unique pour le catalogue.** Un outil est décrit à cinq endroits : schéma zod (`catalog.ts`, 1 315 lignes), `tool_catalog.py` (928), `capability_manifest.py`, `handlers/*.py`, `BaseEngine`. Générer le catalogue Python et le manifeste depuis le JSON Schema produit par TypeScript, et valider les paramètres côté Python avec ce même schéma (règle le point 0.8).
- **Découper les moteurs.** Scinder les deux fichiers de 2 500+ lignes par domaine (session, slides, formes, tableaux, graphiques, texte), avec les helpers OOXML à part.
- **Couplage OOXML des checkers et de l'agent.** `CheckerService`, `SnapshotManager` et `AgentOrchestrator` lisent `engine.sessions[id].extra["prs"]`. `com_engine.py` ne définit jamais `"prs"` : ces outils échoueront en mode COM. Ajouter à `BaseEngine` une méthode qui fournit un `Presentation` python-pptx à jour (export du fichier de travail en COM).
- **Orchestrateur LLM embarqué.** Un client MCP est déjà un LLM : planifier avec un second LLM double le coût et la latence. Constats : `anthropic`/`openai` absents de `requirements.txt`, modèle par défaut périmé (`claude-sonnet-4-20250514`), `execution_timeout_sec` jamais utilisé, étapes du plan exécutées sans validation de schéma.
  Recommandation : le sortir en extra optionnel (`pip install .[agent]`), désactivé par défaut. Garder le rollback et l'exposer comme outil générique (`pptx_checkpoint` / `pptx_restore`).
- **Dépendances.** `fastapi`, `uvicorn`, `websockets` sont dans les dépendances d'exécution alors qu'ils ne servent qu'à l'UI de test : les passer en extra `ui`.
- **Versions déclarées.** README annonce Python ≥ 3.12 et Node ≥ 22 ; `pyproject.toml` cible py311 et tout passe sous Node 20. Aligner.

## 6. Phase 4 — Tests et CI

- **Tests moteur OOXML.** Aucun aujourd'hui. Couvrir chaque outil : créer → muter → sauvegarder → rouvrir → vérifier. Inclure des decks réels (images, graphiques, tableaux fusionnés, groupes imbriqués).
- **Validité des fichiers produits.** Valider chaque sortie contre le schéma OOXML, et l'ouvrir avec LibreOffice en CI.
- **Tests d'intégration MCP.** Client SDK ↔ serveur par stdio : `initialize`, `tools/list`, appels, ressources, arrêt propre.
- **Tests de contrat.** Vérifier que catalogue TS, handlers Python et les deux moteurs exposent exactement les mêmes méthodes.
- **CI Windows.** Ajouter `windows-latest` pour le chemin OOXML et les imports. Le smoke COM reste manuel (licence PowerPoint) ou sur runner auto-hébergé.
- **Qualité.** Ajouter ESLint + Prettier côté TS, `pyright` ou `mypy` côté Python, seuil de couverture, `npm audit` et `pip-audit`.
- **Matrice.** Tester Node 20/22 et Python 3.11/3.12/3.13 selon ce qui est déclaré.

## 7. Phase 5 — Sécurité

- **Périmètre fichiers.** N'importe quel chemin absolu est lisible et inscriptible. Ajouter une liste de dossiers autorisés (variable d'environnement ou `roots` MCP), résoudre les liens symboliques, refuser l'écrasement sans `overwrite: true`.
- **UI web.** `/api/dispatch` exécute toute méthode sans authentification. Ajouter un jeton aléatoire au démarrage, vérifier `Host`/`Origin` (DNS rebinding), refuser tout bind hors `127.0.0.1`.
- **Entrées non fiables.** Limiter taille de fichier, nombre de slides et taux de décompression (zip bomb). Vérifier que tout XML externe passe par `defusedxml`.
- **`pptx_get_slide_xml`.** Tronquer ou paginer : un XML de slide peut dépasser la fenêtre de contexte.
- **Sous-processus.** Lancer `soffice` avec un profil utilisateur isolé et un timeout.
- **Dépendances.** Figer les versions Python (lockfile) : `requirements.txt` n'a que des bornes basses.

## 8. Phase 6 — Fonctionnalités

Par ordre de valeur pour un usage par agent :

1. Export PDF et export image de tout le deck (planche de vignettes pour relecture visuelle).
2. Checkpoint / restore générique (voir phase 3).
3. Import de slides entre decks en conservant la mise en forme source.
4. Édition des masques et layouts, application d'un thème.
5. Liens hypertexte, texte alternatif, vérification d'accessibilité.
6. Graphiques : plus de types, axes, étiquettes, couleurs de séries.
7. Commentaires et propriétés du document.
8. SmartArt et animations : lecture seule en OOXML, édition en COM uniquement.
9. Transport HTTP streamable en plus de stdio.

## 9. Phase 7 — Packaging et documentation

- **Distribution.** `package.json` est `private`, sans `bin` ni `files`. Publier un paquet lançable par `npx`, avec le code Python embarqué et une création automatique de venv (ou `uv`). Proposer aussi un bundle `.mcpb` pour Claude Desktop.
- **README.** Annonce 24 outils et « Phase 1-2 » alors que 68 sont enregistrés. Générer la liste depuis le catalogue. Ajouter des exemples de configuration pour Claude Desktop, Claude Code et Cursor, et un tableau de parité COM / OOXML.
- **`.env.example`.** Documenter toutes les variables (`PPTX_LLM_*`, `PPTX_AGENT_*`, Azure).
- **Versionnage.** CHANGELOG, tags, workflow de release.
- **Notebooks.** Six notebooks de validation manuelle : en convertir l'essentiel en tests automatisés.

## 10. Ordre d'exécution proposé

| Étape | Contenu | Effort indicatif |
|---|---|---|
| 1 | Phase 0 complète + test d'intégration MCP | 2 à 3 jours |
| 2 | Phase 1 (timeouts, redémarrage, logs, temporaires) | 3 jours |
| 3 | Source unique du catalogue + validation Python | 3 à 4 jours |
| 4 | Tests moteur OOXML + CI Windows + linters | 4 à 5 jours |
| 5 | Phase 2 (annotations, sorties, batch, ressources) | 4 jours |
| 6 | Sécurité (périmètre fichiers, UI web, limites) | 2 à 3 jours |
| 7 | Découpage des moteurs, découplage checkers/agent | 4 à 5 jours |
| 8 | Packaging, README, release | 2 jours |
| 9 | Nouvelles fonctionnalités | au fil de l'eau |

Les étapes 3 et 4 passent avant le découpage des moteurs : sans tests moteur, ce refactoring est risqué.

## 11. Non vérifié

- Moteur COM : aucune exécution (pas de Windows ni de PowerPoint).
- Snapshot OOXML : LibreOffice absent de la machine d'audit (wrapper `soffice` cassé), rendu non testé.
- Outils tableaux, graphiques, groupes, dégradés : lus en diagonale, non exécutés un par un.
- Orchestrateur LLM : non exécuté (pas de clé API).
- Snapshot et slides masquées : l'export PDF de LibreOffice omet par défaut les slides masquées, ce qui décalerait la numérotation des pages (slide N rendue ≠ slide N du deck). À confirmer avec LibreOffice installé ; correctif probable : option de filtre `ExportHiddenSlides`.
- Snapshot et instance LibreOffice déjà ouverte : `soffice --headless` peut échouer ou déléguer à l'instance existante. À confirmer ; correctif probable : profil dédié via `-env:UserInstallation`.
