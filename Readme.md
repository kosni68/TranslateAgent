# TranslateAgent (PowerShell + LM Studio)

Petit agent PowerShell qui tourne en arrière-plan et, via un raccourci clavier global, récupère le texte sélectionné, le **corrige**, puis le **traduit fidèlement** vers une langue cible (par défaut **allemand**), et **remplace directement la sélection**.

- **Hotkey global** : `Ctrl + Alt + T`
- **Input** : texte surligné (capturé via `Ctrl+C`)
- **Output** : texte corrigé + traduit, collé via `Ctrl+V`
- **Backend** : LM Studio (API OpenAI-compatible) sur `http://<host>:<port>/v1/...`
