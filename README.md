# CheckAndFixDate
Verifica e aggiorna la data exif confrontandola con la data presente nel json di google takeout o con la data presente nel nome del file

Questo progetto consiste in uno script **PowerShell 5.1** con interfaccia grafica (WinForms) progettato per **Confrontare e aggiornare la data EXIF**

Lo script genera nella castella selezionata un file csv con l'elenco dei mismatch trovati.
I mismatch vengono valutati confrontado la data di creazione exif con la data di creazione estratta dal file json di google takeout oppure la data presente nel nome del file
Lo script ha anche una funzionalita che tenta l'aggiornamento dei dati exif.

Lo script va eseguito con:
powershell -ExecutionPolicy Bypass -File .\CheckAndFixDate.ps1  

# CheckAndFixDate

Questo progetto permette di verificare e correggere le date dei file multimediali utilizzando strumenti esterni come **FFmpeg**.

---

## 📦 Requisiti

Per utilizzare questo progetto è necessario avere installato **FFmpeg**, in particolare:

* `ffmpeg.exe`
* `ffprobe.exe`
* `ffplay.exe` (opzionale)

Questi file **non sono inclusi nel repository** a causa delle dimensioni, ma puoi scaricarli dai repository ufficiali indicati sotto.

---

## 🔗 Download ufficiali di FFmpeg

### 🌐 Sito ufficiale FFmpeg

* [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html)

### 💻 Build consigliate per Windows

* **Gyan.dev (Static Build)**
  [https://www.gyan.dev/ffmpeg/builds/](https://www.gyan.dev/ffmpeg/builds/)

* **BtbN (GitHub Builds)**
  [https://github.com/BtbN/FFmpeg-Builds/releases](https://github.com/BtbN/FFmpeg-Builds/releases)

Queste build sono gratuite, affidabili e comunemente utilizzate dalla community.

---

## 📁 Come installare FFmpeg

1. Scarica la versione per Windows dai link sopra.
2. Estrai lo ZIP.
3. Copia `ffmpeg.exe`, `ffprobe.exe` e (opzionale) `ffplay.exe` nella cartella del progetto

Oppure aggiungi la cartella estratta al tuo **PATH** di sistema.

---

## ▶️ Utilizzo

Una volta posizionati i binari di FFmpeg nella directory corretta, puoi utilizzare gli script del progetto per analizzare e correggere le date dei file.

---
Se hai bisogno di assistenza aggiuntiva o vuoi automatizzare il download dei binari tramite script, contattami!

