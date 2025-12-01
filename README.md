# CheckAndFixDate
Verifica e aggiorna la data exif confrontandola con la data presente nel json di google takeout o con la data presente nel nome del file

Questo progetto consiste in uno script **PowerShell 5.1** con interfaccia grafica (WinForms) progettato per **Confrontare e aggiornare la data EXIF**

Lo script genera nella castella selezionata un file csv con l'elenco dei mismatch trovati.
I mismatch vengono valutati confrontado la data di creazione exif con la data di creazione estratta dal file json di google takeout oppure la data presente nel nome del file
Lo script ha anche una funzionalita che tenta l'aggiornamento dei dati exif.


