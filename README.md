# Stock Gap · Magento

Dashboard interattiva per trovare, negozio per negozio, le SKU con giacenza positiva che non risultano nell’inventario centrale esportato da **Magento**.

## Uso

1. Apri la dashboard online.
2. Carica l’Excel delle giacenze dei negozi.
3. Carica il CSV Magento che elenca le SKU presenti nel magazzino centrale.
4. Carica il file anagrafica con tutte le combinazioni SKU–taglia–barcode disponibili.
5. Scegli la vista: assenti nel centrale, presenti per verificare il matching, match ambigui oppure tutte le SKU.
6. Nella vista di verifica confronta direttamente `SKU Excel` e `SKU CSV`; il Brand mostrato proviene esclusivamente dalla riga CSV abbinata.
7. Filtra per negozio, Brand, quantità minima, SKU o descrizione.
8. Ordina i risultati per giacenza o SKU.
9. Premi **Scarica Excel** per esportare tutte le righe che rispettano i filtri attivi, non soltanto quelle già caricate a schermo. Per ciascuna SKU selezionata, taglie, barcode e descrizione provengono dall’anagrafica; la giacenza proviene dall’Excel negozi e vale `0` quando la combinazione non è presente. Dopo tutti i figli viene aggiunta una riga padre con il solo SKU. Le colonne sono `SKU esploso con i figli`, `taglia`, `barcode`, `giacenza`, `negozio` e `descrizione`. Le taglie testuali seguono l’ordine `XXS`, `XS`, `S`, `M`, `L`, `XL`, `XXL`, `3XL`, `4XL`; quelle numeriche restano in ordine crescente e le mezze taglie indicate con un trattino vengono convertite in decimali (`6-` → `6.5`). Se una SKU filtrata non esiste nell’anagrafica, l’export viene interrotto, la dashboard mostra l’errore e scarica un report Excel con l’elenco completo delle SKU mancanti.

I tre file vengono elaborati soltanto nel browser: SKU e giacenze non vengono caricati su GitHub o su altri server. Per aggiornare i dati basta caricare nuovamente i file.

## Formato del file

La dashboard legge il primo foglio e riconosce automaticamente colonne equivalenti a:

- `Descr. Negozio`
- `Articolo` (SKU)
- `Descrizione Articolo`
- `Descrizione Taglia`
- `BCR` (barcode)
- `Qtà in stock`

Le righe `PUSH` eventualmente presenti nell’Excel vengono ignorate. Il CSV deve contenere almeno la colonna `SKU`; la colonna `Brand` alimenta il filtro marchio.

Il file anagrafica può essere CSV o Excel. La dashboard individua anche un’intestazione preceduta da righe titolo e richiede le colonne `CODICE`, `DESCRIPTION`, `TG` e `BARCODE`. I prefissi `=` usati dal CSV e gli asterischi finali sulle taglie vengono rimossi durante l’elaborazione.

Il collegamento tra Excel negozi e anagrafica usa lo SKU esatto senza distinzione tra maiuscole e minuscole. Punti e spazi interni vengono preservati, così codici distinti come `NBLNBML574OB.D` e `NBLNBML574OBD` non vengono fusi.

Per il confronto, spazi e punti vengono rimossi da entrambi i codici e `OT-` viene rimosso dagli SKU del CSV. La dashboard riconosce match identici, SKU Magento precedute dal comune prefisso di tre caratteri nell’Excel e, come fallback, SKU Magento contenute nel codice Excel. I match multipli della stessa lunghezza sono considerati ambigui e hanno una vista dedicata.

“0 identici senza prefisso” non significa che non sia stato trovato alcun match: significa solo che nessun codice Excel coincide carattere per carattere con quello CSV dopo la normalizzazione. Il totale riconosciuto comprende anche i match ottenuti togliendo il prefisso di tre lettere e quelli contenuti come varianti.

## Avvio locale

Il file delle giacenze è escluso intenzionalmente dal repository. La pagina deve essere servita via HTTP. Per esempio:

```powershell
python -m http.server 8000
```

Poi apri `http://localhost:8000`.
