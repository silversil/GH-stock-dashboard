# Stock Gap · Magento

Dashboard interattiva per trovare, negozio per negozio, le SKU con giacenza positiva che non risultano nell’inventario centrale esportato da **Magento**.

## Uso

1. Apri la dashboard online.
2. Carica l’Excel delle giacenze dei negozi.
3. Carica il CSV Magento che elenca le SKU presenti nel magazzino centrale.
4. Scegli la vista: assenti nel centrale, presenti per verificare il matching, match ambigui oppure tutte le SKU.
5. Nella vista di verifica confronta direttamente `SKU Excel` e `SKU CSV`; il Brand mostrato proviene esclusivamente dalla riga CSV abbinata.
6. Filtra per negozio, Brand, quantità minima, SKU o descrizione.
7. Ordina i risultati per giacenza o SKU.

Entrambi i file vengono elaborati soltanto nel browser: SKU e giacenze non vengono caricati su GitHub o su altri server. Per aggiornare i dati basta caricare nuovamente i file.

## Formato del file

La dashboard legge il primo foglio e riconosce automaticamente colonne equivalenti a:

- `Descr. Negozio`
- `Articolo` (SKU)
- `Descrizione Articolo`
- `Descrizione Taglia`
- `Qtà in stock`

Le righe `PUSH` eventualmente presenti nell’Excel vengono ignorate. Il CSV deve contenere almeno la colonna `SKU`; la colonna `Brand` alimenta il filtro marchio.

Per il confronto, spazi e punti vengono rimossi da entrambi i codici e `OT-` viene rimosso dagli SKU del CSV. La dashboard riconosce match identici, SKU Magento precedute dal comune prefisso di tre caratteri nell’Excel e, come fallback, SKU Magento contenute nel codice Excel. I match multipli della stessa lunghezza sono considerati ambigui e hanno una vista dedicata.

“0 identici senza prefisso” non significa che non sia stato trovato alcun match: significa solo che nessun codice Excel coincide carattere per carattere con quello CSV dopo la normalizzazione. Il totale riconosciuto comprende anche i match ottenuti togliendo il prefisso di tre lettere e quelli contenuti come varianti.

## Avvio locale

Il file delle giacenze è escluso intenzionalmente dal repository. La pagina deve essere servita via HTTP. Per esempio:

```powershell
python -m http.server 8000
```

Poi apri `http://localhost:8000`.
