# Stock Gap · PUSH

Dashboard interattiva per trovare, negozio per negozio, le SKU con giacenza positiva che non sono disponibili nel magazzino centrale **PUSH**.

## Uso

1. Apri la dashboard online.
2. Filtra per negozio, quantità minima, SKU o descrizione.
3. Ordina i risultati per giacenza o SKU.
4. Usa **Carica nuovo Excel** e scegli `smarthub.xlsx` dal computer.
5. Il file viene elaborato soltanto nel browser: SKU e giacenze non vengono caricati su GitHub o su altri server.
6. Dopo aver sostituito il file nella cartella locale, usa di nuovo **Carica nuovo Excel**. Il pulsante **Rileggi file** ricalcola invece il file già selezionato nella sessione corrente.

## Formato del file

La dashboard legge il primo foglio e riconosce automaticamente colonne equivalenti a:

- `Descr. Negozio`
- `Articolo` (SKU)
- `Descrizione Articolo`
- `Descrizione Taglia`
- `Qtà in stock`

Il magazzino centrale viene riconosciuto quando la descrizione del negozio contiene `PUSH` (anche in codici come `PUSH01`). Le quantità positive e negative sono prima sommate per SKU e negozio; solo i totali netti maggiori di zero vengono mostrati.

## Avvio locale

Il file delle giacenze è escluso intenzionalmente dal repository. La pagina deve essere servita via HTTP. Per esempio:

```powershell
python -m http.server 8000
```

Poi apri `http://localhost:8000`.
