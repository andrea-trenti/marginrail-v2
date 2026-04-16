# Input richiesto

## Formato supportato
Solo file `.xlsx` realmente leggibili e compatibili con il template ufficiale MarginRail.

## Regola V2
La V2 non crea più in silenzio colonne core mancanti.  
Se manca una colonna core, l’analisi si ferma con un messaggio leggibile.  
Se manca una colonna opzionale, l’analisi può continuare ma viene mostrato un warning.

## Fogli obbligatori
- Vendite_2025
- Workflow_Deroghe_2025
- Clienti
- Prodotti
- Accordi_Commerciali
- Promo_2025

## Schema colonne

### 1) Vendite_2025

**Colonne core**
- NumeroOrdine
- RigaOrdine
- TipoDocumento
- ClienteID
- CanaleVendita
- ProdottoID
- QtaDocumento
- QtaOrdinata
- PrezzoListinoUnit
- ScontoBasePct
- ScontoContrattualePct
- ScontoPromoPct
- ScontoExtraPct
- ScontoTotalePct
- PrezzoNettoUnit
- FloorPriceUnit
- CostoAttualeUnit
- VariazioneCostoPct
- RicavoRiga
- MargineContributivo
- MarginePct
- PromoID
- AccordoID
- MotivoDeroga
- DataDocumento

**Colonne opzionali**
- Cliente
- Venditore
- Prodotto
- Categoria
- QtaResa
- DataOrdine
- PrioritaOrdine
- GruppoCliente
- Regione
- Provincia

### 2) Workflow_Deroghe_2025

**Colonne core**
- NumeroOrdine
- RigaOrdine
- SogliaScontoRuoloPct
- StatoDeroga

**Colonne opzionali**
- ApprovatoDa
- DataApprovazione
- MotivoDeroga

### 3) Clienti

**Colonne core**
- ClienteID
- RischioCredito

**Colonne opzionali**
- ScontoBase
- AreaCommerciale

### 4) Prodotti

**Colonne core**
- ProdottoID
- MargineTarget

**Colonne opzionali**
- ClasseBrand
- StatoProdotto

### 5) Accordi_Commerciali

**Colonne core**
- AccordoID
- FloorPriceUnit
- StatoAccordo
- ValidoDa
- ValidoA

**Colonne opzionali**
- PrezzoContrattualeUnit
- ScontoContrattualePct

### 6) Promo_2025

**Colonne core**
- PromoID
- DataInizio
- DataFine
- CanaleValido

**Colonne opzionali**
- ScontoExtraPct
- MotivoPromo

## Note operative
- `Vendite_2025` deve avere almeno una riga dati.
- Gli altri fogli possono essere presenti anche senza righe dati, ma in quel caso la V2 mostra un warning perché alcune regole potrebbero non attivarsi.
- Intestazioni duplicate nello stesso foglio sono considerate un errore bloccante.
- Il controllo viene eseguito sia nella web app sia nel motore Python, per evitare bypass.
