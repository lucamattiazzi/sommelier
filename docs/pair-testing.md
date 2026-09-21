# Provare Sommelier

Stato: **preview per test personali**, non una release Marketplace.
I pacchetti beta sono pubblicati su npm. Il servizio va avviato in sviluppo oppure sul proprio server.

## Avvio locale

Prerequisiti: Node.js 22.12+ e pnpm tramite Corepack. Dalla radice del checkout:

```sh
corepack enable
pnpm install --frozen-lockfile
pnpm build
pnpm test:pairing
```

Due terminali, entrambi dalla radice:

```sh
pnpm server:dev
```

```sh
pnpm addin:dev
```

Vite serve il task pane su `https://localhost:3000/taskpane.html`; il relay locale ascolta su `127.0.0.1:3001`.
Il primo avvio può chiedere di installare e considerare attendibile il certificato di sviluppo Office.
Carica in Excel il manifest **`apps/addin/manifest.xml`**.
L'URL del manifest è `https://localhost:3000/manifest.xml` per i percorsi di sideload che lo accettano.
Per il sideload su Mac/Web usa i [passaggi ufficiali Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

Aprendo l'URL in un browser normale usi solo un workbook sintetico in memoria: non è una prova di Office.js o del ricalcolo Excel.

## Collegare il tuo coding agent

1. Apri il TaskPane: mostra le istruzioni e **Copy agent prompt**, senza chiedere un nome.
2. Copia il prompt e incollalo nella sessione del tuo agente. Non serve installare o leggere una skill.
   Il prompt avvia il bridge Node locale e contiene comandi RPC, approvazioni e ciclo di ascolto.
3. Attendi **Connected · end-to-end encrypted**: compare la chat, con il campo di invio sempre visibile.
4. In **Options → Connection name** puoi cambiare il nome predefinito. Contesto, connessioni e
   cronologia sono nelle opzioni; lascia disattivata l'approvazione automatica per questi test.
5. Alla riapertura, riprendi bridge e ascolto nell'agente e usa **Options → Saved terminals → Reconnect**.
   La skill opzionale può gestire questo passaggio senza ripetere il prompt iniziale.

In sviluppo agente e relay condividono la macchina. Per altri dispositivi servono origine HTTPS/WSS,
certificato e manifest coerenti. Un bridge connesso non risveglia un agente che ha smesso di ascoltare.
Il prompt contiene una chiave privata: incollandolo all'agente può finire nel contesto del suo provider.
Non inserirlo in issue, commit o screenshot condivisi. Vedi [architettura e limiti](secure-pairing.md).

Test automatico dell'intero flusso su workbook sintetico, dopo la build:

```sh
pnpm exec playwright install chromium
pnpm test:browser
```

## Workbook sintetico

In un foglio chiamato `Sheet1`, compila:

| Cella | Contenuto |
|---|---|
| A1 | Item |
| B1 | Amount |
| A2 | Example |
| B2 | 10 |
| A3 | Sample |
| B3 | 20 |
| A4 | Total |
| B4 | =SUM(B2:B3) |

Usa la formula equivalente della tua lingua se Excel lo richiede. Non utilizzare workbook aziendali.

## Test manuali e risultati attesi

| Prova | Azione | Risultato atteso |
|---|---|---|
| Lettura | Seleziona A1:B4 e chiedi di descriverla | Valori e formula disponibili; nessuna scrittura |
| Contesto fissato | Premi `Pin selection`, seleziona D1, chiedi `excel.context.get` | Modalità manuale e range A1:B4 conservato |
| Approvazione | Chiedi di cambiare B2 da 10 a 42 via preview/commit | Dialogo prima/dopo; B2 resta 10 fino all'approvazione |
| Rifiuto | Rifiuta una proposta su B3 | B3 resta 20 |
| Conflitto | Fai creare una preview su B2, modifica B2 manualmente, poi chiedi commit | `WORKBOOK_CONFLICT`; la modifica manuale rimane |
| Undo | Approva una modifica e usa `Recent operations` → `Undo` | Ripristino dopo approvazione |
| Formula | Sostituisci B4 via preview/commit, poi undo | Torna la formula, non solo il valore calcolato; verifica dalla barra formule |
| Undo con conflitto | Dopo commit modifica il target manualmente, poi undo | Rifiuto; modifica manuale preservata |
| Limiti | Chiedi una lettura A1:A10001 o clear A1:A1001 | `RANGE_LIMIT_EXCEEDED` prima di accedere ai contenuti |
| Formati | Chiedi clear con applyTo=formats | `UNSUPPORTED_OPERATION`; niente cancellazione di valori |
| Domanda | Chiedi all'agente di usare excel.user.ask con due scelte | Scelta nel task pane e risposta all'agente |
| Riconnessione | Chiudi e riapri il pane, riavvia la skill se necessario, premi `Reconnect` | Stesso terminale senza un nuovo setup; chat e undo locali ricominciano |
| Revoca locale | Premi `Forget` in Excel e usa `forget --name NOME` nel bridge | Credenziali locali rimosse dai due estremi |

Il pin è **contesto, non un permesso di accesso**. Il modello può ancora chiedere altri range espliciti.
I limiti predefiniti del controller relay sono 10.000 celle in lettura, 1.000 in modifica e 100 operazioni
memorizzate per sessione. Il lettore tabelle attuale carica l'intera tabella prima di paginarla: per questo
le tabelle sopra il limite vengono rifiutate, anche chiedendo una pagina piccola.

Il confronto prima/dopo del dialogo mostra fino a 8 righe e 5 colonne; l'approvazione riguarda tutto il range indicato.
L'undo conserva valori e formule per le operazioni preview/commit della sessione; non è un backup del file,
non ripristina formati o strutture e non offre una transazione atomica fra più utenti.
Per le scritture dirette senza operation ID, usare l'undo di Excel quando disponibile; Sommelier non le inserisce nella cronologia operazioni.
Formule dinamiche, celle unite, fogli protetti e workbook con connessioni esterne richiedono ulteriori test Office reali.

## Endpoint diretto e arresto

Il percorso `Disconnect / use a direct agent` accetta un endpoint che implementa **Sommelier 0.1** e CORS.
Non inserire direttamente una normale API `/v1/chat/completions`: serve l'adapter/bridge appropriato.
Il token resta nella memoria della pagina e viene eliminato passando fuori dalla connessione diretta.
`Stop task` annulla il turno in corso; uscire dalla chat annulla anche eventuali approvazioni pendenti.

## Prima di pubblicare

- Verificare l'intero percorso su Excel Mac e Web con un agente reale; Windows se dichiarato supportato.
- Verificare ripristino di formule, stringhe letterali e celle vuote in Excel reale.
- Configurare dominio HTTPS/WSS, manifest di produzione e materiali richiesti dal Marketplace.
- Verificare protezione da abuso del relay, gestione delle sessioni e costi di esercizio.
- Evitare claim di universalità: documentare harness e ambienti effettivamente provati.

La CI include ora anche i test delle applicazioni. Due test avviano un bridge con socket Unix:
un ambiente che vieta tali socket può bloccarli con `EPERM`; non sono stati disabilitati.
