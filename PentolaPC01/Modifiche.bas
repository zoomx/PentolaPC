Attribute VB_Name = "Modifiche"
'2011 04 28
'Nella funzione InputComTimeOut cambiato il codice nella parte OnComm=False
'in modo che restituisca sempre un valore.
'Aggiunto NewTextVal = "" in SwapBytes
'Aggiunto tStringa = "" in Char2ascii


'2011 04 06
'In long record corretto errore premendo il pulsante annulla.

'2011 04 05
'Corretto l'errore di lettura dei valori nel file .ini
'Se non c'era il file o mancavano valori andava in errore


'2011 04 04
'Nella routine InputComTimeOutSantino corretto un errore grave.
'Gli imput dalla COM non controllavano se non c'erano tutti gli 8 valori così
'spesso l'ottavo veniva letto troppo presto e non riusciva a sincronizzarsi mai
'o perdeva la sincronizzazione specialmente con il convertitore USB-seriale.

'2011 03 22
'Aggiunta la variabile globale Kchamber
'Aggiunto il salvataggio e il recupero automatico dei setting
'Ricerca del perchè con la scheda di santino non funziona dopo la prima misura
'Trovato Usavo InputModeBinary con le stringhe unicode.
'Ricerca del perchè non funziona con la gascard Trovato: in Initcard eliminato il vbcrlf iniziale!
'L'azzeramento della matrice provoca problemi con l'assegnamento al grafico. Elimino il redim?
'No invece dell'erase uso il redim e ottengo lo stesso la cancellazione!!!!


'2011 03 07
'Continua l'azzeramento di variabili prima della misura
'Magari ci vorrebbe un controllo hai salvato i dati?
'Aggiunta una label per il tipo di scheda di misura attiva

'2011 03 03
'Inizio a lavorare con Erase per azzerare tutti i dati ad ogni misura


'2010 04 15
'Lo sfondo della casella di testo adesso è esplicitamente bianco &H00FFFFFF&
'perchè in alcuni sistemi appariva nero

'2010 03 15
'Aggiunta la gestione dell'interfaccia di Santino

'2006 04 19
'Aggiunta la gestione della scheda 3000ppm

'2006 05 17
'Aggiunta la gestione dell'acquisizione per tempi lunghissimi
'non ancora perfettamente operante.

'2006 10 03
'Completata la gestione dei tempi lunghi

'2006 10 18
'Cambiato il controllo di timeout con lo 0 che è la stringa restituita dalla routine in caso di timeoit

'2007 03 21
'Aggiunta la gestione della scheda 5%
'In MSChart1 RowCount è stato posto a 1

'2008 01 22
'Aggiunta la gestione della scheda 10%

'2008 03 19
'Inizio Aggiunta la gestione della scheda Mastrolia

'2009 10 20
'Aggiunta gestione errori in
'xmax xmin ymax e ymin del grafico
'I grafici non vengono ricaricati!!!!
'Errore semicorretto. LA gestione prevede un formato file diverso da
'quello che viene salvato
'No, viene salvato il file giusto forse la colpa è in long record
'Tutto corretto, ora carica sia i file corti che quelli long
'Aggiunta la generazione del nome del file di salvataggio in automatico

'2009 10 26
'Aggiunto il 5% a long record. Non so perchè ma mancava!
