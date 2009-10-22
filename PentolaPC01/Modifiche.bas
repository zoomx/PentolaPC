Attribute VB_Name = "Modifiche"
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
