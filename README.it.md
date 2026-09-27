# eFootball Wi-Fi Monitor

[English](README.md) | **Italiano**

Script Python per Windows che controlla la connessione mentre giochi a eFootball via Wi-Fi.
A fine sessione salva tutto in un **file Excel colorato**. Con `--sniff` ti avvisa anche con un suono, prima del calcio d'inizio, se la partita avrà una connessione buona o scarsa.

## Cosa misura

| Voce | Cosa ti dice |
|---|---|
| **router** (ping, sbalzi, pacchetti persi) | qualità del Wi-Fi tra PC e router: deve stare sotto 5 ms, con pochi sbalzi e 0% di pacchetti persi |
| **internet** (default 1.1.1.1) | qualità della linea del provider |
| **Wi-Fi** (segnale %, dBm, banda, canale) | potenza del segnale, con avvisi se sei sulla banda 2.4 GHz o se il PC cambia access point |
| **partita** (solo con `--sniff`) | server o avversario, nazione, ping, pacchetti al secondo, pacchetti persi stimati e "scatti" (pause oltre 150 ms nei dati in arrivo = lag) |

Verde = buono, giallo = al limite, rosso = problema.

## Installazione (una volta sola)

1. Installa Python 3 da https://www.python.org/downloads/ e spunta **"Add python.exe to PATH"**.
2. Copia `efootball_monitor.py` in una cartella, per esempio `C:\efootball-monitor`.
3. Nel Prompt dei comandi scrivi `pip install openpyxl scapy`.
4. Installa **Npcap** da https://npcap.com lasciando le opzioni di default. Serve solo per `--sniff`.
5. Tieni **attiva la posizione di Windows** (Impostazioni > Privacy e sicurezza > Posizione). Senza la posizione, Windows non lascia leggere il segnale Wi-Fi.

## Versione .exe (per chi non ha Python)

**Il modo più semplice:** scarica `eFootballMonitor.exe` dalla pagina [Releases](../../releases). Lo crea GitHub in automatico da questo codice.

In alternativa puoi crearlo tu:

1. Sul **tuo** PC, dove Python c'è già, metti `build_exe.bat` nella stessa cartella dello script e fai **doppio clic**.
2. Dopo un paio di minuti trovi il programma in `dist\eFootballMonitor.exe`.
3. Chi lo usa deve solo installare **Npcap** e tenere attiva la posizione di Windows.

Con un doppio clic sull'exe, Windows chiede i permessi di amministratore e l'analisi della partita (`--sniff`) parte già attiva. I file Excel vengono salvati accanto all'exe. Alla fine si preme Ctrl+C e poi Invio per chiudere.

L'exe non è firmato, quindi al primo avvio Windows può mostrare "Windows ha protetto il PC": clicca **Ulteriori informazioni > Esegui comunque**. Qualche antivirus può segnalarlo per errore, succede spesso con i programmi creati così.

## Come si usa

Apri il **Prompt dei comandi come amministratore**, entra nella cartella e avvia lo script:

```
cd C:\efootball-monitor
python efootball_monitor.py --sniff
```

Lascialo aperto mentre giochi: ogni 5 secondi scrive una riga. Quando hai finito, torna sulla finestra e premi **Ctrl+C**.

Senza `--sniff` (e senza amministratore) funziona lo stesso, ma misura solo Wi-Fi, router e internet, non la partita.

## Avviso prima della partita (con `--sniff`)

Appena il gioco si collega al server della partita, in circa 1 secondo il PC suona e scrive una riga come questa:

`>>> PARTITA IN ARRIVO (0.1 s)  server dedicato  34.154.0.13:5735  Milano, Italia  ping 8 ms -> connessione BUONA`

- **1 bip acuto** = connessione buona
- **2 bip** = così così, oppure non misurabile
- **3 bip gravi** = scarsa
- **+ 1 bip lungo** = l'avversario (o il server) è in **un'altra nazione**

A fine partita il prompt scrive anche la pagella della partita.

**Server dedicato** vuol dire che la partita passa da un server (per esempio Google Cloud a Milano). In questo caso l'IP dell'avversario non si vede, quindi la nazione indicata è quella del server.
**P2P** vuol dire che sei collegato direttamente all'avversario, quindi la nazione è la sua.

La nazione viene chiesta al servizio gratuito ip-api.com. Se hai Discord o altre chiamate aperte possono scattare avvisi falsi. Per togliere il suono aggiungi `--muto`.

## I file Excel

Quando premi **Ctrl+C** lo script crea due cose nella cartella dello script.

**1. Il file della sessione**, per esempio `efootball_log_20260927_214045.xlsx`, con tre fogli:
- **Riepilogo**: per che percentuale del tempo la connessione è stata OK, ATTENZIONE o PROBLEMA, le medie della sessione e i problemi più frequenti.
- **Partite**: le partite trovate in questa sessione, con server, nazione e ping.
- **Andamento**: una riga ogni 5 secondi. Le colonne **Giudizio** e **Cosa non va** spiegano a parole cosa succedeva (per esempio "Wi-Fi instabile, il ping al router balla di 12 ms"). Ogni numero è colorato e i filtri sono già attivi: nella colonna Giudizio tieni solo `PROBLEMA` per vedere subito i momenti peggiori.

**2. Lo storico `efootball_match_history.xlsx`**, un file unico a cui ogni sessione aggiunge le sue partite in fondo. Oltre a server, nazione e ping all'avvio, ogni partita ha una **pagella**: durata, ping medio durante la partita, scatti, pausa più lunga, pacchetti persi stimati e un giudizio finale (OK / ATTENZIONE / PROBLEMA). Col tempo ti fa vedere quali server e quali nazioni ti danno davvero lag. Tienilo **chiuso** quando premi Ctrl+C, altrimenti non si può aggiornare.


**Se chiudi la finestra invece di premere Ctrl+C**, l'Excel non viene creato. Nella cartella resta un file `.csv` di appoggio con lo stesso nome, e puoi trasformarlo in Excel così:

```
python efootball_monitor.py --excel efootball_log_20260927_214045.csv
```

## Come leggere i risultati

- **Router con sbalzi o pacchetti persi**: il problema è il Wi-Fi di casa. Avvicinati al router, usa la 5 GHz o un cavo.
- **Router ok ma internet male**: il problema è la linea o il provider.
- **Tutto ok ma tanti "scatti" in partita**: il server o l'avversario è lontano o instabile. Non è colpa tua.
- **Tutto verde ma senti ritardo nei comandi**: di solito dipende dall'avversario lontano, perché il gioco rallenta entrambi per tenervi sincronizzati. Può dipendere anche dal PC: prova a disattivare il V-Sync, a usare un controller col cavo e a mettere lo schermo in "modalità gioco".
- **Pacchetti persi in partita**: il server manda a ritmo fisso (circa 55 pacchetti al secondo). Lo script prende come riferimento il ritmo dei momenti migliori e conta quelli che mancano. È una stima: se la perdita è costante per tutta la partita viene sottostimata.
- **Server "no ping"**: molti server di gioco non rispondono al ping, è normale. In quel caso guarda gli "scatti".

## Opzioni

| Opzione | A cosa serve |
|---|---|
| `--sniff` | analizza la partita e dà l'avviso prima del calcio d'inizio (serve l'amministratore) |
| `--muto` | nessun suono quando viene trovata una partita |
| `--window 2` | una riga ogni 2 secondi invece di 5 |
| `--gateway 192.168.1.1` | indica tu l'IP del router, se non viene trovato da solo |
| `--internet 8.8.8.8` | pinga un altro host internet |
| `--excel FILE.csv` | crea l'Excel da un file `.csv` di appoggio rimasto |
