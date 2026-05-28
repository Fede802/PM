#!/usr/bin/env python3
"""Genera la sezione Stime di Effort in formato per-attivita' nel notebook.tex.
Legge i dati dal JSON e usa nomi/ordini concordati con il PDF ufficiale."""
import json
import re

JSON_PATH = "stime_delphi.json"
TEX_PATH  = "report_build/notebook.tex"

# (id, nome) in ordine, raggruppati per macro-sezione
GROUPS = [
    ("1 Marketing", [
        ("1.1.1.1",  "Ricerca Organizzazioni"),
        ("1.1.1.2",  "Contatto Organizzazioni"),
        ("1.1.1.3",  "Onboarding Organizzazioni"),
        ("1.1.2",    "Gestione Contatto Organizzazioni già Registrate"),
        ("1.2.1.1",  "Individuazione Profilo medio utente"),
        ("1.2.1.2",  "Analisi tipologia di Eventi più remunerativi"),
        ("1.2.1.3",  "Analisi Provenienza geografica delle visualizzazioni della piattaforma"),
        ("1.2.2.1.1","Ricerca e selezione influencer"),
        ("1.2.2.1.2","Contatto e accordi"),
        ("1.2.2.1.3","Briefing contenuti"),
        ("1.2.2.1.4","Monitoraggio e ottimizzazione pubblicazioni"),
        ("1.2.2.2.1","Ricerca e selezione piattaforme ads"),
        ("1.2.2.2.2","Creazione contenuti (grafiche/video)"),
        ("1.2.2.2.3","Setup campagna"),
        ("1.2.2.2.4","Monitoraggio e ottimizzazione"),
    ]),
    ("2 App Mobile", [
        ("2.1.1.1.1","Design della schermata Home per utenti non registrati"),
        ("2.1.1.1.2","Implementazione della schermata Home per utenti non registrati"),
        ("2.1.1.2.1","Design della schermata Home per utenti registrati"),
        ("2.1.1.2.2","Implementazione della schermata Home per utenti registrati"),
        ("2.1.2.1.1","Design del sistema di ricerca per la schermata Esplora"),
        ("2.1.2.1.2","Implementazione del sistema di ricerca per la schermata Esplora"),
        ("2.1.2.2.1","Design della schermata Esplora per utenti non registrati"),
        ("2.1.2.2.2","Implementazione della schermata Esplora per utenti non registrati"),
        ("2.1.2.3.1","Design della schermata Esplora per utenti registrati"),
        ("2.1.2.3.2","Implementazione della schermata Esplora per utenti registrati"),
        ("2.1.2.4.1","Design dei filtri per gli eventi nella schermata Esplora"),
        ("2.1.2.4.2","Implementazione dei filtri per gli eventi nella schermata Esplora"),
        ("2.1.3.1.1","Design della schermata Pagina Evento"),
        ("2.1.3.1.2","Implementazione della schermata Pagina Evento"),
        ("2.1.3.2.1","Design della sezione Recensioni"),
        ("2.1.3.2.2","Implementazione della sezione Recensioni"),
        ("2.1.3.3.1","Design della navigazione verso la schermata di checkout"),
        ("2.1.3.3.2","Implementazione della schermata di checkout"),
        ("2.1.3.4.1","Design della funzionalità di like"),
        ("2.1.3.4.2","Implementazione della funzionalità di like"),
        ("2.1.4.1.1","Design della sezione dei biglietti nella schermata di checkout"),
        ("2.1.4.1.2","Implementazione della sezione dei biglietti nella schermata di checkout"),
        ("2.1.4.2.1","Design del sistema di completamento dell'acquisto"),
        ("2.1.4.2.2","Implementazione del sistema di completamento dell'acquisto"),
        ("2.1.5.1.1","Design della schermata delle recensioni"),
        ("2.1.5.1.2","Implementazione della schermata delle recensioni"),
        ("2.1.5.2.1","Design del resoconto del rating dell'organizzazione"),
        ("2.1.5.2.2","Implementazione del resoconto del rating"),
        ("2.1.5.3.1","Design del sistema di inserimento di una recensione"),
        ("2.1.5.3.2","Implementazione del sistema di inserimento di una recensione"),
        ("2.1.6.1.1","Design della sezione delle informazioni del profilo"),
        ("2.1.6.1.2","Implementazione della sezione delle informazioni del profilo"),
        ("2.1.6.2.1","Design del sistema di modifica del profilo"),
        ("2.1.6.2.2","Implementazione del sistema di modifica del profilo"),
        ("2.1.6.3.1","Design del sistema di modifica delle impostazioni dell'account"),
        ("2.1.6.3.2","Implementazione del sistema di modifica delle impostazioni dell'account"),
        ("2.1.6.4.1","Design del sistema di follow/unfollow"),
        ("2.1.6.4.2","Implementazione del sistema di follow/unfollow"),
        ("2.1.6.5.1","Design del sistema di contatto nei profili delle organizzazioni"),
        ("2.1.6.5.2","Implementazione del sistema di contatto nei profili delle organizzazioni"),
        ("2.1.6.6.1","Design della sezione degli eventi a cui l'utente ha messo like"),
        ("2.1.6.6.2","Implementazione della sezione degli eventi a cui l'utente ha messo like"),
        ("2.1.6.7.1","Design della sezione degli eventi a cui l'utente ha partecipato"),
        ("2.1.6.7.2","Implementazione della sezione degli eventi a cui l'utente ha partecipato"),
        ("2.1.6.8.1","Design della sezione degli eventi creati da un'organizzazione"),
        ("2.1.6.8.2","Implementazione della sezione degli eventi creati da un'organizzazione"),
        ("2.1.6.9.1","Design della sezione delle bozze degli eventi"),
        ("2.1.6.9.2","Implementazione della sezione delle bozze degli eventi"),
        ("2.1.6.10.1","Design del sistema di visualizzazione delle statistiche"),
        ("2.1.6.10.2","Implementazione del sistema di visualizzazione delle statistiche"),
        ("2.1.6.11.1","Design della sezione delle recensioni"),
        ("2.1.6.11.2","Implementazione della sezione delle recensioni"),
        ("2.1.7.1.1","Design della schermata dei biglietti"),
        ("2.1.7.1.2","Implementazione della schermata dei biglietti"),
        ("2.1.8.1.1","Design della sezione delle conversazioni"),
        ("2.1.8.1.2","Implementazione della sezione delle conversazioni"),
        ("2.1.8.2.1","Design della schermata delle chat"),
        ("2.1.8.2.2","Implementazione della schermata delle chat"),
        ("2.1.8.3.1","Design della funzionalità di ricerca delle conversazioni"),
        ("2.1.8.3.2","Implementazione della schermata di ricerca delle conversazioni"),
        ("2.1.9.1.1","Design della sezione delle informazioni dell'evento"),
        ("2.1.9.1.2","Implementazione della sezione delle informazioni dell'evento"),
        ("2.1.9.2.1","Design delle funzionalità di gestione dell'evento"),
        ("2.1.9.2.2","Implementazione delle funzionalità di gestione dell'evento"),
        ("2.1.10.1.1","Design della schermata di registrazione/login"),
        ("2.1.10.1.2","Implementazione della schermata di registrazione/login"),
        ("2.1.11.1.1","Design della schermata di scansione biglietti"),
        ("2.1.11.1.2","Implementazione della schermata di scansione biglietti"),
        ("2.1.12.1.1","Design del sistema di navigazione"),
        ("2.1.12.1.2","Implementazione del sistema di navigazione"),
        ("2.1.13.1.1","Verificare che l'app sia responsive"),
        ("2.2.1",    "Studio Funzionamento Notifiche push"),
        ("2.2.2",    "Implementazione Notifiche push"),
        ("2.3.1.1",  "Individuazione dei Principali Task Per Valutare l'usabilità"),
        ("2.3.1.2",  "Contatto Utenti Target e Organizzazione Sessioni di Usability Test"),
        ("2.3.1.3",  "Svolgimento Sessioni di Usability Test"),
        ("2.3.1.4",  "Valutazione Risultati Dei Test Di Usabilità"),
        ("2.4.1",    "Configurazione account Google Play"),
        ("2.4.2",    "Configurazione account App Store"),
        ("2.4.3",    "Preparazione asset (icone, screenshot, descrizioni)"),
        ("2.4.4",    "Build e firma dell'app"),
        ("2.4.5",    "Pubblicazione App"),
        ("2.4.6",    "Gestione Review"),
    ]),
    ("3 Aggiornamento Sistema Esistente", [
        ("3.1.1","Scelta provider OAuth2 (e.g. Google)"),
        ("3.1.2","Aggiornamento configurazione servizio Autenticazione"),
        ("3.1.3","Testing servizio Autenticazione"),
        ("3.1.4","Integrazione aggiornamenti nel servizio Utenti"),
        ("3.1.5","Aggiornamento Documentazione API del servizio Utenti"),
        ("3.1.6","Testing servizio Utenti"),
        ("3.2.1.1","Sviluppo endpoint RBS 3.2.1 nel servizio Geolocalizzazione"),
        ("3.2.1.2","Documentazione API con endpoint RBS 3.2.1"),
        ("3.2.1.3","Testing della funzionalità RBS 3.2.1"),
        ("3.2.2.1","Sviluppo endpoint RBS 3.2.2 nel servizio Geolocalizzazione"),
        ("3.2.2.2","Integrazione campo posizione in tutti i servizi che lo richiedono"),
        ("3.2.2.3","Documentazione API con endpoint RBS 3.2.2"),
        ("3.2.2.4","Testing della funzionalità RBS 3.2.2"),
        ("3.3.1.1","Aggiunta notifica RBS 3.3.1 al servizio Notifiche"),
        ("3.3.1.2","Aggiornamento Documentazione API con notifica RBS 3.3.1"),
        ("3.3.1.3","Testing della funzionalità di notifica RBS 3.3.1"),
        ("3.3.2.1","Aggiunta notifica RBS 3.3.2 al servizio Notifiche"),
        ("3.3.2.2","Aggiornamento Documentazione API con notifica RBS 3.3.2"),
        ("3.3.2.3","Testing della funzionalità di notifica RBS 3.3.2"),
        ("3.3.3.1","Aggiunta notifica RBS 3.3.3 al servizio Notifiche"),
        ("3.3.3.2","Aggiornamento Documentazione API con notifica RBS 3.3.3"),
        ("3.3.3.3","Testing della funzionalità di notifica RBS 3.3.3"),
        ("3.3.4.1","Aggiunta notifica RBS 3.3.4 al servizio Notifiche"),
        ("3.3.4.2","Aggiornamento Documentazione API con notifica RBS 3.3.4"),
        ("3.3.4.3","Testing della funzionalità di notifica RBS 3.3.4"),
        ("3.3.5.1","Aggiunta notifica RBS 3.3.5 al servizio Notifiche"),
        ("3.3.5.2","Aggiornamento Documentazione API con notifica RBS 3.3.5"),
        ("3.3.5.3","Testing della funzionalità di notifica RBS 3.3.5"),
        ("3.3.6.1","Aggiunta notifica RBS 3.3.6 al servizio Notifiche"),
        ("3.3.6.2","Aggiornamento Documentazione API con notifica RBS 3.3.6"),
        ("3.3.6.3","Testing della funzionalità di notifica RBS 3.3.6"),
        ("3.3.7.1","Aggiunta notifica RBS 3.3.7 al servizio Notifiche"),
        ("3.3.7.2","Aggiornamento Documentazione API con notifica RBS 3.3.7"),
        ("3.3.7.3","Testing della funzionalità di notifica RBS 3.3.7"),
        ("3.3.8.1","Aggiunta notifica RBS 3.3.8 al servizio Notifiche"),
        ("3.3.8.2","Aggiornamento Documentazione API con notifica RBS 3.3.8"),
        ("3.3.8.3","Testing della funzionalità di notifica RBS 3.3.8"),
        ("3.4.1","Sviluppo servizio Subscription per la gestione dei piani di abbonamento"),
        ("3.4.2","Integrazione servizio Subscription con il resto del sistema"),
        ("3.4.3","Documentazione API del servizio Subscription"),
        ("3.4.4","Testing servizio Subscription"),
        ("3.5.1.1","Aggiunta endpoint RBS 3.5.1 al servizio Ticketing"),
        ("3.5.1.2","Aggiornamento Documentazione API con endpoint RBS 3.5.1"),
        ("3.5.1.3","Testing della funzionalità RBS 3.5.1"),
        ("3.5.2.1","Aggiornamento endpoint di checkout per supportare RBS 3.5.2 nel servizio Ticketing"),
        ("3.5.2.2","Aggiornamento Documentazione API endpoint di checkout"),
        ("3.5.2.3","Testing della funzionalità di checkout aggiornata con RBS 3.5.2"),
        ("3.5.3.1","Aggiunta endpoint RBS 3.5.3 al servizio Ticketing"),
        ("3.5.3.2","Aggiornamento Documentazione API con endpoint RBS 3.5.3"),
        ("3.5.3.3","Testing della funzionalità RBS 3.5.3"),
        ("3.6.1","Aggiornamento schermata di registrazione/login con supporto OAuth2"),
        ("3.6.2","Integrazione API servizio di geolocalizzazione"),
        ("3.6.3","Supporto nuove notifiche"),
        ("3.6.4","Supporto inserimento codice sconto nella schermata di checkout"),
        ("3.6.5","Design schermata per i piani di abbonamento"),
        ("3.6.6","Implementazione schermata per i piani di abbonamento"),
        ("3.6.7","Integrazione sistema di abbonamento nelle schermate del sito, con aggiunta schermata ad hoc per le statistiche di performance per le organizzazioni"),
        ("3.7.1.1","Verificare che i nuovi endpoint e quelli aggiornati rispettino il requisito di performance"),
        ("3.7.2.1","Verificare che il frontend web sia responsive"),
    ]),
]

def tex_escape(s):
    repl = {"&":r"\&","%":r"\%","$":r"\$","#":r"\#","_":r"\_","{":r"\{","}":r"\}","~":r"\textasciitilde{}","^":r"\textasciicircum{}"}
    return "".join(repl.get(c,c) for c in s)

def fmt(v):
    return ("%.1f" % v) if isinstance(v,(int,float)) else str(v)

state = json.load(open(JSON_PATH, encoding="utf-8"))
out = []
out.append(r"\footnotesize")
for group_name, acts in GROUPS:
    out.append(r"\subsection*{%s}" % tex_escape(group_name))
    out.append(r"\addcontentsline{toc}{subsection}{%s}" % tex_escape(group_name))
    for act_id, act_name in acts:
        rounds = state.get(act_id, [])
        out.append(r"\medskip\noindent\textbf{Attività %s -- %s}\par\nopagebreak\smallskip"
                   % (tex_escape(act_id), tex_escape(act_name)))
        out.append(r"\nopagebreak\noindent\begin{tabular}"
                   r"{|>{\centering\arraybackslash}p{1.0cm}|"
                   r">{\centering\arraybackslash}p{2.5cm}|"
                   r">{\centering\arraybackslash}p{2.5cm}|"
                   r">{\centering\arraybackslash}p{2.5cm}|"
                   r">{\centering\arraybackslash}p{2.1cm}|"
                   r">{\centering\arraybackslash}p{2.3cm}|}")
        out.append(r"\hline")
        out.append(r"\rowcolor{brandbg}\textbf{Round} & "
                   r"\textbf{Alice Alfonsi\newline(ore/u)} & "
                   r"\textbf{Federico Bravetti\newline(ore/u)} & "
                   r"\textbf{Tommaso Brini\newline(ore/u)} & "
                   r"\textbf{Media\newline(ore/u)} & "
                   r"\textbf{Mediana\newline(ore/u)} \\\hline")
        for r in rounds:
            out.append(r"%d & %s & %s & %s & %s & %s \\\hline" % (
                r["round"], fmt(r["alice"]), fmt(r["federico"]),
                fmt(r["tommaso"]), fmt(r["mean"]), fmt(r["median"])))
        out.append(r"\end{tabular}\par")
out.append(r"\normalsize")
new_block = "\n".join(out)

# Sostituisco la longtable corrente nel tex
tex = open(TEX_PATH, encoding="utf-8").read()
pattern = re.compile(
    r"\\footnotesize\s*\n\\begin\{longtable\}\{@\{\}p\{6\.4cm\}cccccc@\{\}\}.*?"
    r"\\end\{longtable\}\s*\n\\normalsize.*?"
    r"(?=\\medskip|\\clearpage)", re.S)
m = pattern.search(tex)
if not m:
    raise SystemExit("longtable STIME non trovata")
new_tex = tex[:m.start()] + new_block + "\n" + tex[m.end():]
open(TEX_PATH, "w", encoding="utf-8").write(new_tex)
print("OK, sostituito blocco STIME")
