<%

' Language definition - Use ISO 2 chr country value
Const INFO_Langset_Tip_Warning = "it"

' Warning
Dim TIP_Warning_t(1)
Dim TIP_Warning_c(1)

TIP_Warning_t(0) = "Blocco installazione non attivo"
TIP_Warning_c(0) = "Si consiglia l'attivazione del blocco dell'installazione per evitare modifiche non autorizzate alla configurazione del programma. <br />Il blocco consiste nella creazione automatica di un file specifico che impedisce l'avvio di qualsiasi procedura di installazione o aggiornamento.</p><p style='text-align: center;'><a href='setup_lock.asp?lock=1' title='Attiva blocco installazione'><img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/setuplock.png' alt='Attiva blocco installazione' border='0' align='middle' /> Attiva blocco installazione</a>"
TIP_Warning_t(1) = "Superamento della soglia massima"
TIP_Warning_c(1) = "Si consiglia di procedere con l'eliminazioni di alcuni dati dalla tabella corrente per ridurre il volume occupato. <br />Un eccesso di dati può rallentare notevolmente l'esecuzione del programma."

%>