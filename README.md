<p align="center">
  <img src="https://github.com/user-attachments/assets/4b5847a5-3753-416a-b3fd-42857bfa6699" width="200" />
</p>

# 📊 CVV SCRP

Uno script Python per estrarre i voti dal registro elettronico *Spaggiari* tramite scraping, e salvarli in un file Excel ordinato e leggibile.

> ⚠️ **Nota.** Questo progetto è in fase iniziale di sviluppo, aspettati molti cambiamenti.

---

## 🔍 Caratteristiche

- Effettua scraping dei voti dal registro elettronico Spaggiari
- Organizza i dati per materia, tipo, data e periodo
- Esporta i voti in un file Excel (`registro_voti.xlsx`) con formattazione chiara

---

## ⚙️ Requisiti

- Python 3.7+
- Le seguenti librerie Python (installabili con `pip`):

```bash
pip install requests beautifulsoup4 pandas xlsxwriter
```

---

## 🛠️ Installazione e Uso

1. **Clona il repository** o scarica lo script:
   ```bash
   git clone https://github.com/notswagale/cvvscrp.git
   cd cvvscrp
   ```

2. **Inserisci i tuoi cookie di sessione:**
   Apri lo script Python e **sostituisci**:
   ```python
   'webidentity': 'INCOLLA webidentity QUI',
   'PHPSESSID': 'INCOLLA PHPSESSID QUI'
   ```

   Puoi recuperare questi cookie accedendo al registro elettronico dal browser, ispezionando i cookie della sessione.

3. **Esegui lo script:**
   ```bash
   python registro_scraper.py
   ```

4. Il file `registro_voti.xlsx` verrà generato nella cartella del progetto e, se possibile, aperto automaticamente.

---

## 📌 Note Importanti

- **Verifica i cookie**: i cookie di sessione possono scadere. Se il file non si genera, controlla che siano aggiornati.
- **Non usare questo script per automatizzare accessi non autorizzati.**
- **La struttura del sito potrebbe cambiare**: in tal caso, potrebbe essere necessario aggiornare lo script.

---

## 📬 Contatti

Hai suggerimenti, problemi? Non esitare a contattarmi:

🗨️ Discord: `attempingtoindexnil`

✉️ E-Mail: alessio@notswaggames.com
