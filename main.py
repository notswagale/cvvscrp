import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import warnings
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import os

warnings.filterwarnings('ignore', category=InsecureRequestWarning)

url = 'https://web.spaggiari.eu/cvv/app/default/genitori_note.php?ordine=materia&filtro=tutto'
cookies = {
    'webrole': 'gen',
    'webidentity': 'INCOLLA webidentity QUI',
    'PHPSESSID': 'INCOLLA PHPSESSID QUI'
}

try:
    resp = requests.get(url, cookies=cookies, verify=False)
    resp.raise_for_status()
except requests.RequestException as e:
    print(f"Errore nella richiesta HTTP: {e}")
    exit(1)

soup = BeautifulSoup(resp.text, 'html.parser')

q_map = {
    'q1': 'Trimestre 1',
    'q2': 'Trimestre 2',
    'q3': 'Trimestre 3',
    'q4': 'Pentamestre'
}

dati = []
current_subject = None

for row in soup.find_all('tr'):

    subject_td = row.find('td', class_=lambda x: x and 'registro redtext' in x)
    if subject_td:
        current_subject = subject_td.get_text(strip=True)
        continue
    
    if any('griglia_sep_darkgray_top' in td.get('class', []) for td in row.find_all('td')):
        cells = row.find_all('td')
        if len(cells) < 6:
            continue
            
        try:
            date_span = cells[1].find('span', class_='voto_data')
            data = date_span.get_text(strip=True) if date_span else ''
            
            voto_div = cells[2].find('div', class_='cella_div')
            voto = voto_div.find('p', class_='s_reg_testo').get_text(strip=True) if voto_div else ''
            
            tipo_span = cells[3].find('span', class_='voto_data')
            if tipo_span:
                tipo_text = tipo_span.get_text('\n', strip=True)
                tipo = tipo_text.split('\n')[0] if '\n' in tipo_text else tipo_text
                peso = tipo_text.split('Peso:')[-1].strip() if 'Peso:' in tipo_text else ''
            else:
                tipo = peso = ''
            
            periodo = 'Non specificato'
            grade_td = cells[2]
            for cls in grade_td.get('class', []):
                if cls in q_map:
                    periodo = q_map[cls]
                    break
            
            desc_div = cells[5].find('div')
            descrizione = desc_div.get_text(strip=True) if desc_div else ''
            
            dati.append({
                'Materia': current_subject,
                'Data': data,
                'Voto': voto,
                'Tipo': tipo,
                'Peso': peso,
                'Periodo': periodo,
                'Descrizione': descrizione
            })
        except Exception as e:
            print(f"Errore nell'elaborazione di una riga: {e}")
            continue

if not dati:
    print("❌ Nessun dato trovato. Possibili cause:")
    print("- Cookie scaduti/non validi")
    print("- Modifiche alla struttura della pagina")
    print("- Nessun voto presente")
    exit(1)



df = pd.DataFrame(dati)


try:
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True)
    df = df.sort_values(['Materia', 'Periodo', 'Data'])
except Exception as e:
    print(f"Errore nella conversione delle date: {e}")
    df = df.sort_values(['Materia', 'Periodo'])

try:
    with pd.ExcelWriter('registro_voti.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Voti')
        
    
        workbook = writer.book
        worksheet = writer.sheets['Voti']
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#FFD700',
            'border': 1
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        

        worksheet.set_column('A:A', 35) 
        worksheet.set_column('B:B', 12) 
        worksheet.set_column('C:C', 8)   
        worksheet.set_column('D:D', 15) 
        worksheet.set_column('E:E', 10) 
        worksheet.set_column('F:F', 15) 
        worksheet.set_column('G:G', 60)  
        
    print("✅ File Excel creato con successo: registro_voti.xlsx")

    try:
        os.startfile("registro_voti.xlsx")
    except Exception as e:
         print(f"Impossibile aprire automaticamente il file: {e}")

except Exception as e:
    print(f"Errore nella creazione del file Excel: {e}")
