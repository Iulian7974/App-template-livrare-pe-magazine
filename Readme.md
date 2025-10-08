# Generator template N-ERP (Streamlit)

Aplicație Streamlit care generează template-uri N-ERP dintr-un fișier de exemplu, cu opțiuni:
- Un **singur Excel** cu **foi separate** (una per depozit, foaia denumită cu ID-ul depozitului).
- **Fișiere separate** (câte un Excel per depozit) într-un **ZIP**.
- Descărcare pentru un **depozit individual** (opțional).

## 1) Cerințe de intrare
Fișier Excel (.xlsx) cu coloanele:
- `Warehouse`
- `Material Code`
- `Quantity`
- `New Price`

Pot fi acceptate și variații minore de denumiri (ex. „Cantitate” pentru `Quantity`, „Cod material” pentru `Material Code`), aplicația normalizează automat.

## 2) Reguli de generare template
Fiecare foaie/fișier respectă exact structura:
- `Material Code` (din intrare)
- `Quantity` (din intrare)
- `Val. Type` = **A**
- `New Price` (din intrare)
- `I/C New Price` = gol
- `I/C New Cur.` = gol
- `Plant` = **L402**
- `S / L` = gol
- `Biz.ModelGroup` = gol
- `Biz.Category` = gol

## 3) Rulare locală
```bash
git clone <repo-url>
cd nerp-generator
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
