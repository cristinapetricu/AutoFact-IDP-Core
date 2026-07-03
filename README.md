# AutoFact IDP Core

Robot RPA pentru procesarea automată a facturilor, construit în UiPath Studio. Sistemul citește facturi (PDF/imagini) dintr-un folder de intrare, extrage datele relevante folosind Document Understanding (AI), convertește sumele în RON, generează un raport Excel centralizat, poate genera facturi noi (Human-in-the-Loop) și trimite automat un email de confirmare.

**Candidat:** Cristina Petricu
**Coordonator științific:** Conf. dr. ing. Mircea Rădac
**Universitatea Politehnica Timișoara** — Facultatea de Automatică și Calculatoare, Ingineria Sistemelor

Repository: https://github.com/cristinapetricu/AutoFact-IDP-Core

---

## Structura proiectului

```
AutoFactIDPCore/
├── Data/
│   ├── Input/              ← facturi PDF de procesat (se pun aici de utilizator)
│   ├── Output/              ← rapoarte Excel generate (rezultat)
│   ├── Processed/          ← facturi procesate cu succes (mutate automat)
│   ├── Failed/              ← facturi cu erori de procesare (mutate automat)
│   ├── Templates/
│   │   └── TemplateFactura.xlsx   ← șablon pentru facturi generate (M4)
│   └── config.xlsx          ← configurare robot (foldere, prag confidence, email destinatar)
├── Framework/
│   └── InitAllSettings.xaml ← încarcă config.xlsx la pornire
├── Main.xaml                ← orchestrator principal (REFramework)
├── Process.xaml
├── M1_ExtractInvoiceData.xaml    ← extragere date din PDF via Document Understanding
├── M2_WriteToExcel.xaml          ← generare raport Excel (statistici + grafic)
├── M3_CurrencyConversion.xaml    ← conversie valutară (Frankfurter API, cache 1h)
├── M4_GenerateInvoice.xaml       ← generare factură nouă (Human-in-the-Loop)
├── M5_SendConfirmationEmail.xaml ← trimitere email cu raportul atașat
├── ParseAmount.xaml          ← utilitar: normalizare sume (format RO/EN)
├── ParseAndFormatDate.xaml   ← utilitar: normalizare date
├── project.json              ← manifest proiect UiPath
└── project.uiproj
```

---

## Cerințe / Dependențe

- **UiPath Studio** (Community sau Enterprise Edition) — versiune 2024.x sau ulterioară
- Conexiune la internet (necesară pentru Document Understanding AI și pentru API-ul de curs valutar)
- Microsoft Excel instalat local (activitățile de scriere Excel folosesc automatizare Excel Interop)
- Pachete UiPath (se instalează automat la deschiderea proiectului, din `project.json`):
  - `UiPath.DocumentUnderstanding.Activities` ≥ 3.2.0-preview
  - `UiPath.Excel.Activities` ≥ 3.5.1
  - `UiPath.Mail.Activities` ≥ 2.9.0-preview
  - `UiPath.MicrosoftOffice365.Activities` ≥ 3.9.0-preview
  - `UiPath.GSuite.Activities` ≥ 3.9.0-preview
  - `UiPath.System.Activities` ≥ 26.6.0
  - `UiPath.UIAutomation.Activities` ≥ 26.4.2-preview
  - `UiPath.WebAPI.Activities` ≥ 2.5.0-preview
  - `UiPath.Word.Activities` ≥ 2.5.1

---

## Pași de instalare

1. **Clonează repository-ul:**
   ```
   git clone https://github.com/cristinapetricu/AutoFact-IDP-Core.git
   ```

2. **Deschide proiectul în UiPath Studio:**
   - Lansează UiPath Studio
   - `Open` → selectează folderul clonat → fișierul `project.json`
   - Studio va instala automat dependențele lipsă (bara de progres la deschidere)

3. **Configurează `Data\config.xlsx`** (sheet `Settings`), completează cheile:

   | Cheie | Descriere | Exemplu |
   |---|---|---|
   | `Input_Folder` | folder facturi de intrare | `Data\Input` |
   | `Output_Folder` | folder rapoarte generate | `Data\Output` |
   | `Processed_Folder` | folder facturi procesate cu succes | `Data\Processed` |
   | `Failed_Folder` | folder facturi cu erori | `Data\Failed` |
   | `TemplateFactura` | șablon pentru facturi generate | `Data\Templates\TemplateFactura.xlsx` |
   | `EmailTo` | adresa destinatar raport | *(adresa ta)* |

4. **Conectează un cont de email** (necesar pentru M5) — la prima rulare, activitatea de trimitere email va cere autentificare (Outlook/Gmail, în funcție de pachetul folosit).

---

## Compilare

Proiectul este interpretat direct de UiPath Studio/Robot — nu necesită un pas de compilare separat în sensul clasic. Pentru a genera un pachet publicabil (`.nupkg`), din UiPath Studio:

```
Design → Publish → Publish Process
```

Aceasta creează un pachet versionat, pregătit pentru rulare via UiPath Robot/Orchestrator (opțional — nu e necesar pentru rularea locală, descrisă mai jos).

---

## Lansare (rulare locală)

1. Pune una sau mai multe facturi (PDF/imagine) în `Data\Input\`
2. În UiPath Studio, deschide `Main.xaml`
3. Apasă **Run** (▶) din bara de sus, sau `F5`
4. Robotul:
   - extrage automat datele din fiecare factură (Document Understanding)
   - convertește sumele în RON, dacă e cazul
   - clasifică automat fiecare factură ca **Succes** sau **Eroare** (pe baza pragului de încredere și a datelor extrase)
   - mută facturile procesate în `Data\Processed` sau `Data\Failed`, după caz
   - generează raportul Excel în `Data\Output` (statistici, grafic, listă facturi)
   - întreabă opțional dacă se generează o factură nouă (M4 — Human-in-the-Loop)
   - trimite un email cu raportul atașat (M5)

---

## Note

- Fișierele binare compilate (pachete `.nupkg`, cache-uri locale `.local/`, `.settings/`) sunt excluse din repository conform `.gitignore` — nu sunt necesare pentru a rula/compila proiectul din sursă.
- Pentru testare, se pot folosi facturi de test în diverse formate (PDF nativ, scanat) pentru a observa comportamentul de clasificare Succes/Eroare al robotului.
