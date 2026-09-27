# ICG Drive Test – programator vehicule demo FOTON / CAVAN

Aplicație web instalabilă pe telefon (PWA) pentru programarea drive test-urilor pe vehiculele demonstrative
Tunland V9, Tunland G7, eToano și Cavan. Același stack ca „Foton by Inter Cargo”:
**GitHub Pages** (găzduire) + **Google Sheets / Apps Script** (date, Drive, e-mail).

## Conținutul pachetului

| Fișier | Rol |
|---|---|
| `index.html` | Aplicația completă |
| `Code.gs` | Backend-ul Apps Script |
| `manifest.json`, `sw.js` | Instalare pe telefon și pornire rapidă |
| `icons/brand/foton.png`, `intercargo.png` | Logo-uri albe pentru antet (aceleași ca în fotonagenti) |
| `icons/brand/*-dark.png` | Logo-uri închise la culoare pentru contract și procese-verbale |
| `icons/icon-192.png`, `icon-512.png` | Iconița aplicației |

## Instalare (aprox. 20 de minute)

> **Important:** creează Sheet-ul din **contul Google al firmei**. Scriptul rulează „ca tine”,
> deci e-mailurile către clienți pleacă de pe contul care deține Sheet-ul.

1. **Google Sheet** nou, „ICG Drive Test” → Extensii → Apps Script → șterge conținutul și lipește `Code.gs`.
2. Setări proiect (rotița) → Fus orar: *(GMT+02:00) Bucharest*.
3. Selectează funcția **`setup`** → Rulează → acceptă permisiunile. Setup-ul creează:
   foile, folderul **„ICG Drive Test – Documente”** în Drive, cele 4 vehicule,
   contul **Florin Loghin / PIN 1234 (admin)** și alerta zilnică de la ora 8.
4. **Implementare → Implementare nouă → Aplicație web** · Execută ca: **Eu** · Acces: **Oricine** → copiază URL-ul `/exec`.
5. În `index.html`, sus în script: `const API_URL = 'URL-ul copiat';`
6. **GitHub**: repo nou (ex. `florinx1/drivetest`), urcă toate fișierele păstrând folderul `icons/`,
   apoi Settings → Pages → Branch `main` / root.
7. Intră în aplicație ca Florin Loghin (PIN 1234), apoi din **Admin**:
   - **Contul meu** (atingi numele din antet) → schimbă PIN-ul;
   - **Date firmă și e-mail**: sunt precompletate de pe fotonromania.com (sediu social Calea Moșilor, punct de lucru Bragadiru, CUI, J). Completezi reprezentantul legal, telefon asistență, penalitate, franșiză,
     e-mail pentru rezumatul zilnic și e-mail de copie;
   - **Vehicule și documente**: VIN, nr. înmatriculare, valoare, km curent, revizie, ITP / RCA / CASCO / rovinietă;
   - **Utilizatori**: fiecare agent primește un PIN (generat automat) și, opțional, e-mail pentru notificări;
   - **Zile ocupate**: perioadele deja angajate înainte de lansare.
8. Partajează folderul „ICG Drive Test – Documente” cu colegii care trebuie să vadă contractele.

La modificarea ulterioară a `Code.gs`: Implementare → Gestionează implementările → Editează → *Versiune nouă*.
Dacă adaugi funcții noi care cer permisiuni noi, rulează o dată `setup()` din editor.

## Ce face aplicația

- **Autentificare cu PIN** pe fiecare agent; secțiunea Admin apare doar pentru rolul admin, iar serverul verifică rolul la fiecare operațiune.
- **Programare** din calendar, cu zilele ocupate marcate; nicio limită de zile, doar recomandarea din subsol
  (≤ 3 zile V9/G7, ≤ 7 zile eToano/Cavan) și avertisment la depășire.
- **Date client** (PJ/PF), cu buton **Completează din ANAF** după CUI (denumire, adresă, Reg. Com., TVA).
- **Predare** în ziua programării: șofer, km, nivel, dotări, documente, **5 poze obligatorii** (+ suplimentare), semnături →
  **contract de comodat + Anexa 1 + anexă foto**, salvat automat ca PDF în Drive (folder pe client) și trimis pe e-mail clientului
  (copie la agent și la adresa de copie).
- Contractul include **Anexa 3 – Nota de informare GDPR** pentru drive test (aliniată cu Politica de confidențialitate de pe fotonromania.com). La predare, confirmarea că clientul a primit-o este obligatorie; acordul de marketing rămâne opțional și separat.
- Predarea este **blocată** dacă ITP, RCA sau rovinieta sunt expirate.
- Predarea este **blocată** și cât timp vehiculul nu a fost returnat de clientul anterior. Dacă un test întârzie, agentul cu următoarea programare pe acel vehicul (în următoarele 7 zile) primește e-mail în dimineața respectivă, iar programarea lui apare marcată în aplicație și în rezumatul zilnic.
- **Drive test terminat**: km, nivel, daune, poze de retur, feedback, semnături → **proces-verbal de restituire** cu poze predare/retur
  puse una lângă alta, salvat în Drive și trimis pe e-mail. Kilometrajul vehiculului se actualizează automat.
- **Reamintiri**: e-mail automat clientului cu o zi înainte de retur; butoane WhatsApp și apel direct din fișa programării.
- **Rezumat zilnic la ora 8** pe e-mail: retururi întârziate, predări și retururi de azi, programări fără date client,
  documente care expiră în 30 de zile, revizii în mai puțin de 1.000 km, cereri eliberate din lista de așteptare.
- **Listă de așteptare**: când perioada cerută se eliberează (anulare, retur mai devreme, deblocare), agentul primește e-mail
  și poate programa direct.
- **Rezultate**: teste finalizate, km, grad de utilizare pe vehicul, performanță pe agent, rezultat comercial (în lucru / ofertă / vândut / pierdut), intenția clientului.

## Note

- Contractul trebuie **validat de avocatul ICG** înainte de utilizare (penalitate, franșiză, răspundere, semnătura electronică).
- Pozele se partajează în Drive „oricine are linkul poate vedea”, ca miniaturile să apară în aplicație. Dacă domeniul firmei
  blochează partajarea prin link, pozele se salvează normal, dar miniaturile nu se afișează în aplicație.
- Limită Gmail: aproximativ 100 de destinatari pe zi pentru un cont gratuit și 1.500 pentru Google Workspace.
