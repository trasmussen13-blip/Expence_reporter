# Lokal server

Denne distribution kører uden OpenAI, cloud OCR, cloud-nøgler eller GPU-krav i
selve Flask-appen. Billedanalyse udføres af Ollama på den lokale server med
modellen `llama3.2-vision`.

## 1. Installer Python-afhængigheder

Python 3.11 eller nyere anbefales:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Sæt også en stabil, lang `SESSION_SECRET` i serverens miljø, så login-sessioner
ikke bliver ugyldige ved en genstart.

## 2. Installer og start Ollama

Installer Ollama på den lokale server, og hent vision-modellen:

```bash
ollama pull llama3.2-vision
ollama serve
```

Ollama skal være tilgængelig på `http://127.0.0.1:11434`, eller også skal
`OLLAMA_HOST` sættes til den korrekte lokale adresse:

```bash
export OLLAMA_HOST=http://127.0.0.1:11434
```

## 3. Start kvitteringsappen

```bash
./run_local_server.sh
```

Åbn derefter `http://127.0.0.1:5000`.

Porten kan ændres uden kodeændringer:

```bash
PORT=8080 ./run_local_server.sh
```

## Login og flere brugere

Appen understøtter direkte OAuth-login med Google og Microsoft. Loginudbyderne
skal konfigureres på den lokale server; klient-id'er og hemmeligheder skal ikke
lægges i kildekoden eller i den eksporterede ZIP.

1. Opret en OAuth-klient hos Google og/eller Microsoft.
2. Sæt de relevante `GOOGLE_*` og/eller `MICROSOFT_*` miljøvariabler fra
   `.env.example` på serveren.
3. Registrér disse callback-adresser hos udbyderen:

   ```text
   http://127.0.0.1:5000/auth/callback/google
   http://127.0.0.1:5000/auth/callback/microsoft
   ```

4. Hvis appen tilgås via et fast hostname, sæt også
   `OAUTH_REDIRECT_BASE_URL` til samme baseadresse.

Efter login ser hver bruger kun sine egne kvitteringer, skabelonkopi, rapporter
og ZIP-eksporter. Navnet under **Rapportmaster** skrives i Excel-skabelonens
`NAME`-felt.

## Data og sikkerhed

- SQLite-databasen og uploadede billeder gemmes lokalt i `receipt_data/`.
- `receipt_data/base_template.xlsx` er den officielle Excel-skabelon.
- ZIP-eksporten indeholder den udfyldte Excel-fil og originale kvitteringsbilleder.
- Den lokale eksportpakke indeholder ikke aktive kvitteringer eller eksisterende
  SQLite-data fra udviklingsmiljøet.
- Ingen API-nøgle er nødvendig for AI-analysen.

## Ny lokal eksport

Kør fra projektroden for at oprette en ren ZIP-distribution:

```bash
python3 make_local_export.py
```

Skriptet tager kun appens kildekode, UI, afhængigheder, startvejledning og den
officielle skabelon med. Midlertidige uploads, databaser og genererede rapporter
bliver ikke pakket.