# Nordnet Excel-projekt (Power Query + VBA)

Dette projekt giver en robust skabelon til at hente kontodata fra Nordnet endpointet:

`GET https://public.nordnet.se/api/2/accounts/{accid}/info`

Målet er en Excel-fil, der kan opdatere data ved åbning og via **Refresh**.

## Indhold i projektet

- `powerquery/Nordnet_Account_Info.pq` – M-kode (Power Query) med:
  - parametre for `accid`, base-URL, timeout og auth/session-header-værdier
  - API-kald via `Web.Contents`
  - robust fejlhåndtering
  - parsing af JSON til tabel
  - felt for sidste opdateringstidspunkt (`LastRefreshUtc`)
- `vba/ThisWorkbook.bas` – VBA-kode til auto-refresh ved åbning
- `vba/modNordnetRefresh.bas` – VBA-makro til manuel refresh-knap

---

## Opsætning i Excel (anbefalet)

> Nedenfor bruges danske menunavne i fri oversættelse; i engelsk Excel hedder menuen typisk **Data > Get Data > Launch Power Query Editor**.

### 1) Opret workbook og ark

1. Opret en ny Excel-fil (`.xlsm` hvis du vil bruge VBA).
2. Omdøb første ark til **Nordnet Konto**.

### 2) Opret parametre i Power Query

1. Gå til **Data → Hent data → Start Power Query Editor**.
2. Opret disse parametre (Manage Parameters):
   - `pAccId` (Text) – fx `12345678`
   - `pBaseUrl` (Text) – `https://public.nordnet.se`
   - `pApiVersionPath` (Text) – `/api/2`
   - `pAuthHeaderName` (Text) – fx `Authorization`
   - `pAuthHeaderValue` (Text) – fx `Bearer <token>`
   - `pSessionHeaderName` (Text) – fx `Cookie` (kan være tom)
   - `pSessionHeaderValue` (Text) – fx `JSESSIONID=<value>` (kan være tom)
   - `pTimeoutSeconds` (Number) – fx `30`

> Hvis API'et ikke kræver visse headers, kan navn/værdi stå tom.


### Her skifter du account id og auth/session

I **Manage Parameters** er det her, du skifter til rigtige værdier:

- `pAccId` = dit rigtige account id
- `pAuthHeaderName` = fx `Authorization`, hvis Nordnet kræver den header
- `pAuthHeaderValue` = fx `Bearer DIN_TOKEN`
- `pSessionHeaderName` = fx `Cookie`
- `pSessionHeaderValue` = fx `JSESSIONID=...`

Hvis endpointet ikke kræver auth/session, kan disse parametre være tomme (navn + værdi).

### Sådan gør du i praksis

1. Åbn Excel og gå til **Data → Hent data → Start Power Query Editor**.
2. Åbn **Manage Parameters** og find/opret parametrene fra listen ovenfor.
3. Erstat eksempelværdierne med dine egne rigtige værdier (især `pAccId`).
4. Opdater queryen med **Refresh Preview** i Power Query og derefter **Close & Load**.

### 3) Indsæt M-kode

1. Opret en **Blank Query**.
2. Åbn **Advanced Editor**.
3. Indsæt indholdet fra `powerquery/Nordnet_Account_Info.pq`.
4. Gem query som `NordnetAccountInfo`.

### 4) Load til arket

1. **Close & Load To...**
2. Vælg **Table** i eksisterende ark **Nordnet Konto** (fx celle `A1`).
3. Tabellen vil indeholde felter fra API'et + `LastRefreshUtc`.

### 5) Seneste opdatering i separat felt (valgfrit)

Hvis du vil vise opdateringstid i en dedikeret celle:

- Lav en ny query, der refererer til `NordnetAccountInfo` og returnerer kun `LastRefreshUtc`.
- Load den til fx celle `H2` og navngiv feltet “Sidste opdatering (UTC)”.

### 6) Auto-refresh ved åbning

1. Gem som `.xlsm`.
2. Importér `vba/ThisWorkbook.bas` og `vba/modNordnetRefresh.bas` i VBA-editoren.
3. I Excel: højreklik query → **Properties**:
   - Sæt flueben i **Refresh data when opening the file**.
   - (Valgfrit) **Refresh every X minutes**.

### 7) Manuel refresh-knap

1. Indsæt en Form Control-knap i arket.
2. Tildel makroen `RefreshNordnetData`.

---

## Sikkerhed og best practice

- Gem aldrig rigtige tokens i versionskontrol.
- Brug parametre til credentials; helst via sikker hemmelighedshåndtering internt.
- Undgå at logge rå auth-data i fejlbokse.
- Hold endpoint-path i parameter (`pApiVersionPath`) så opgradering er nem.
- Bevar timeout og fejlhåndtering for stabil drift.

---

## Fejlsøgning

- **401/403**: Tjek auth/session-headere.
- **404**: Tjek `accid` og endpoint-version.
- **Formula.Firewall**: Sæt passende privacy-niveau eller kombiner queries korrekt.
- **Ingen data i tabel**: Bekræft at respons er JSON-record og at feltnavne findes.

