# Sletteproces – Personaleindsigt

Automatiseret oprydnings- og anonymiseringsrobot for afsluttede aktindsigtssager i
personalemapper. Robotten identificerer sager, hvis opbevaringsperiode er udløbet,
sletter de tilhørende data i SharePoint og GetOrganized (GO), markerer sletningen i
databasen og anonymiserer efterfølgende sager, der er fuldt slettet i begge systemer.

Procesnavn i OpenOrchestrator: `PersonaleIndsigtSletteprocess`

---

## Hvad robotten gør

Kørslen behandler sager i databasen `AKTINDSIGTERPERSONALEMAPPER` i tre trin:

1. **Identificér udløbne sager.** Finder afsluttede sager, hvor seneste fuldførte
   gennemløb (`last_run_complete`) ligger mere end opbevaringsgrænsen tilbage i tid
   (aktuelt 31 dage), og som endnu ikke er markeret slettet i både SharePoint og GO.
2. **Slet data.** For hver udløben sag slettes SharePoint-mappen (rekursivt, inkl.
   filer og undermapper) og GO-sagen via REST-API'et. Hver vellykket sletning
   markeres i databasen (`slettet_sharepoint = 1`, `slettet_go = 1`).
3. **Anonymisér.** Sager, der er slettet i både SharePoint og GO, kan herefter
   anonymiseres — persondata i `dbo.cases` og `dbo.case_journal_items` overskrives
   med teksten `ANONYMISERET`, og sagens status sættes til `ANONYMISERET`.

Hvis ingen sager opfylder kriterierne, afslutter robotten uden ændringer.

---

## Arkitektur og flow

```
process()
│
├─ Opret DB-forbindelse (pyodbc, Trusted_Connection, autocommit = False)
│
├─ tjek_case(cur)                  → dict af udløbne sager (nøgle = aktid)
│
└─ for hver udløben sag:
   ├─ delete_sharepoint_folder()   → sletter mappe rekursivt + marker_slettet_sharepoint()
   ├─ delete_case_go()             → sletter GO-sag + marker_slettet_go()
   └─ conn.commit()                → gemmer markeringerne for sagen
```

Sletningerne sker mod eksterne systemer (SharePoint via `office365-rest-python-client`,
GO via NTLM-autentificeret REST), mens markeringer og anonymisering skrives til
SQL Server. Da forbindelsen kører med `autocommit = False`, skal databaseændringer
committes eksplicit, før de er permanente.

---

## Afhængigheder

- `OpenOrchestrator` – proceskørsel, constants og credentials
- `pyodbc` – forbindelse til SQL Server (FDW-/sagsdatabasen)
- `office365-rest-python-client` – SharePoint-adgang (certifikatbaseret auth)
- `requests` + `requests_ntlm` – NTLM-autentificerede kald mod GO's REST-API

---

## Konfiguration i OpenOrchestrator

### Constants

| Navn | Beskrivelse |
|------|-------------|
| `SqlServer` | Servernavn for SQL Server-forbindelsen |
| `GOApiURL` | Base-URL for GetOrganized REST-API'et |
| `PersonaleIndsigtSharepointUrl` | URL til SharePoint-sitet med personaleindsigtsmapper |

### Credentials

| Navn | username | password |
|------|----------|----------|
| `GOAktApiUser` | GO API-bruger | GO API-adgangskode (NTLM) |
| `SharePointAPI` | tenant | client_id |
| `SharePointCert` | certifikat-thumbprint | sti til certifikat |

SharePoint-autentificering sker certifikatbaseret via `with_client_certificate`.



## Funktionsoversigt

| Funktion | Ansvar |
|----------|--------|
| `tjek_case(cur)` | Finder udløbne, afsluttede sager og returnerer dem som dict nøglet på `aktid` |
| `delete_sharepoint_folder(...)` | Sletter en SharePoint-mappe rekursivt og markerer sagen slettet i SharePoint |
| `delete_case_go(...)` | Sletter en GO-sag via REST-API og markerer sagen slettet i GO |
| `marker_slettet_sharepoint(cur, aktid)` | Sætter `slettet_sharepoint = 1` for en sag |
| `marker_slettet_go(cur, aktid)` | Sætter `slettet_go = 1` for en sag |
| `tjek_anonym(cur)` | Sætter `status = 'ANONYMISERET'` for sager slettet i både SharePoint og GO |
| `anonymiser_sag(cur, aktid)` | Overskriver persondata i `dbo.cases` og `dbo.case_journal_items` |
| `sharepoint_client(...)` | Opretter en autentificeret SharePoint `ClientContext` |
| `create_ntlm_session(...)` | Opretter en NTLM-autentificeret `requests.Session` til GO |
| `process(orchestrator_connection)` | Indgangspunkt der orkestrerer hele kørslen |

---



## Opbevaringsgrænse

Udløbsgrænsen beregnes i `tjek_case` som `dags dato − 31 dage`. En sag regnes som
udløben, når `last_run_complete` ligger før denne grænse. Sager uden
`last_run_complete` (NULL) springes over. Justér antallet af dage i `tjek_case`,
hvis opbevaringspolitikken ændres.
