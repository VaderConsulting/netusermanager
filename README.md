# netusermanager

TSAI (Telefonica Servicios Avanzados de Informacion) VB6 ActiveX DLL `TSAI_NT.dll` (project TSAI_NT, class `NTUsers`) for WinNT 4.0 network user administration. Exposes add/delete/rename user, change password, bulk add from a text file, and local-group membership helpers over NetAPI32, plus Advapi32 ACL helpers for folder permissions. VBP is supplied as `Tsai_nt.vbp.example` (restore to `Tsai_nt.vbp` before opening). Copyright note in the VBP points to 1998.

**Source last updated:** 1998-06-01 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** ActiveX DLL

_Note: original OneDrive LastWriteTime values were wiped to 2026-08-27 by a zip transfer; date above uses best available evidence (headers/copyright where helpful)._

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `TSAI_NT` (`Tsai_nt.vbp.example`) | VB6 | ActiveX DLL | WinNT 4.0 user add/delete/rename, password, local groups, folder ACLs |

## How to open

Rename `Tsai_nt.vbp.example` to `Tsai_nt.vbp`, then open in Visual Basic 6.0 IDE:
- `Tsai_nt.vbp`

## Requirements

- Visual Basic 6.0 IDE (originally built against VB5 extensibility references)
- Windows NT 4.0-era NetAPI32 / Advapi32 privileges for user and ACL calls
- Help file path in the VBP may need adjusting (`TSAI_NTUSERS.HLP`)

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/Old/netusermanager`.
Company names in project files: TSAI - Telefonica, Servicios Avanzados de Informacion.
Third-party attribution: TSAI - Telefonica, Servicios Avanzados de Informacion. See `THIRD_PARTY_NOTICES.md`.

## License

Third-party code remains under its original terms (or none, where none were supplied). See `THIRD_PARTY_NOTICES.md`. Do not treat this tree as VaderConsulting original MIT-licensed work.
