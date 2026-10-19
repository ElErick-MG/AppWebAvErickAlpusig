# PoliMusic — notas de proyecto

- **Stack**: ASP Clásico (VBScript, `CodePage 65001`, UTF-8 sin BOM) + ADO + SQL Server.
  Sin frameworks ni build step; el sitio se copia y funciona.
- **BDD**: `BDD_PoliMusic` en `localhost,14333`, usuario `usr_polimusic` / `Politecnica1`,
  `Encrypt=Yes;TrustServerCertificate=Yes`. Driver preferido `MSOLEDBSQL`.
- **Roles**: `TBL_USER.USER_TYPE` 0 = usuario, 1 = admin. Toda restricción se valida en
  servidor (`RequireAdmin`, `CanManagePlaylist`), no solo ocultando botones.
- **Rutas de assets**: la BDD guarda `~/songFiles/...`, `~/userImages/...`; se resuelven con
  `AssetUrl()` (quita `~/`, normaliza separadores, evita `../`).
- **Convención**: cada página declara sus variables con prefijo (`sg`, `hm`, `ar`, `pl`, `pd`,
  `pf`, `us`, `uf`, `sf`, `af`, `bf`) para no colisionar en el scope compartido de los includes.
- **Idioma**: interfaz en inglés (requisito de la rúbrica: "Songs list", "Welcome …");
  documentación y comentarios en español.
