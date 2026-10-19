# Guía de despliegue en IIS — PoliMusic

Dos caminos: **automático** (script incluido) o **manual** (IIS Manager). Ambos dejan la
aplicación publicada en `C:\inetpub\wwwroot`.

---

## 0. Requisitos

- Windows con IIS y la característica **ASP** (`IIS-ASP`) instalada.
- SQL Server escuchando en `localhost,14333` con la base `BDD_PoliMusic` y el login
  `usr_polimusic` creados (los crea `BDD_PoliMusic.sql`).
- Driver **Microsoft OLE DB Driver for SQL Server** (`MSOLEDBSQL`). Si no está, la
  aplicación cae automáticamente a `SQLNCLI11` / `SQLOLEDB`.

Comprobar la característica ASP:

```powershell
Get-WindowsOptionalFeature -Online -FeatureName IIS-ASP
```

Si está `Disabled`:

```powershell
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All
```

---

## 1. Despliegue automático (recomendado)

1. Clic derecho en `deploy\deploy.bat` → **Ejecutar como administrador**
   (el propio `.bat` pide elevación si hace falta).
2. El script:
   - verifica/activa `IIS-WebServer`, `IIS-StaticContent`, `IIS-DefaultDocument`, `IIS-ASP`;
   - copia el proyecto a `C:\inetpub\wwwroot` excluyendo `deploy/`,
     `SongFilesAndImages_resources/`, `.sql`, `.py`, `.md`, `.ps1` y `.bat`;
   - concede lectura a `IIS_IUSRS`;
   - crea el **Application Pool** `PoliMusic` en modo *No Managed Code*, 64 bits;
   - lo asigna al sitio `Default Web Site` y recicla el pool;
   - hace una petición de prueba a `http://localhost/login.asp`.

Si quieres publicar en otra ruta o sitio:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy.ps1 `
    -Destination "C:\inetpub\wwwroot\PoliMusic" `
    -SiteName "PoliMusic" -AppPool "PoliMusicPool"
```

> Nota: al copiar con `robocopy /E` no se borra el contenido previo de `wwwroot`
> (`iisstart.htm` sigue ahí, pero `default.asp` tiene prioridad en los documentos
> predeterminados de IIS).

---

## 2. Despliegue manual (IIS Manager)

1. **Copiar archivos**
   Copia el contenido del proyecto a `C:\inetpub\wwwroot` **sin** las carpetas
   `deploy` ni `SongFilesAndImages_resources`.
   Deben quedar en la raíz: `login.asp`, `default.asp`, `songs.asp`, `includes\`, `assets\`,
   `songFiles\`, `albumImages\`, `ArtistImages\`, `userImages\`, `playlistImages\`.

2. **Application Pool**
   - IIS Manager → *Application Pools* → **Add Application Pool**
   - Nombre: `PoliMusic`
   - *.NET CLR Version*: **No Managed Code**
   - *Managed pipeline mode*: Integrated
   - Aceptar y luego **Recycle**.

3. **Asignar el pool al sitio**
   - IIS Manager → *Sites* → `Default Web Site` → clic derecho → **Advanced Settings**
   - *Application Pool*: `PoliMusic`

4. **Permisos**
   ```powershell
   icacls "C:\inetpub\wwwroot" /grant "IIS_IUSRS:(OI)(CI)RX" /T
   ```

5. **Verificar ASP**
   - Selecciona el sitio → **ASP**
   - *Compilation* → *Debugging Properties* → **Send Errors To Browser** = `True`
     (útil durante la revisión; desactívalo al terminar)

6. **Probar**
   Abre `http://localhost/login.asp`.
   - `userAdmin` / `12345678` → administrador
   - `Erick.Alpusig` / `12345678` → usuario normal

---

## 3. Comprobaciones posteriores

```powershell
# ¿El sitio responde?
Invoke-WebRequest http://localhost/login.asp -UseBasicParsing | Select-Object StatusCode

# ¿El pool está en No Managed Code?
Get-ItemProperty IIS:\AppPools\PoliMusic | Select-Object managedRuntimeVersion

# ¿La base responde?
sqlcmd -S "localhost,14333" -U usr_polimusic -P "Politecnica1" -d BDD_PoliMusic -C -Q "SELECT COUNT(*) FROM TBL_SONG"
```

---

## 4. Solución de problemas

| Error | Qué revisar |
|---|---|
| `500.21 – Handler "ASPClassic" has a bad module` | La característica ASP no está instalada → `Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All` |
| `500 – Cannot reach SQL Server` (página propia) | Instancia/puerto, login `usr_polimusic`, driver OLE DB |
| `500.19` | Permisos NTFS sobre la carpeta o `web.config` heredado del sitio padre |
| El login devuelve “Invalid credentials” con datos correctos | Revisa que `Encrypt`/`TrustServerCertificate` coincidan con la configuración del servidor |
| Los MP3 no suenan | `songFiles/` debe contener los archivos que indica `TBL_SONG.SONG_PATH` |
| Imágenes rotas | Deben existir `albumImages/`, `ArtistImages/`, `userImages/`, `playlistImages/` junto a los `.asp` |

Para ver el error real de ASP en el navegador:

```
IIS Manager → sitio → ASP → Debugging Properties → Send Errors To Browser = True
IIS Manager → sitio → Error Pages → Edit Feature Settings → Detailed errors
```
