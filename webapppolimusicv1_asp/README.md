# PoliMusic

Clon de un reproductor de streaming musical construido con **ASP Clásico (VBScript) + HTML5 + CSS3 + JavaScript vanilla** sobre **SQL Server**, con una identidad visual propia (negro carbón, terracota y verde jade) en lugar de los colores de Spotify.

> Escuela Politécnica Nacional · Aplicaciones Web · `webapppolimusicv1_asp`

---

## 1. Stack y decisiones técnicas

| Capa | Tecnología |
|---|---|
| Servidor | IIS 10 (Windows) + ASP Clásico (VBScript, `CodePage 65001`) |
| Datos | SQL Server 2025 Express en `localhost,14333` · base `BDD_PoliMusic` |
| Acceso a datos | ADO / OLE DB con `ADODB.Command` y **consultas 100 % parametrizadas** |
| Front | HTML5 semántico, CSS3 (grid/flex, variables, `clamp()`), JS ES5+ sin dependencias |
| Audio | `<audio controls>` nativo por fila + barra de reproducción persistente |

Sin frameworks, sin CDN, sin build step: el proyecto se copia y funciona.

### Paleta

| Rol | Colores |
|---|---|
| Fondos | `#0D0D0D` · `#161616` · `#1F1F1F` |
| Primario (terracota) | `#B03A2E` · `#C0392B` · `#922B21` |
| Acento (verde jade) | `#00A86B` · `#148F77` · `#1ABC9C` |
| Texto | `#F2F4F4` · `#BDC3C7` |

Todas las combinaciones están declaradas como variables CSS en `assets/css/polimusic.css` (`:root`).

---

## 2. Estructura del proyecto

```
webapppolimusicv1_asp/
├─ login.asp                 Autenticación (crea Session username / id_user / user_type)
├─ logout.asp                Session.Abandon + redirección a login
├─ register.asp              Registro de usuarios normales (USER_TYPE = 0)
├─ default.asp               Pantalla principal: bienvenida, métricas, novedades
├─ songs.asp                 Listado de canciones + reproductor + CRUD admin
├─ artists.asp               Catálogo de artistas
├─ artist_form.asp           Alta / edición de artista (admin)
├─ albums.asp                Catálogo de álbumes
├─ album_detail.asp          Portada + tracklist reproducible
├─ album_form.asp            Alta / edición de álbum (admin)
├─ playlists.asp             Playlists propias, públicas y (admin) todas
├─ playlist_detail.asp       Canciones de la playlist + añadir / quitar
├─ playlist_form.asp         Alta / edición de playlist
├─ users.asp                 Administración de usuarios (solo admin)
├─ user_form.asp             Alta / edición de usuario (solo admin)
├─ delete.asp                Endpoint único de borrado (POST + CSRF + rol)
├─ api_play.asp              Incrementa TBL_SONG.PLAYS y responde JSON
├─ includes/
│   ├─ config.asp            Constantes, cadena de conexión, variables de página
│   ├─ lib_util.asp          Codificación HTML, assets, formato, flash, CSRF, validación
│   ├─ lib_db.asp            Conexión ADO con fallback de proveedor, helpers
│   ├─ lib_auth.asp          Sesión, roles, guardas RequireLogin / RequireAdmin
│   ├─ lib_ui.asp            Set de iconos SVG + componentes (badges, audio, avatar)
│   ├─ layout_top.asp        Shell: sidebar, topbar, barra de reproducción
│   ├─ layout_bottom.asp     Cierre, scripts y limpieza de conexión
│   ├─ auth_top.asp          Shell de las pantallas de autenticación
│   └─ auth_bottom.asp
├─ assets/
│   ├─ css/polimusic.css     Sistema de diseño completo (≈2.400 líneas)
│   ├─ js/polimusic.js       Reproductor, filtro, toasts, menú móvil
│   └─ img/                  favicon + carátula de reserva
├─ songFiles/                Audio referenciado por SONG_PATH
├─ albumImages/  ArtistImages/  userImages/  playlistImages/
├─ deploy/
│   ├─ deploy.ps1            Despliegue automático en IIS
│   ├─ deploy.bat            Lanzador con elevación de privilegios
│   ├─ GUIA_DESPLIEGUE.md    Guía manual paso a paso
│   ├─ validate_sql.sql      Prueba de todas las consultas (rollback)
│   ├─ check_asp.py          Verificación estática de los .asp
│   ├─ check_calls.py        Aridad de llamadas y retornos mal asignados
│   ├─ smoke_test.py         Login + recorrido de las 15 páginas por HTTP
│   ├─ functional_test.py    Roles, contenido renderizado y assets
│   ├─ update.bat            Recopia a la subcarpeta publicada
│   └─ make_artwork.py       Generador de carátulas de reserva
└─ BDD_PoliMusic.sql         Script de la base de datos (proporcionado)
```

---

## 3. Base de datos

Cadena de conexión usada (definida en `includes/config.asp`):

```
Server=localhost,14333;Database=BDD_PoliMusic;User Id=usr_polimusic;
Password=Politecnica1;Encrypt=true;TrustServerCertificate=true;
```

Traducida a OLE DB:

```
Provider=MSOLEDBSQL;Server=localhost,14333;Database=BDD_PoliMusic;User Id=usr_polimusic;
Password=Politecnica1;Encrypt=Yes;TrustServerCertificate=Yes;Connect Timeout=8;
```

`lib_db.asp` intenta en orden `MSOLEDBSQL → SQLNCLI11 → SQLNCLI → SQLOLEDB` y memoriza
el proveedor que funciona en `Application("PM_PROVIDER")`, de modo que el sitio arranca
aunque el servidor solo tenga el driver antiguo.

Tablas utilizadas: `TBL_USER`, `TBL_SONG`, `TBL_ARTIST`, `TBL_ALBUM`, `TBL_PLAYER`,
`TBL_TRACKLIST`, `TBL_PLAYLIST`, `TBL_PLAYLIST_SONG`, `TBL_CATALOG` (tipos de playlist).

---

## 4. Roles y permisos

| Acción | `USER_TYPE = 0` (usuario) | `USER_TYPE = 1` (admin) |
|---|:--:|:--:|
| Iniciar / cerrar sesión, registrarse | ✅ | ✅ |
| Reproducir canciones | ✅ | ✅ |
| Consultar álbumes y artistas | ✅ | ✅ |
| Crear / editar / borrar **sus** playlists | ✅ | ✅ |
| Ver playlists públicas ajenas | ✅ | ✅ |
| CRUD de canciones | ❌ | ✅ |
| CRUD de artistas | ❌ | ✅ |
| CRUD de álbumes | ❌ | ✅ |
| CRUD de usuarios | ❌ | ✅ |
| Editar / borrar playlists ajenas | ❌ | ✅ |

La restricción **no es solo visual**: cada página de administración llama a
`RequireAdmin()` y `delete.asp` vuelve a comprobar el rol antes de ejecutar el borrado.
Un usuario normal que escriba `users.asp` en la barra de direcciones recibe un mensaje de
acceso denegado y es redirigido.

Las contraseñas se guardan como **SHA1 hexadecimal en mayúsculas** calculado por SQL Server
(`HASHBYTES('SHA1', …)`). El login acepta tanto ese formato como las filas antiguas en texto
plano del script de siembra (`userNormal`, `userAdmin`, `Erick.Alpusig`).

---

## 5. Pantallas y flujo

1. **login.asp** → valida contra `TBL_USER`, guarda `Session("username")`, `Session("id_user")`,
   `Session("user_type")` y `Session("user_photo")`, y redirige a `default.asp`.
2. **register.asp** → crea siempre `USER_TYPE = 0` (imposible auto-promoverse).
3. **default.asp** → banner *“Welcome {usuario}”*, métricas del catálogo, novedades
   reproducibles, álbumes, artistas y playlists; botón de **Log out** visible.
4. **songs.asp** → encabezado **“Songs list”**, tabla con `ID`, `Song Name`, `Play`,
   `Reproductions` y `Actions` (esta última solo si el usuario es admin), un `<audio controls>`
   por fila y contador de reproducciones que se actualiza en vivo.
5. **logout.asp** → `Session.Abandon` y vuelta a `login.asp`.

### Detalles de interfaz

- **Barra de reproducción persistente**: cualquier `<audio>` de la página alimenta la barra
  inferior (título, artista, portada, transporte, seek y volumen). Al reproducir se hace
  `POST api_play.asp` y el contador de la fila se actualiza sin recargar.
- **Filtro instantáneo** en la tabla + búsqueda en servidor (`songs.asp?q=…`).
- **Marcado “File not deployed”** cuando `SONG_PATH` apunta a un archivo que no está en disco.
- **Responsive**: sidebar off-canvas < 980 px, tabla desplazable y columnas reducidas < 760 px.
- **Accesibilidad**: foco visible, `aria-label` en controles, `aria-live` en toasts,
  respeto a `prefers-reduced-motion`.

---

## 6. Seguridad implementada

- Todas las consultas usan parámetros ADO (`?`), nunca concatenación de entrada.
- Salida siempre codificada con `Server.HTMLEncode` (`H()`).
- Token CSRF por sesión en todos los formularios POST y en `api_play.asp`.
- `delete.asp` solo acepta POST y valida el retorno (sin *open redirect*).
- Guardas de rol en servidor (`RequireLogin`, `RequireAdmin`, `CanManagePlaylist`).
- Un administrador no puede degradarse ni borrarse a sí mismo.

---

## 7. Puesta en marcha

### 7.1 Base de datos

Ejecutar `BDD_PoliMusic.sql` en la instancia de SQL Server del puerto 14333
(crea el login `usr_polimusic`, la base y los datos de ejemplo).

### 7.2 IIS

Opción automática (recomendada):

```bat
deploy\deploy.bat        REM pide elevación y publica en C:\inetpub\wwwroot
```

Opción manual: ver **`deploy/GUIA_DESPLIEGUE.md`**.

### 7.3 Credenciales de demostración

| Usuario | Contraseña | Rol |
|---|---|---|
| `userAdmin` | `12345678` | Administrador (`USER_TYPE = 1`) |
| `Erick.Alpusig` | `12345678` | Usuario (`USER_TYPE = 0`) |
| `userNormal` | `12345678` | Usuario (`USER_TYPE = 0`) |
| `username1` | `12345678` | Usuario (contraseña almacenada como SHA1) |

---

## 8. Recursos multimedia

- `songFiles/bensound-adventure.mp3` proviene del paquete original.
- Los temas `dreams`, `energy`, `scifi` y `tomorrow` **no venían** en
  `SongFilesAndImages_resources`; para que las cinco filas reproduzcan se generaron tonos
  de relleno de 28 s con el mismo nombre. Reemplázalos por los MP3 originales cuando los tengas.
- Las carpetas `albumImages/` y `ArtistImages/` no existían en el paquete: las carátulas y
  retratos son **arte generado** con `deploy/make_artwork.py` usando la paleta del proyecto.
- Las rutas de la base (`~/songFiles/…`, `~/userImages/…`) se resuelven con `AssetUrl()`,
  que quita el prefijo `~/`, normaliza separadores y evita salir de la raíz de la aplicación.

---

## 9. Verificación realizada

Todo se comprobó contra la aplicación desplegada en `http://localhost/webapppolimusicv1_asp`.

| Prueba | Resultado |
|---|---|
| `deploy/validate_sql.sql` — 19 bloques con todas las consultas reales dentro de una transacción con `ROLLBACK` | Sin errores; la base queda intacta |
| `deploy/check_asp.py` — delimitadores, balance de bloques, comillas y variables sin `Dim` (28 archivos) | 0 hallazgos |
| `deploy/check_calls.py` — aridad de las 70 funciones y retornos mal asignados | 0 problemas |
| `deploy/smoke_test.py` — login + 15 páginas por HTTP | **15/15 OK** |
| `deploy/functional_test.py` — roles, contenido y assets | **31/31 OK** |
| Reproducción real en Chromium: play, barra inferior, contador `api_play` | `{"ok":true,"id":1,"plays":1}` |
| Capturas reales de la aplicación en ejecución | `docs/preview/` |

`diagnostico.asp` es una página de diagnóstico (texto plano) accesible **solo desde
`localhost`**; sirve para verificar la conexión y el estado del objeto `Request` durante el
despliegue.

---

## 10. Problemas frecuentes

| Síntoma | Causa / solución |
|---|---|
| Página “Cannot reach SQL Server” | Instancia apagada, puerto 14333 cerrado, login inexistente o driver OLE DB ausente. El mensaje incluye el error del proveedor. |
| Error 500 al abrir cualquier `.asp` | La característica *ASP* de IIS no está instalada, o el **Application Pool** usa .NET CLR en vez de *No Managed Code*. |
| Error 500.19 de configuración | El sitio no tiene permiso de lectura sobre la carpeta o hay un `web.config` heredado conflictivo. |
| El audio no suena | El archivo de `SONG_PATH` no está en `songFiles/`. La interfaz lo marca como *File not deployed*. |
| Se ven detalles del error ASP | IIS Manager → sitio → *ASP* → *Debugging Properties* → *Send Errors To Browser* = `True` |
