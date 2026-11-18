# Inventario Web Independiente

Aplicación web de inventario (PWA) para ordenador y móvil. **No instala ejecutables** ni utiliza los archivos de Inventario Gambirasi.

## Acceso a la aplicación

Una vez que GitHub Pages esté habilitado y publicado, la dirección será:

**https://roserocristhian2211.github.io/inventario-web/**

El código de este repositorio es público; los productos, salidas, fotografías y otros datos de inventario **no se guardan en GitHub**.

## Activar GitHub Pages (una sola vez)

1. Ve a [Settings → Pages](https://github.com/roserocristhian2211/inventario-web/settings/pages).
2. En **Build and deployment → Source**, selecciona **GitHub Actions**.
3. Guarda los cambios si GitHub lo solicita.
4. Abre [Actions](https://github.com/roserocristhian2211/inventario-web/actions/workflows/pages.yml) y ejecuta o vuelve a ejecutar el flujo **Publicar Inventario Web** (Run workflow / Re-run jobs).
5. Confirma que el despliegue sea satisfactorio antes de utilizar la dirección web.

## Funciones

- Productos y unidades, materiales directos e indirectos; búsqueda, edición y eliminación lógica.
- Entradas y salidas con guías/documentos, traspasos por piscina, conteos y ajustes.
- Fotos de productos y alertas de agotamiento y existencia mínima.
- Funcionamiento sin internet tras la primera visita y almacenamiento en **IndexedDB** del navegador.
- Copia de seguridad JSON, exportación de productos CSV e importación CSV.
- Sincronización opcional con una **cuenta diferente de OneDrive**, usando Microsoft Graph y OAuth 2.0 + PKCE.

## Conectar otro OneDrive

La sincronización no se activa por defecto. Debes autorizar una cuenta Microsoft distinta y usar una carpeta nueva.

1. Entra en https://entra.microsoft.com/ con la cuenta que autorizará la aplicación, o pide ayuda al administrador de Microsoft 365 si no puedes registrar aplicaciones.
2. Registra una aplicación Microsoft Entra, con plataforma **Aplicación de página única (SPA)** y dirección de redirección exactamente igual a **https://roserocristhian2211.github.io/inventario-web/** (barra final incluida).
3. Solicita permisos delegados de Microsoft Graph `User.Read` y `Files.ReadWrite`. La app usa también `openid`, `profile`, `offline_access`.
4. Abre la página web del inventario → **OneDrive**. Escribe el **identificador de aplicación (Client ID)**, elige el tipo de cuenta y escribe el nombre de una carpeta **nueva y exclusiva**. No escribas secretos de cliente.
5. Selecciona **Conectar Microsoft**, autoriza la cuenta deseada y después **Sincronizar ahora**.

El código está diseñado para crear la carpeta autorizada, un archivo `identidad-inventario-web.json`, subcarpetas `eventos` y `fotos`. La base local queda vinculada al identificador del OneDrive usado por primera vez; cambiar a otra cuenta requiere un perfil limpio del navegador para evitar mezclar datos.

**Advertencia:** `Files.ReadWrite` es un permiso delegado amplio; no es un permiso limitado por el servidor a la carpeta de esta app. El propio código restringe sus operaciones a la carpeta exclusiva. La sincronización de registros independientes no sustituye una base de datos transaccional cuando varias personas registran salidas simultáneas.

## Uso sin conexión

Primero entra con conexión para instalar el Service Worker. La app guarda los datos en el navegador y permite trabajar offline. Una vez vuelva internet, conecta OneDrive y pulsa Sincronizar. Los datos no sincronizados pueden perderse si borras datos del navegador; exporta respaldos periódicamente.

## Estructura

- `index.html`: interfaz en español.
- `styles.css`: interfaz responsive para computadoras y móviles.
- `app.js`: gestión de inventario, IndexedDB, OAuth, fotos y sincronización Microsoft Graph.
- `manifest.webmanifest`, `sw.js`, `icon.svg`: PWA y offline.
- `.github/workflows/pages.yml`: despliegue automático con GitHub Actions.

**No se importaron ni modificaron datos de Inventario Gambirasi.** La conexión real a OneDrive aún requiere autorización y pruebas con el destino elegido.
