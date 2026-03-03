# 🔐 Gestor de Certificados Digitales

Herramienta para Windows que permite **activar y desactivar certificados digitales** de forma rápida, sin tener que entrar al gestor de certificados de Windows (`certmgr.msc`).

Especialmente útil para **gestores, administradores de fincas, asesores y cualquier profesional** que trabaje con múltiples certificados digitales de diferentes clientes o comunidades de propietarios.

---

## ¿Qué problema resuelve?

Cuando tienes muchos certificados digitales instalados en Windows, al acceder a una sede electrónica el navegador te muestra **todos los certificados a la vez**, sin posibilidad de filtrar. Esto obliga a identificar manualmente el certificado correcto cada vez.

Esta herramienta permite **ocultar temporalmente** los certificados que no necesitas, de forma que el navegador solo muestre el que quieres usar en cada momento.

---

## ¿Cómo funciona?

Windows tiene dos almacenes de certificados relevantes:

- **Personal** → los certificados aquí son **visibles** en el selector del navegador
- **Otras personas** → los certificados aquí están **ocultos** para el navegador

La herramienta mueve certificados entre estos dos almacenes con un solo clic.

> 💡 **Consejo:** Para sacar el máximo partido, edita el "Nombre descriptivo" de cada certificado en `certmgr.msc` y escribe ahí el nombre de la empresa y el CIF. El buscador de esta herramienta busca exactamente en ese campo.

---

## Características

- 🔍 **Buscador en tiempo real** por nombre de empresa o CIF
- ✅ **Activar** certificado con un clic (lo mueve a Personal, visible en el navegador)
- ⏸ **Desactivar** certificado con un clic (lo mueve a Otras personas, oculto en el navegador)
- 🔴 Certificados **caducados** marcados visualmente en rojo
- Muestra la **fecha de vencimiento** de cada certificado
- Filtro opcional para mostrar **solo los activos**
- Confirmación antes de cada acción para evitar errores

---

## Requisitos

- Windows 10 o Windows 11
- PowerShell 5.1 (incluido por defecto en Windows 10/11)
- Los certificados deben estar instalados como **archivo de software** (`.p12` / `.pfx`). No funciona con tarjetas criptográficas físicas ni DNIe.

---

## Instalación

No requiere instalación. Descarga los dos archivos y colócalos en la misma carpeta:

- `GestorCertificados.ps1`
- `Abrir_GestorCertificados.bat`

---

## Uso

1. Haz doble clic en **`Abrir_GestorCertificados.bat`**
2. Busca el certificado por nombre o CIF
3. Selecciónalo y pulsa **ACTIVAR** o **DESACTIVAR**
4. Abre (o recarga) el navegador

> ⚠️ Si el navegador ya estaba abierto, ciérralo y vuelve a abrirlo para que detecte el cambio.

---

## Aviso de SmartScreen al ejecutar

Al descargar el `.bat` desde internet, Windows mostrará este aviso la primera vez:

> *"Windows protegió su PC — SmartScreen de Microsoft Defender impidió el inicio de una aplicación desconocida"*

Es normal para cualquier archivo descargado que no tenga firma de código comercial. Para ejecutarlo:

1. Haz clic en **"Más información"**
2. Haz clic en **"Ejecutar de todas formas"**

Solo ocurre una vez. A partir de ahí el archivo se ejecuta directamente.

> El código fuente está disponible en este repositorio para que puedas revisar exactamente qué hace antes de ejecutarlo.

---

## ¿Es seguro?

Sí. El código fuente está disponible aquí para que cualquiera pueda revisarlo. La herramienta únicamente mueve certificados entre almacenes del propio usuario en Windows, sin enviar ningún dato a internet ni modificar ningún certificado.

---

## Licencia

MIT — Uso libre. Puedes usar, copiar, modificar y distribuir este proyecto sin restricciones, incluso en proyectos comerciales. Solo se pide mantener el aviso de autoría original.
