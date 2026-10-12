# Receptor temporal de leads de JBD

Este Apps Script se instala dentro de la hoja de Google Workspace `Leads JBD Web`.

## Publicación

1. Crear la hoja `Leads JBD Web` usando la cuenta empresarial que enviará los correos.
2. Abrir **Extensiones → Apps Script**.
3. Reemplazar el contenido de `Code.gs` por el archivo de esta carpeta y guardar.
4. En Apps Script, abrir **Configuración del proyecto** y elegir `America/Lima` como zona horaria.
5. Pulsar **Implementar → Nueva implementación**.
6. Seleccionar **Aplicación web**.
7. En **Ejecutar como**, seleccionar `Yo`.
8. En **Quién tiene acceso**, seleccionar `Cualquier usuario`.
9. Autorizar el acceso a la hoja y el envío de correo.
10. Copiar la URL que termina en `/exec`. No usar la URL `/dev`.
11. Pegar esa URL en `data-lead-endpoint` del formulario en `src/pages/index.astro`.
12. Compilar, probar y publicar la web.

Cada lead consume dos destinatarios de la cuota diaria: el aviso interno y el acuse al cliente.
