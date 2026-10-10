# La Tiendita del Fisio

Tienda web de una sola página: el cliente ve todo el catálogo, arma su pedido y lo envía por WhatsApp.

- `index.html`: toda la web (HTML, CSS y JS en un archivo, sin dependencias).
- `datos.json`: copia del catálogo para la primera carga y como respaldo si Google no responde.
- `img/p/`: fotos de productos (WebP 480 px).
- `apps-script/Code.gs`: código de la Google Sheet (catálogo, tasa BCV, pedidos).

## Cómo funciona

1. La dueña edita la Google Sheet (pestañas Productos, Tasa, Pedidos y Ajustes).
2. La web lee la hoja a través de la aplicación web de Apps Script (`TDF.API` en `index.html`).
3. Al enviar un pedido, la web lo registra en la pestaña Pedidos y abre WhatsApp con el mensaje escrito.

## Tasa BCV

`actualizarTasa()` corre cada hora: lee bcv.org.ve, y si falla usa ve.dolarapi.com. Guarda un historial por
"fecha valor" y la web usa la última cuya fecha ya llegó (el BCV publica por adelantado la del próximo día hábil).
Una tasa manual en la pestaña Tasa manda sobre la automática.

## Configuración

En `index.html`, bloque `window.TDF`: `API` (URL `/exec` de Apps Script), `PIXEL` (ID del píxel de Meta) y `WHATSAPP`.
