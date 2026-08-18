# Por qué no te mandé un link altiro

Intenté exponer un link público directo desde mi sandbox (con `localtunnel`, la forma más rápida
que conozco), pero el ambiente donde corro tiene el acceso a internet restringido a
repositorios de paquetes (npm, pip) por seguridad — no puede abrir túneles hacia servicios
externos como localtunnel, ngrok o Cloudflare Tunnel. No es algo que pueda saltarme.

Lo que sí hice: probé el formulario web completo (el mismo código de esta carpeta) corriendo acá
mismo, subiendo tu excel real de julio, y dio exactamente los mismos números que ya validamos:
135 en disputa con Alas, 71 de ManualQuotes desactualizada, 97 cancelados, 5 a revisar. Es decir,
el código funciona — lo único que falta es un lugar público donde alojarlo para que tú (o
finanzas) lo abran desde un link.

## La forma más rápida y gratis de conseguir ese link: Render.com

Render tiene un plan gratuito genuino (no es un trial que se acaba), no pide tarjeta de crédito,
y soporta desplegar directo desde un `Dockerfile` — que es justo lo que hay en esta carpeta. Son
~10 minutos, casi todo a clicks:

### Paso 1 — Subir esta carpeta a GitHub
Render despliega desde un repositorio de Git. Si no tienes GitHub, crea una cuenta gratis en
github.com (2 minutos), y sube esta carpeta (`webdemo/`) como un repositorio nuevo:

- Vía la web de GitHub: **New repository** → arrastra estos archivos con el botón "uploading an
  existing file" → Commit.
- O si tienes Git instalado: `git init && git add . && git commit -m "demo" && git remote add
  origin <url-de-tu-repo> && git push -u origin main`.

### Paso 2 — Crear el servicio en Render
1. Entra a render.com y crea una cuenta gratis (puedes usar tu login de GitHub directamente).
2. **New +** → **Web Service**.
3. Conecta el repositorio que acabas de crear.
4. Render va a detectar el `Dockerfile` automáticamente — déjalo así (no cambies el "Build
   Command" ni el "Start Command", ya están definidos en el Dockerfile).
5. En **Instance Type**, elige **Free**.
6. Click **Create Web Service**.

### Paso 3 — Esperar el build y probar
Render va a construir la imagen (2-4 minutos la primera vez) y te va a dar una URL del estilo
`https://tu-servicio.onrender.com`. Ábrela, sube el excel de julio, y deberías ver la misma
pantalla de resultados que probamos acá.

**Nota:** en el plan gratuito, el servicio "se duerme" después de 15 minutos sin uso, y la
primera visita después de eso tarda ~1 minuto en despertar. Para una demo de prueba no es
problema; si más adelante lo quieren usar en serio todo el tiempo, ahí sí conviene un plan pago
(o volver a la opción de correo con n8n que ya armamos, que no tiene este problema).

## Archivos de esta carpeta

- `app.py`: la aplicación web (formulario de subida + resultado), ya probada.
- `Dockerfile`, `requirements.txt`: para que Render (o cualquier otro proveedor con soporte
  Docker) sepa cómo construir y correr la app.
- `scripts/reconciliar_alas.py`, `assets/rates_table.json`: el mismo motor de conciliación ya
  validado.

Esto es un ambiente de **prueba**, no reemplaza el flujo de correo/n8n que armamos antes para el
proceso mensual real — es solo para que puedas ver la interfaz funcionando con un link de verdad.
