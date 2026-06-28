# Arquitectura de Despliegue - CelYa

Este documento registra de forma permanente la configuración de infraestructura del proyecto para guiar a los agentes de IA y desarrolladores en futuras modificaciones.

## 🌐 1. Páginas de Frontend en GitHub Pages (Público Principal)
* **Hosting:** GitHub Pages.
* **Archivos:** `index.html`, `dashboard.html`, `formulario.html`, `config.js` y recursos estáticos.
* **Dominio:** [estrenacelya.com](https://estrenacelya.com) gestionado mediante Cloudflare.
* **🚨 Regla de Despliegue:** **NUNCA** hacer despliegues por SCP o SSH de estos archivos al VPS. Para publicar cambios, únicamente se debe realizar un `git push` a la rama `main` de este repositorio en GitHub. Pages se actualiza en automático tras 1-2 minutos.

## 🖥️ 2. Portal de Administración y Clientes en el VPS (Subdominio Portal)
* **Hosting:** VPS de Racknerd (`107.175.122.33`).
* **Archivos:** `portal.html` (y recursos relacionados).
* **Dominio:** [portal.estrenacelya.com](https://portal.estrenacelya.com)
* **🚨 Regla de Despliegue:** Los cambios en `portal.html` **SÍ** requieren deploy al VPS a través de SCP/SSH para actualizar la web de producción, además de guardarse en GitHub. Las rutas de destino en el VPS son:
  * Respaldos en `/root/`
  * Nginx Proxy en `/root/nginx-proxy/data/html/`
  * Volumen Docker en `/var/lib/docker/volumes/nginx-proxy_data/_data/html/`
  * Copia activa al contenedor: `docker cp /root/portal.html nginx-proxy_app_1:/var/www/html/portal.html`

## ⚙️ 3. Servidor de Backend y Automatizaciones
* **Hosting:** VPS de Racknerd (`107.175.122.33`).
* **Servicios:** Automatizaciones en **n8n**, base de datos SQLite de n8n, plantillas de contrato PDF.
* **Dominio:** `n8n.estrenacelya.com` y endpoints de API internos.
* **Regla de Despliegue:** Solo interactuar mediante SSH/Docker cuando se modifiquen flujos de n8n, bases de datos o configuraciones del servidor de backend.
