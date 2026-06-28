# Guía de Actualización y Mantenimiento de CelYa

Esta guía detalla las reglas de oro y pasos a seguir cada vez que realices cambios o actualizaciones en los diferentes componentes del ecosistema (Google Sheets, GitHub Pages y n8n).

---

## 1. Cambios en Google Apps Script (`script.gs`)
Siempre que modifiques el código de `script.gs`:

1. **Guardar cambios:** Asegúrate de guardar los cambios en el editor de Apps Script.
2. **Crear nueva implementación (Nueva versión):** 
   * Haz clic en **Implementar (Deploy)** -> **Administrar implementaciones (Manage deployments)**.
   * Haz clic en el icono de **lápiz (Editar)** de tu implementación activa.
   * En la sección "Versión", selecciona **Nueva versión (New version)**.
   * Haz clic en **Implementar (Deploy)**.
   > [!IMPORTANT]
   > Si no creas una "Nueva versión", Google seguirá ejecutando el código viejo de la Web App, ignorando los cambios que hayas hecho en el archivo.
3. **Validar URL:** Si por alguna razón tuviste que crear una implementación desde cero y el ID de la URL cambió, recuerda actualizarla de inmediato en `config.js` y en los flujos de n8n en el VPS.

---

## 2. Cambios en el Frontend (Páginas Web: `formulario.html`, `dashboard.html`, `config.js`)
Siempre que modifiques o actualices algún archivo de la web:

1. **Subir cambios a GitHub (Push):**
   Abre una terminal en la carpeta del formulario y ejecuta:
   ```bash
   git add .
   git commit -m "Descripción de lo que actualizaste"
   git push origin main
   ```
2. **Forzar recarga en el navegador (Evitar Caché):**
   GitHub Pages puede tardar hasta 1 minuto en desplegar los cambios. Cuando entres a la web en producción, **limpia la caché del navegador**:
   * En Windows/Linux: Presiona `Ctrl + F5` o abre una ventana de incógnito.
   * En Mac: Presiona `Cmd + Shift + R`.
   > [!NOTE]
   > Si no limpias la caché, el navegador cargará la versión guardada en tu disco y verás errores de red o pantallas antiguas.

---

## 3. Cambios en los flujos de n8n o Servidor (VPS)
Siempre que modifiques un workflow de n8n o requieras actualizar plantillas:

1. **Plantilla HTML del Contrato:** Si cambias el diseño de la plantilla HTML del contrato, debes reemplazar el archivo en el VPS en la ruta:
   `/root/.n8n/templates/Contrato_y_Pagare_CelYa.html`
2. **Re-importar workflows:** Si actualizas los archivos JSON de los flujos de n8n por consola en el VPS, recuerda que n8n los importa desactivados por defecto. Debes correr la query de activación en SQLite y reiniciar n8n:
   ```bash
   # Comando para activar y reiniciar (usando nuestro script automatizado en el VPS)
   python3 /root/activate.py
   docker restart n8n_n8n_1
   ```
3. **Publicación desde n8n Editor:** Si creas un nuevo webhook o trigger de tiempo directo en la web de n8n, recuerda hacer clic en el botón **`Publish`** en la esquina superior derecha para habilitarlo en producción.
4. **Zona Horaria del Servidor n8n (Importante para Triggers y Resúmenes):**
   * El archivo de configuración de Docker Compose de n8n está en `/root/n8n/docker-compose.yaml` en el VPS.
   * Para asegurar que los triggers programados (ej. 8:00 AM) se ejecuten en la hora de México, el contenedor de n8n debe tener las variables de entorno de zona horaria en su sección `environment`:
     ```yaml
     - GENERIC_TIMEZONE=America/Mexico_City
     - TZ=America/Mexico_City
     ```
   * Si necesitas reiniciar n8n para aplicar cambios del Compose, usa `docker-compose` (con guion medio):
     ```bash
     cd /root/n8n
     docker-compose down
     docker-compose up -d
     ```
   * Para verificar la hora interna en n8n: `docker exec n8n_n8n_1 date`.

---

## 4. Checklist Rápido de Validación tras un cambio
Para confirmar que todo funciona correctamente tras una actualización grande:
- [ ] ¿El formulario abre sin errores de red en incógnito?
- [ ] ¿El catálogo web muestra los precios actualizados?
- [ ] ¿Al registrar un cliente de prueba, el PDF se genera y se guarda en Google Drive?
- [ ] ¿Te llega el correo de confirmación de Gmail con el PDF adjunto?
- [ ] ¿Los reportes diarios de Telegram se enviaron correctamente a las 8:00 AM?
