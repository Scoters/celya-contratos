(function() {
    let scriptUrl = "https://script.google.com/macros/s/AKfycbx_DHa9zNyGFZrfr6dS4NgoW9xbArfpskejYi_UP_o3NDP6ovIx7m_Aj_KIp0DGNAHJ/exec";
    let apiToken = "";
    
    // Cargar variables de configuración desde el archivo .env en el servidor web en tiempo de ejecución
    const xhr = new XMLHttpRequest();
    xhr.open('GET', '/.env', false); // Carga síncrona
    try {
        xhr.send(null);
        if (xhr.status === 200) {
            const lines = xhr.responseText.split('\n');
            lines.forEach(line => {
                const parts = line.split('=');
                if (parts.length >= 2) {
                    const key = parts[0].trim();
                    const val = parts.slice(1).join('=').trim();
                    if (key === 'GOOGLE_SCRIPT_URL') {
                        scriptUrl = val;
                    } else if (key === 'API_TOKEN') {
                        apiToken = val;
                    }
                }
            });
        }
    } catch(e) {
        console.warn("No se pudo cargar el archivo .env localmente, usando fallback de Google Script URL.", e);
    }
    
    window.CONFIG = {
        GOOGLE_SCRIPT_URL: scriptUrl,
        API_TOKEN: apiToken,
        PRELOADED_CATALOG: []
    };
})();
