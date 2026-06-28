import os

def optimizar_nginx():
    path = "/root/nginx-proxy/data/nginx/proxy_host/4.conf"
    if not os.path.exists(path):
        print(f"Error: No se encontro el archivo en {path}")
        return
        
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    # Si ya se optimizó previamente, removemos el bloque anterior para inyectar el nuevo con CORS
    if "gzip_types" in content:
        print("Restaurando bloque de Nginx previo para actualizar con politicas CORS...")
        # Limpiamos el bloque anterior
        import re
        content = re.sub(
            r'# Compresion Gzip para portal\.estrenacelya\.com\s+gzip on;.*?# Cache agresiva para recursos estaticos \(Add Expires headers\)\s+location ~\* \\\.(\(js\|css\|png\|jpg\|jpeg\|gif\|ico\|svg\|webp\|woff\|woff2\|ttf\|otf\|json\)\|\(js\|css\|png\|jpg\|jpeg\|gif\|ico\|svg\|webp\|woff\|woff2\|ttf\|otf\|json\))\$ \{.*?\}',
            "",
            content,
            flags=re.DOTALL
        )
        # Quitar duplicados si hay
        content = content.replace('  gzip on;\n  gzip_types text/plain text/css application/json application/javascript text/xml application/xml application/xml+rss text/javascript image/svg+xml;\n  gzip_proxied any;\n  gzip_min_length 1000;\n  gzip_comp_level 6;', '')
        content = content.replace('  location ~* \.(js|css|png|jpg|jpeg|gif|ico|svg|webp|woff|woff2|ttf|otf|json)$ {\n    root /data/html;\n    expires 1y;\n    add_header Cache-Control "public, no-transform";\n    log_not_found off;\n    access_log off;\n  }', '')

    # Bloque de reemplazo optimizado con Gzip, Expires y CORS
    target = """  location / {
    root /data/html;
    index portal.html index.html;
    try_files $uri $uri/ /portal.html;
  }"""

    replacement = """  # Compresion Gzip para portal.estrenacelya.com
  gzip on;
  gzip_types text/plain text/css application/json application/javascript text/xml application/xml application/xml+rss text/javascript image/svg+xml;
  gzip_proxied any;
  gzip_min_length 1000;
  gzip_comp_level 6;

  # Cache agresiva y politicas CORS para recursos estaticos (Add Expires headers & CORS)
  location ~* \.(js|css|png|jpg|jpeg|gif|ico|svg|webp|woff|woff2|ttf|otf|json)$ {
    root /data/html;
    expires 1y;
    add_header Cache-Control "public, no-transform";
    add_header Access-Control-Allow-Origin "*";
    add_header Access-Control-Allow-Methods "GET, POST, OPTIONS";
    add_header Access-Control-Allow-Headers "*";
    log_not_found off;
    access_log off;
  }

  location / {
    root /data/html;
    index portal.html index.html;
    try_files $uri $uri/ /portal.html;
  }"""

    # Hacemos reemplazo flexible
    import re
    # Primero removemos cualquier bloque previo de location de estáticos que hayamos agregado
    content = re.sub(r'location ~\* \\\.(\(js\|css\|png\|jpg\|jpeg\|gif\|ico\|svg\|webp\|woff\|woff2\|ttf\|otf\|json\)\|\(js\|css\|png\|jpg\|jpeg\|gif\|ico\|svg\|webp\|woff\|woff2\|ttf\|otf\|json\))\$ \{.*?\}', '', content, flags=re.DOTALL)
    # También removemos cualquier directiva gzip repetida
    content = re.sub(r'gzip on;.*?gzip_comp_level\s+\d+;', '', content, flags=re.DOTALL)

    content_mod, count = re.subn(
        r'location\s+/\s+\{\s+root\s+/data/html;\s+index\s+portal\.html\s+index\.html;\s+try_files\s+\$uri\s+\$uri/\s+/portal\.html;\s+\}',
        replacement,
        content
    )

    if count > 0:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content_mod)
        print("Archivo 4.conf de Nginx optimizado y con politicas CORS exitosamente.")
    else:
        print("Error: No se pudo inyectar el bloque de optimizacion en Nginx.")

if __name__ == "__main__":
    optimizar_nginx()
