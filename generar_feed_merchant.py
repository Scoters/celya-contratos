import urllib.request
import urllib.parse
import json
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import os
import math

def obtener_marca(modelo):
    model_lower = (modelo or "").lower()
    if "iphone" in model_lower or "apple" in model_lower:
        return "Apple"
    elif "samsung" in model_lower or "galaxy" in model_lower:
        return "Samsung"
    elif "motorola" in model_lower or "moto" in model_lower:
        return "Motorola"
    elif "xiaomi" in model_lower or "redmi" in model_lower or "poco" in model_lower:
        return "Xiaomi"
    elif "oppo" in model_lower:
        return "Oppo"
    elif "vivo" in model_lower:
        return "Vivo"
    elif "huawei" in model_lower:
        return "Huawei"
    elif "realme" in model_lower:
        return "Realme"
    elif "honor" in model_lower:
        return "Honor"
    elif "infinix" in model_lower:
        return "Infinix"
    elif "tecno" in model_lower:
        return "Tecno"
    elif "zte" in model_lower:
        return "ZTE"
    elif "oneplus" in model_lower:
        return "OnePlus"
    else:
        return "CelYa!"

def calcular_cuota_unificada(precio_contado, factor_eng, factor_int, num_pagos):
    try:
        eng = math.ceil((precio_contado * factor_eng) / 10) * 10
        total_financiado = (precio_contado * (1 + factor_int)) - eng
        cuota_sin_redondear = total_financiado / num_pagos
        return int(round(cuota_sin_redondear)), eng
    except Exception:
        return 0, 0

def generar_feed():
    script_url_base = "https://script.google.com/macros/s/AKfycbx_DHa9zNyGFZrfr6dS4NgoW9xbArfpskejYi_UP_o3NDP6ovIx7m_Aj_KIp0DGNAHJ/exec"
    url_tasas = f"{script_url_base}?action=obtenerTasasConfig"
    url_catalogo = f"{script_url_base}?action=buscarModelosCatalogo&q="
    
    tasas = {
        "interes26": 0.66,
        "interes52": 0.90,
        "engancheSemanal": 0.165,
        "engancheComodo": 0.11,
        "interesQuin26": 0.45,
        "interesQuin52": 0.65,
        "engancheQuin26": 0.25,
        "engancheQuin52": 0.35
    }

    try:
        # 1. Obtener Tasas en tiempo real
        req_tasas = urllib.request.Request(url_tasas, headers={'User-Agent': 'Mozilla/5.0'})
        try:
            with urllib.request.urlopen(req_tasas, timeout=10) as res:
                res_data = json.loads(res.read().decode('utf-8'))
                if isinstance(res_data, dict) and "interes26" in res_data:
                    tasas.update(res_data)
                    print("Tasas obtenidas en tiempo real del backend.")
        except Exception as e:
            print(f"Advertencia: No se pudieron obtener tasas en tiempo real ({e}). Usando fallbacks.")

        # 2. Obtener Catálogo
        req_cat = urllib.request.Request(url_catalogo, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req_cat, timeout=15) as res:
            raw_response = res.read().decode('utf-8')
            data = json.loads(raw_response)
            
        if data.get("result") != "success" or "resultados" not in data:
            print("Error: No se obtuvieron resultados válidos del catálogo.")
            return

        productos = data["resultados"]
        
        # --- TAREA A: ESCRIBIR CATALOGO.JSON ESTÁTICO EN EL VPS ---
        dest_json_paths = [
            "/root/nginx-proxy/data/html/catalogo.json",
            "/var/lib/docker/volumes/nginx-proxy_data/_data/html/catalogo.json"
        ]
        
        json_escrito = False
        for path in dest_json_paths:
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(raw_response)
                print(f"Catálogo JSON estático generado en: {path}")
                json_escrito = True
            except Exception as e:
                print(f"Error escribiendo JSON en {path}: {e}")

        if json_escrito:
            try:
                os.system("docker cp /root/nginx-proxy/data/html/catalogo.json nginx-proxy_app_1:/var/www/html/catalogo.json")
                print("Catálogo JSON copiado al contenedor Docker exitosamente.")
            except Exception as e:
                print(f"Error al copiar catalogo.json al contenedor docker: {e}")

        # --- TAREA B: CREAR Y ESCRIBIR XML DE MERCHANT FEED ---
        rss = ET.Element("rss", {
            "xmlns:g": "http://base.google.com/ns/1.0",
            "version": "2.0"
        })
        
        channel = ET.SubElement(rss, "channel")
        
        title = ET.SubElement(channel, "title")
        title.text = "CelYa! - Catálogo Oficial de Celulares"
        
        link_web = ET.SubElement(channel, "link")
        link_web.text = "https://estrenacelya.com"
        
        desc = ET.SubElement(channel, "description")
        desc.text = "Estrena el celular de tus sueños hoy mismo con enganches bajos y cómodos pagos semanales. Entregas en CDMX y EdoMex."
        
        for prod in productos:
            precio_contado = prod.get("precioContado")
            if not precio_contado:
                continue
            
            try:
                precio_num = float(precio_contado)
            except ValueError:
                continue

            # Calcular cuotas y enganches exactos de forma idéntica al JS
            cuota_sem_26, eng_sem_26 = calcular_cuota_unificada(precio_num, tasas.get("engancheSemanal", 0.165), tasas.get("interes26", 0.66), 26)
            cuota_sem_52, eng_sem_52 = calcular_cuota_unificada(precio_num, tasas.get("engancheComodo", 0.11), tasas.get("interes52", 0.90), 52)
            enganche_minimo = min(eng_sem_26, eng_sem_52)

            item = ET.SubElement(channel, "item")
            
            # g:id
            g_id = ET.SubElement(item, "g:id")
            g_id.text = str(prod.get("id", prod.get("modelo", "")))
            
            # g:title
            g_title = ET.SubElement(item, "g:title")
            g_title.text = f"{prod.get('modelo', 'Celular')} a Crédito"
            
            # g:description
            g_desc = ET.SubElement(item, "g:description")
            g_desc.text = (
                f"Estrena tu {prod.get('modelo')} hoy mismo en CelYa! "
                f"Pago inicial con enganche desde ${enganche_minimo} MXN y cómodos pagos a plazos de "
                f"${cuota_sem_52} MXN semanales (Plan Cómodo de 52 semanas) o "
                f"${cuota_sem_26} MXN semanales (Plan Rápido de 26 semanas). "
                f"Mínimos requisitos, trámite inmediato sin aval y entrega personal en CDMX y Estado de México."
            )
            
            # g:link
            g_link = ET.SubElement(item, "g:link")
            modelo_escapado = urllib.parse.quote(prod.get("modelo", ""))
            g_link.text = f"https://estrenacelya.com/?modelo={modelo_escapado}"
            
            # g:image_link
            g_image_link = ET.SubElement(item, "g:image_link")
            g_image_link.text = prod.get("foto", "https://estrenacelya.com/logo-modified.webp")
            
            # g:condition
            g_cond = ET.SubElement(item, "g:condition")
            es_usado = prod.get("esUsado")
            if es_usado is True or str(es_usado).lower() == 'true':
                g_cond.text = "refurbished"
            else:
                g_cond.text = "new"
                
            # g:availability
            g_avail = ET.SubElement(item, "g:availability")
            estado = str(prod.get("estado", "")).upper()
            if estado in ["DISPONIBLE", "MAS VENDIDO"]:
                g_avail.text = "in_stock"
            elif estado == "BAJO DEMANDA":
                g_avail.text = "preorder"
            else:
                g_avail.text = "out_of_stock"
                
            # g:price
            g_price = ET.SubElement(item, "g:price")
            g_price.text = f"{int(round(precio_num))} MXN"
            
            # g:brand
            g_brand = ET.SubElement(item, "g:brand")
            g_brand.text = obtener_marca(prod.get("modelo", ""))
            
        # Formatear el XML de forma estética
        xml_str = ET.tostring(rss, encoding='utf-8')
        parsed = minidom.parseString(xml_str)
        pretty_xml = parsed.toprettyxml(indent="  ", encoding="utf-8")
        
        # Escribir el archivo XML
        dest_xml_paths = [
            "/root/nginx-proxy/data/html/merchant_feed.xml",
            "/var/lib/docker/volumes/nginx-proxy_data/_data/html/merchant_feed.xml"
        ]
        
        xml_escrito = False
        for path in dest_xml_paths:
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "wb") as f:
                    f.write(pretty_xml)
                print(f"Feed XML generado exitosamente en: {path}")
                xml_escrito = True
            except Exception as e:
                print(f"Error escribiendo XML en {path}: {e}")
                
        if xml_escrito:
            try:
                os.system("docker cp /root/nginx-proxy/data/html/merchant_feed.xml nginx-proxy_app_1:/var/www/html/merchant_feed.xml")
                print("Feed XML copiado al contenedor Docker exitosamente.")
            except Exception as e:
                print(f"Error al copiar XML al contenedor docker: {e}")
                
    except Exception as e:
        print(f"Error general en generar_feed: {e}")

if __name__ == "__main__":
    generar_feed()
