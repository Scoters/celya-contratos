import paramiko
import sys

try:
    hostname = "107.175.122.33"
    port = 22
    username = "root"
    password = "NYsSy39vvYB7690Whm"
    local_file = "portal.html"
    remote_path = "/root/nginx-proxy/data/html/portal.html"
    
    print(f"Conectando a {hostname}...")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname, port, username, password)
    print("Conexión SSH exitosa.")
    
    print("Subiendo portal.html...")
    sftp = ssh.open_sftp()
    sftp.put(local_file, remote_path)
    sftp.close()
    print("Subida completa.")
    
    print("Copiando al contenedor Docker...")
    stdin, stdout, stderr = ssh.exec_command("docker cp /root/nginx-proxy/data/html/portal.html nginx-proxy_app_1:/var/www/html/portal.html")
    
    out = stdout.read().decode().strip()
    err = stderr.read().decode().strip()
    if out: print(out)
    if err: print(err)
    print("Comando Docker ejecutado exitosamente.")
    
    ssh.close()
    print("Despliegue finalizado.")
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
