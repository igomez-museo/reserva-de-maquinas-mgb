import http.server
import socketserver
import json
import os

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class MyHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_POST(self):
        if self.path == '/api/log_user':
            try:
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                data = json.loads(post_data.decode('utf-8'))
                
                device_id = data.get('deviceId', 'desconocido')
                username = data.get('username', 'desconocido')
                timestamp = data.get('timestamp', '')
                user_agent = self.headers.get('User-Agent', 'Unknown')
                
                # Intentar obtener el nombre de red (hostname) o la IP del cliente
                client_ip = self.client_address[0]
                try:
                    import socket
                    client_hostname = socket.getfqdn(client_ip)
                except Exception:
                    client_hostname = client_ip
                
                # Escribir en el archivo de logs incluyendo el nombre de red/IP
                log_line = f"[{timestamp}] ID Dispositivo: {device_id} | Cuenta: {username} | Nombre de Red/IP: {client_hostname} | Agente: {user_agent}\n"
                log_file_path = os.path.join(DIRECTORY, "usuarios_conectados.txt")
                
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(log_line)
                
                # Responder al cliente
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success", "message": "Logged"}).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

if __name__ == "__main__":
    os.chdir(DIRECTORY)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), MyHandler) as httpd:
        print(f"Servidor de Reserva de Maquinas ejecutandose en: http://localhost:{PORT}")
        print(f"Sirviendo archivos desde: {DIRECTORY}")
        print("Registros de inicio de sesion se guardaran en: usuarios_conectados.txt")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido.")
