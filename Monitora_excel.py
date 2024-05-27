import os
import time
import psutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

continua = True

class ArquivoFechadoHandler(FileSystemEventHandler):
    def __init__(self, arquivo_monitorado):
        super().__init__()
        self.arquivo_monitorado = arquivo_monitorado
        self.ultima_mensagem = {}

    def imprimir_mensagem(self, caminho_arquivo):
        tempo_atual = time.time()
        if caminho_arquivo in self.ultima_mensagem:
            tempo_ultima_mensagem = self.ultima_mensagem[caminho_arquivo]
            if tempo_atual - tempo_ultima_mensagem < 5:  # Alterado para 5 segundos
                return False
        self.ultima_mensagem[caminho_arquivo] = tempo_atual
        return True

    def on_modified(self, event):
        global continua  # Declaração da variável global
        if event.is_directory:
            return
        caminho_arquivo = event.src_path
        if os.path.basename(caminho_arquivo) == self.arquivo_monitorado:
            if self.imprimir_mensagem(caminho_arquivo):
                # Encontra e fecha o processo do Excel
                for proc in psutil.process_iter():
                    if "excel.exe" in proc.name().lower():
                        time.sleep(.1)  # Aguarda 1 segundos antes de finalizar
                        proc.kill()

                continua = False
       

# Nome do arquivo a ser monitorado
arquivo_monitorado = "0002.xlsx"

# Caminho completo do arquivo a ser monitorado
caminho_arquivo_monitorado = os.path.abspath(arquivo_monitorado)

# Cria um observador
observer = Observer()

# Define o manipulador de eventos
event_handler = ArquivoFechadoHandler(arquivo_monitorado)

# Adiciona o manipulador de eventos ao observador
observer.schedule(event_handler, path=os.path.dirname(caminho_arquivo_monitorado))

# Inicia o observador em segundo plano
observer.start()

while continua:
    time.sleep(.1)

observer.stop()
observer.join()

print("Excel Fechado!")
