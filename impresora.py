import os
import tempfile
import subprocess

class ImpresoraTermica:
    def __init__(self, nombre="POS-80-Series", ancho_caracteres=48):
        self.nombre = nombre
        self.ancho = ancho_caracteres
        self.usar_win32 = False
        try:
            import win32print
            self.win32print = win32print
            impresoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            nombres = [p[2] for p in impresoras]
            if self.nombre in nombres:
                self.nombre_impresora = self.nombre
            else:
                self.nombre_impresora = win32print.GetDefaultPrinter()
                if not self.nombre_impresora:
                    raise Exception("No hay impresora predeterminada.")
            self.usar_win32 = True
            print(f"✅ Impresora usando win32print: {self.nombre_impresora}")
        except ImportError:
            print("⚠️ win32print no disponible. Se usará método alternativo.")
        except Exception as e:
            print(f"⚠️ Error con win32print: {e}. Se usará método alternativo.")

    def imprimir(self, texto):
        if self.usar_win32:
            self._imprimir_win32(texto)
        else:
            self._imprimir_alternativo(texto)

    def _imprimir_win32(self, texto):
        try:
            inicio = b'\x1B\x40'          # Inicializar impresora
            alinear_izq = b'\x1B\x61\x00' # Alinear izquierda
            fin = b'\n\n\n\n\x1D\x56\x00' # Avance y corte
            texto_bytes = texto.encode('cp858', errors='replace')
            data = inicio + alinear_izq + texto_bytes + fin
            
            hprinter = self.win32print.OpenPrinter(self.nombre_impresora)
            try:
                self.win32print.StartDocPrinter(hprinter, 1, ("Ticket", None, "RAW"))
                try:
                    self.win32print.StartPagePrinter(hprinter)
                    self.win32print.WritePrinter(hprinter, data)
                    self.win32print.EndPagePrinter(hprinter)
                finally:
                    self.win32print.EndDocPrinter(hprinter)
            finally:
                self.win32print.ClosePrinter(hprinter)
            print("✅ Ticket enviado a la impresora (ESC/POS).")
        except Exception as e:
            print(f"❌ Error en win32print: {e}. Usando método alternativo.")
            self._imprimir_alternativo(texto)

    def _imprimir_alternativo(self, texto):
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='cp1252') as f:
                f.write(texto)
                temp_path = f.name
            subprocess.run(['notepad', '/p', temp_path], check=True, shell=True)
            os.unlink(temp_path)
            print("✅ Ticket enviado a la impresora (notepad /p).")
        except Exception as e:
            print(f"❌ Error al imprimir con notepad: {e}")
            print("--- TICKET (impreso en consola) ---")
            print(texto)
            print("------------------------------------")
            raise Exception(f"No se pudo imprimir: {e}")