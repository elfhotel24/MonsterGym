# 🏋️ MonsterGym POS

Sistema de gestión integral para gimnasios, desarrollado en **Python** con almacenamiento local en **SQLite** y configuración en **JSON**. Diseñado para funcionar de forma local y ligera, sin necesidad de servidores externos.

---

## 📋 Características

- 📦 **Manejo de inventario** — Control de productos y stock del gimnasio.
- 🧾 **Facturación** — Emisión de facturas para clientes.
- 💰 **Cuadre de caja** — Cierre y arqueo de caja diario.
- 📊 **Reportes de ventas** — Reportes por rango de fechas y por producto.
- 👥 **Suscripciones de clientes** — Gestión de membresías y vencimientos.
- 🖨️ **Impresión de tickets** — Soporte para impresoras térmicas.

---

## 🛠️ Tecnologías utilizadas

| Componente       | Tecnología          |
|------------------|---------------------|
| Lenguaje         | Python 3.x          |
| Base de datos    | SQLite              |
| Configuración    | JSON                |
| Empaquetado      | PyInstaller         |
| Interfaz         | Tkinter / CustomTkinter |

---

## 📁 Estructura del proyecto
MonsterGym/
├── main.py # Punto de entrada de la aplicación
├── impresora.py # Módulo para manejo de impresión
├── monster.ico # Ícono de la aplicación
├── main.spec # Configuración de PyInstaller
├── MonsterGymPOS.spec # Configuración de PyInstaller (POS)
├── .gitignore # Archivos ignorados por Git
└── README.md # Este archivo



---

## 🚀 Instalación y ejecución

### Requisitos previos

- Python 3.8 o superior
- pip

### Pasos

1. **Clona el repositorio:**
   ```bash
   git clone https://github.com/elfhotel24/MonsterGym.git
   cd MonsterGym

Crea un entorno virtual (recomendado):

bash
python -m venv venv
# En Windows:
venv\Scripts\activate
# En Linux/Mac:
source venv/bin/activate
Instala las dependencias:

bash
pip install -r requirements.txt
Ejecuta la aplicación:

bash
python main.py
📦 Generar el ejecutable (.exe)
Si deseas empaquetar la aplicación como un ejecutable de Windows:

bash
pyinstaller --onefile --windowed --icon=monster.ico --name=MonsterGymPOS main.py
El ejecutable se generará en la carpeta dist/.

⚙️ Configuración
La aplicación utiliza un archivo de configuración en formato JSON para parámetros locales.

⚠️ Importante: Los archivos .db, .sqlite y config_local.json no se incluyen en el repositorio por contener datos sensibles. Se generan automáticamente en la primera ejecución.

🤝 Contribuciones
Este es un proyecto personal, pero cualquier sugerencia o mejora es bienvenida. Puedes abrir un issue o enviar un pull request.

👤 Autor
elfhotel24

GitHub: @elfhotel24

   
