# Extractor de Datos de Clientes desde Facturas PDF

Este proyecto es una aplicación de escritorio desarrollada en Python que permite extraer datos de clientes desde facturas en formato PDF y guardarlos en un archivo Excel. La aplicación utiliza bibliotecas como `pdfplumber`, `openpyxl` y `PyQt5` para procesar los archivos y proporcionar una interfaz gráfica de usuario.

## Características

- Extrae datos como Razón Social, RUT, Giro, Dirección, Comuna, Ciudad, Nombre de Contacto y Teléfono desde facturas PDF.
- Guarda los datos en un archivo Excel (`clientes.xlsx`).
- Evita duplicados verificando si el RUT ya existe en el archivo Excel.
- Permite procesar un único archivo PDF o múltiples archivos a la vez.
- Incluye una barra de progreso para mostrar el estado del procesamiento.
- Interfaz gráfica amigable desarrollada con PyQt5.

## Requisitos

- Python 3.10 o superior.
- Las siguientes bibliotecas de Python:
  - `pdfplumber`
  - `openpyxl`
  - `PyQt5`

## Instalación

1. Clona este repositorio en tu máquina local:

   ```bash
   git clone https://github.com/C1ZC/lector_facturas.git
   cd lector_facturas
   ```

2. Crea un entorno virtual:

   ```bash
   python -m venv .venv
   ```

3. Activa el entorno virtual:

   - En Windows:
     ```bash
     .venv\Scripts\activate
     ```
   - En macOS/Linux:
     ```bash
     source .venv/bin/activate
     ```

4. Instala las dependencias:

   ```bash
   pip install -r requirements.txt
   ```

## Uso

1. Ejecuta la aplicación:

   ```bash
   python main.py
   ```

2. Usa la interfaz gráfica para:
   - Seleccionar un archivo PDF o múltiples archivos PDF.
   - Procesar los archivos y extraer los datos.
   - Ver el archivo Excel generado.

## Comandos útiles

- **Crear un ejecutable con PyInstaller**:
  ```bash
  pyinstaller --onefile --windowed main.py
  ```
  El ejecutable se generará en la carpeta `dist/`.

- **Actualizar las dependencias**:
  ```bash
  pip install --upgrade -r requirements.txt
  ```

## Estructura del Proyecto

```
lector_facturas/
├── main.py               # Código principal de la aplicación
├── requirements.txt      # Dependencias del proyecto
├── .gitignore            # Archivos y carpetas ignorados por Git
└── README.md             # Documentación del proyecto
```

## Autor

**Camilo Zavala - C1ZC**  
Portafolio: [https://c1zc.github.io/CamiloZavala/](https://c1zc.github.io/CamiloZavala/)

## Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.