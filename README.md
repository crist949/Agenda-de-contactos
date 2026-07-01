# Agenda-de-contactos
Una herramienta utilizada para organizar y almacenar la información de personas o empresas con las que se mantiene comunicación.


LIBRERÍAS NECESARIAS

Instalar con pip:
    pip install flask openpyxl

CÓMO EJECUTAR

1. Abrir una terminal en la carpeta del proyecto.
2. Ejecutar:
       python app.py
3. Abrir el navegador en:
       http://127.0.0.1:5000

CREDENCIALES DE ACCESO

    Usuario:    admin
    Contraseña: 1234

ESTRUCTURA DEL PROYECTO

proyecto/
├── app.py                  
├── contactos.xlsx          
├── README.txt
├── static/
│   ├── css/
│   │   ├── style.css
│   │   ├── agregar.css
│   │   ├── editar_contac.css
│   │   ├── reporte_contac.css
│   │   ├── login.css
│   │   └── inicio.css
│   └── img/
│       └── logotipo.png
└── templates/
    ├── inicio.html         
    ├── login.html          
    ├── index.html          
    ├── agregar.html        
    ├── editar_contac.html  
    ├── editar_formulario.html 
    ├── detalle.html        
    ├── reporte_contac.html 
    └── cuenta_usuario.html 

FUNCIONALIDADES IMPLEMENTADAS

 Login (admin / 1234) con mensaje emergente de error
 CRUD completo de contactos (Agregar, Ver, Editar, Eliminar)
 Búsqueda por nombre
 Ordenar alfabéticamente A-Z
 Detalle completo del contacto (Más información)
 Confirmación modal al editar y al eliminar
 Validación frontend (JS) y backend (Python):
    -Correo con formato válido (usuario@dominio.com)
    -Teléfono de exactamente 8 dígitos
    -Campos obligatorios: nombre, teléfono, correo
 Campo Categoría (Familia, Trabajo, Amigos, Otro)
 Campo Favorito (toggle switch)
 Reporte con total de contactos y favoritos
 Toda la manipulación del Excel con openpyxl
 7 pantallas distintas
