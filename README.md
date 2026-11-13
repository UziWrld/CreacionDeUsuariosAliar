# Herramienta de Gestión y Aprovisionamiento en Active Directory - Aliar S.A.

Aplicación corporativa para la administración, consulta, clonación de permisos y aprovisionamiento individual y masivo de colaboradores en Active Directory bajo la ruta oficial:

`OU=USUARIOS,OU=E_ALIAR,DC=aliar,DC=com,DC=co`

---

## Inicio Rápido

Para iniciar la aplicación:
1. Ejecutar **`Iniciar-AppWinForms.vbs`** (inicio silencioso en segundo plano) o **`Iniciar-AppWinForms.bat`**.
2. Para generar el acceso directo en el Escritorio del usuario, ejecutar **`Crear-AccesoDirectoEscritorio.bat`**.

---

## Actualización del Repositorio

Para sincronizar la herramienta con la última versión publicada en GitHub, ejecutar dentro de esta carpeta:

```bash
git pull
```

---

## Funcionalidades Principales

### 1. Creación Individual (Nuevo Colaborador)
- **Pegado Inteligente:** Procesa cadenas de texto recibidas de Gestión Humana (ejemplo: `1013687966 JOSE ANGEL ROMERO GUERRERO OPERARIO PLANTA`), separando automáticamente la cédula, asignando nombres y cargo en Tipo Título (`Jose Angel Romero Guerrero`, `Operario Planta`), identificando la Sede en mayúscula sostenida (`CUN_PLENG`) y generando el usuario de dominio (`jose.romero`).
- **Clonación desde Modelo:** Permite buscar a un colaborador existente en el Directorio Activo para heredar en tiempo real su Sede, Departamento, Perfil Web (`GR_Perfil_*`), acceso VPN (`GR_VPN_SSL`) y la totalidad de sus membresías y carpetas departamentales.
- **Copia de Credenciales:** Genera en el portapapeles el formato directo para notificación:
  ```text
  Usuario de dominio: nombre.apellido
  Contraseña: clave
  ```

### 2. Gestión de Permisos (Colaborador Existente)
- Consulta en vivo contra el Controlador de Dominio de los grupos actuales del colaborador.
- Cálculo de diferenciales automáticos (grupos a agregar vs. grupos a remover).
- Homologación de permisos tomando como referencia a otro colaborador del área o sede.

### 3. Carga Masiva (Lotes de Personal)
- Cuadrícula interactiva para aprovisionar lotes de colaboradores.
- Casilla de acceso VPN individual por fila (habilitada por defecto).
- Copia consolidada de credenciales para entrega a Gestión Humana.
- Exportación del reporte final a formato CSV compatible con Microsoft Excel.

### 4. Soporte Oficial de Sedes y Unidades Organizativas (OU)
- Enlace directo con las 33 Unidades Organizativas oficiales del dominio `aliar.com.co` (Cundinamarca, Meta, Santander, Antioquia, Atlántico y demás regionales).

---

## Estructura de Archivos

- `App-CreacionUsuariosAliar.ps1`: Lógica principal del sistema e interfaz gráfica en Windows Forms.
- `MANUAL_DE_USO.txt`: Manual operativo de referencia en texto plano.
- `README.md`: Documentación técnica del proyecto.
- `aliar_ad_data.json`: Base de datos local de respaldo con estructura de grupos y usuarios.
- `Logo-web-1.png`: Imagen corporativa institucional.
- `Iniciar-AppWinForms.vbs`: Lanzador silencioso de la aplicación.
- `Iniciar-AppWinForms.bat`: Lanzador estándar por consola de comandos.
- `Reiniciar-AppWinForms.bat`: Utilidad de desbloqueo y reinicio de procesos.
- `Crear-AccesoDirectoEscritorio.bat`: Generador de acceso directo en el Escritorio.
- `.gitignore`: Configuración de exclusión para temporales y reportes de datos sensibles.
