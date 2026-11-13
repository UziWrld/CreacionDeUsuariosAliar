# Herramienta de Gestión y Aprovisionamiento en Active Directory - Aliar S.A.

Aplicación corporativa de producción para la gestión, consulta, clonación de permisos y creación individual y masiva de colaboradores en Active Directory bajo el contenedor oficial:

`OU=USUARIOS,OU=E_ALIAR,DC=aliar,DC=com,DC=co`

---

## 🚀 Inicio Rápido

Para ejecutar la aplicación en cualquier equipo:
1. Haz doble clic sobre **`Iniciar-AppWinForms.vbs`** o **`Iniciar-AppWinForms.bat`**.
2. La ventana se abrirá de manera silenciosa, nativa y sin consolas PowerShell visibles.
3. Para crear un acceso directo en tu Escritorio, ejecuta **`Crear-AccesoDirectoEscritorio.bat`**.

---

## 🔄 Cómo Actualizar la Herramienta (Para el Equipo)

Cualquier miembro del equipo con acceso al repositorio puede actualizar la aplicación a la última versión ejecutando en una terminal dentro de esta carpeta:

```bash
git pull
```

---

## ✨ Características Principales

1. **Pestaña 1: Creación Individual (Nuevo Colaborador)**
   - **Pegado Inteligente:** Pega la información enviada por Gestión Humana (ejemplo: `1013687966 JOSE ANGEL ROMERO GUERRERO OPERARIO PLANTA`) y el sistema auto-clasifica los nombres en Tipo Título (`Jose Angel Romero Guerrero`), el cargo en Tipo Título (`Operario Planta`), asigna automáticamente la Sede (`CUN_PLENG`) y genera el usuario de dominio (`jose.romero`).
   - **Clonación de Modelo:** Permite buscar a un colaborador modelo para heredar en vivo su Sede, Departamento, Perfil Web (`GR_Perfil_*`), acceso a VPN (`GR_VPN_SSL`) y todas sus membresías departamentales y de carpetas.
   - **Copia de Credenciales:** Botón dedicado para copiar al portapapeles en formato limpio:
     ```text
     Usuario de dominio: nombre.apellido
     Contraseña: clave
     ```

2. **Pestaña 2: Gestión de Permisos (Colaborador Existente)**
   - Consulta instantánea y en tiempo real de colaboradores y sus membresías en Active Directory.
   - Cálculo automático de diferencial de grupos a agregar y grupos a remover.
   - Clonación directa de permisos entre colaboradores del mismo rol o sede.

3. **Pestaña 3: Carga Masiva (Lotes de Personal)**
   - Tabla interactiva editable para procesar decenas o cientos de colaboradores en segundos.
   - Casilla de VPN habilitada por defecto y configurable por fila.
   - Copia masiva de credenciales para entrega a Gestión Humana.
   - Exportación de reporte formal a CSV / Excel (`Reporte_Aprovisionamiento_GH_*.csv`).

4. **Sincronización de Sedes y OUs Oficiales**
   - Cobertura completa de las 33 Unidades Organizativas (OUs) oficiales de Aliar S.A. distribuidas en Cundinamarca, Meta, Santander, Antioquia, Atlántico y demás regionales.

---

## 📋 Estructura del Repositorio

- `App-CreacionUsuariosAliar.ps1`: Script principal con interfaz gráfica Windows Forms y lógica de dominio.
- `MANUAL_DE_USO.txt`: Manual operativo detallado en texto plano.
- `aliar_ad_data.json`: Base de datos local de respaldo con estructura de grupos y usuarios.
- `Logo-web-1.png`: Logotipo corporativo de Aliar S.A.
- `Iniciar-AppWinForms.vbs`: Lanzador silencioso sin consola para usuarios finales.
- `Iniciar-AppWinForms.bat`: Lanzador alternativo por consola.
- `Reiniciar-AppWinForms.bat`: Cierra instancias colgadas y reinicia la herramienta.
- `Crear-AccesoDirectoEscritorio.bat`: Crea el acceso directo en el Escritorio.
- `.gitignore`: Excluye archivos temporales y reportes de personal generados localmente.
