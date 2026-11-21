# Instrucciones de Configuración de Supabase para Crono-SYS

Este proyecto está diseñado para conectarse a tu proyecto de Supabase con seguridad a nivel de filas (**Row Level Security - RLS**), garantizando que ningún usuario pueda ver o modificar el horario ni las tareas de otro usuario.

---

## Paso 1: Abrir el editor SQL en Supabase

1. Entra a tu panel de control de Supabase: [https://supabase.com/dashboard](https://supabase.com/dashboard)
2. Selecciona tu proyecto (`hplvbifzybthoclneumn`).
3. En el menú lateral izquierdo, haz clic en el icono **SQL Editor** (o pulsa la combinación de teclas para buscar "SQL").
4. Haz clic en **+ New Query** (Nueva consulta).

---

## Paso 2: Ejecutar el script SQL

1. Abre el archivo [`supabase/schema.sql`](./schema.sql).
2. Copia todo su contenido y pégalo en el editor SQL de Supabase.
3. Haz clic en el botón verde **Run** (Ejecutar) en la esquina inferior derecha.

El script creará automáticamente:
- **`profiles`**: Tabla para los datos de los usuarios registrados.
- **`schedules`**: Tabla para guardar los horarios importados desde Excel.
- **`tasks`**: Tabla para las tareas y tiempos de estudio organizados.
- **`schedule_blocks`**: Tabla para los bloques interactivos visibles en el calendario estilo Outlook/Teams.
- **Políticas RLS**: Cada tabla tendrá `ALTER TABLE ... ENABLE ROW LEVEL SECURITY;` y políticas estrictas vinculadas a `auth.uid() = user_id`.
- **Triggers**: Creación automática del perfil al registrar un nuevo usuario con Supabase Auth.

---

## Paso 3: Configuración de Autenticación en Supabase

1. En el panel de Supabase, ve a **Authentication** -> **Providers**.
2. Asegúrate de que **Email** esté activado (`Enabled`).
3. Activa **Confirm email** en **Authentication -> Providers -> Email**.
4. En **Authentication -> URL Configuration**, configura la URL del sitio (`https://<tu-dominio>/`) y añade esa misma URL a **Redirect URLs**. El enlace de confirmación debe abrir Crono-SYS para iniciar la sesión confirmada.

---

## Credenciales Configura das en Crono-SYS

- **URL del Proyecto**: `https://hplvbifzybthoclneumn.supabase.co`
- **Clave Pública (Publishable Key)**: `sb_publishable_FEyFH-AVquoJFfYoO40UBw_GaryX6C4`

Ya están integradas en `js/supabase-client.js`. En cuanto ejecutes el SQL, la sincronización en la nube funcionará en tiempo real.
