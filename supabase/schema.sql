-- ==============================================================
-- CRONO-SYS - DATABASE SCHEMA & ROW LEVEL SECURITY (RLS)
-- Compatible con Supabase PostgreSQL
-- ==============================================================

-- 1. Habilitar extensión UUID
create extension if not exists "uuid-ossp";

-- 2. TABLA DE PERFILES (Vinculada a auth.users de Supabase)
create table if not exists public.profiles (
    id uuid primary key references auth.users(id) on delete cascade,
    email text,
    full_name text,
    avatar_url text,
    created_at timestamptz default now() not null,
    updated_at timestamptz default now() not null
);

-- 3. TABLA DE HORARIOS (Almacena cada versión o horario importado)
create table if not exists public.schedules (
    id uuid primary key default gen_random_uuid(),
    user_id uuid not null references auth.users(id) on delete cascade,
    title text not null default 'Horario Base Semanal',
    file_name text,
    is_active boolean default true not null,
    metadata jsonb default '{}'::jsonb,
    created_at timestamptz default now() not null,
    updated_at timestamptz default now() not null
);

-- 4. TABLA DE TAREAS Y TIEMPOS DE ESTUDIO
create table if not exists public.tasks (
    id uuid primary key default gen_random_uuid(),
    user_id uuid not null references auth.users(id) on delete cascade,
    schedule_id uuid references public.schedules(id) on delete set null,
    title text not null,
    category text default 'estudio' check (category in ('estudio', 'tarea', 'trabajo', 'personal')),
    duration_minutes integer not null default 60 check (duration_minutes > 0),
    preferred_days text[] default array['lunes','martes','miercoles','jueves','viernes'],
    time_preference text default 'any' check (time_preference in ('any', 'morning', 'afternoon', 'evening', 'custom')),
    custom_start_time time,
    custom_end_time time,
    max_session_minutes integer default 60 check (max_session_minutes > 0),
    rest_between_minutes integer default 15 check (rest_between_minutes >= 0),
    priority text default 'medium' check (priority in ('low', 'medium', 'high')),
    status text default 'pending' check (status in ('pending', 'scheduled', 'completed')),
    created_at timestamptz default now() not null,
    updated_at timestamptz default now() not null
);

-- 5. TABLA DE BLOQUES DEL CALENDARIO (Horario Outlook / Teams)
-- Contiene tanto las horas base del Excel como las tareas y descansos auto-organizados
create table if not exists public.schedule_blocks (
    id uuid primary key default gen_random_uuid(),
    schedule_id uuid not null references public.schedules(id) on delete cascade,
    user_id uuid not null references auth.users(id) on delete cascade,
    task_id uuid references public.tasks(id) on delete set null,
    day_of_week text not null check (day_of_week in ('lunes','martes','miercoles','jueves','viernes','sabado','domingo')),
    start_time time not null,
    end_time time not null,
    title text not null,
    block_type text not null default 'base' check (block_type in ('base', 'study', 'task', 'break')),
    color text default '#0284c7',
    notes text,
    created_at timestamptz default now() not null
);

-- Índices de alto rendimiento
create index if not exists idx_schedules_user on public.schedules(user_id);
create index if not exists idx_blocks_schedule on public.schedule_blocks(schedule_id);
create index if not exists idx_blocks_user_day on public.schedule_blocks(user_id, day_of_week);
create index if not exists idx_tasks_user on public.tasks(user_id);

-- ==============================================================
-- 6. POLÍTICAS DE SEGURIDAD ROW LEVEL SECURITY (RLS)
-- Asegura que CADA USUARIO SOLO PUEDA VER Y MODIFICAR SUS DATOS
-- ==============================================================

alter table public.profiles enable row level security;
alter table public.schedules enable row level security;
alter table public.schedule_blocks enable row level security;
alter table public.tasks enable row level security;

-- Políticas para Profiles
create policy "Usuarios pueden ver su propio perfil"
    on public.profiles for select
    using (auth.uid() = id);

create policy "Usuarios pueden insertar su propio perfil"
    on public.profiles for insert
    with check (auth.uid() = id);

create policy "Usuarios pueden actualizar su propio perfil"
    on public.profiles for update
    using (auth.uid() = id);

-- Políticas para Schedules
create policy "Usuarios pueden ver sus propios horarios"
    on public.schedules for select
    using (auth.uid() = user_id);

create policy "Usuarios pueden crear sus propios horarios"
    on public.schedules for insert
    with check (auth.uid() = user_id);

create policy "Usuarios pueden actualizar sus propios horarios"
    on public.schedules for update
    using (auth.uid() = user_id);

create policy "Usuarios pueden borrar sus propios horarios"
    on public.schedules for delete
    using (auth.uid() = user_id);

-- Políticas para Schedule Blocks
create policy "Usuarios pueden ver sus propios bloques de calendario"
    on public.schedule_blocks for select
    using (auth.uid() = user_id);

create policy "Usuarios pueden crear sus propios bloques"
    on public.schedule_blocks for insert
    with check (auth.uid() = user_id);

create policy "Usuarios pueden actualizar sus propios bloques"
    on public.schedule_blocks for update
    using (auth.uid() = user_id);

create policy "Usuarios pueden borrar sus propios bloques"
    on public.schedule_blocks for delete
    using (auth.uid() = user_id);

-- Políticas para Tasks
create policy "Usuarios pueden ver sus propias tareas"
    on public.tasks for select
    using (auth.uid() = user_id);

create policy "Usuarios pueden crear sus propias tareas"
    on public.tasks for insert
    with check (auth.uid() = user_id);

create policy "Usuarios pueden actualizar sus propias tareas"
    on public.tasks for update
    using (auth.uid() = user_id);

create policy "Usuarios pueden borrar sus propias tareas"
    on public.tasks for delete
    using (auth.uid() = user_id);

-- ==============================================================
-- 7. TRIGGER AUTOMÁTICO PARA NUEVOS USUARIOS DE AUTH
-- ==============================================================
create or replace function public.handle_new_user()
returns trigger as $$
begin
    insert into public.profiles (id, email, full_name, avatar_url)
    values (
        new.id,
        new.email,
        coalesce(new.raw_user_meta_data->>'full_name', split_part(new.email, '@', 1)),
        coalesce(new.raw_user_meta_data->>'avatar_url', '')
    )
    on conflict (id) do nothing;
    return new;
end;
$$ language plpgsql security definer;

-- Trigger tras registro en Supabase Auth
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
    after insert on auth.users
    for each row execute function public.handle_new_user();
