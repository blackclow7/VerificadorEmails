-- Ejecuta esto en Supabase: Dashboard -> SQL Editor -> New query -> pega y RUN

create table if not exists users (
    id uuid primary key default gen_random_uuid(),
    email text unique not null,
    credits integer not null default 20,
    created_at timestamptz default now()
);

create table if not exists jobs (
    id uuid primary key default gen_random_uuid(),
    user_email text references users(email),
    filename text,
    total_emails integer,
    results jsonb,              -- solo se llena para verificaciones hechas EN LA NUBE (web)
    status_counts jsonb,        -- conteo agregado (ej. {"Accepted": 40, "Rejected": 5}), usado por la app de escritorio
    source text default 'web',  -- 'web' o 'desktop'
    created_at timestamptz default now()
);

create index if not exists idx_jobs_user_email on jobs(user_email);

-- Si ya habías corrido este schema antes (versión anterior sin status_counts/source),
-- ejecuta también estas dos líneas para actualizar la tabla existente sin perder datos:
alter table jobs add column if not exists status_counts jsonb;
alter table jobs add column if not exists source text default 'web';
