alter table public.tasks
  drop constraint if exists tasks_category_check;

alter table public.tasks
  add constraint tasks_category_check
  check (category in ('estudio', 'tarea', 'proyecto', 'trabajo', 'personal'));
