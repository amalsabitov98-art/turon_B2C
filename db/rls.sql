-- =============================================================================
-- Turon B2C — политики доступа (Row Level Security)
--
-- Принятое решение по приватности:
--   • РОП (role in 'rop','admin') — видит все сделки, всех менеджеров, все планы.
--   • Менеджер                    — видит ТОЛЬКО свои сделки, свои платежи,
--                                   свой план и свой профиль. Цифры коллег
--                                   недоступны ни в каком виде, включая агрегаты.
--
-- Разграничение живёт здесь, а не во фронтенде. Даже если UI покажет лишнее
-- или кто-то дёрнет API напрямую — база не отдаст чужие строки.
--
-- Запись во все таблицы идёт только сервисным ключом (импорт из Google Sheets),
-- поэтому политик insert/update для обычных пользователей нет: service_role
-- обходит RLS по определению.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Хелперы. security definer — чтобы проверка роли сама не упёрлась в RLS.
-- -----------------------------------------------------------------------------

create or replace function public.is_rop()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from profiles p
    where p.id = auth.uid()
      and p.role in ('rop', 'admin')
      and p.is_active
  );
$$;

revoke all on function public.is_rop() from public;
grant execute on function public.is_rop() to authenticated;


-- -----------------------------------------------------------------------------
-- Включаем RLS
-- -----------------------------------------------------------------------------

alter table profiles            enable row level security;
alter table clients             enable row level security;
alter table deals               enable row level security;
alter table payments            enable row level security;
alter table targets             enable row level security;
alter table dictionaries        enable row level security;
alter table dictionary_aliases  enable row level security;
alter table profile_aliases     enable row level security;
alter table sync_runs           enable row level security;
alter table sync_errors         enable row level security;


-- -----------------------------------------------------------------------------
-- profiles
-- -----------------------------------------------------------------------------

-- Менеджер видит только себя. РОП — всю команду (нужно для рейтинга и фильтров).
create policy profiles_select on profiles
  for select to authenticated
  using (id = auth.uid() or public.is_rop());

-- Правку профиля (аватар, телефон) разрешаем себе, но роль менять нельзя:
-- смена role делается только сервисным ключом.
create policy profiles_update_self on profiles
  for update to authenticated
  using (id = auth.uid())
  with check (id = auth.uid() and role = (select p.role from profiles p where p.id = auth.uid()));


-- -----------------------------------------------------------------------------
-- deals — ключевая политика
-- -----------------------------------------------------------------------------

create policy deals_select on deals
  for select to authenticated
  using (manager_id = auth.uid() or public.is_rop());


-- -----------------------------------------------------------------------------
-- payments — доступ наследуется от сделки
-- -----------------------------------------------------------------------------

create policy payments_select on payments
  for select to authenticated
  using (
    public.is_rop()
    or exists (select 1 from deals d where d.id = payments.deal_id and d.manager_id = auth.uid())
  );


-- -----------------------------------------------------------------------------
-- clients — менеджер видит только тех клиентов, по которым у него есть сделка
-- -----------------------------------------------------------------------------

create policy clients_select on clients
  for select to authenticated
  using (
    public.is_rop()
    or exists (select 1 from deals d where d.client_id = clients.id and d.manager_id = auth.uid())
  );


-- -----------------------------------------------------------------------------
-- targets — свой план виден, чужой нет
-- -----------------------------------------------------------------------------

create policy targets_select on targets
  for select to authenticated
  using (manager_id = auth.uid() or public.is_rop());


-- -----------------------------------------------------------------------------
-- Справочники — читают все авторизованные (это не чувствительные данные)
-- -----------------------------------------------------------------------------

create policy dictionaries_select on dictionaries
  for select to authenticated using (true);

-- Псевдонимы сотрудников — служебные данные импорта, нужны только РОПу
-- на экране «Команда».
create policy profile_aliases_select on profile_aliases
  for select to authenticated using (public.is_rop());

create policy dictionary_aliases_select on dictionary_aliases
  for select to authenticated using (true);


-- -----------------------------------------------------------------------------
-- Журнал импорта — только РОП/админ
-- -----------------------------------------------------------------------------

create policy sync_runs_select on sync_runs
  for select to authenticated using (public.is_rop());

create policy sync_errors_select on sync_errors
  for select to authenticated using (public.is_rop());


-- -----------------------------------------------------------------------------
-- ВАЖНО про вьюхи
--
-- Вьюхи в Postgres по умолчанию исполняются с правами владельца и обходят RLS
-- базовых таблиц. Поэтому все аналитические вьюхи объявляются с
-- security_invoker — тогда RLS применяется к тому, кто делает запрос, и
-- менеджер, читая v_manager_month, физически видит только свою строку.
-- -----------------------------------------------------------------------------

alter view v_deal_facts      set (security_invoker = on);
alter view v_manager_month   set (security_invoker = on);
alter view v_attention_items set (security_invoker = on);
