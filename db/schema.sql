-- =============================================================================
-- Turon B2C — дашборд продаж турагентства
-- Схема БД: PostgreSQL 15+ / Supabase
--
-- Порядок применения:
--   1) schema.sql   — этот файл (типы, таблицы, индексы, вьюхи)
--   2) rls.sql      — политики доступа (роли РОП / менеджер)
--   3) seed.sql     — справочники и тестовые данные (опционально)
-- =============================================================================

create extension if not exists "pgcrypto";

-- -----------------------------------------------------------------------------
-- 1. Перечисления
-- -----------------------------------------------------------------------------

-- Роли пользователей системы.
create type user_role as enum ('admin', 'rop', 'manager');

-- Статусы сделки. Порядок = порядок движения по воронке.
--   new         — заявка поступила, в работу не взята
--   qualified   — связались с клиентом, потребность выяснена
--   offer_sent  — отправлен расчёт / подборка туров
--   booked      — бронь подтверждена оператором  <-- момент засчёта продажи
--   prepaid     — внесена предоплата
--   paid        — оплачено полностью
--   departed    — тур состоялся (клиент вылетел)
--   cancelled   — аннуляция на любом этапе
create type deal_status as enum (
  'new', 'qualified', 'offer_sent', 'booked', 'prepaid', 'paid', 'departed', 'cancelled'
);

-- Направление платежа: in — клиент платит агентству, out — агентство платит оператору.
create type payment_direction as enum ('in', 'out');

-- Тип справочника.
create type dictionary_kind as enum ('direction', 'operator', 'source');


-- -----------------------------------------------------------------------------
-- 2. Пользователи
-- -----------------------------------------------------------------------------

-- Профиль пользователя. id совпадает с auth.users.id (Supabase Auth).
create table profiles (
  id           uuid primary key references auth.users(id) on delete cascade,
  full_name    text        not null,
  role         user_role   not null default 'manager',
  email        text,
  phone        text,
  avatar_url   text,
  is_active    boolean     not null default true,

  hired_at     date,
  -- Дата ухода. Сотрудник никогда не удаляется: его сделки остаются в
  -- отчётах закрытых периодов. Уход снимает доступ и убирает человека из
  -- текущих рейтингов, но история продаж не меняется задним числом.
  left_at      date,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);

-- Написания имени сотрудника в исходной таблице. Их бывает несколько:
-- в реальном отчёте одна и та же Сарвиноз встречается как «Сарвиноз»,
-- «сарвиноз» и «SARVINOZ». Отдельная таблица, а не поле в profiles,
-- потому что псевдонимов у одного человека много и они добавляются
-- со временем.
create table profile_aliases (
  alias      text primary key,          -- нормализованный: lower(trim(...))
  profile_id uuid not null references profiles(id) on delete cascade
);

create index profile_aliases_profile_idx on profile_aliases (profile_id);

comment on table profile_aliases is
  'Написания ФИО менеджера в отчёте. По ним импорт связывает строку с профилем.';


-- -----------------------------------------------------------------------------
-- 3. Справочники (направления, туроператоры, источники лидов)
-- -----------------------------------------------------------------------------

-- Канонические значения. То, что в итоге видно в интерфейсе.
create table dictionaries (
  id         bigserial primary key,
  kind       dictionary_kind not null,
  name       text            not null,
  sort_order int             not null default 100,
  is_active  boolean         not null default true,
  unique (kind, name)
);

-- Синонимы из таблицы: «турция», «Turkey», «ТУРЦИЯ  » → одно направление.
-- Без этого справочники засоряются опечатками и аналитика разъезжается.
create table dictionary_aliases (
  kind          dictionary_kind not null,
  alias         text            not null,  -- нормализованный: lower(trim(...))
  dictionary_id bigint          not null references dictionaries(id) on delete cascade,
  primary key (kind, alias)
);


-- -----------------------------------------------------------------------------
-- 4. Клиенты
-- -----------------------------------------------------------------------------

create table clients (
  id         bigserial primary key,
  full_name  text not null,
  phone      text,
  email      text,
  birth_date date,
  comment    text,
  created_at timestamptz not null default now()
);

-- Телефон — естественный ключ дедупликации при импорте.
create unique index clients_phone_uniq on clients (phone) where phone is not null;


-- -----------------------------------------------------------------------------
-- 5. Сделки — центральная таблица
-- -----------------------------------------------------------------------------

create table deals (
  id            bigserial primary key,

  -- Ключ строки из источника. Для Google Sheets — значение колонки «ID заявки».
  -- Обеспечивает идемпотентность импорта: повторный прогон не плодит дубли.
  external_id   text unique,

  manager_id    uuid   not null references profiles(id) on delete restrict,
  client_id     bigint references clients(id) on delete set null,

  status        deal_status not null default 'new',

  direction_id  bigint references dictionaries(id),
  operator_id   bigint references dictionaries(id),
  source_id     bigint references dictionaries(id),

  tourists_count int not null default 1 check (tourists_count > 0),

  -- Деньги. gross — сумма от клиента, net — себестоимость (перевод оператору).
  -- Комиссия агентства = gross - net и считается базой для всех метрик маржи.
  currency      char(3)       not null default 'USD',
  fx_rate       numeric(14,6) not null default 1 check (fx_rate > 0),  -- курс к базовой валюте
  gross         numeric(14,2) not null default 0 check (gross >= 0),
  net           numeric(14,2) not null default 0 check (net   >= 0),

  commission      numeric(14,2) generated always as (gross - net) stored,
  gross_base      numeric(14,2) generated always as (gross * fx_rate) stored,
  commission_base numeric(14,2) generated always as ((gross - net) * fx_rate) stored,

  -- Даты жизненного цикла. Разделены намеренно: по ним строятся разные метрики.
  inquiry_date     date not null,   -- дата обращения → знаменатель конверсии
  booked_at        date,            -- дата подтверждения брони → период засчёта продажи
  paid_at          date,            -- дата полной оплаты → факт для кассы
  payment_due_date date,            -- дедлайн доплаты оператору → «горящие» доплаты
  depart_date      date,            -- дата вылета → ближайшие вылеты
  return_date      date,
  cancelled_at     date,
  cancel_reason    text,

  comment       text,

  -- Служебное для импорта.
  source_system text        not null default 'google_sheets',
  source_row    int,                  -- номер строки в листе (для сообщений об ошибках)
  row_hash      text,                 -- хеш строки: пропускаем неизменившиеся
  synced_at     timestamptz,

  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now(),

  -- Инварианты: бронь/оплата/аннуляция не могут быть раньше обращения.
  constraint deals_booked_after_inquiry  check (booked_at    is null or booked_at    >= inquiry_date),
  constraint deals_cancel_has_date       check (status <> 'cancelled' or cancelled_at is not null),
  constraint deals_return_after_depart   check (return_date  is null or depart_date is null or return_date >= depart_date)
);

create index deals_manager_idx      on deals (manager_id);
create index deals_inquiry_idx      on deals (inquiry_date desc);
create index deals_booked_idx       on deals (booked_at desc) where booked_at is not null;
create index deals_status_idx       on deals (status);
create index deals_depart_idx       on deals (depart_date) where depart_date is not null;
create index deals_due_idx          on deals (payment_due_date) where payment_due_date is not null;
create index deals_manager_booked_idx on deals (manager_id, booked_at desc);


-- -----------------------------------------------------------------------------
-- 6. Платежи
-- -----------------------------------------------------------------------------

create table payments (
  id          bigserial primary key,
  external_id text unique,
  deal_id     bigint            not null references deals(id) on delete cascade,
  direction   payment_direction not null default 'in',
  amount      numeric(14,2)     not null check (amount > 0),
  currency    char(3)           not null default 'USD',
  fx_rate     numeric(14,6)     not null default 1 check (fx_rate > 0),
  due_date    date,
  paid_at     date,
  method      text,
  comment     text,
  created_at  timestamptz not null default now()
);

create index payments_deal_idx on payments (deal_id);
create index payments_due_idx  on payments (due_date) where paid_at is null;


-- -----------------------------------------------------------------------------
-- 7. Планы (таргеты)
-- -----------------------------------------------------------------------------

-- manager_id = null → план на весь отдел.
create table targets (
  id           bigserial primary key,
  manager_id   uuid references profiles(id) on delete cascade,
  period_month date          not null,          -- всегда первое число месяца
  metric       text          not null default 'commission_base'
                 check (metric in ('gross_base', 'commission_base', 'deals_count', 'tourists_count')),
  plan_value   numeric(14,2) not null check (plan_value >= 0),
  created_at   timestamptz   not null default now(),
  constraint targets_month_is_first check (date_trunc('month', period_month)::date = period_month),
  unique nulls not distinct (manager_id, period_month, metric)
);


-- -----------------------------------------------------------------------------
-- 8. Журнал синхронизации с Google Sheets
-- -----------------------------------------------------------------------------

create table sync_runs (
  id            bigserial primary key,
  source        text        not null default 'google_sheets',
  started_at    timestamptz not null default now(),
  finished_at   timestamptz,
  status        text        not null default 'running'
                  check (status in ('running', 'success', 'partial', 'failed')),
  rows_read     int not null default 0,
  rows_upserted int not null default 0,
  rows_skipped  int not null default 0,   -- не изменились (совпал row_hash)
  rows_failed   int not null default 0,
  error         text
);

-- Построчные ошибки импорта. Показываются РОПу плашкой «12 строк не загружено».
create table sync_errors (
  id          bigserial primary key,
  run_id      bigint not null references sync_runs(id) on delete cascade,
  row_number  int,
  external_id text,
  column_name text,
  raw_value   text,
  message     text not null
);

create index sync_errors_run_idx on sync_errors (run_id);


-- -----------------------------------------------------------------------------
-- 9. Вьюхи для аналитики
-- -----------------------------------------------------------------------------

-- Плоский факт по сделке: все справочники раскрыты, метрики нормализованы.
-- Дашборд читает только эту вьюху, а не deals напрямую.
create view v_deal_facts as
select
  d.id,
  d.external_id,
  d.manager_id,
  p.full_name           as manager_name,
  d.client_id,
  c.full_name           as client_name,
  d.status,
  dir.name              as direction,
  op.name               as operator,
  src.name              as source,
  d.tourists_count,
  d.currency,
  d.gross_base,
  d.commission_base,
  case when d.gross_base > 0
       then round(d.commission_base / d.gross_base * 100, 2)
  end                   as margin_pct,
  d.inquiry_date,
  d.booked_at,
  d.paid_at,
  d.payment_due_date,
  d.depart_date,
  d.cancelled_at,
  -- Продажа засчитана: бронь подтверждена и не аннулирована.
  (d.booked_at is not null and d.status <> 'cancelled') as is_won,
  (d.status = 'cancelled')                              as is_cancelled,
  -- Открыта: в работе, ещё не бронь и не аннуляция.
  (d.status in ('new', 'qualified', 'offer_sent'))      as is_open
from deals d
  join profiles p        on p.id = d.manager_id
  left join clients c    on c.id = d.client_id
  left join dictionaries dir on dir.id = d.direction_id
  left join dictionaries op  on op.id  = d.operator_id
  left join dictionaries src on src.id = d.source_id;

-- Помесячный агрегат по менеджеру. Основа для карточек KPI и рейтинга.
create view v_manager_month as
select
  f.manager_id,
  f.manager_name,
  date_trunc('month', f.booked_at)::date        as period_month,
  count(*) filter (where f.is_won)              as deals_won,
  count(*) filter (where f.is_cancelled)        as deals_cancelled,
  coalesce(sum(f.gross_base)      filter (where f.is_won), 0) as gross_base,
  coalesce(sum(f.commission_base) filter (where f.is_won), 0) as commission_base,
  coalesce(sum(f.tourists_count)  filter (where f.is_won), 0) as tourists,
  case when count(*) filter (where f.is_won) > 0
       then round(sum(f.gross_base) filter (where f.is_won)
                  / count(*) filter (where f.is_won), 2)
  end as avg_check
from v_deal_facts f
where f.booked_at is not null
group by 1, 2, 3;

-- «Горящее»: то, что требует действия сегодня. Правая колонка дашборда.
create view v_attention_items as
-- Просроченная или подходящая доплата
select
  d.manager_id,
  'payment_due'::text as kind,
  d.id                as deal_id,
  d.payment_due_date  as due_date,
  (d.payment_due_date - current_date) as days_left
from deals d
where d.status in ('booked', 'prepaid')
  and d.payment_due_date is not null
  and d.payment_due_date <= current_date + 5
union all
-- Вылет в ближайшие 3 дня
select d.manager_id, 'departure', d.id, d.depart_date, (d.depart_date - current_date)
from deals d
where d.status in ('booked', 'prepaid', 'paid')
  and d.depart_date between current_date and current_date + 3
union all
-- Заявка без движения больше 2 дней
select d.manager_id, 'stale_lead', d.id, d.inquiry_date, (current_date - d.inquiry_date)
from deals d
where d.status in ('new', 'qualified')
  and d.updated_at < now() - interval '2 days';


-- -----------------------------------------------------------------------------
-- 10. Триггер updated_at
-- -----------------------------------------------------------------------------

create or replace function set_updated_at() returns trigger
language plpgsql as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

create trigger deals_updated_at    before update on deals    for each row execute function set_updated_at();
create trigger profiles_updated_at before update on profiles for each row execute function set_updated_at();
