-- 工程表ツールの利用状況カウンタ（Supabase）
-- Supabase の SQL Editor に貼り付けて実行する。
-- 匿名キーは taskpane.js に書くため公開される。RLS で「書き込みのみ」に制限する。

-- ① 記録用のテーブル
create table if not exists public.usage_log (
  id         bigint generated always as identity primary key,
  created_at timestamptz not null default now(),
  install_id uuid not null,                     -- 端末ごとの匿名ID（個人情報ではない）
  event      text not null check (event in ('open', 'generate')),
  version    text check (char_length(version) <= 32),
  host       text check (char_length(host) <= 32)
);

create index if not exists usage_log_created_at_idx on public.usage_log (created_at desc);
create index if not exists usage_log_install_id_idx on public.usage_log (install_id);

-- ② 匿名キーでできるのは INSERT だけ。SELECT のポリシーは作らないので読み取りは不可。
alter table public.usage_log enable row level security;

drop policy if exists "anon insert only" on public.usage_log;
create policy "anon insert only"
  on public.usage_log for insert to anon
  with check (true);

-- ③ API（PostgREST）から届く匿名ロールに、書き込みの権限だけを明示的に与える。
--    プロジェクト作成時の「新しいテーブルに自動で API 権限を与える」の設定に関係なく動くようにする。
grant usage on schema public to anon;
grant insert on public.usage_log to anon;
revoke select, update, delete on public.usage_log from anon;


-- ========================================================================
-- 集計用のクエリ（SQL Editor で実行して確認する）
-- ========================================================================

-- 直近30日に使われた端末の数＝「今どれくらいの人が使っているか」
select count(distinct install_id) as machines_30d
from public.usage_log
where created_at > now() - interval '30 days';

-- 日ごとの台数・起動回数・ガントチャート生成回数
select (created_at at time zone 'Asia/Tokyo')::date            as day,
       count(distinct install_id)                              as machines,
       count(*) filter (where event = 'open')                  as opens,
       count(*) filter (where event = 'generate')              as generates
from public.usage_log
group by 1
order by 1 desc
limit 60;

-- バージョン別の台数（自動更新されていない端末の確認）
select version,
       count(distinct install_id) as machines,
       max(created_at)            as last_seen
from public.usage_log
where created_at > now() - interval '30 days'
group by 1
order by 2 desc;

-- 端末ごとの最終利用日（何台が今も使い続けているかの確認）
select install_id,
       min(created_at) as first_seen,
       max(created_at) as last_seen,
       count(*)        as events
from public.usage_log
group by 1
order by last_seen desc;
