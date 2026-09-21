-- =====================================================================
-- MIGRATION: Buat bucket 'documents' khusus untuk file PDF/Word/Excel
-- Jalankan di Supabase SQL Editor
-- =====================================================================

-- 1. Buat bucket baru khusus dokumen (public agar bisa diakses siapa saja)
insert into storage.buckets (id, name, public)
values ('documents', 'documents', true)
on conflict (id) do nothing;

-- 2. Policy: Semua orang bisa BACA/DOWNLOAD dokumen
create policy "public_read_documents"
  on storage.objects for select
  using (bucket_id = 'documents');

-- 3. Policy: Hanya admin (user yang login) yang bisa UPLOAD/EDIT/HAPUS
create policy "admin_write_documents"
  on storage.objects for all
  using (bucket_id = 'documents' and auth.role() = 'authenticated')
  with check (bucket_id = 'documents' and auth.role() = 'authenticated');
