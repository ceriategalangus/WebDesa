-- =====================================================================
-- Website Informasi Desa - Aktifkan Realtime
-- =====================================================================
-- Jalankan script ini di: Supabase Dashboard > SQL Editor > New query

DO $$ 
BEGIN
  -- Menambahkan tabel ke dalam supabase_realtime agar bisa realtime update
  -- Jika tabel sudah ada di dalam publication, error akan diabaikan
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE umkm;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE berita;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE agenda;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE dokumen;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE perangkat_desa;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE galeri;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE pengaduan;
  EXCEPTION WHEN OTHERS THEN END;
  
  BEGIN
    ALTER PUBLICATION supabase_realtime ADD TABLE settings;
  EXCEPTION WHEN OTHERS THEN END;

END $$;
