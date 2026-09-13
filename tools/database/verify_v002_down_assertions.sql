DO $$
DECLARE
    table_count integer;
BEGIN
    SELECT count(*) INTO table_count
    FROM information_schema.tables
    WHERE table_schema = 'public'
      AND table_name IN ('admin_integration_idempotency', 'admin_integration_audit');
    IF table_count <> 0 THEN
        RAISE EXCEPTION 'V002 down migration left owned tables: %', table_count;
    END IF;
END $$;
