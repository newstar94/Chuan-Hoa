DO $$
DECLARE
    remaining_objects text;
BEGIN
    SELECT string_agg(object_name, ', ' ORDER BY object_name)
      INTO remaining_objects
      FROM (
          SELECT tablename AS object_name
            FROM pg_tables
           WHERE schemaname = 'public'
          UNION ALL
          SELECT typname AS object_name
            FROM pg_type
           WHERE typnamespace = 'public'::regnamespace
             AND typtype = 'e'
          UNION ALL
          SELECT proname AS object_name
            FROM pg_proc
           WHERE pronamespace = 'public'::regnamespace
             AND proname = 'reject_immutable_update'
      ) AS owned_objects;

    IF remaining_objects IS NOT NULL THEN
        RAISE EXCEPTION 'V001 down migration left owned objects: %', remaining_objects;
    END IF;
END;
$$;
