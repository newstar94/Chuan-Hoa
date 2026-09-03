BEGIN;

DO $$
DECLARE
    expected_tables text[] := ARRAY[
        'organizations',
        'users',
        'identities',
        'organization_memberships',
        'devices',
        'device_challenges',
        'refresh_token_families',
        'products',
        'features',
        'trial_campaigns',
        'trial_grants',
        'entitlement_grants',
        'license_leases',
        'execution_grants',
        'offers',
        'quotes',
        'orders',
        'payment_events',
        'outbox_messages',
        'idempotency_records',
        'approval_requests',
        'audit_events'
    ];
    missing_table text;
BEGIN
    SELECT table_name
      INTO missing_table
      FROM unnest(expected_tables) AS table_name
     WHERE to_regclass('public.' || table_name) IS NULL
     LIMIT 1;

    IF missing_table IS NOT NULL THEN
        RAISE EXCEPTION 'required table is missing: %', missing_table;
    END IF;
END;
$$;

INSERT INTO users (
    id,
    normalized_email,
    display_name,
    status,
    email_verified_at_utc,
    created_at_utc,
    updated_at_utc
) VALUES (
    '00000000-0000-0000-0000-000000000001',
    'migration@example.invalid',
    'Migration Test User',
    'ACTIVE',
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:00:00Z'
);

INSERT INTO products (
    id,
    code,
    display_name,
    status,
    created_at_utc
) VALUES (
    '10000000-0000-0000-0000-000000000001',
    'CHUAN_HOA_WORD',
    'Chuẩn Hóa Word',
    'ACTIVE',
    '2026-09-01T00:00:00Z'
);

INSERT INTO devices (
    id,
    user_id,
    display_name,
    public_key_jwk,
    public_key_thumbprint,
    assurance,
    status,
    enrolled_at_utc
) VALUES (
    '20000000-0000-0000-0000-000000000001',
    '00000000-0000-0000-0000-000000000001',
    'Migration Test Device',
    '{"kty":"EC","crv":"P-256","x":"test-x","y":"test-y"}'::jsonb,
    'migration-test-thumbprint',
    'SOFTWARE_KEY',
    'ACTIVE',
    '2026-09-01T00:00:00Z'
);

INSERT INTO trial_campaigns (
    id,
    product_id,
    kind,
    status,
    starts_at_utc,
    ends_at_utc,
    published_at_utc,
    created_at_utc
) VALUES (
    '30000000-0000-0000-0000-000000000001',
    '10000000-0000-0000-0000-000000000001',
    'LAUNCH',
    'ACTIVE',
    '2026-09-01T00:00:00Z',
    '2026-10-01T00:00:00Z',
    '2026-08-31T00:00:00Z',
    '2026-08-31T00:00:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO trial_campaigns (
            id,
            product_id,
            kind,
            status,
            starts_at_utc,
            ends_at_utc,
            published_at_utc,
            created_at_utc
        ) VALUES (
            '30000000-0000-0000-0000-000000000002',
            '10000000-0000-0000-0000-000000000001',
            'LAUNCH',
            'SCHEDULED',
            '2026-09-15T00:00:00Z',
            '2026-10-15T00:00:00Z',
            '2026-08-31T00:00:00Z',
            '2026-08-31T00:00:00Z'
        );
        RAISE EXCEPTION 'overlapping launch trial campaign was accepted';
    EXCEPTION
        WHEN exclusion_violation THEN NULL;
    END;
END;
$$;

INSERT INTO offers (
    id,
    product_id,
    version,
    audience,
    channel,
    currency,
    amount_minor,
    billing_term_seconds,
    feature_codes,
    status,
    effective_from_utc,
    effective_until_utc,
    published_at_utc,
    created_at_utc
) VALUES (
    '40000000-0000-0000-0000-000000000001',
    '10000000-0000-0000-0000-000000000001',
    1,
    'INDIVIDUAL',
    'DIRECT',
    'VND',
    100000,
    2592000,
    ARRAY['COMPLIANCE'],
    'PUBLISHED',
    '2026-09-01T00:00:00Z',
    '2026-10-01T00:00:00Z',
    '2026-08-31T00:00:00Z',
    '2026-08-31T00:00:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO offers (
            id,
            product_id,
            version,
            audience,
            channel,
            currency,
            amount_minor,
            billing_term_seconds,
            feature_codes,
            status,
            effective_from_utc,
            effective_until_utc,
            published_at_utc,
            created_at_utc
        ) VALUES (
            '40000000-0000-0000-0000-000000000002',
            '10000000-0000-0000-0000-000000000001',
            2,
            'INDIVIDUAL',
            'DIRECT',
            'VND',
            120000,
            2592000,
            ARRAY['COMPLIANCE'],
            'PUBLISHED',
            '2026-09-15T00:00:00Z',
            '2026-10-15T00:00:00Z',
            '2026-09-01T00:00:00Z',
            '2026-09-01T00:00:00Z'
        );
        RAISE EXCEPTION 'overlapping published offer was accepted';
    EXCEPTION
        WHEN exclusion_violation THEN NULL;
    END;
END;
$$;

INSERT INTO trial_grants (
    id,
    user_id,
    product_id,
    campaign_id,
    kind,
    status,
    starts_at_utc,
    ends_at_utc,
    source_reference,
    created_at_utc
) VALUES (
    '50000000-0000-0000-0000-000000000001',
    '00000000-0000-0000-0000-000000000001',
    '10000000-0000-0000-0000-000000000001',
    '30000000-0000-0000-0000-000000000001',
    'LAUNCH',
    'ACTIVE',
    '2026-09-01T00:00:00Z',
    '2026-10-01T00:00:00Z',
    'launch-campaign-migration-test',
    '2026-09-01T00:00:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO trial_grants (
            id,
            user_id,
            product_id,
            campaign_id,
            kind,
            status,
            starts_at_utc,
            ends_at_utc,
            source_reference,
            created_at_utc
        ) VALUES (
            '50000000-0000-0000-0000-000000000002',
            '00000000-0000-0000-0000-000000000001',
            '10000000-0000-0000-0000-000000000001',
            NULL,
            'PERSONAL',
            'ACTIVE',
            '2026-10-01T00:00:00Z',
            '2026-10-31T00:00:00Z',
            'personal-trial-should-not-stack',
            '2026-10-01T00:00:00Z'
        );
        RAISE EXCEPTION 'second trial grant for the same user and product was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;
END;
$$;

INSERT INTO execution_grants (
    id,
    user_id,
    device_id,
    command_id,
    client_release_id,
    document_fingerprint,
    operation_scope,
    nonce_hash,
    jti,
    key_id,
    issued_at_utc,
    not_before_utc,
    expires_at_utc
) VALUES (
    '60000000-0000-0000-0000-000000000001',
    '00000000-0000-0000-0000-000000000001',
    '20000000-0000-0000-0000-000000000001',
    'OnFormatDocument',
    'migration-test-release',
    'migration-test-document',
    'DOCUMENT_WRITE',
    decode('01020304', 'hex'),
    'migration-test-jti',
    'migration-test-key',
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:05:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO execution_grants (
            id,
            user_id,
            device_id,
            command_id,
            client_release_id,
            document_fingerprint,
            operation_scope,
            nonce_hash,
            jti,
            key_id,
            issued_at_utc,
            not_before_utc,
            expires_at_utc
        ) VALUES (
            '60000000-0000-0000-0000-000000000002',
            '00000000-0000-0000-0000-000000000001',
            '20000000-0000-0000-0000-000000000001',
            'OnFormatDocument',
            'migration-test-release',
            'migration-test-document',
            'DOCUMENT_WRITE',
            decode('01020304', 'hex'),
            'migration-test-jti-2',
            'migration-test-key',
            '2026-09-01T00:00:00Z',
            '2026-09-01T00:00:00Z',
            '2026-09-01T00:05:00Z'
        );
        RAISE EXCEPTION 'duplicate execution grant nonce was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;

    BEGIN
        INSERT INTO execution_grants (
            id,
            user_id,
            device_id,
            command_id,
            client_release_id,
            document_fingerprint,
            operation_scope,
            nonce_hash,
            jti,
            key_id,
            issued_at_utc,
            not_before_utc,
            expires_at_utc
        ) VALUES (
            '60000000-0000-0000-0000-000000000003',
            '00000000-0000-0000-0000-000000000001',
            '20000000-0000-0000-0000-000000000001',
            'OnFormatDocument',
            'migration-test-release',
            'migration-test-document',
            'DOCUMENT_WRITE',
            decode('05060708', 'hex'),
            'migration-test-jti',
            'migration-test-key',
            '2026-09-01T00:00:00Z',
            '2026-09-01T00:00:00Z',
            '2026-09-01T00:05:00Z'
        );
        RAISE EXCEPTION 'duplicate execution grant jti was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;
END;
$$;

INSERT INTO quotes (
    id,
    user_id,
    offer_id,
    offer_version,
    product_id,
    currency,
    amount_minor,
    billing_term_seconds,
    feature_codes,
    status,
    integrity_binding,
    created_at_utc,
    expires_at_utc
) VALUES (
    '70000000-0000-0000-0000-000000000001',
    '00000000-0000-0000-0000-000000000001',
    '40000000-0000-0000-0000-000000000001',
    1,
    '10000000-0000-0000-0000-000000000001',
    'VND',
    100000,
    2592000,
    ARRAY['COMPLIANCE'],
    'ACTIVE',
    decode('11121314', 'hex'),
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:30:00Z'
);

INSERT INTO orders (
    id,
    user_id,
    quote_id,
    stable_order_code,
    currency,
    amount_minor,
    status,
    created_at_utc,
    updated_at_utc
) VALUES (
    '80000000-0000-0000-0000-000000000001',
    '00000000-0000-0000-0000-000000000001',
    '70000000-0000-0000-0000-000000000001',
    'MIGRATION-ORDER-001',
    'VND',
    100000,
    'CREATED',
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:00:00Z'
);

INSERT INTO payment_events (
    id,
    order_id,
    provider,
    provider_event_id,
    event_type,
    signature_valid,
    currency,
    amount_minor,
    payload_hash,
    received_at_utc
) VALUES (
    '90000000-0000-0000-0000-000000000001',
    '80000000-0000-0000-0000-000000000001',
    'MIGRATION_TEST',
    'provider-event-001',
    'PAYMENT_RECEIVED',
    TRUE,
    'VND',
    100000,
    decode('21222324', 'hex'),
    '2026-09-01T00:00:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO payment_events (
            id,
            order_id,
            provider,
            provider_event_id,
            event_type,
            signature_valid,
            payload_hash,
            received_at_utc
        ) VALUES (
            '90000000-0000-0000-0000-000000000002',
            '80000000-0000-0000-0000-000000000001',
            'MIGRATION_TEST',
            'provider-event-001',
            'PAYMENT_RECEIVED',
            TRUE,
            decode('31323334', 'hex'),
            '2026-09-01T00:01:00Z'
        );
        RAISE EXCEPTION 'duplicate provider event id was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;

    BEGIN
        UPDATE payment_events
           SET processing_result = 'mutation must fail'
         WHERE id = '90000000-0000-0000-0000-000000000001';
        RAISE EXCEPTION 'payment event update was accepted';
    EXCEPTION
        WHEN SQLSTATE '55000' THEN NULL;
    END;

    BEGIN
        DELETE FROM payment_events
         WHERE id = '90000000-0000-0000-0000-000000000001';
        RAISE EXCEPTION 'payment event delete was accepted';
    EXCEPTION
        WHEN SQLSTATE '55000' THEN NULL;
    END;
END;
$$;

INSERT INTO audit_events (
    event_id,
    occurred_at_utc,
    actor_user_id,
    environment,
    action_code,
    target_type,
    target_id,
    correlation_id,
    event_hash
) VALUES (
    'a0000000-0000-0000-0000-000000000001',
    '2026-09-01T00:00:00Z',
    '00000000-0000-0000-0000-000000000001',
    'MIGRATION_TEST',
    'MIGRATION_APPLIED',
    'DATABASE',
    'V001',
    'b0000000-0000-0000-0000-000000000001',
    decode('41424344', 'hex')
);

DO $$
BEGIN
    BEGIN
        UPDATE audit_events
           SET reason = 'mutation must fail'
         WHERE event_id = 'a0000000-0000-0000-0000-000000000001';
        RAISE EXCEPTION 'audit event update was accepted';
    EXCEPTION
        WHEN SQLSTATE '55000' THEN NULL;
    END;

    BEGIN
        DELETE FROM audit_events
         WHERE event_id = 'a0000000-0000-0000-0000-000000000001';
        RAISE EXCEPTION 'audit event delete was accepted';
    EXCEPTION
        WHEN SQLSTATE '55000' THEN NULL;
    END;
END;
$$;

INSERT INTO idempotency_records (
    id,
    scope_key,
    idempotency_key_hash,
    request_hash,
    method,
    path,
    status,
    owner_token,
    expires_at_utc,
    created_at_utc
) VALUES (
    'c0000000-0000-0000-0000-000000000001',
    'migration-test-user',
    decode('00112233445566778899aabbccddeeff00112233445566778899aabbccddeeff', 'hex'),
    decode('102132435465768798a9bacbdcedfe0f102132435465768798a9bacbdcedfe0f', 'hex'),
    'POST',
    '/v1/migration-test',
    'IN_PROGRESS',
    'd0000000-0000-0000-0000-000000000001',
    '2026-09-02T00:00:00Z',
    '2026-09-01T00:00:00Z'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO idempotency_records (
            id,
            scope_key,
            idempotency_key_hash,
            request_hash,
            method,
            path,
            status,
            owner_token,
            expires_at_utc,
            created_at_utc
        ) VALUES (
            'c0000000-0000-0000-0000-000000000002',
            'migration-test-user',
            decode('00112233445566778899aabbccddeeff00112233445566778899aabbccddeeff', 'hex'),
            decode('ffeeddccbbaa99887766554433221100ffeeddccbbaa99887766554433221100', 'hex'),
            'POST',
            '/v1/migration-test',
            'IN_PROGRESS',
            'd0000000-0000-0000-0000-000000000002',
            '2026-09-02T00:00:00Z',
            '2026-09-01T00:00:00Z'
        );
        RAISE EXCEPTION 'duplicate scoped idempotency key was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;
END;
$$;

INSERT INTO outbox_messages (
    message_id,
    aggregate_type,
    aggregate_id,
    event_type,
    payload,
    headers,
    occurred_at_utc,
    available_at_utc,
    idempotency_key
) VALUES (
    'e0000000-0000-0000-0000-000000000001',
    'Order',
    '80000000-0000-0000-0000-000000000001',
    'OrderCreated',
    '{"orderId":"80000000-0000-0000-0000-000000000001"}'::jsonb,
    '{"correlationId":"b0000000-0000-0000-0000-000000000001"}'::jsonb,
    '2026-09-01T00:00:00Z',
    '2026-09-01T00:00:00Z',
    'migration-order-created-001'
);

DO $$
BEGIN
    BEGIN
        INSERT INTO outbox_messages (
            message_id,
            aggregate_type,
            aggregate_id,
            event_type,
            payload,
            headers,
            occurred_at_utc,
            available_at_utc,
            idempotency_key
        ) VALUES (
            'e0000000-0000-0000-0000-000000000002',
            'Order',
            '80000000-0000-0000-0000-000000000001',
            'OrderCreated',
            '{"orderId":"80000000-0000-0000-0000-000000000001"}'::jsonb,
            '{}'::jsonb,
            '2026-09-01T00:00:01Z',
            '2026-09-01T00:00:01Z',
            'migration-order-created-001'
        );
        RAISE EXCEPTION 'duplicate outbox idempotency key was accepted';
    EXCEPTION
        WHEN unique_violation THEN NULL;
    END;
END;
$$;

ROLLBACK;
