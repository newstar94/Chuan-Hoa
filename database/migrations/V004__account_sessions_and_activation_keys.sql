BEGIN;

ALTER TABLE users ADD COLUMN password_hash text;

-- Account authentication is deliberately separate from the existing device
-- enrollment/lease tables.  A session has an absolute lifetime and never
-- slides merely because the client is active.
CREATE TABLE account_sessions (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    device_thumbprint text NOT NULL,
    token_hash bytea NOT NULL UNIQUE,
    issued_at_utc timestamptz NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    revoked_at_utc timestamptz,
    last_seen_at_utc timestamptz,
    CHECK (octet_length(token_hash) = 32),
    CHECK (expires_at_utc = issued_at_utc + interval '72 hours'),
    CHECK (expires_at_utc > issued_at_utc)
);
CREATE INDEX ix_account_sessions_user_active
    ON account_sessions(user_id, expires_at_utc)
    WHERE revoked_at_utc IS NULL;

-- One current binding per account. Releasing a device updates this row rather
-- than creating a second active binding, making concurrent login decisions
-- deterministic under a row lock.
CREATE TABLE account_device_bindings (
    user_id uuid PRIMARY KEY REFERENCES users(id),
    device_thumbprint text NOT NULL,
    bound_at_utc timestamptz NOT NULL,
    last_seen_at_utc timestamptz NOT NULL,
    released_at_utc timestamptz,
    release_reason text
);
ALTER TABLE account_device_bindings
    ADD CONSTRAINT uq_account_device_binding_device UNIQUE (user_id, device_thumbprint);
CREATE INDEX ix_account_device_bindings_active_device
    ON account_device_bindings(device_thumbprint)
    WHERE released_at_utc IS NULL;

CREATE TABLE activation_keys (
    id uuid PRIMARY KEY,
    key_prefix text NOT NULL,
    key_hash bytea NOT NULL UNIQUE,
    status text NOT NULL CHECK (status IN ('ACTIVE', 'REVOKED', 'EXPIRED')),
    product_id uuid NOT NULL REFERENCES products(id),
    feature_codes text[] NOT NULL,
    max_devices integer NOT NULL CHECK (max_devices > 0),
    created_at_utc timestamptz NOT NULL,
    created_by uuid REFERENCES users(id),
    created_by_external_actor text,
    expires_at_utc timestamptz,
    revoked_at_utc timestamptz,
    note text,
    CHECK (octet_length(key_hash) = 32),
    CHECK (key_prefix <> ''),
    CHECK (cardinality(feature_codes) > 0)
);
CREATE INDEX ix_activation_keys_hash ON activation_keys(key_hash);
CREATE INDEX ix_activation_keys_status_expiry ON activation_keys(status, expires_at_utc);

CREATE TABLE activation_key_devices (
    id uuid PRIMARY KEY,
    activation_key_id uuid NOT NULL REFERENCES activation_keys(id),
    device_thumbprint text NOT NULL,
    activated_at_utc timestamptz NOT NULL,
    released_at_utc timestamptz,
    release_reason text,
    UNIQUE (activation_key_id, device_thumbprint)
);
CREATE INDEX ix_activation_key_devices_active
    ON activation_key_devices(activation_key_id)
    WHERE released_at_utc IS NULL;

COMMIT;
