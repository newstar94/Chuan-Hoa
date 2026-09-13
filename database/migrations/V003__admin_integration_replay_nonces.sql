BEGIN;

CREATE TABLE admin_integration_replay_nonces (
    client_id text NOT NULL,
    nonce text NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    PRIMARY KEY (client_id, nonce)
);

CREATE INDEX ix_admin_integration_replay_nonces_expiry
    ON admin_integration_replay_nonces(expires_at_utc);

COMMIT;
