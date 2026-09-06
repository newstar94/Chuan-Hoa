# Development 1.0.0.135

- Installer: artifacts/installers/development/ChuanHoa_Development_Test_Setup_1.0.0.135.exe
- SHA256: C1D501F5BA611E113B03F88BE8AB949F1EF51ABA9DB689BB6D38AF37E36BC7F6
- Size: 352552 bytes.
- Build and Development artifact audit passed; pinned self-signed certificate, not production trust.
- Installer tamper tests passed 8/8 (installer_tamper_rejection_135.json).
- Silent installation returned 0; installed registry version 1.0.0.135.
- Installed DevelopmentAccessSmoke --missing-external-trust passed: the signed embedded public key works with a nonexistent external trust path.
- About dialog source now includes the loaded assembly version.
- Installer delegates its transaction through the installed Office AppVLP on Click-to-Run systems; MSI path unchanged. This requires independent user launch validation, not merely external registry verification.

Pending: user launches Word from their ordinary shortcut and checks Ribbon and About. Do not infer this passes from the automated startup smoke alone. Full lifecycle/rollback under the new Office context has not yet been rerun.
