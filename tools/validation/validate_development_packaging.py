from __future__ import annotations

import re
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
PROPS = ROOT / "Directory.Build.props"
VSTO_PROJECT = ROOT / "src" / "ChuanHoa.AddIn.Vsto" / "ChuanHoa.AddIn.Vsto.csproj"
BUILD_CONTRACT = ROOT / "tools" / "vsto" / "BuildContract.ps1"
PUBLISH = ROOT / "tools" / "vsto" / "publish_development_test.ps1"
LOCAL_PUBLISH = ROOT / "tools" / "vsto" / "publish_local_development.ps1"
BUILDER = ROOT / "tools" / "vsto" / "build_development_test_exe.ps1"
BOOTSTRAPPER_PROJECT = (
    ROOT
    / "tools"
    / "vsto"
    / "DevelopmentTestBootstrapper"
    / "DevelopmentTestBootstrapper.csproj"
)
BOOTSTRAPPER = BOOTSTRAPPER_PROJECT.with_name("Program.cs")
AUDIT = ROOT / "tools" / "validation" / "audit_development_installer.ps1"
LIFECYCLE = ROOT / "tools" / "vsto" / "verify_development_installer_lifecycle.ps1"
TAMPER_REJECTION = (
    ROOT / "tools" / "validation" / "verify_development_installer_tamper_rejection.ps1"
)


def text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def require(source: str, fragments: tuple[str, ...], message: str) -> None:
    missing = [fragment for fragment in fragments if fragment not in source]
    if missing:
        raise RuntimeError(f"{message}: {missing}")


def quoted_set_between(source: str, start: str, end: str) -> set[str]:
    try:
        block = source.split(start, 1)[1].split(end, 1)[0]
    except IndexError as error:
        raise RuntimeError(f"Could not parse payload allowlist after {start!r}.") from error
    return {
        value.replace("\\", "/")
        for _, value in re.findall(r"(['\"])([^'\"]+)\1", block)
        if "/" in value or "." in value
    }


def validate() -> None:
    props = ET.parse(PROPS).getroot()
    values: dict[str, list[str]] = {}
    for group in props.findall("PropertyGroup"):
        for child in group:
            values.setdefault(child.tag, []).append((child.text or "").strip())
    product_versions = values.get("ProductVersion", [])
    if len(product_versions) != 1 or not re.fullmatch(
        r"\d+\.\d+\.\d+\.\d+", product_versions[0]
    ):
        raise RuntimeError("Directory.Build.props must contain one four-part ProductVersion.")
    for name in ("Version", "AssemblyVersion", "FileVersion", "InformationalVersion"):
        if values.get(name) != ["$(ProductVersion)"]:
            raise RuntimeError(f"{name} must be exactly $(ProductVersion).")
    if values.get("IncludeSourceRevisionInInformationalVersion") != ["false"]:
        raise RuntimeError(
            "IncludeSourceRevisionInInformationalVersion must be false to prevent "
            "ProductVersion split-brain in signed SDK assemblies."
        )

    project = text(VSTO_PROJECT)
    require(
        project,
        (
            '<ProjectReference Include="..\\ChuanHoa.Client.Core\\ChuanHoa.Client.Core.csproj">',
            '<Target Name="ValidateUnifiedVersionContract" BeforeTargets="PrepareForBuild">',
            "'$(ApplicationVersion)' != '$(ProductVersion)'",
        ),
        "VSTO clean graph/unified version gate is incomplete",
    )
    if "<HintPath>$(ClientCoreOutput)</HintPath>" in project or "--no-restore" in project:
        raise RuntimeError("VSTO project still depends on a stale Client.Core binary/asset file.")

    contract = text(BUILD_CONTRACT)
    require(
        contract,
        (
            "Resolve-ChuanHoaApplicationVersion",
            "does not match ProductVersion",
            "Assert-ChuanHoaTrustedPublicKey",
            "private RSA material is forbidden",
            "Get-ChuanHoaSigningCertificate",
            "Assert-ChuanHoaManagedAssemblyVersion",
            "Assert-ChuanHoaManifestVersion",
            "Get-ChuanHoaMsBuild",
            "Get-ChuanHoaSignTool",
            "Invoke-ChuanHoaAuthenticodeSign",
        ),
        "Shared packaging contract is incomplete",
    )

    for path in (PUBLISH, LOCAL_PUBLISH):
        source = text(path)
        require(
            source,
            (
                "Resolve-ChuanHoaApplicationVersion",
                "CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256",
                "Get-ChuanHoaSigningCertificate",
                "Assert-ChuanHoaManagedAssemblyVersion",
                "Assert-ChuanHoaManifestVersion",
                "Get-ChuanHoaMsBuild",
                "Invoke-ChuanHoaAuthenticodeSign",
                "/p:TrustedSigningCertificateSha256=$signingCertificatePin",
                "/t:PublishOnly",
                "/p:BuildProjectReferences=false",
                "ChuanHoa.AddIn.Vsto\\bin\\",
                "ChuanHoa.Client.Core.dll",
            ),
            f"Publish script lacks version/signing gates: {path.name}",
        )
        if "Where-Object Subject -eq" in source:
            raise RuntimeError(f"Publish script trusts a certificate by subject: {path.name}")

    builder = text(BUILDER)
    require(
        builder,
        (
            "[string]$TrustedPublicKeyPath",
            "[string]$TrustedPublicKeySha256",
            "[string]$SigningCertificateSha256",
            "Assert-ChuanHoaTrustedPublicKey",
            "Get-ChuanHoaSigningCertificate",
            "/p:TrustedPublicKeySha256=",
            "/p:SigningCertificateSha256=",
            "Outer installer signer is not the pinned Development certificate",
            "$accessSmokeSignature = Invoke-ChuanHoaAuthenticodeSign",
        ),
        "Single-file Development builder does not use explicit pinned inputs",
    )
    if "LocalApplicationData')) 'ChuanHoa\\Development\\trusted-key.xml'" in builder:
        raise RuntimeError("Builder still silently consumes a machine-local trusted key.")

    bootstrapper_project = text(BOOTSTRAPPER_PROJECT)
    require(
        bootstrapper_project,
        (
            "ChuanHoa.DevelopmentInstaller.TrustedPublicKey.sha256",
            "ChuanHoa.DevelopmentInstaller.SigningCertificate.sha256",
            "TrustedPublicKeySha256 must be an explicit",
            "SigningCertificateSha256 must be an explicit",
        ),
        "Bootstrapper does not embed both build pins",
    )

    bootstrapper = text(BOOTSTRAPPER)
    require(
        bootstrapper,
        (
            "VerifyRunningInstallerSigner(signingCertificatePin)",
            "VerifyStagedPayload(stagingDirectory, version",
            "using (var transaction = new InstallTransaction(",
            "transaction.Rollback();",
            'InjectFault("after-trusted-certificate")',
            'InjectFault("after-trusted-key")',
            'InjectFault("after-access-smoke")',
            'InjectFault("after-current-switch")',
            'InjectFault("after-registry")',
            'InjectFault("after-verification")',
            "RestoreRegistryValues(",
            "cacheSnapshot.Restore",
            "RestoreDirectoryState",
            "does not match the pinned SHA-256",
            "EnsureNoLegacyClickOnceDevelopmentAddIn",
            "RunCertificateUtility(\"-f -user -addstore ",
            "RunCertificateUtility(\"-f -user -delstore ",
            "Environment.SpecialFolder.System), \"certutil.exe\"",
            "CreateNoWindow = true",
            "WindowStyle = ProcessWindowStyle.Hidden",
            "RunCapturedProcess(startInfo, 15000, \"certutil.exe\")",
            "process.BeginOutputReadLine()",
            "process.BeginErrorReadLine()",
            "process.WaitForExit(timeoutMilliseconds)",
            "HasPinnedCertificate(storeName, thumbprint, expectedSha256)",
            "RegisterAppsFeatures(version, installDirectory, cachedInstallerPath)",
            "AppsFeaturesRegistryPath",
            "QuietUninstallString",
            "ModifyPath",
            "CacheRunningInstaller",
            "GetInstallerCacheRoot",
            "ScheduleInstallerCacheCleanup",
            "AssertDirectChildDirectory",
            "VerifyInstallerSigner(cachedInstallerPath, signingCertificateSha256)",
            "VerifyAuthenticodeSignature(installerPath)",
            'DllImport("wintrust.dll"',
            "VerifySignedManifest(path, certificatePath)",
            "signedXml.CheckSignature(certificate, true)",
            "VerifyManifestFileHashes(path, Path.GetDirectoryName(path))",
            "SHA-256 payload không khớp manifest",
            '"ChuanHoa.Client.Core.dll"',
            '"ChuanHoa.DevelopmentAccessSmoke.exe"',
            "_appsFeaturesRegistrySnapshot",
            "RestoreInstallerCache",
        ),
        "Bootstrapper transaction/pinning contract is incomplete",
    )
    install_flow = bootstrapper.split("private static void UninstallDevelopmentChannel", 1)[0]
    if "UninstallClickOnceDevelopmentAddIn(baseDirectory);" in install_flow:
        raise RuntimeError("Install flow still performs a non-rollbackable ClickOnce uninstall.")
    if "store.Remove(certificate)" in bootstrapper or "store.Add(certificate)" in bootstrapper:
        raise RuntimeError("Development certificate operations can still trigger a Windows certificate-store dialog.")

    audit = text(AUDIT)
    require(
        audit,
        (
            "Resolve-ChuanHoaApplicationVersion",
            "Installer embedded pins do not match",
            "Installer payload does not match allowlist",
            "BEGIN PRIVATE KEY",
            "Assert-ManifestBytesVersion",
            "Assert-ChuanHoaManagedAssemblyVersion",
            "INNER_PE_AUTHENTICODE_NOT_VALID",
            "OWNED_INNER_PE_AUTHENTICODE_VALID",
            "APPS_FEATURES_AND_CACHED_UNINSTALL_IMPLEMENTED",
            "productionReady = $false",
        ),
        "Installer audit is missing a required release gate/gap",
    )
    lifecycle = text(LIFECYCLE)
    require(
        lifecycle,
        (
            "Get-Process WINWORD",
            "Assert-DictionaryPreserved",
            "TrustedKey = Get-FileSnapshot $trustedKeyPath",
            "AccessCache = @($accessCachePaths",
            "PersonalDictionary = Get-FileInventory $dictionaryDirectory",
            "after-trusted-certificate",
            "after-trusted-key",
            "after-access-smoke",
            "after-current-switch",
            "after-registry",
            "after-verification",
            "StateRestored = $true",
            "Invoke-RegisteredUninstall",
            "QuietUninstallString",
            "InstallerCache = Get-FileInventory $installerCacheRoot",
            "AppsFeaturesRegistry = Get-RegistryInventory $appsFeaturesRegistryPath",
            "Installed owned PE signer pin mismatch",
            "ChuanHoa.DevelopmentAccessSmoke.exe",
            "PASS_DEVELOPMENT_LIFECYCLE",
        ),
        "Reusable Development installer lifecycle verification is incomplete",
    )
    tamper_rejection = text(TAMPER_REJECTION)
    require(
        tamper_rejection,
        (
            "PASS_TAMPER_REJECTED_ON_TEMP_COPY",
            "VerifyStagedPayload",
            "VerifyInstallerSigner",
            "ChuanHoa.AddIn.Vsto.dll",
            "ChuanHoa.Client.Core.dll",
            "ChuanHoa.DevelopmentAccessSmoke.exe",
            "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
            "ChuanHoa.AddIn.Vsto.vsto",
            "ChuanHoa.AddIn.Vsto.dll.manifest",
            "installedStateTouched = $false",
        ),
        "Tamper-negative verification does not cover the signed installer/payload",
    )
    bootstrapper_allowlist = quoted_set_between(
        bootstrapper,
        "var expectedFiles = new HashSet<string>",
        "var actualFiles = new HashSet<string>",
    )
    audit_allowlist = quoted_set_between(
        audit,
        "$expectedEntries = @(",
        ")\n$forbiddenPatterns",
    )
    if bootstrapper_allowlist != audit_allowlist:
        raise RuntimeError(
            "Bootstrapper and installer-audit payload allowlists differ. "
            f"BootstrapperOnly={sorted(bootstrapper_allowlist - audit_allowlist)} "
            f"AuditOnly={sorted(audit_allowlist - bootstrapper_allowlist)}"
        )


if __name__ == "__main__":
    validate()
    print("DEVELOPMENT_PACKAGING: PASS")
