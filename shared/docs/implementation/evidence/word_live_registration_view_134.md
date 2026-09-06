# Word live registration discrepancy — 2026-09-06

Observed the user's failing Word process 20164 (Explorer parent 10432), without closing the document.

- Installed release: 1.0.0.134.
- External HKCU registry query: ChuanHoa.AddIn.Vsto LoadBehavior = 3.
- Word COMAddIn.Connect: false; accessibility tree and screenshot confirm no Chuẩn hóa tab.
- Word Options lists Chuẩn hóa under Inactive, COM Add-ins checkbox unchecked.
- Read the identical HKCU key through the existing Word process using System.PrivateProfileString (COM indexed property, not VBA): LoadBehavior = 2.
- Manifest from inside Word matches the external registry's Current deployment URI.
- Set only that LoadBehavior through Word's System.PrivateProfileString to 3; immediate readback = 3.
- Explicit COMAddIn.Connect = true then succeeded. Readback remained 3.
- Accessibility tree and screenshot of the same Document1 window now show the Chuẩn hóa tab.

No document content read for diagnosis or changed; Word not closed. No security policy changed. No installer or product DLL rebuilt.

The differing registry views are proven; the underlying virtualization mechanism is not yet proven. Persistence after the user closes and relaunches Word from Explorer remains unverified. Do not label this a completed permanent fix until that path passes.
