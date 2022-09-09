### PSWindowsUpdate (unofficial fork)

PSWindowsUpdate is written by Michal Gajda ([mgajda83/PSWindowsUpdate](https://github.com/mgajda83/PSWindowsUpdate)).

This "fork" is a **decompiled** version of the original PSWindowsUpdate v2.2.0.3 DLLs for the sole purpose of fixing bugs, improving security, and allowing transparency to an otherwise **awesome** PowerShell module used by [millions of sysadmins](https://www.powershellgallery.com/packages/PSWindowsUpdate). Currently, this repo contains a bug fix for `Remove-WindowsUpdate -UpdateID` (please see the git history for more information). 

The project builds without any issues in Rider however its use in production environments is discouraged at this time. Additional source code reviews & cleanup of leftover artifacts are still needed.

For anyone wishing to build this project, one must first [compile](https://docs.microsoft.com/en-us/windows/win32/bits/bits-dot-net) a `BITSReference4_0` reference DLL using Type Libraries from the Windows 10+ SDK. For convenience, I have included the build script featured on Microsoft's page.
