### PSWindowsUpdate (unofficial fork)

PSWindowsUpdate is written by Michal Gajda ([mgajda83/PSWindowsUpdate](https://github.com/mgajda83/PSWindowsUpdate)).

This "fork" is a decompiled version of the original PSWindowsUpdate DLLs for the sole purpose of fixing bugs, improving security, and allowing transparency to an otherwise widely used awesome PowerShell module.

The project currently builds in Rider (and probably VS) however one must first compile a `BITSReference4_0` reference DLL. More information can be found on [Microsoft's page](https://docs.microsoft.com/en-us/windows/win32/bits/bits-dot-net) on the matter. I have included the build script featured on that page for convenience.