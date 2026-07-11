@echo off
REM =========================================================================
REM  compile_tlb.bat — Compile czWebview.odl to czWebview.tlb
REM  Uses MIDL.EXE from Windows SDK (/mktyplib203 mode)
REM =========================================================================
setlocal

pushd "%~dp0"

echo.
echo  czWebview TLB Compiler
echo  ======================
echo.

REM --- Find MIDL.EXE ---
set "MIDL_EXE="

where midl.exe >nul 2>&1 && for /f "tokens=*" %%i in ('where midl.exe') do set "MIDL_EXE=%%i"
if defined MIDL_EXE goto :found_midl

for %%V in (10.0.28000.0 10.0.26100.0 10.0.22621.0 10.0.22000.0 10.0.20348.0 10.0.19041.0 10.0.18362.0 10.0.17763.0 10.0.17134.0) do (
    if exist "C:\Program Files (x86)\Windows Kits\10\bin\%%V\x86\midl.exe" (
        set "MIDL_EXE=C:\Program Files (x86)\Windows Kits\10\bin\%%V\x86\midl.exe"
        goto :found_midl
    )
)

echo  [ERROR] MIDL.EXE not found!
echo  Install Windows SDK from:
echo  https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
pause
goto :eof

:found_midl
echo  MIDL: %MIDL_EXE%

if not exist czWebview.odl (
    echo  [ERROR] czWebview.odl not found in %CD%
    pause
    goto :eof
)

REM --- Create minimal stubs so MIDL can resolve IUnknown without cl.exe ---
echo  Creating temporary stubs...

echo typedef long HRESULT; > oaidl.idl
echo typedef unsigned long ULONG; >> oaidl.idl
echo typedef hyper CURRENCY; >> oaidl.idl
echo typedef long VARIANT_BOOL; >> oaidl.idl
echo typedef unsigned short USHORT; >> oaidl.idl
echo typedef struct tagVARIANT { USHORT vt; USHORT r1; USHORT r2; USHORT r3; hyper v; } VARIANT; >> oaidl.idl
echo typedef long BSTR; >> oaidl.idl
echo typedef long LPSTR; >> oaidl.idl
echo typedef long LPWSTR; >> oaidl.idl
echo typedef long SAFEARRAY; >> oaidl.idl
echo [local, object, uuid(00000000-0000-0000-C000-000000000046)] interface IUnknown { HRESULT QueryInterface([in] long riid, [out] long* ppv); ULONG AddRef(); ULONG Release(); } >> oaidl.idl
echo [local, object, uuid(00020400-0000-0000-C000-000000000046)] interface IDispatch : IUnknown { HRESULT GetTypeInfoCount([out] long* p); } >> oaidl.idl

echo // stub > ocidl.idl

REM --- Compile ---
echo  Compiling...
echo.

"%MIDL_EXE%" /mktyplib203 /nocpp /I . /no_def_idir /tlb czWebview.tlb czWebview.odl

echo.

REM --- Verify ---
if exist czWebview.tlb (
    for %%F in (czWebview.tlb) do (
        echo  [SUCCESS] czWebview.tlb created!
        echo  Size: %%~zF bytes
    )
) else (
    echo  [ERROR] Compilation failed! TLB was not created.
)

REM --- Cleanup ---
if exist oaidl.idl del oaidl.idl
if exist ocidl.idl del ocidl.idl

echo.
popd
pause
