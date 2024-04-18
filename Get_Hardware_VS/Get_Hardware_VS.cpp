#pragma comment(lib, "wbemuuid.lib")

#include <iostream>
#include <Windows.h>
#include <WbemIdl.h>
#include <combaseapi.h>


#define LOG(x) \
    std::cout << x << std::endl;
#define error(code, stage) \
    if (FAILED(code)) \
    { \
        LOG("Failed at " << stage << ": 0x" << hex << code); \
        exit(-1); \
    } \
    else { \
        LOG("PASSED: " << stage); \
    };
#define _WIN32_DCOM

using namespace std;

void SartComConnection() {
    HRESULT status = CoInitializeEx(0, COINIT_MULTITHREADED);
    error(status, "initialize COM library");
}

void SetDefaultSecurity() {
    HRESULT status = CoInitializeSecurity(
        NULL,                        // Security descriptor    
        -1,                          // COM negotiates authentication service
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication level for proxies
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation level for proxies
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities of the client or server
        NULL);                       // Reserved

    error(status, "initialize security");
    if (FAILED(status)) CoUninitialize();
}

IWbemLocator* GrabInterfacePointer() {
    IWbemLocator *pLoc;
    HRESULT status = CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);
    error(status, "create IWbemLocator object");
    if (FAILED(status)) CoUninitialize();

    return pLoc;
}

IWbemServices* ConnectToProxy(IWbemLocator* pLoc) {
    IWbemServices* pSvc = 0;
    auto status = pLoc->ConnectServer(
        BSTR(L"ROOT\\DEFAULT"),  //namespace
        NULL,       // User name 
        NULL,       // User password
        0,         // Locale 
        NULL,     // Security flags
        0,         // Authority 
        0,        // Context object 
        &pSvc);   // IWbemServices proxy
    error(status, "Connect to WMI");
    if (FAILED(status)) {
        pLoc->Release();
        CoUninitialize();
    }
    
    return pSvc;
}

void SetProxySecurity(IWbemLocator* pLoc, IWbemServices* pSvc) {
    HRESULT status = CoSetProxyBlanket(pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );
    error(status, "set proxy blanket");
    if (FAILED(status)) {
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
    }
}

int main()
{
    SartComConnection();
    
    SetDefaultSecurity();
    
    IWbemLocator* pLoc = GrabInterfacePointer();
    IWbemServices* pSvc = ConnectToProxy(pLoc);

    SetProxySecurity(pLoc, pSvc);
    
    LOG("\n");

    IEnumWbemClassObject* pEnum = NULL;
    BSTR Language = SysAllocString(L"WQL");
    BSTR Query = SysAllocString(L"SELECT * FROM Win32_PnPEntity");
    auto status = pSvc->ExecQuery(Language, Query, WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnum);
    error(status, "Querying");

    

    SysFreeString(Query);
    SysFreeString(Language);

    
}