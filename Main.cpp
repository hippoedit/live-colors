
#include "PlugIn.h"
#include "..\shared\PluginFactory.h"

class LiveColorsModule : public CAtlDllModuleT<LiveColorsModule>
{
} _AtlModule;

// DLL Entry Point
EXTERN_C BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    return _AtlModule.DllMain(dwReason, lpReserved);
}

STDAPI DllCanUnloadNow()
{
    return _AtlModule.DllCanUnloadNow();
}

STDAPI DllGetPluginFactory(LPVOID* ppv)
{
	return CPluginFactoryExt::CreateInstanceEx()->QueryInterface(__uuidof(IPluginFactory), ppv);
}
