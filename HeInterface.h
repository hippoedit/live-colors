#pragma once

#include <atlbase.h>
#include <atlcom.h>
#include <atlstr.h>

using namespace ATL;
#include <dispex.h>
#if _WIN64
#import ".\..\HippoEdit\HippoEdit64.tlb" raw_interfaces_only, raw_native_types, no_namespace, no_smart_pointers, named_guids, no_implementation, auto_search exclude("IDispatchEx", "IServiceProvider")
//#import ".\..\bin64\HippoEdit.exe" raw_interfaces_only, raw_native_types, no_namespace, no_smart_pointers, named_guids, no_implementation, auto_search exclude("IDispatchEx", "IServiceProvider")
#else
#import ".\..\HippoEdit\HippoEdit.tlb" raw_interfaces_only, raw_native_types, no_namespace, no_smart_pointers, named_guids, no_implementation, auto_search exclude("IDispatchEx", "IServiceProvider")
//#import ".\..\binU\Hippoedit.exe" raw_interfaces_only, raw_native_types, no_namespace, no_smart_pointers, named_guids, no_implementation, auto_search exclude("IDispatchEx", "IServiceProvider")
#endif

typedef long LINE_T;