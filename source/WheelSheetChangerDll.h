#ifndef HOOK_H
#define HOOK_H

#include <windows.h>

#ifdef DLL_EXPORTS
#define DLL_DECLSPEC extern "C" __declspec(dllexport)
#else
#define DLL_DECLSPEC extern "C" __declspec(dllimport)
#endif

#define IS_ENABLE_EXCEL 0
#define IS_ENABLE_EXPLORER 1

DLL_DECLSPEC BOOL __stdcall StartHook();
DLL_DECLSPEC BOOL __stdcall EndHook();
DLL_DECLSPEC void __stdcall HookSettingChange(int, BOOL);

#endif
