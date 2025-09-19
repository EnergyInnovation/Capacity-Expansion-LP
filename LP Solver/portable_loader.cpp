// portable_loader.cpp
// Make the add-on portable: search the model folder (current directory)
// and the DLL's own folder for dependent DLLs like glpk.dll.

#define UNICODE
#define _UNICODE
#include <windows.h>
#include <Shlwapi.h>
#include <string>          // <-- needed for std::wstring
#pragma comment(lib, "Shlwapi.lib")

#ifndef LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
#define LOAD_LIBRARY_SEARCH_DEFAULT_DIRS 0x00001000
#define LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR 0x00000100
#define LOAD_LIBRARY_SEARCH_USER_DIRS    0x00000400
#endif

// Add a directory to Windows' safe DLL search path if available; otherwise fall back.
static void add_search_dir(const wchar_t* dir) {
    if (!dir || !*dir) return;

    HMODULE k32 = GetModuleHandleW(L"kernel32.dll");
    auto pAddDllDirectory =
        (decltype(&AddDllDirectory))GetProcAddress(k32, "AddDllDirectory");
    auto pSetDefaultDllDirectories =
        (decltype(&SetDefaultDllDirectories))GetProcAddress(k32, "SetDefaultDllDirectories");

    if (pSetDefaultDllDirectories && pAddDllDirectory) {
        // Use modern, safe search semantics and allow user dirs.
        pSetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS |
                                  LOAD_LIBRARY_SEARCH_USER_DIRS);
        pAddDllDirectory(dir);
    } else {
        // Legacy fallback: prepend to PATH for this process and set the DLL dir.
        wchar_t oldPath[32767] = {0};
        GetEnvironmentVariableW(L"PATH", oldPath, 32767);
        std::wstring newPath = std::wstring(dir) + L";" + oldPath;
        SetEnvironmentVariableW(L"PATH", newPath.c_str());
        SetDllDirectoryW(dir);
    }
}

// Try to load a DLL from a specific directory if present (ignore failures).
static void preload_if_present(const wchar_t* dir, const wchar_t* dllName) {
    if (!dir || !*dir || !dllName || !*dllName) return;
    wchar_t full[MAX_PATH] = {0};
    PathCombineW(full, dir, dllName);
    LoadLibraryW(full);
}

static void init_portable_loader(HINSTANCE hInst) {
    // Folder containing THIS add-on DLL (…\LP Solver or wherever the DLL sits)
    wchar_t selfPath[MAX_PATH] = {0};
    GetModuleFileNameW(hInst, selfPath, MAX_PATH);
    PathRemoveFileSpecW(selfPath);

    // Model folder: Vensim normally sets CWD to the folder containing the .mdl
    wchar_t modelDir[MAX_PATH] = {0};
    GetCurrentDirectoryW(MAX_PATH, modelDir);

    // Add both directories to the loader's search list
    add_search_dir(selfPath);
    add_search_dir(modelDir);

    // Opportunistically load GLPK from either folder if present
    preload_if_present(selfPath, L"glpk.dll");
    preload_if_present(modelDir, L"glpk.dll");

    // Also ask the loader to resolve glpk.dll using the new safe search paths
    LoadLibraryExW(L"glpk.dll", NULL,
        LOAD_LIBRARY_SEARCH_DEFAULT_DIRS |
        LOAD_LIBRARY_SEARCH_USER_DIRS   |
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR);
}

// IMPORTANT: Only one DllMain in the whole DLL.
// If you already define DllMain elsewhere, remove it there and keep this one.
extern "C" BOOL WINAPI DllMain(HINSTANCE hInst, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        init_portable_loader(hInst);
    }
    return TRUE;
}
