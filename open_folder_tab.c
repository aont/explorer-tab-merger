#include "explorer_tabs.h"

#include <shellapi.h>
#include <stdio.h>
#include <stdlib.h>

static char *normalize_folder_path(const char *input) {
    DWORD required, written;
    char *path;
    if (!input || !input[0]) return NULL;
    required = GetFullPathNameA(input, 0, NULL, NULL);
    if (!required) return _strdup(input);
    path = (char *)malloc(required);
    if (!path) return NULL;
    written = GetFullPathNameA(input, required, path, NULL);
    if (!written || written >= required) {
        free(path);
        return _strdup(input);
    }
    return path;
}

int main(int argc, char **argv) {
    ExplorerTabList tabs = {0};
    WindowList windows = {0};
    char *target_path;
    HWND first_window, tab_host;
    HRESULT hr;
    size_t known_count = 0;
    HINSTANCE shell_result;

    if (argc < 2) { fprintf(stderr, "Usage: open_folder_tab.exe <folder path>\n"); return 1; }
    target_path = normalize_folder_path(argv[1]);
    if (!target_path) { fprintf(stderr, "Empty folder path or insufficient memory.\n"); return 1; }
    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { fprintf(stderr, "CoInitializeEx failed: 0x%08lx\n", (unsigned long)hr); free(target_path); return 1; }

    if (!collect_explorer_tabs(&tabs, &windows) || !windows.count) {
        printf("No Explorer window found; launching folder via ShellExecute.\n");
        shell_result = ShellExecuteA(NULL, "open", target_path, NULL, NULL, SW_SHOWNORMAL);
        explorer_tab_list_free(&tabs); window_list_free(&windows);
        CoUninitialize(); free(target_path);
        return (INT_PTR)shell_result <= 32 ? 2 : 0;
    }
    first_window = windows.items[0];
    explorer_tab_list_free(&tabs);
    window_list_free(&windows);
    tab_host = find_shell_tab_host(first_window);
    if (!tab_host) {
        fprintf(stderr, "Could not find ShellTabWindowClass in the first window.\n");
        CoUninitialize(); free(target_path); return 3;
    }
    if (!create_tab_and_navigate(first_window, tab_host, target_path, &known_count, FALSE)) {
        fprintf(stderr, "Failed to create or navigate new tab; falling back to ShellExecute.\n");
        ShellExecuteA(NULL, "open", target_path, NULL, NULL, SW_SHOWNORMAL);
    }
    CoUninitialize();
    free(target_path);
    return 0;
}
