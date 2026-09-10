#include "explorer_tabs.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef struct StringList { char **items; size_t count; } StringList;

static void string_list_free(StringList *list) {
    size_t i;
    for (i = 0; i < list->count; ++i) free(list->items[i]);
    free(list->items);
}

static BOOL string_list_append(StringList *list, const char *value) {
    char **items = (char **)realloc(list->items, (list->count + 1) * sizeof(*items));
    char *copy;
    if (!items) return FALSE;
    list->items = items;
    copy = _strdup(value);
    if (!copy) return FALSE;
    list->items[list->count++] = copy;
    return TRUE;
}

int main(void) {
    ExplorerTabList tabs = {0};
    WindowList windows = {0}, close_windows = {0};
    StringList urls = {0};
    HWND first_window, tab_host;
    size_t known_count = 0, success_count = 0, i, j;
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    int result = 0;

    if (FAILED(hr)) { fprintf(stderr, "CoInitializeEx failed: 0x%08lx\n", (unsigned long)hr); return 1; }
    if (!collect_explorer_tabs(&tabs, &windows)) {
        fprintf(stderr, "Failed to enumerate Explorer tabs.\n"); result = 2; goto cleanup;
    }
    if (!windows.count) { printf("No Explorer windows detected.\n"); goto cleanup; }
    first_window = windows.items[0];
    for (i = 0; i < tabs.count; ++i) {
        if (tabs.items[i].top_level == first_window) {
            ++known_count;
        } else {
            BOOL found = FALSE;
            if (tabs.items[i].url[0] && !string_list_append(&urls, tabs.items[i].url)) {
                fprintf(stderr, "Out of memory.\n"); result = 4; goto cleanup;
            }
            for (j = 0; j < close_windows.count; ++j)
                if (close_windows.items[j] == tabs.items[i].top_level) found = TRUE;
            if (!found) {
                HWND *items = (HWND *)realloc(close_windows.items, (close_windows.count + 1) * sizeof(*items));
                if (!items) { result = 4; goto cleanup; }
                close_windows.items = items;
                close_windows.items[close_windows.count++] = tabs.items[i].top_level;
            }
        }
    }
    explorer_tab_list_free(&tabs);
    if (!urls.count) { printf("Nothing to merge.\n"); goto cleanup; }
    tab_host = find_shell_tab_host(first_window);
    if (!tab_host) { fprintf(stderr, "Could not find ShellTabWindowClass in the first window.\n"); result = 3; goto cleanup; }
    printf("Merging %zu tab(s) into the first window...\n", urls.count);
    for (i = 0; i < urls.count; ++i) {
        if (create_tab_and_navigate(first_window, tab_host, urls.items[i], &known_count, TRUE)) ++success_count;
        else fprintf(stderr, "[warn] Failed to create tab for: %s\n", urls.items[i]);
    }
    for (i = 0; i < close_windows.count; ++i)
        if (close_windows.items[i] && close_windows.items[i] != first_window)
            SendMessageA(close_windows.items[i], WM_CLOSE, 0, 0);
    printf("Completed. %zu tab(s) moved.\n", success_count);

cleanup:
    explorer_tab_list_free(&tabs);
    window_list_free(&windows);
    window_list_free(&close_windows);
    string_list_free(&urls);
    CoUninitialize();
    return result;
}
