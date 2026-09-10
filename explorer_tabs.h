#ifndef EXPLORER_TABS_H
#define EXPLORER_TABS_H

#define CINTERFACE
#define COBJMACROS
#define _WIN32_IE 0x0700
#define _WIN32_DCOM

#include <windows.h>
#include <exdisp.h>

typedef struct ExplorerTab {
    IWebBrowser2 *browser;
    char *url;
    HWND top_level;
} ExplorerTab;

typedef struct ExplorerTabList {
    ExplorerTab *items;
    size_t count;
    size_t capacity;
} ExplorerTabList;

typedef struct WindowList {
    HWND *items;
    size_t count;
    size_t capacity;
} WindowList;

void explorer_tab_list_free(ExplorerTabList *tabs);
void window_list_free(WindowList *windows);
BOOL collect_explorer_tabs(ExplorerTabList *tabs, WindowList *window_order);
HWND find_shell_tab_host(HWND top_level);
BOOL create_tab_and_navigate(HWND first_window, HWND tab_host, const char *url,
                             size_t *known_tab_count, BOOL debug_output);

#endif
