#include "explorer_tabs.h"

#include <oleauto.h>
#include <servprov.h>
#include <shlguid.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define WM_COMMAND_ID_NEW_TAB 0xA21B

static BOOL reserve_tabs(ExplorerTabList *list, size_t needed) {
    ExplorerTab *items;
    size_t capacity;
    if (needed <= list->capacity) return TRUE;
    capacity = list->capacity ? list->capacity * 2 : 8;
    while (capacity < needed) capacity *= 2;
    items = (ExplorerTab *)realloc(list->items, capacity * sizeof(*items));
    if (!items) return FALSE;
    list->items = items;
    list->capacity = capacity;
    return TRUE;
}

static BOOL window_list_contains(const WindowList *list, HWND window) {
    size_t i;
    for (i = 0; i < list->count; ++i) {
        if (list->items[i] == window) return TRUE;
    }
    return FALSE;
}

static BOOL window_list_append(WindowList *list, HWND window) {
    HWND *items;
    size_t capacity;
    if (window_list_contains(list, window)) return TRUE;
    if (list->count == list->capacity) {
        capacity = list->capacity ? list->capacity * 2 : 8;
        items = (HWND *)realloc(list->items, capacity * sizeof(*items));
        if (!items) return FALSE;
        list->items = items;
        list->capacity = capacity;
    }
    list->items[list->count++] = window;
    return TRUE;
}

void explorer_tab_list_free(ExplorerTabList *tabs) {
    size_t i;
    if (!tabs) return;
    for (i = 0; i < tabs->count; ++i) {
        if (tabs->items[i].browser) IWebBrowser2_Release(tabs->items[i].browser);
        free(tabs->items[i].url);
    }
    free(tabs->items);
    memset(tabs, 0, sizeof(*tabs));
}

void window_list_free(WindowList *windows) {
    if (!windows) return;
    free(windows->items);
    memset(windows, 0, sizeof(*windows));
}

static char *bstr_to_ansi(BSTR value) {
    UINT wide_length;
    int bytes;
    char *result;
    if (!value) return NULL;
    wide_length = SysStringLen(value);
    bytes = WideCharToMultiByte(CP_ACP, 0, value, (int)wide_length, NULL, 0, NULL, NULL);
    result = (char *)malloc((size_t)bytes + 1);
    if (!result) return NULL;
    if (bytes) WideCharToMultiByte(CP_ACP, 0, value, (int)wide_length, result, bytes, NULL, NULL);
    result[bytes] = '\0';
    return result;
}

static BSTR ansi_to_bstr(const char *value) {
    int length;
    BSTR result;
    if (!value) return NULL;
    length = MultiByteToWideChar(CP_ACP, 0, value, -1, NULL, 0);
    if (length <= 0) return NULL;
    result = SysAllocStringLen(NULL, (UINT)(length - 1));
    if (!result) return NULL;
    MultiByteToWideChar(CP_ACP, 0, value, -1, result, length);
    return result;
}

static BOOL get_dispatch_property(IDispatch *dispatch, const WCHAR *name, VARIANT *result) {
    LPOLESTR names[1];
    DISPID id;
    DISPPARAMS parameters;
    HRESULT hr;
    if (!dispatch || !name || !result) return FALSE;
    VariantInit(result);
    names[0] = (LPOLESTR)name;
    hr = IDispatch_GetIDsOfNames(dispatch, &IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return FALSE;
    memset(&parameters, 0, sizeof(parameters));
    hr = IDispatch_Invoke(dispatch, id, &IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYGET, &parameters, result, NULL, NULL);
    if (FAILED(hr)) {
        VariantClear(result);
        return FALSE;
    }
    return TRUE;
}

static char *extract_explorer_url(IWebBrowser2 *browser) {
    BSTR location = NULL;
    char *url = NULL;
    IDispatch *document = NULL, *folder = NULL, *self = NULL;
    VARIANT value;

    if (SUCCEEDED(IWebBrowser2_get_LocationURL(browser, &location)) && location) {
        url = bstr_to_ansi(location);
        SysFreeString(location);
    }
    if (url && url[0]) return url;
    free(url);
    url = NULL;

    if (FAILED(IWebBrowser2_get_Document(browser, &document)) || !document) return NULL;
    if (!get_dispatch_property(document, L"Folder", &value)) goto cleanup;
    if (V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value)) {
        folder = V_DISPATCH(&value);
        IDispatch_AddRef(folder);
    }
    VariantClear(&value);
    if (!folder || !get_dispatch_property(folder, L"Self", &value)) goto cleanup;
    if (V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value)) {
        self = V_DISPATCH(&value);
        IDispatch_AddRef(self);
    }
    VariantClear(&value);
    if (!self || !get_dispatch_property(self, L"Path", &value)) goto cleanup;
    if (V_VT(&value) == VT_BSTR && V_BSTR(&value)) {
        char *path = bstr_to_ansi(V_BSTR(&value));
        if (path && strncmp(path, "::", 2) == 0) {
            size_t length = strlen(path) + 7;
            url = (char *)malloc(length);
            if (url) snprintf(url, length, "shell:%s", path);
        } else if (path && strncmp(path, "shell::", 7) == 0) {
            url = path;
            path = NULL;
        }
        free(path);
    }
    VariantClear(&value);

cleanup:
    if (self) IDispatch_Release(self);
    if (folder) IDispatch_Release(folder);
    IDispatch_Release(document);
    return url;
}

static HRESULT navigate_browser(IWebBrowser2 *browser, const char *url) {
    VARIANT target, empty;
    HRESULT hr;
    if (!browser) return E_POINTER;
    VariantInit(&target);
    VariantInit(&empty);
    V_VT(&target) = VT_BSTR;
    V_BSTR(&target) = ansi_to_bstr(url);
    if (!V_BSTR(&target)) return E_OUTOFMEMORY;
    hr = IWebBrowser2_Navigate2(browser, &target, &empty, &empty, &empty, &empty);
    VariantClear(&target);
    return hr;
}

BOOL collect_explorer_tabs(ExplorerTabList *tabs, WindowList *window_order) {
    IShellWindows *shell_windows = NULL;
    long count = 0, i;
    HRESULT hr;
    explorer_tab_list_free(tabs);
    window_list_free(window_order);
    hr = CoCreateInstance(&CLSID_ShellWindows, NULL, CLSCTX_ALL,
                          &IID_IShellWindows, (void **)&shell_windows);
    if (FAILED(hr)) return FALSE;
    if (FAILED(IShellWindows_get_Count(shell_windows, &count))) goto failure;

    for (i = 0; i < count; ++i) {
        VARIANT index;
        IDispatch *dispatch = NULL;
        IWebBrowser2 *browser = NULL;
        IServiceProvider *provider = NULL;
        IShellBrowser *shell_browser = NULL;
        SHANDLE_PTR handle = 0;
        HWND top_level;
        char *url;
        VariantInit(&index);
        V_VT(&index) = VT_I4;
        V_I4(&index) = i;
        if (FAILED(IShellWindows_Item(shell_windows, index, &dispatch)) || !dispatch) continue;
        if (FAILED(IDispatch_QueryInterface(dispatch, &IID_IWebBrowser2, (void **)&browser)) || !browser) {
            IDispatch_Release(dispatch);
            continue;
        }
        IDispatch_Release(dispatch);
        if (FAILED(IWebBrowser2_QueryInterface(browser, &IID_IServiceProvider, (void **)&provider)) ||
            !provider || FAILED(IServiceProvider_QueryService(provider, &SID_STopLevelBrowser,
                                                               &IID_IShellBrowser, (void **)&shell_browser))) {
            if (provider) IServiceProvider_Release(provider);
            IWebBrowser2_Release(browser);
            continue;
        }
        IShellBrowser_Release(shell_browser);
        IServiceProvider_Release(provider);
        if (FAILED(IWebBrowser2_get_HWND(browser, &handle)) || !handle) {
            IWebBrowser2_Release(browser);
            continue;
        }
        top_level = (HWND)(INT_PTR)handle;
        url = extract_explorer_url(browser);
        if (!url) url = _strdup("");
        if (!url || !window_list_append(window_order, top_level) || !reserve_tabs(tabs, tabs->count + 1)) {
            free(url);
            IWebBrowser2_Release(browser);
            goto failure;
        }
        tabs->items[tabs->count].browser = browser;
        tabs->items[tabs->count].url = url;
        tabs->items[tabs->count].top_level = top_level;
        ++tabs->count;
    }
    IShellWindows_Release(shell_windows);
    return TRUE;

failure:
    IShellWindows_Release(shell_windows);
    explorer_tab_list_free(tabs);
    window_list_free(window_order);
    return FALSE;
}

typedef struct FindTabHostData { HWND target; } FindTabHostData;

static BOOL CALLBACK enum_find_tab_host(HWND window, LPARAM parameter) {
    FindTabHostData *data = (FindTabHostData *)parameter;
    char class_name[256] = {0};
    if (GetClassNameA(window, class_name, 255) && strcmp(class_name, "ShellTabWindowClass") == 0) {
        data->target = window;
        return FALSE;
    }
    return TRUE;
}

HWND find_shell_tab_host(HWND top_level) {
    FindTabHostData data = {0};
    EnumChildWindows(top_level, enum_find_tab_host, (LPARAM)&data);
    return data.target;
}

static BOOL browser_is_known(IWebBrowser2 *browser, IWebBrowser2 **known, size_t count) {
    size_t i;
    for (i = 0; i < count; ++i) if (known[i] == browser) return TRUE;
    return FALSE;
}

BOOL create_tab_and_navigate(HWND first_window, HWND tab_host, const char *url,
                             size_t *known_tab_count, BOOL debug_output) {
    ExplorerTabList baseline = {0};
    WindowList windows = {0};
    IWebBrowser2 **known = NULL;
    size_t baseline_count = known_tab_count ? *known_tab_count : 0;
    size_t known_count = 0, i;
    DWORD waited = 0;
    BOOL success = FALSE;

    if (!first_window || !tab_host || !url || !url[0]) return FALSE;
    if (!collect_explorer_tabs(&baseline, &windows)) return FALSE;
    window_list_free(&windows);
    known = (IWebBrowser2 **)malloc(baseline.count * sizeof(*known));
    if (baseline.count && !known) goto cleanup;
    baseline_count = 0;
    for (i = 0; i < baseline.count; ++i) {
        if (baseline.items[i].top_level == first_window) {
            ++baseline_count;
            known[known_count++] = baseline.items[i].browser;
        }
    }
    if (known_tab_count) *known_tab_count = baseline_count;
    if (debug_output) printf("[debug] Baseline tab count for first window: %zu\n", baseline_count);
    SendMessageA(tab_host, WM_COMMAND, (WPARAM)WM_COMMAND_ID_NEW_TAB, 0);

    while (waited <= 8000) {
        ExplorerTabList current = {0};
        IWebBrowser2 *candidate = NULL;
        size_t current_count = 0;
        if (!collect_explorer_tabs(&current, &windows)) {
            Sleep(300); waited += 300; continue;
        }
        window_list_free(&windows);
        for (i = 0; i < current.count; ++i) {
            if (current.items[i].top_level == first_window) {
                ++current_count;
                if (!candidate && !browser_is_known(current.items[i].browser, known, known_count))
                    candidate = current.items[i].browser;
            }
        }
        if (candidate && current_count > baseline_count) {
            HRESULT hr = navigate_browser(candidate, url);
            if (SUCCEEDED(hr)) {
                success = TRUE;
                if (known_tab_count) *known_tab_count = current_count;
            }
            explorer_tab_list_free(&current);
            break;
        }
        explorer_tab_list_free(&current);
        Sleep(300);
        waited += 300;
    }

cleanup:
    free(known);
    explorer_tab_list_free(&baseline);
    window_list_free(&windows);
    return success;
}
