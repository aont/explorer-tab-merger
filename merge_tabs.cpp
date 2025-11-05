// merge_tabs.cpp - Merge Explorer tabs into the first window (ANSI, MinGW-w64 friendly)
// Build: g++ merge_tabs.cpp -std=c++17 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32

#define _WIN32_IE 0x0700
#define _WIN32_DCOM

#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlguid.h>
#include <exdisp.h>
#include <shldisp.h>
#include <servprov.h>
#include <oleauto.h>
#include <shlwapi.h>

#include <vector>
#include <string>
#include <algorithm>
#include <iostream>
#include <iomanip>

static const UINT WM_COMMAND_ID_NEW_TAB = 0xA21B; // same as newtab.cpp (undocumented)

struct TabInfo {
    IWebBrowser2* browser; // holds one reference; caller must Release
    std::string parsingName;
    std::string locationUrl;
    HWND topLevel;
};

// --- Helpers for BSTR/ANSI ---
static BSTR AnsiToBSTR(const char* s) {
    if (!s) return nullptr;
    int wlen = MultiByteToWideChar(CP_ACP, 0, s, -1, nullptr, 0);
    if (wlen <= 0) return nullptr;
    BSTR b = SysAllocStringLen(nullptr, (UINT)(wlen - 1));
    if (!b) return nullptr;
    MultiByteToWideChar(CP_ACP, 0, s, -1, b, wlen);
    return b;
}

static std::string BSTRtoAnsi(BSTR b) {
    if (!b) return std::string();
    UINT lenW = SysStringLen(b);
    if (lenW == 0) return std::string();
    int bytes = WideCharToMultiByte(CP_ACP, 0, b, lenW, nullptr, 0, nullptr, nullptr);
    std::string out(bytes, '\0');
    WideCharToMultiByte(CP_ACP, 0, b, lenW, &out[0], bytes, nullptr, nullptr);
    return out;
}

static std::string WideToAnsi(const wchar_t* ws) {
    if (!ws || !*ws) return std::string();
    int bytes = WideCharToMultiByte(CP_ACP, 0, ws, -1, nullptr, 0, nullptr, nullptr);
    if (bytes <= 0) return std::string();
    std::string out(bytes - 1, '\0');
    WideCharToMultiByte(CP_ACP, 0, ws, -1, &out[0], bytes, nullptr, nullptr);
    return out;
}

static std::string GetParsingNameFromShellBrowser(IShellBrowser* sb) {
    if (!sb) return std::string();

    IShellView* sv = nullptr;
    if (FAILED(sb->QueryActiveShellView(&sv)) || !sv) {
        return std::string();
    }

    IFolderView* fv = nullptr;
    HRESULT hr = sv->QueryInterface(IID_PPV_ARGS(&fv));
    if (FAILED(hr) || !fv) {
        sv->Release();
        return std::string();
    }

    IPersistFolder2* pf2 = nullptr;
    hr = fv->GetFolder(IID_PPV_ARGS(&pf2));
    fv->Release();
    if (FAILED(hr) || !pf2) {
        sv->Release();
        return std::string();
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    hr = pf2->GetCurFolder(&pidl);
    pf2->Release();
    sv->Release();
    if (FAILED(hr) || !pidl) {
        if (pidl) {
            CoTaskMemFree(pidl);
        }
        return std::string();
    }

    std::wstring buffer(32768, L'\0');
    hr = PidlToParsingName(pidl, buffer.data(), static_cast<UINT>(buffer.size()));
    std::string result;
    if (SUCCEEDED(hr) && !buffer.empty() && buffer[0] != L'\0') {
        result = WideToAnsi(buffer.data());
    }

    CoTaskMemFree(pidl);
    return result;
}

static HRESULT NavigateBrowser(IWebBrowser2* wb, const std::string& url) {
    if (!wb) return E_POINTER;
    VARIANT vURL; VariantInit(&vURL);
    VARIANT vEmpty; VariantInit(&vEmpty);

    vURL.vt = VT_BSTR;
    vURL.bstrVal = AnsiToBSTR(url.c_str());
    if (!vURL.bstrVal) {
        VariantClear(&vURL);
        return E_OUTOFMEMORY;
    }

    HRESULT hr = wb->Navigate2(&vURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
    VariantClear(&vURL);
    VariantClear(&vEmpty);
    return hr;
}

// --- Collect Explorer tabs ---
static bool CollectExplorerTabs(std::vector<TabInfo>& tabs, std::vector<HWND>& windowOrder) {
    tabs.clear();
    windowOrder.clear();

    IShellWindows* pSW = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW)))) {
        return false;
    }

    long count = 0;
    if (FAILED(pSW->get_Count(&count))) {
        pSW->Release();
        return false;
    }

    for (long i = 0; i < count; ++i) {
        VARIANT vIdx; VariantInit(&vIdx);
        vIdx.vt = VT_I4;
        vIdx.lVal = i;

        IDispatch* pDisp = nullptr;
        if (FAILED(pSW->Item(vIdx, &pDisp)) || !pDisp) {
            VariantClear(&vIdx);
            continue;
        }
        VariantClear(&vIdx);

        IWebBrowser2* pWB = nullptr;
        if (FAILED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB)) || !pWB) {
            pDisp->Release();
            continue;
        }
        pDisp->Release();

        bool isExplorer = false;
        std::string parsingName;
        IServiceProvider* sp = nullptr;
        if (SUCCEEDED(pWB->QueryInterface(IID_IServiceProvider, (void**)&sp)) && sp) {
            IShellBrowser* sb = nullptr;
            if (SUCCEEDED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb))) && sb) {
                isExplorer = true;
                parsingName = GetParsingNameFromShellBrowser(sb);
                sb->Release();
            }
            sp->Release();
        }
        if (!isExplorer) {
            pWB->Release();
            continue;
        }

        SHANDLE_PTR handle = 0;
        HWND topLevel = nullptr;
        if (SUCCEEDED(pWB->get_HWND(&handle))) {
            topLevel = (HWND)handle;
        }
        if (!topLevel) {
            pWB->Release();
            continue;
        }

        BSTR bUrl = nullptr;
        std::string locationUrl;
        if (SUCCEEDED(pWB->get_LocationURL(&bUrl)) && bUrl) {
            locationUrl = BSTRtoAnsi(bUrl);
            SysFreeString(bUrl);
        }

        if (std::find(windowOrder.begin(), windowOrder.end(), topLevel) == windowOrder.end()) {
            windowOrder.push_back(topLevel);
        }

        std::string target = !parsingName.empty() ? parsingName : locationUrl;
        std::cout << "[debug] Explorer tab found: top-level HWND=0x" << std::hex << std::setw(0)
                  << reinterpret_cast<uintptr_t>(topLevel)
                  << ", IWebBrowser2=" << pWB
                  << ", ParsingName=" << parsingName
                  << ", LocationURL=" << locationUrl
                  << ", Target=" << target << std::dec << "\n";

        tabs.push_back({ pWB, parsingName, locationUrl, topLevel });
    }

    pSW->Release();
    return true;
}

// --- Find ShellTabWindowClass inside a top-level Explorer window ---
struct FindTabHostData {
    HWND target;
};

static BOOL CALLBACK EnumFindTabHost(HWND hwnd, LPARAM lParam) {
    auto* data = reinterpret_cast<FindTabHostData*>(lParam);
    if (!data) return FALSE;

    char cls[256] = {0};
    if (GetClassNameA(hwnd, cls, 255) && std::string(cls) == "ShellTabWindowClass") {
        data->target = hwnd;
        return FALSE; // stop enumeration
    }

    EnumChildWindows(hwnd, EnumFindTabHost, lParam);
    return data->target ? FALSE : TRUE;
}

static HWND FindShellTabHost(HWND topLevel) {
    FindTabHostData data{};
    EnumChildWindows(topLevel, EnumFindTabHost, reinterpret_cast<LPARAM>(&data));
    return data.target;
}

// --- Create new tab in the first window and navigate ---
static bool CreateTabAndNavigate(HWND firstWindow, HWND tabHost, const std::string& url, size_t& knownTabCount) {
    if (!firstWindow || !tabHost || url.empty()) return false;

    size_t baselineCount = knownTabCount;
    {
        std::vector<TabInfo> beforeTabs;
        std::vector<HWND> beforeWindows;
        if (CollectExplorerTabs(beforeTabs, beforeWindows)) {
            size_t currentCount = 0;
            for (auto& t : beforeTabs) {
                if (t.topLevel == firstWindow) {
                    ++currentCount;
                }
            }
            baselineCount = currentCount;
            knownTabCount = currentCount;
            std::cout << "[debug] Baseline tab count for first window: " << baselineCount << "\n";
        }
        for (auto& t : beforeTabs) {
            if (t.browser) t.browser->Release();
        }
    }

    std::cout << "[debug] Sending WM_COMMAND to create new tab in HWND=0x" << std::hex
              << reinterpret_cast<uintptr_t>(tabHost) << std::dec << "\n";
    SendMessageA(tabHost, WM_COMMAND, (WPARAM)WM_COMMAND_ID_NEW_TAB, 0);

    const DWORD timeoutMs = 8000;
    const DWORD retryMs = 300;
    DWORD waited = 0;

    while (waited <= timeoutMs) {
        std::vector<TabInfo> tabs;
        std::vector<HWND> windows;
        if (!CollectExplorerTabs(tabs, windows)) {
            Sleep(retryMs);
            waited += retryMs;
            continue;
        }

        std::vector<TabInfo*> firstWindowTabs;
        for (auto& t : tabs) {
            if (t.topLevel == firstWindow) {
                firstWindowTabs.push_back(&t);
            }
        }

        size_t currentCount = firstWindowTabs.size();
        if (currentCount > baselineCount && !firstWindowTabs.empty()) {
            TabInfo* newestTab = firstWindowTabs.back();
            IWebBrowser2* newBrowser = newestTab->browser;
            std::cout << "[debug] Identified new tab by count increase (" << baselineCount
                      << " -> " << currentCount << ") in HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(firstWindow) << std::dec << "\n";

            HRESULT navHr = NavigateBrowser(newBrowser, url);

            for (auto& t : tabs) {
                if (t.browser && t.browser != newBrowser) {
                    t.browser->Release();
                }
            }

            if (SUCCEEDED(navHr)) {
                knownTabCount = currentCount;
                baselineCount = currentCount;
                std::cout << "[debug] Navigation succeeded for new tab.\n";
                newBrowser->Release();
                return true;
            }

            newBrowser->Release();
            return false;
        }

        for (auto& t : tabs) {
            if (t.browser) {
                t.browser->Release();
            }
        }

        Sleep(retryMs);
        waited += retryMs;
    }

    return false;
}

int main() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::cerr << "CoInitializeEx failed: 0x" << std::hex << hr << "\n";
        return 1;
    }

    std::vector<TabInfo> tabs;
    std::vector<HWND> windowOrder;
    if (!CollectExplorerTabs(tabs, windowOrder)) {
        std::cerr << "Failed to enumerate Explorer tabs.\n";
        CoUninitialize();
        return 2;
    }

    if (windowOrder.empty()) {
        std::cout << "No Explorer windows detected.\n";
        for (auto& t : tabs) if (t.browser) t.browser->Release();
        CoUninitialize();
        return 0;
    }

    HWND firstWindow = windowOrder.front();
    size_t knownTabCount = 0;
    std::vector<std::string> urlsToMerge;
    std::vector<HWND> windowsToClose;

    for (auto& t : tabs) {
        if (t.topLevel == firstWindow) {
            ++knownTabCount;
            std::cout << "[debug] Known tab in first window on startup: HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(t.topLevel)
                      << ", IWebBrowser2=" << t.browser << std::dec << "\n";
        } else {
            std::string navigateTarget = !t.parsingName.empty() ? t.parsingName : t.locationUrl;
            if (!navigateTarget.empty()) {
                urlsToMerge.push_back(navigateTarget);
                std::cout << "[debug] Tab queued for merge: HWND=0x" << std::hex
                          << reinterpret_cast<uintptr_t>(t.topLevel)
                          << ", IWebBrowser2=" << t.browser << std::dec
                          << ", ParsingName=" << t.parsingName
                          << ", LocationURL=" << t.locationUrl
                          << ", Target=" << navigateTarget << "\n";
            }
            if (std::find(windowsToClose.begin(), windowsToClose.end(), t.topLevel) == windowsToClose.end()) {
                windowsToClose.push_back(t.topLevel);
            }
        }
    }

    for (auto& t : tabs) {
        if (t.browser) t.browser->Release();
    }

    if (urlsToMerge.empty()) {
        std::cout << "Nothing to merge.\n";
        CoUninitialize();
        return 0;
    }

    HWND tabHost = FindShellTabHost(firstWindow);
    if (!tabHost) {
        std::cerr << "Could not find ShellTabWindowClass in the first window.\n";
        CoUninitialize();
        return 3;
    }

    std::cout << "Merging " << urlsToMerge.size() << " tab(s) into the first window...\n";

    size_t successCount = 0;
    for (const auto& url : urlsToMerge) {
        if (CreateTabAndNavigate(firstWindow, tabHost, url, knownTabCount)) {
            ++successCount;
        } else {
            std::cerr << "[warn] Failed to create tab for: " << url << "\n";
        }
    }

    for (HWND h : windowsToClose) {
        if (h && h != firstWindow) {
            PostMessageA(h, WM_CLOSE, 0, 0);
        }
    }

    std::cout << "Completed. " << successCount << " tab(s) moved.\n";

    CoUninitialize();
    return 0;
}
