// merge_tabs.cpp - Merge Explorer tabs into the first window (ANSI, MinGW-w64 friendly)
// Build: g++ merge_tabs.cpp -std=c++17 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32

#define _WIN32_IE 0x0700
#define _WIN32_DCOM
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif

#include <windows.h>
#include <shlobj.h>
#include <shlguid.h>
#include <exdisp.h>
#include <shldisp.h>
#include <servprov.h>
#include <oleauto.h>

#include <vector>
#include <string>
#include <set>
#include <algorithm>
#include <iostream>
#include <iomanip>
#include <mutex>
#include <new>

static const UINT WM_COMMAND_ID_NEW_TAB = 0xA21B; // same as newtab.cpp (undocumented)

#ifndef DISPID_WINDOWREGISTERED
#define DISPID_WINDOWREGISTERED 200
#endif

struct TabInfo {
    IWebBrowser2* browser; // holds one reference; caller must Release
    std::string url;
    HWND topLevel;
};

class ShellWindowsEventSink : public IDispatch {
public:
    ShellWindowsEventSink(IShellWindows* shellWindows, HWND targetWindow)
        : m_refCount(1),
          m_shellWindows(shellWindows),
          m_targetWindow(targetWindow),
          m_event(CreateEvent(nullptr, TRUE, FALSE, nullptr)),
          m_capturedBrowser(nullptr) {
        if (m_shellWindows) {
            m_shellWindows->AddRef();
        }
    }

    ~ShellWindowsEventSink() {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_capturedBrowser) {
            m_capturedBrowser->Release();
            m_capturedBrowser = nullptr;
        }
        if (m_shellWindows) {
            m_shellWindows->Release();
            m_shellWindows = nullptr;
        }
        if (m_event) {
            CloseHandle(m_event);
            m_event = nullptr;
        }
    }

    void BeginListening(HWND targetWindow) {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_targetWindow = targetWindow;
        if (m_capturedBrowser) {
            m_capturedBrowser->Release();
            m_capturedBrowser = nullptr;
        }
        if (m_event) {
            ResetEvent(m_event);
        }
    }

    bool WaitForNewTab(DWORD timeoutMs, IWebBrowser2** outBrowser) {
        if (!outBrowser) return false;
        *outBrowser = nullptr;

        {
            std::lock_guard<std::mutex> lock(m_mutex);
            if (!m_event) {
                return false;
            }
        }

        ULONGLONG startTick = GetTickCount64();

        while (true) {
            DWORD remaining = timeoutMs;
            if (timeoutMs != INFINITE) {
                ULONGLONG elapsed = GetTickCount64() - startTick;
                if (elapsed >= timeoutMs) {
                    return false;
                }
                remaining = static_cast<DWORD>(timeoutMs - elapsed);
            }

            HANDLE handles[1] = { m_event };
            DWORD waitResult = MsgWaitForMultipleObjects(1, handles, FALSE, remaining, QS_ALLINPUT);
            if (waitResult == WAIT_OBJECT_0) {
                std::lock_guard<std::mutex> lock(m_mutex);
                if (m_event) {
                    ResetEvent(m_event);
                }
                if (m_capturedBrowser) {
                    IWebBrowser2* browser = m_capturedBrowser;
                    browser->AddRef();
                    m_capturedBrowser->Release();
                    m_capturedBrowser = nullptr;
                    *outBrowser = browser;
                    return true;
                }
                return false;
            } else if (waitResult == WAIT_OBJECT_0 + 1) {
                MSG msg;
                while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            } else if (waitResult == WAIT_TIMEOUT) {
                return false;
            } else {
                return false;
            }
        }
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override {
        if (!ppvObject) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDispatch) {
            *ppvObject = static_cast<IDispatch*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&m_refCount));
    }

    ULONG STDMETHODCALLTYPE Release() override {
        ULONG count = static_cast<ULONG>(InterlockedDecrement(&m_refCount));
        if (count == 0) {
            delete this;
        }
        return count;
    }

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override {
        if (pctinfo) *pctinfo = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override {
        return DISP_E_UNKNOWNNAME;
    }

    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS* pDispParams,
                                      VARIANT*, EXCEPINFO*, UINT*) override {
        if (dispIdMember == DISPID_WINDOWREGISTERED && pDispParams && pDispParams->cArgs == 1) {
            VARIANTARG& arg = pDispParams->rgvarg[0];
            if (arg.vt == VT_I4 || arg.vt == VT_INT) {
                HandleWindowRegistered(arg.lVal);
            }
        }
        return S_OK;
    }

private:
    void HandleWindowRegistered(LONG cookie) {
        IShellWindows* shell = nullptr;
        HWND targetWindow = nullptr;
        HANDLE eventHandle = nullptr;
        {
            std::lock_guard<std::mutex> lock(m_mutex);
            shell = m_shellWindows;
            targetWindow = m_targetWindow;
            eventHandle = m_event;
            if (shell) {
                shell->AddRef();
            }
        }

        if (!shell || !eventHandle) {
            if (shell) shell->Release();
            return;
        }

        VARIANT vIndex;
        VariantInit(&vIndex);
        vIndex.vt = VT_I4;
        vIndex.lVal = cookie;

        IDispatch* disp = nullptr;
        if (FAILED(shell->Item(vIndex, &disp)) || !disp) {
            shell->Release();
            return;
        }

        IWebBrowser2* browser = nullptr;
        HRESULT hr = disp->QueryInterface(IID_IWebBrowser2, (void**)&browser);
        disp->Release();
        if (FAILED(hr) || !browser) {
            shell->Release();
            return;
        }

        SHANDLE_PTR handle = 0;
        HWND topLevel = nullptr;
        if (SUCCEEDED(browser->get_HWND(&handle))) {
            topLevel = (HWND)handle;
        }

        if (topLevel != targetWindow) {
            browser->Release();
            shell->Release();
            return;
        }

        {
            std::lock_guard<std::mutex> lock(m_mutex);
            if (!m_event || m_targetWindow != targetWindow) {
                browser->Release();
                shell->Release();
                return;
            }

            if (m_capturedBrowser) {
                m_capturedBrowser->Release();
                m_capturedBrowser = nullptr;
            }
            m_capturedBrowser = browser;
            m_capturedBrowser->AddRef();
            std::cout << "[debug] WindowRegistered captured tab: HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(topLevel)
                      << ", IWebBrowser2=" << browser
                      << std::dec << "\n";
            SetEvent(m_event);
        }

        browser->Release();
        shell->Release();
    }

    volatile LONG m_refCount;
    IShellWindows* m_shellWindows;
    HWND m_targetWindow;
    HANDLE m_event;
    IWebBrowser2* m_capturedBrowser;
    std::mutex m_mutex;
};

static uintptr_t GetBrowserUnknownPointer(IWebBrowser2* browser) {
    if (!browser) return 0;
    IUnknown* unk = nullptr;
    uintptr_t value = 0;
    if (SUCCEEDED(browser->QueryInterface(IID_IUnknown, (void**)&unk)) && unk) {
        value = reinterpret_cast<uintptr_t>(unk);
        unk->Release();
    }
    return value;
}

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
static bool CollectExplorerTabsFromShellWindows(IShellWindows* pSW, std::vector<TabInfo>& tabs, std::vector<HWND>& windowOrder) {
    if (!pSW) return false;

    tabs.clear();
    windowOrder.clear();

    long count = 0;
    if (FAILED(pSW->get_Count(&count))) {
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
        IServiceProvider* sp = nullptr;
        if (SUCCEEDED(pWB->QueryInterface(IID_IServiceProvider, (void**)&sp)) && sp) {
            IShellBrowser* sb = nullptr;
            if (SUCCEEDED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb))) && sb) {
                isExplorer = true;
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
        std::string url;
        if (SUCCEEDED(pWB->get_LocationURL(&bUrl)) && bUrl) {
            url = BSTRtoAnsi(bUrl);
            SysFreeString(bUrl);
        }

        if (std::find(windowOrder.begin(), windowOrder.end(), topLevel) == windowOrder.end()) {
            windowOrder.push_back(topLevel);
        }

        uintptr_t unkPtr = GetBrowserUnknownPointer(pWB);
        std::cout << "[debug] Explorer tab found: top-level HWND=0x" << std::hex << std::setw(0)
                  << reinterpret_cast<uintptr_t>(topLevel)
                  << ", IWebBrowser2=" << pWB
                  << ", IUnknown=0x" << unkPtr
                  << ", URL=" << url << std::dec << "\n";

        tabs.push_back({ pWB, url, topLevel });
    }

    return true;
}

static bool CollectExplorerTabs(std::vector<TabInfo>& tabs, std::vector<HWND>& windowOrder) {
    IShellWindows* pSW = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW)))) {
        return false;
    }

    bool result = CollectExplorerTabsFromShellWindows(pSW, tabs, windowOrder);
    pSW->Release();
    return result;
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
static bool CreateTabAndNavigate(HWND firstWindow, HWND tabHost, const std::string& url, std::set<uintptr_t>& knownTabs) {
    if (!firstWindow || !tabHost || url.empty()) return false;

    IShellWindows* pSW = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW)))) {
        return false;
    }

    std::set<uintptr_t> baseline = knownTabs;
    {
        std::vector<TabInfo> beforeTabs;
        std::vector<HWND> beforeWindows;
        if (CollectExplorerTabsFromShellWindows(pSW, beforeTabs, beforeWindows)) {
            for (auto& t : beforeTabs) {
                if (t.topLevel == firstWindow) {
                    uintptr_t key = GetBrowserUnknownPointer(t.browser);
                    if (key) {
                        std::cout << "[debug] Baseline tab in first window: HWND=0x" << std::hex
                                  << reinterpret_cast<uintptr_t>(t.topLevel)
                                  << ", IUnknown=0x" << key << std::dec << "\n";
                        baseline.insert(key);
                        knownTabs.insert(key);
                    }
                }
                if (t.browser) t.browser->Release();
            }
        }
    }

    ShellWindowsEventSink* sink = new (std::nothrow) ShellWindowsEventSink(pSW, firstWindow);
    if (!sink) {
        pSW->Release();
        return false;
    }

    IConnectionPointContainer* cpc = nullptr;
    HRESULT hr = pSW->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc);
    if (FAILED(hr) || !cpc) {
        sink->Release();
        pSW->Release();
        return false;
    }

    IConnectionPoint* cp = nullptr;
    hr = cpc->FindConnectionPoint(DIID_DShellWindowsEvents, &cp);
    if (FAILED(hr) || !cp) {
        cpc->Release();
        sink->Release();
        pSW->Release();
        return false;
    }

    DWORD adviseCookie = 0;
    hr = cp->Advise(static_cast<IUnknown*>(sink), &adviseCookie);
    if (FAILED(hr)) {
        cp->Release();
        cpc->Release();
        sink->Release();
        pSW->Release();
        return false;
    }

    sink->BeginListening(firstWindow);

    std::cout << "[debug] Sending WM_COMMAND to create new tab in HWND=0x" << std::hex
              << reinterpret_cast<uintptr_t>(tabHost) << std::dec << "\n";
    SendMessageA(tabHost, WM_COMMAND, (WPARAM)WM_COMMAND_ID_NEW_TAB, 0);

    const DWORD timeoutMs = 8000;
    IWebBrowser2* newBrowser = nullptr;
    bool captured = sink->WaitForNewTab(timeoutMs, &newBrowser);

    bool result = false;
    if (!captured || !newBrowser) {
        if (newBrowser) {
            newBrowser->Release();
            newBrowser = nullptr;
        }
        std::cerr << "[warn] Timed out waiting for WindowRegistered event. Falling back to polling.\n";

        const DWORD pollIntervalMs = 200;
        ULONGLONG startTick = GetTickCount64();
        while (GetTickCount64() - startTick < timeoutMs) {
            std::vector<TabInfo> pollTabs;
            std::vector<HWND> pollWindows;
            if (CollectExplorerTabsFromShellWindows(pSW, pollTabs, pollWindows)) {
                for (auto& t : pollTabs) {
                    if (t.topLevel == firstWindow) {
                        uintptr_t key = GetBrowserUnknownPointer(t.browser);
                        if (key && baseline.find(key) == baseline.end()) {
                            newBrowser = t.browser;
                            newBrowser->AddRef();
                            captured = true;
                            for (auto& releaseTab : pollTabs) {
                                if (releaseTab.browser) releaseTab.browser->Release();
                            }
                            break;
                        }
                    }
                }
                if (!captured) {
                    for (auto& releaseTab : pollTabs) {
                        if (releaseTab.browser) releaseTab.browser->Release();
                    }
                }
            }

            if (captured && newBrowser) {
                std::cout << "[debug] Fallback polling captured new tab instance.\n";
                break;
            }

            Sleep(pollIntervalMs);
        }
    }

    if (captured && newBrowser) {
        uintptr_t newKey = GetBrowserUnknownPointer(newBrowser);
        if (newKey && baseline.find(newKey) == baseline.end()) {
            std::cout << "[debug] Navigating captured tab: HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(firstWindow)
                      << ", IUnknown=0x" << newKey
                      << std::dec << ", URL=" << url << "\n";
            HRESULT navHr = NavigateBrowser(newBrowser, url);
            if (SUCCEEDED(navHr)) {
                knownTabs.insert(newKey);
                baseline.insert(newKey);
                std::cout << "[debug] Navigation succeeded for tab IUnknown=0x" << std::hex
                          << newKey << std::dec << "\n";
                result = true;
            } else {
                std::cerr << "[warn] Navigate2 failed for new tab: 0x" << std::hex
                          << navHr << std::dec << "\n";
            }
        } else {
            std::cerr << "[warn] Captured tab was already known or invalid.\n";
        }
        newBrowser->Release();
    } else {
        std::cerr << "[warn] Failed to detect the new tab after polling.\n";
    }

    if (adviseCookie) {
        cp->Unadvise(adviseCookie);
    }
    cp->Release();
    cpc->Release();
    sink->Release();
    pSW->Release();

    return result;
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
    std::set<uintptr_t> knownTabs;
    std::vector<std::string> urlsToMerge;
    std::vector<HWND> windowsToClose;

    for (auto& t : tabs) {
        uintptr_t key = GetBrowserUnknownPointer(t.browser);
        if (t.topLevel == firstWindow) {
            if (key) {
                std::cout << "[debug] Known tab in first window on startup: HWND=0x" << std::hex
                          << reinterpret_cast<uintptr_t>(t.topLevel)
                          << ", IUnknown=0x" << key << std::dec << "\n";
                knownTabs.insert(key);
            }
        } else {
            if (!t.url.empty()) {
                urlsToMerge.push_back(t.url);
                std::cout << "[debug] Tab queued for merge: HWND=0x" << std::hex
                          << reinterpret_cast<uintptr_t>(t.topLevel)
                          << ", IUnknown=0x" << key << std::dec
                          << ", URL=" << t.url << "\n";
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
        if (CreateTabAndNavigate(firstWindow, tabHost, url, knownTabs)) {
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
