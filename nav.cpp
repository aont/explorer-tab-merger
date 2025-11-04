// nav_by_index.cpp  (MinGW-w64 / ANSI / no ATL)
// Build: g++ nav_by_index.cpp -o nav_by_index.exe -lole32 -loleaut32 -luuid -luser32 -lshell32 -loleacc

#define _WIN32_DCOM
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <exdisp.h>    // IWebBrowser2
#include <shldisp.h>   // IShellWindows
#include <initguid.h>
#include <iostream>
#include <string>

// --- Utilities (ANSI <-> BSTR) ---
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

// --- Get IWebBrowser2 by IShellWindows index ---
static HRESULT GetWebBrowserByIndex(long index, IWebBrowser2** ppwb) {
    if (!ppwb) return E_POINTER;
    *ppwb = nullptr;

    IShellWindows* pSW = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW));
    if (FAILED(hr)) return hr;

    long count = 0;
    hr = pSW->get_Count(&count);
    if (FAILED(hr)) { pSW->Release(); return hr; }
    if (index < 0 || index >= count) { pSW->Release(); return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX); }

    VARIANT vIdx; VariantInit(&vIdx);
    vIdx.vt = VT_I4; vIdx.lVal = index;

    IDispatch* pDisp = nullptr;
    hr = pSW->Item(vIdx, &pDisp);
    VariantClear(&vIdx);
    if (FAILED(hr) || !pDisp) { pSW->Release(); return FAILED(hr) ? hr : E_FAIL; }

    IWebBrowser2* pWB = nullptr;
    hr = pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB);
    pDisp->Release();
    pSW->Release();
    if (FAILED(hr) || !pWB) return FAILED(hr) ? hr : E_NOINTERFACE;

    *ppwb = pWB;
    return S_OK;
}

// --- Navigate2 (must use Navigate2) ---
static HRESULT Navigate2ByIndex(long index, const char* urlOrPathAnsi) {
    if (!urlOrPathAnsi || !*urlOrPathAnsi) return E_INVALIDARG;

    IWebBrowser2* pWB = nullptr;
    HRESULT hr = GetWebBrowserByIndex(index, &pWB);
    if (FAILED(hr)) return hr;

    VARIANT vURL; VariantInit(&vURL);
    vURL.vt = VT_BSTR;
    vURL.bstrVal = AnsiToBSTR(urlOrPathAnsi);
    if (!vURL.bstrVal) { pWB->Release(); return E_OUTOFMEMORY; }

    VARIANT vEmpty; VariantInit(&vEmpty);
    hr = pWB->Navigate2(&vURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);

    VariantClear(&vURL);
    pWB->Release();
    return hr;
}

// --- List IShellWindows items (indexes you can use) ---
static void ListWindowsByIndex() {
    IShellWindows* pSW = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW)))) {
        std::cout << "ShellWindows unavailable.\n";
        return;
    }
    long count = 0;
    if (FAILED(pSW->get_Count(&count))) { pSW->Release(); return; }

    std::cout << "IShellWindows items (use these indexes with 'navigate'):\n";
    if (count == 0) {
        std::cout << "  (none)\n";
        pSW->Release(); return;
    }

    for (long i = 0; i < count; ++i) {
        VARIANT vIdx; VariantInit(&vIdx);
        vIdx.vt = VT_I4; vIdx.lVal = i;

        IDispatch* pDisp = nullptr;
        if (SUCCEEDED(pSW->Item(vIdx, &pDisp)) && pDisp) {
            IWebBrowser2* pWB = nullptr;
            if (SUCCEEDED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB)) && pWB) {
                SHANDLE_PTR h{};
                HWND hw{};
                if (SUCCEEDED(pWB->get_HWND(&h))) hw = (HWND)h;

                char title[512] = {};
                if (hw) GetWindowTextA(hw, title, sizeof(title)-1);

                BSTR locB = nullptr;
                pWB->get_LocationURL(&locB);
                std::string loc = BSTRtoAnsi(locB);
                if (locB) SysFreeString(locB);

                DWORD pid = 0;
                if (hw) GetWindowThreadProcessId(hw, &pid);

                std::cout << "  index: " << i
                          << "  pid: " << pid
                          << "  title: " << (title[0] ? title : "(no title)")
                          << "  url: " << (loc.empty() ? "(none)" : loc)
                          << "\n";

                pWB->Release();
            }
            pDisp->Release();
        }
        VariantClear(&vIdx);
    }
    pSW->Release();
}

static void PrintUsage() {
    std::cout <<
R"(Usage:
  nav_by_index.exe list
    - Lists IShellWindows items with their indexes.

  nav_by_index.exe navigate <index> <url_or_path>
    - Example: nav_by_index.exe navigate 0 "C:\Windows\System32"
    - Example: nav_by_index.exe navigate 1 "https://example.com/"
)";
}

int main(int argc, char** argv) {
    if (argc < 2) { PrintUsage(); return 1; }

    std::string cmd = argv[1];

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::cerr << "CoInitializeEx failed: 0x" << std::hex << hr << "\n";
        return 2;
    }

    if (cmd == "list") {
        ListWindowsByIndex();
        CoUninitialize();
        return 0;
    }

    if (cmd == "navigate") {
        if (argc < 4) { PrintUsage(); CoUninitialize(); return 1; }
        long index = strtol(argv[2], nullptr, 10);
        const char* target = argv[3];

        hr = Navigate2ByIndex(index, target);
        if (FAILED(hr)) {
            std::cerr << "Navigate2 failed: 0x" << std::hex << hr << "\n";
            CoUninitialize();
            return 3;
        }
        CoUninitialize();
        return 0;
    }

    PrintUsage();
    CoUninitialize();
    return 1;
}
