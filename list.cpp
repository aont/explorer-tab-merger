// list_explorer_tabs.cpp  (MinGW-w64 / C++17, 64bit推奨)
// g++ list.cpp -std=c++17 -O0 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid
#define UNICODE
#define _UNICODE
#define _WIN32_IE 0x0700
#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <exdisp.h>
#include <servprov.h>   // IServiceProvider
#include <shlguid.h>    // SID_STopLevelBrowser
#include <cstdio>
#include <string>

template<class T> static void SafeRelease(T*& p){ if(p){ p->Release(); p=nullptr; } }

static std::wstring UrlToPath(const std::wstring& url){
    if (url.rfind(L"file://", 0) == 0){
        DWORD cch = (DWORD)url.size() + 1;
        std::wstring path(cch, L'\0');
        if (S_OK == PathCreateFromUrlW(url.c_str(), path.data(), &cch, 0)){
            path.resize(cch ? cch - 1 : 0);
            return path;
        }
    }
    return L"";
}

static std::wstring PidlToParsingName(PCIDLIST_ABSOLUTE pidl, bool* isFS){
    PWSTR psz = nullptr;
    if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_FILESYSPATH, &psz)) && psz){
        std::wstring s(psz); CoTaskMemFree(psz);
        if (isFS) *isFS = true;
        return s; // 通常のファイルシステムパス
    }
    if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &psz)) && psz){
        std::wstring s(psz); CoTaskMemFree(psz);
        if (isFS) *isFS = false;
        return s; // 仮想フォルダは ::{GUID}\... 等
    }
    if (isFS) *isFS = false;
    return L"";
}

static std::wstring GetCurrentFolderViaBrowser(IWebBrowser2* wb, bool* isFS){
    *isFS = false;
    IServiceProvider* sp = nullptr;
    if (FAILED(wb->QueryInterface(IID_IServiceProvider, (void**)&sp))) return L"";
    IShellBrowser* sb = nullptr;
    HRESULT hr = sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb));
    SafeRelease(sp);
    if (FAILED(hr)) return L""; // Explorer でなければ失敗

    IShellView* sv = nullptr;
    if (FAILED(sb->QueryActiveShellView(&sv))){ SafeRelease(sb); return L""; }
    IFolderView* fv = nullptr;
    if (FAILED(sv->QueryInterface(IID_PPV_ARGS(&fv)))){ SafeRelease(sv); SafeRelease(sb); return L""; }
    IShellFolder* sf = nullptr;
    if (FAILED(fv->GetFolder(IID_PPV_ARGS(&sf)))){ SafeRelease(fv); SafeRelease(sv); SafeRelease(sb); return L""; }

    std::wstring out;
    IPersistFolder2* pf2 = nullptr;
    if (SUCCEEDED(sf->QueryInterface(IID_PPV_ARGS(&pf2)))){
        PIDLIST_ABSOLUTE pidl = nullptr;
        if (SUCCEEDED(pf2->GetCurFolder(&pidl)) && pidl){
            out = PidlToParsingName(pidl, isFS);
            CoTaskMemFree(pidl);
        }
        SafeRelease(pf2);
    }
    SafeRelease(sf); SafeRelease(fv); SafeRelease(sv); SafeRelease(sb);

    // 失敗時のフォールバック（URL→パス解決）
    if (out.empty()){
        BSTR b = nullptr;
        if (SUCCEEDED(wb->get_LocationURL(&b)) && b){
            std::wstring url(b, SysStringLen(b)); SysFreeString(b);
            auto path = UrlToPath(url);
            if (!path.empty()){ *isFS = true; return path; }
            return url; // shell: 等のURL
        }
    }
    return out;
}

int main(){
    // wprintf(L"Hello\n");
    // SetConsoleOutputCP(CP_UTF8);
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    IShellWindows* sw = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&sw)))){
        wprintf(L"IShellWindows 取得失敗\n"); CoUninitialize(); return 1;
    }

    LONG count=0; sw->get_Count(&count);
    for (LONG i=0; i<count; ++i){
        VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = i;
        IDispatch* disp = nullptr;
        if (S_OK == sw->Item(v, &disp) && disp){
            IWebBrowser2* wb = nullptr;
            if (SUCCEEDED(disp->QueryInterface(IID_PPV_ARGS(&wb)))){
                // Explorer かどうかを SID_STopLevelBrowser で判定
                IServiceProvider* sp=nullptr; IShellBrowser* sb=nullptr;
                bool isExplorer=false;
                if (SUCCEEDED(wb->QueryInterface(IID_PPV_ARGS(&sp)))){
                    if (SUCCEEDED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb)))) isExplorer=true;
                }
                SafeRelease(sb); SafeRelease(sp);
                if (isExplorer){
                    LONG_PTR h=0; HWND hwnd=nullptr;
                    IWebBrowserApp* wba=nullptr;
                    if (SUCCEEDED(wb->QueryInterface(IID_PPV_ARGS(&wba)))){
                        wba->get_HWND(&h); hwnd=(HWND)h; SafeRelease(wba);
                    }
                    bool isFS=false;
                    auto cur = GetCurrentFolderViaBrowser(wb, &isFS);

                    BSTR burl=nullptr; std::wstring url;
                    if (SUCCEEDED(wb->get_LocationURL(&burl)) && burl){
                        url.assign(burl, SysStringLen(burl)); SysFreeString(burl);
                    }
                    wprintf(L"#%ld hwnd=0x%p\n  URL: %ls\n  Path: %ls%ls\n",
                        i, hwnd, url.c_str(), cur.c_str(), isFS? L"" : L"  (virtual)");
                    // wprintf(L"#%ld hwnd=0x%p\n  URL: %s\n  Path: %s%s\n",
                    //     i, hwnd, url.c_str(), cur.c_str(), isFS? L"" : L"  (virtual)");
                }
                wb->Release();
            }
            disp->Release();
        }
        VariantClear(&v);
    }

    sw->Release();
    CoUninitialize();
    return 0;
}
