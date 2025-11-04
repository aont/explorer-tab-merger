// list_explorer_tabs_mb.cpp (MinGW-w64 / C++17)
// g++ list_explorer_tabs_mb.cpp -std=c++17 -O0 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid

// ★ UNICODE 定義を削除してマルチバイト版にする
#undef UNICODE
#undef _UNICODE
#define _WIN32_IE 0x0700

#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <exdisp.h>
#include <servprov.h>
#include <shlguid.h>
#include <cstdio>
#include <string>

template<class T> static void SafeRelease(T*& p){ if(p){ p->Release(); p=nullptr; } }

// URL からパスへ変換（マルチバイト）
static std::string UrlToPath(const std::string& url){
    if (url.rfind("file://", 0) == 0){
        DWORD cch = (DWORD)url.size() + 1;
        std::string path(cch, '\0');
        if (S_OK == PathCreateFromUrlA(url.c_str(), path.data(), &cch, 0)){
            path.resize(cch ? cch - 1 : 0);
            return path;
        }
    }
    return "";
}

// PIDL からパス名（ANSI）を取得
static std::string PidlToParsingName(PCIDLIST_ABSOLUTE pidl, bool* isFS){
    CHAR sz[MAX_PATH];
    if (SHGetPathFromIDListA(pidl, sz)){
        if (isFS) *isFS = true;
        return sz; // 通常のファイルシステムパス
    }

    LPSTR psz = nullptr;
    if (SUCCEEDED(SHGetNameFromIDListA(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &psz)) && psz){
        std::string s(psz);
        CoTaskMemFree(psz);
        if (isFS) *isFS = false;
        return s; // 仮想フォルダなど
    }
    if (isFS) *isFS = false;
    return "";
}

// IWebBrowser2 から現在フォルダを取得（ANSI版）
static std::string GetCurrentFolderViaBrowser(IWebBrowser2* wb, bool* isFS){
    *isFS = false;
    IServiceProvider* sp = nullptr;
    if (FAILED(wb->QueryInterface(IID_IServiceProvider, (void**)&sp))) return "";
    IShellBrowser* sb = nullptr;
    HRESULT hr = sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb));
    SafeRelease(sp);
    if (FAILED(hr)) return "";

    IShellView* sv = nullptr;
    if (FAILED(sb->QueryActiveShellView(&sv))){ SafeRelease(sb); return ""; }
    IFolderView* fv = nullptr;
    if (FAILED(sv->QueryInterface(IID_PPV_ARGS(&fv)))){ SafeRelease(sv); SafeRelease(sb); return ""; }
    IShellFolder* sf = nullptr;
    if (FAILED(fv->GetFolder(IID_PPV_ARGS(&sf)))){ SafeRelease(fv); SafeRelease(sv); SafeRelease(sb); return ""; }

    std::string out;
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

    // URL→パス解決のフォールバック
    if (out.empty()){
        BSTR b = nullptr;
        if (SUCCEEDED(wb->get_LocationURL(&b)) && b){
            _bstr_t bt(b, false);
            std::string url = (const char*)bt;
            auto path = UrlToPath(url);
            if (!path.empty()){ *isFS = true; return path; }
            return url;
        }
    }
    return out;
}

int main(){
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    IShellWindows* sw = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&sw)))){
        printf("IShellWindows 取得失敗\n"); 
        CoUninitialize(); 
        return 1;
    }

    LONG count=0; sw->get_Count(&count);
    for (LONG i=0; i<count; ++i){
        VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = i;
        IDispatch* disp = nullptr;
        if (S_OK == sw->Item(v, &disp) && disp){
            IWebBrowser2* wb = nullptr;
            if (SUCCEEDED(disp->QueryInterface(IID_PPV_ARGS(&wb)))){
                IServiceProvider* sp=nullptr; IShellBrowser* sb=nullptr;
                bool isExplorer=false;
                if (SUCCEEDED(wb->QueryInterface(IID_PPV_ARGS(&sp)))){
                    if (SUCCEEDED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb)))) 
                        isExplorer=true;
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

                    BSTR burl=nullptr; std::string url;
                    if (SUCCEEDED(wb->get_LocationURL(&burl)) && burl){
                        _bstr_t bt(burl, false);
                        url = (const char*)bt;
                    }
                    printf("#%ld hwnd=0x%p\n  URL: %s\n  Path: %s%s\n",
                        i, hwnd, url.c_str(), cur.c_str(), isFS? "" : "  (virtual)");
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
