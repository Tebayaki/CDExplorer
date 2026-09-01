#include <locale.h>
#include <stdio.h>
#include <Windows.h>
#include <atlsafe.h>
#include <ExDisp.h>
#include <Shldisp.h>
#include <shobjidl_core.h>
#include <string>

#define FAILED_OR_NULLPTR(hr, p) (FAILED(hr) || p == nullptr)

// 用于安全输出 Unicode 路径：控制台 -> WriteConsoleW；否则输出 UTF-8 bytes
void print_path(BSTR bstr_path) {
    if (bstr_path == nullptr) {
        return;
    }
    LPCWSTR wpath = (LPCWSTR)bstr_path;
    HANDLE hout = GetStdHandle(STD_OUTPUT_HANDLE);
    DWORD mode;
    if (hout != INVALID_HANDLE_VALUE && GetConsoleMode(hout, &mode)) {
        // 直接写宽字符到控制台（适用于 Windows 原生控制台）
        DWORD written;
        WriteConsoleW(hout, wpath, static_cast<DWORD>(SysStringLen(bstr_path)), &written, NULL);
        WriteConsoleW(hout, L"\n", 1, &written, NULL);
    } else {
        // stdout 被重定向或不是原生控制台：转换为 UTF-8 写字节流（适用于 pipes/files/Git Bash 等）
        int required = WideCharToMultiByte(CP_UTF8, 0, wpath, static_cast<int>(SysStringLen(bstr_path)), NULL, 0, NULL, NULL);
        if (required > 0) {
            std::string utf8;
            utf8.resize(required);
            WideCharToMultiByte(CP_UTF8, 0, wpath, static_cast<int>(SysStringLen(bstr_path)), &utf8[0], required, NULL, NULL);
            fwrite(utf8.data(), 1, utf8.size(), stdout);
        }
        fputc('\n', stdout);
    }
};

int main() {
    setlocale(LC_CTYPE, "");

    HWND topExplorerHwnd = FindWindowW(L"CabinetWClass", NULL);
    if (topExplorerHwnd == NULL) {
        return 1;
    }
    HWND selectedTabHwnd = FindWindowExW(topExplorerHwnd, NULL, L"ShellTabWindowClass", NULL);

    HRESULT hr = CoInitialize(NULL);
    CComPtr<IShellWindows> shellWindows;
    hr = shellWindows.CoCreateInstance(CLSID_ShellWindows);
    if (FAILED_OR_NULLPTR(hr, shellWindows)) {
        return 2;
    }

    CComPtr<IWebBrowser2> target;
    CComVariant index = 0;
    long count = 0;
    shellWindows->get_Count(&count);
    for (; index.intVal < count; index.intVal++) {
        CComPtr<IDispatch> disp;
        hr = shellWindows->Item(index, &disp);
        if (FAILED_OR_NULLPTR(hr, disp)) {
            break;
        }
        CComPtr<IWebBrowser2> webBrowser;
        hr = disp->QueryInterface(&webBrowser);
        if (FAILED_OR_NULLPTR(hr, webBrowser)) {
            continue;
        }
        SHANDLE_PTR explorerHwnd = NULL;
        webBrowser->get_HWND(&explorerHwnd);
        if (explorerHwnd != (SHANDLE_PTR)topExplorerHwnd) {
            continue;
        }
        if (selectedTabHwnd != NULL) {
            CComPtr<IServiceProvider> serviceProvider;
            hr = webBrowser->QueryInterface(&serviceProvider);
            if (FAILED_OR_NULLPTR(hr, serviceProvider)) {
                continue;
            }
            CComPtr<IOleWindow> oleWindow;
            hr = serviceProvider->QueryService(IID_IOleWindow, &oleWindow);
            if (FAILED_OR_NULLPTR(hr, oleWindow)) {
                continue;
            }
            HWND tabHwnd = NULL;
            oleWindow->GetWindow(&tabHwnd);
            if (tabHwnd != selectedTabHwnd) {
                continue;
            }
        }
        target = webBrowser;
        break;
    }
    if (target == nullptr) {
        return 3;
    }

    CComPtr<IDispatch> disp;
    hr = target->get_Document(&disp);
    if (FAILED_OR_NULLPTR(hr, disp)) {
        return 4;
    }
    CComPtr<IShellFolderViewDual> folderView;
    hr = disp->QueryInterface(&folderView);
    if (FAILED_OR_NULLPTR(hr, folderView)) {
        return 4;
    }
    CComPtr<Folder> folder;
    hr = folderView->get_Folder(&folder);
    if (FAILED_OR_NULLPTR(hr, folder)) {
        return 4;
    }
    CComPtr<Folder2> folder2;
    hr = folder->QueryInterface(&folder2);
    if (FAILED_OR_NULLPTR(hr, folder2)) {
        return 4;
    }
    CComPtr<FolderItem> folderItem;
    hr = folder2->get_Self(&folderItem);
    if (FAILED_OR_NULLPTR(hr, folderItem)) {
        return 4;
    }
    CComBSTR path;
    hr = folderItem->get_Path(&path);
    if (FAILED_OR_NULLPTR(hr, path)) {
        return 4;
    }

    if (!PathFileExistsW(path)) {
        return 5;
    }

    print_path(path);
    return 0;
}