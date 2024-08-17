#include <locale.h>
#include <stdio.h>
#include <Windows.h>
#include <atlsafe.h>
#include <ExDisp.h>
#include <Shldisp.h>

#define FAILED_OR_NULLPTR(hr, p) (FAILED(hr) || p == nullptr)

int main() {
    setlocale(LC_CTYPE, "");

    HWND topExplorerHwnd = FindWindowW(L"CabinetWClass", NULL);
    if (topExplorerHwnd == NULL) {
        return 1;
    }

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
        SHANDLE_PTR hwnd = NULL;
        webBrowser->get_HWND(&hwnd);
        if (hwnd == (SHANDLE_PTR)topExplorerHwnd) {
            target = webBrowser;
            break;
        }
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

    wprintf(L"%s\n", (LPWSTR)path);
    return 0;
}