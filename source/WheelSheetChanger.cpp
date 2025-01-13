/*
WheelSheetChanger.cpp
use command prompt
g++ -o WheelSheetChanger.exe WheelSheetChanger.cpp WheelSheetChanger.o -mwindows -municode -luser32 -lshell32 -lWheelSheetChanger -Wl,-rpath=. -L.
cl /FeWheelSheetChanger.exe WheelSheetChanger.cpp WheelSheetChanger.lib User32.lib Shell32.lib WheelSheetChanger.res
*/

#ifndef UNICODE
#define UNICODE
#endif

#include <windows.h>
#include <shellapi.h>
#include <tchar.h>
#include <wchar.h>
#include <stdio.h>
#include "WheelSheetChangerDll.h"

#define TRAY_MENU_OPEN (WM_APP + 1)
#define TRAY_MENU_ITEM_EXIT (WM_APP + 2)
#define TRAY_MENU_ITEM_EXCEL (WM_APP + 3)
#define TRAY_MENU_ITEM_EXPLORER (WM_APP + 4)

// グローバル変数
NOTIFYICONDATA nid;
HMENU hMenu;
BOOL hooked;
HANDLE hMutex;
wchar_t inifile[MAX_PATH];

// 終了メニューが選ばれた時の処理
void OnExit()
{
    // アイコンをタスクトレイから削除
    Shell_NotifyIcon(NIM_DELETE, &nid);

    // フック解除
    if (hooked)
    {
        EndHook();
    }

    // ミューテックスを解放
    CloseHandle(hMutex);

    // アプリケーションを終了
    PostQuitMessage(0);
}

// 設定切り替え
void SettingChange(UINT itemName)
{
    // メニューのチェック状態を取得
    BOOL nextChecked = !(GetMenuState(hMenu, itemName, MF_BYCOMMAND) & MF_CHECKED);

    // メニューのチェック状態を反転
    MENUITEMINFO info = {sizeof(info)};
    info.fMask = MIIM_STATE;
    info.fState = (nextChecked ? MFS_CHECKED : MFS_UNCHECKED);
    SetMenuItemInfo(hMenu, itemName, FALSE, &info);

    // 設定切り替えとiniファイルへの書き込み
    switch (itemName)
    {
    case TRAY_MENU_ITEM_EXCEL:
        WritePrivateProfileString(L"setting", L"Excel", (nextChecked ? L"1" : L"0"), inifile);
        HookSettingChange(IS_ENABLE_EXCEL, nextChecked);
        break;
    case TRAY_MENU_ITEM_EXPLORER:
        WritePrivateProfileString(L"setting", L"Explorer", (nextChecked ? L"1" : L"0"), inifile);
        HookSettingChange(IS_ENABLE_EXPLORER, nextChecked);
        break;
    }
}

// ウィンドウプロシージャ
LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case TRAY_MENU_OPEN: // タスクトレイアイコンからのメッセージ
        if (lParam == WM_RBUTTONDOWN)
        { // 右クリックされた時
            POINT pt;
            SetForegroundWindow(hwnd);
            SetFocus(hwnd);
            GetCursorPos(&pt);
            TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, hwnd, NULL); // コンテキストメニューを表示
        }
        break;

    case WM_COMMAND: // メニュー項目の選択
        switch (LOWORD(wParam))
        {
        case TRAY_MENU_ITEM_EXIT: // "Exit"メニューが選ばれた場合
            OnExit();
            break;
        case TRAY_MENU_ITEM_EXCEL: // "Excel"メニューが選ばれた場合
            SettingChange(TRAY_MENU_ITEM_EXCEL);
            break;
        case TRAY_MENU_ITEM_EXPLORER: // "エクスプローラ"メニューが選ばれた場合
            SettingChange(TRAY_MENU_ITEM_EXPLORER);
            break;
        }
        break;

    case WM_DESTROY:
        OnExit();
        break;

    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    // 二重起動禁止
    wchar_t AppName[] = L"WheelSheetChanger";

    hMutex = CreateMutex(NULL, TRUE, AppName);  // ミューテックスを作成
    if (GetLastError() == ERROR_ALREADY_EXISTS) // ミューテックスがすでに存在する場合
    {
        MessageBox(NULL, L"アプリケーションはすでに起動しています。", L"エラー", MB_ICONERROR);
        // ミューテックスを解放して終了
        CloseHandle(hMutex);
        return 1;
    }

    // ウィンドウクラスの設定
    WNDCLASSEX wc = {sizeof(WNDCLASSEX), CS_CLASSDC, WndProc, 0L, 0L, hInstance, NULL, NULL, NULL, NULL, AppName, NULL};
    RegisterClassEx(&wc);

    // メインウィンドウを作成（非表示）
    HWND hwnd = CreateWindow(wc.lpszClassName, AppName, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, NULL, NULL, wc.hInstance, NULL);

    // ウィンドウを非表示にする
    ShowWindow(hwnd, SW_HIDE); // ここでウィンドウを非表示にする

    // トレイアイコンの設定
    nid = {sizeof(nid)};
    nid.hWnd = hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = TRAY_MENU_OPEN;

    // リソースからアイコンをロード
    HICON hIcon = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(100), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR);
    if (hIcon)
    {
        nid.hIcon = hIcon;
    }
    else
    {
        return 1;
    }

    wcscpy_s(nid.szTip, AppName);

    // アイコンをタスクトレイに追加
    Shell_NotifyIcon(NIM_ADD, &nid);

    // 設定確認
    wchar_t filepath[MAX_PATH];
    BOOL isEnable_Excel = FALSE;
    BOOL isEnable_Explorer = FALSE;
    if (GetModuleFileName(NULL, filepath, _countof(filepath)) > 0) // 実行ファイルのパス
    {
        // 実行ファイル名を分割
        wchar_t drive[_MAX_DRIVE];
        wchar_t dir[_MAX_DIR];
        wchar_t fname[_MAX_FNAME];
        wchar_t ext[_MAX_EXT];
        _wsplitpath_s(filepath, drive, dir, fname, ext);

        // iniファイルパス取得
        swprintf_s(inifile, L"%s%s%s.ini", drive, dir, fname);

        // 設定取得
        int settingCount = 0;
        if (_waccess_s(inifile, 0) == 0) // iniファイルが存在する場合
        {
            wchar_t setting[256];

            // Excelの設定
            if (GetPrivateProfileString(L"setting", L"Excel", L"1", setting, _countof(setting), inifile))
            {
                isEnable_Excel = (wcscmp(setting, L"1") == 0);
                HookSettingChange(IS_ENABLE_EXCEL, isEnable_Excel);
                settingCount++;
            }

            // Explorerの設定
            if (GetPrivateProfileString(L"setting", L"Explorer", L"1", setting, _countof(setting), inifile))
            {
                isEnable_Explorer = (wcscmp(setting, L"1") == 0);
                HookSettingChange(IS_ENABLE_EXPLORER, isEnable_Explorer);
                settingCount++;
            }
        }

        // 取得できなかった場合
        if (settingCount != 2)
        {
            WritePrivateProfileString(L"setting", L"Excel", L"1", inifile);
            WritePrivateProfileString(L"setting", L"Explorer", L"1", inifile);
            isEnable_Excel = TRUE;
            isEnable_Explorer = TRUE;
            HookSettingChange(IS_ENABLE_EXCEL, isEnable_Excel);
            HookSettingChange(IS_ENABLE_EXPLORER, isEnable_Explorer);
        }
    }

    // コンテキストメニューの作成
    hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING | (isEnable_Excel ? MF_CHECKED : 0), TRAY_MENU_ITEM_EXCEL, L"Excelでホイールによるシート変更、横スクロール");
    AppendMenu(hMenu, MF_STRING | (isEnable_Explorer ? MF_CHECKED : 0), TRAY_MENU_ITEM_EXPLORER, L"エクスプローラの余白ダブルクリックで上の階層へ");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, TRAY_MENU_ITEM_EXIT, L"終了");

    // フック開始
    hooked = StartHook();

    // メッセージループ
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}
