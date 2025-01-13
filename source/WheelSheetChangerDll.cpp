/*
WheelSheetChangerDll.cpp
use command prompt
g++ -shared -o WheelSheetChanger.dll WheelSheetChangerDll.cpp -luser32 -lgdi32 -Wl,--out-implib,libWheelSheetChanger.a
cl /LD /FeWheelSheetChanger.dll WheelSheetChangerDll.cpp User32.lib Gdi32.lib
*/

#ifndef UNICODE
#define UNICODE
#endif
#define PSAPI_VERSION 2
#include <windows.h>
#include <psapi.h>

#define DLL_EXPORTS
#include "WheelSheetChangerDll.h"

// グローバル変数をプロセス間で共有
#ifdef __MINGW64__
// MinGW用
HHOOK g_hHook __attribute__((section("shared"), shared)) = NULL;
BOOL isEnable_Excel __attribute__((section("shared"), shared)) = TRUE;
BOOL isEnable_Explorer __attribute__((section("shared"), shared)) = TRUE;
#else
// VC用
#pragma data_seg("Shared")
HHOOK g_hHook = NULL;
BOOL isEnable_Excel = TRUE;
BOOL isEnable_Explorer = TRUE;
#pragma data_seg()
#pragma comment(linker, "/Section:Shared,rws")
#endif

HINSTANCE g_hInst;
HANDLE hTimerQueue = NULL;
HANDLE hTimer = NULL;

// マウスカーソル位置の色
COLORREF CursorPointColor()
{
    // マウスカーソル位置
    POINT point;
    GetCursorPos(&point);

    // デバイス コンテキストを取得
    HDC hdc = GetDC(NULL);

    // マウスカーソル位置の色を取得
    COLORREF color = GetPixel(hdc, point.x, point.y);

    // デバイス コンテキストを解放
    ReleaseDC(NULL, hdc);

    return color;
}

// ダブルクリック時のアクション
void DCAction()
{
    // マウスカーソル位置
    POINT point;
    GetCursorPos(&point);

    // マウスカーソル位置のウィンドウ
    HWND hWnd = WindowFromPoint(point);

    // クラス名
    wchar_t className[MAX_PATH];
    GetClassName(hWnd, className, _countof(className));

    // エクスプローラのファイルビュー
    if (wcscmp(className, L"DirectUIHWND") == 0)
    {
        // プロセスIDを取得
        DWORD processId;
        GetWindowThreadProcessId(hWnd, &processId);

        // プロセスオブジェクトを開く
        HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId);

        if (hProcess != NULL)
        {
            // 実行ファイルパスを取得
            wchar_t lpFilename[MAX_PATH];
            DWORD lpdwSize = 0;
            GetModuleFileNameEx(hProcess, NULL, lpFilename, _countof(lpFilename));

            // 実行ファイル名を取得
            wchar_t drive[_MAX_DRIVE];
            wchar_t dir[_MAX_DIR];
            wchar_t fname[_MAX_FNAME];
            wchar_t ext[_MAX_EXT];
            _wsplitpath_s(lpFilename, drive, dir, fname, ext);

            // 小文字化
            _wcslwr_s(fname);
            _wcslwr_s(ext);

            // explorer.exeの場合
            if (wcscmp(fname, L"explorer") == 0 && wcscmp(ext, L".exe") == 0)
            {
                // マウスカーソル位置の色が白の場合（余白の場合）
                if (CursorPointColor() == 0x00FFFFFF)
                {
                    // Alt+↑を送信
                    INPUT input[4] = {};

                    // Altを押す
                    input[0].type = INPUT_KEYBOARD;
                    input[0].ki.wVk = VK_MENU;
                    input[0].ki.wScan = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);
                    input[0].ki.dwExtraInfo = GetMessageExtraInfo();

                    // UPを押す
                    input[1].type = INPUT_KEYBOARD;
                    input[1].ki.wVk = VK_UP;
                    input[1].ki.wScan = MapVirtualKey(VK_UP, MAPVK_VK_TO_VSC);
                    input[1].ki.dwExtraInfo = GetMessageExtraInfo();

                    // UPを離す
                    input[2].type = INPUT_KEYBOARD;
                    input[2].ki.wVk = VK_UP;
                    input[2].ki.wScan = MapVirtualKey(VK_UP, MAPVK_VK_TO_VSC);
                    input[2].ki.dwFlags = KEYEVENTF_KEYUP;
                    input[2].ki.dwExtraInfo = GetMessageExtraInfo();

                    // Altを離す
                    input[3].type = INPUT_KEYBOARD;
                    input[3].ki.wVk = VK_MENU;
                    input[3].ki.wScan = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);
                    input[3].ki.dwFlags = KEYEVENTF_KEYUP;
                    input[3].ki.dwExtraInfo = GetMessageExtraInfo();

                    SendInput(4, input, sizeof(INPUT));
                }
            }

            // オブジェクトハンドルを閉じる
            CloseHandle(hProcess);
        }
    }
}

// 水平スクロールバーの領域を取得
RECT scrollRECT = {};
BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam)
{
    wchar_t title[MAX_PATH];
    GetWindowText(hwnd, title, _countof(title));
    if (wcscmp(title, L"水平方向") == 0)
    {
        // ウィンドウ領域取得
        GetWindowRect(hwnd, &scrollRECT);
        return FALSE;
    }
    return TRUE;
}

// down:1->up:2->down:3->up:4
int mode = 0;
VOID CALLBACK TimerRoutine(PVOID lpParam, BOOLEAN TimerOrWaitFired)
{
    mode = 0;
}
void SetTimer()
{
    // タイマー有効であればリセット
    if (hTimerQueue != NULL && hTimer != NULL)
    {
        DeleteTimerQueueTimer(hTimerQueue, hTimer, NULL);
    }
    // タイマー開始
    if (hTimerQueue != NULL)
    {
        CreateTimerQueueTimer(&hTimer, hTimerQueue, (WAITORTIMERCALLBACK)TimerRoutine, NULL, 500, 0, 0);
    }
}

LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    // マウス移動回数
    static int counter = 0;

    if (nCode == HC_ACTION)
    {
        // Explorer
        if (isEnable_Explorer)
        {
            switch (wParam)
            {
            case WM_MOUSEMOVE: // マウス移動
                // マウス移動をカウントしていく
                if (counter <= 255 && mode > 0)
                {
                    counter++;
                }

                // 移動5回以上でmodeリセット
                if (counter >= 5)
                {
                    mode = 0;
                }
                break;
            case WM_LBUTTONDOWN: // 左ボタンダウン
                if (mode == 0)
                {
                    // シングルクリックダウン
                    mode = 1;
                    counter = 0;
                    SetTimer();
                }
                else if (mode == 2 && counter < 5)
                {
                    // マウスカーソル位置の色が白の場合
                    if (CursorPointColor() == 0x00FFFFFF)
                    {
                        // シングルクリックアップ
                        mode = 3;
                    }
                }
                break;
            case WM_LBUTTONUP: // 左ボタンアップ
                if (mode == 1 && counter < 5)
                {
                    // ダブルクリックダウン
                    mode = 2;
                }
                else if (mode == 3 && counter < 5)
                {
                    // ダブルクリックアップ->ダブルクリック判定
                    DCAction();
                    mode = 0;
                }
                break;
            }
        }

        // Excel
        if (isEnable_Excel && wParam == WM_MOUSEWHEEL) // マウスホイール
        {
            // マウスカーソル位置
            POINT point;
            GetCursorPos(&point);

            // マウスカーソル位置のウィンドウ
            HWND hWnd = WindowFromPoint(point);

            // プロセスIDを取得
            DWORD processId;
            GetWindowThreadProcessId(hWnd, &processId);

            // プロセスオブジェクトを開く
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId);

            if (hProcess != NULL)
            {
                // シート選択領域かどうか
                BOOL sheetArea = FALSE;

                // 実行ファイルパスを取得
                wchar_t lpFilename[MAX_PATH];
                DWORD lpdwSize = 0;
                GetModuleFileNameEx(hProcess, NULL, lpFilename, _countof(lpFilename));

                // 実行ファイル名を取得
                wchar_t drive[_MAX_DRIVE];
                wchar_t dir[_MAX_DIR];
                wchar_t fname[_MAX_FNAME];
                wchar_t ext[_MAX_EXT];
                _wsplitpath_s(lpFilename, drive, dir, fname, ext);

                // 小文字化
                _wcslwr_s(fname);
                _wcslwr_s(ext);

                // EXCEL.EXEの場合
                if (wcscmp(fname, L"excel") == 0 && wcscmp(ext, L".exe") == 0)
                {
                    // ルートウィンドウ
                    HWND roothWnd = GetAncestor(hWnd, GA_ROOT);

                    // 水平スクロールバーの領域を取得
                    scrollRECT = {};
                    EnumChildWindows(roothWnd, EnumChildProc, (LPARAM)0);

                    // 水平スクロールバーの領域を取得できた場合
                    if ((scrollRECT.top - scrollRECT.bottom) != 0 && (scrollRECT.right - scrollRECT.left) != 0)
                    {
                        // ルートウィンドウのクライアント領域を取得
                        RECT windowSize;
                        GetWindowRect(roothWnd, &windowSize);

                        // シート選択領域の場合（簡易判定）
                        if (windowSize.left <= point.x && point.x <= scrollRECT.left && scrollRECT.top - 5 <= point.y && point.y <= scrollRECT.bottom + 5)
                        {
                            sheetArea = TRUE;
                        }
                        // 水平スクロールバー領域の場合
                        else if (scrollRECT.left <= point.x && point.x <= scrollRECT.right && scrollRECT.top - 5 <= point.y && point.y <= scrollRECT.bottom + 5)
                        {
                            // ホイールの回転方向確認
                            LPMSLLHOOKSTRUCT param;
                            param = (LPMSLLHOOKSTRUCT)lParam;
                            WORD delta = HIWORD(param->mouseData);

                            // 水平スクロールメッセージの送信
                            PostMessageW(hWnd, WM_MOUSEHWHEEL, MAKEWPARAM(0, -delta), MAKELPARAM(point.x, point.y));
                            return 1;
                        }
                    }
                }

                // オブジェクトハンドルを閉じる
                CloseHandle(hProcess);

                // シート選択領域の場合
                if (sheetArea)
                {
                    // ホイールの回転方向確認
                    LPMSLLHOOKSTRUCT param;
                    param = (LPMSLLHOOKSTRUCT)lParam;
                    INT16 delta = (INT16)(HIWORD(param->mouseData));

                    // PageUp/PageDownの選択
                    WORD key = delta > 0 ? VK_PRIOR : VK_NEXT;

                    INPUT input[4] = {};

                    // Ctrlを押す
                    input[0].type = INPUT_KEYBOARD;
                    input[0].ki.wVk = VK_CONTROL;
                    input[0].ki.wScan = MapVirtualKey(VK_CONTROL, MAPVK_VK_TO_VSC);
                    input[0].ki.dwExtraInfo = GetMessageExtraInfo();

                    // PageUp/PageDownを押す
                    input[1].type = INPUT_KEYBOARD;
                    input[1].ki.wVk = key;
                    input[1].ki.wScan = MapVirtualKey(key, MAPVK_VK_TO_VSC);
                    input[1].ki.dwExtraInfo = GetMessageExtraInfo();

                    // PageUp/PageDownを離す
                    input[2].type = INPUT_KEYBOARD;
                    input[2].ki.wVk = key;
                    input[2].ki.wScan = MapVirtualKey(key, MAPVK_VK_TO_VSC);
                    input[2].ki.dwFlags = KEYEVENTF_KEYUP;
                    input[2].ki.dwExtraInfo = GetMessageExtraInfo();

                    // Ctrlを離す
                    input[3].type = INPUT_KEYBOARD;
                    input[3].ki.wVk = VK_CONTROL;
                    input[3].ki.wScan = MapVirtualKey(VK_CONTROL, MAPVK_VK_TO_VSC);
                    input[3].ki.dwFlags = KEYEVENTF_KEYUP;
                    input[3].ki.dwExtraInfo = GetMessageExtraInfo();

                    SendInput(4, input, sizeof(INPUT));

                    return 1;
                }
            }
        }
    }
    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

// フック開始
DLL_DECLSPEC BOOL __stdcall StartHook()
{
    g_hHook = SetWindowsHookEx(WH_MOUSE_LL, (HOOKPROC)HookProc, g_hInst, 0);
    return g_hHook != NULL;
}

// フック終了
DLL_DECLSPEC BOOL __stdcall EndHook()
{
    return UnhookWindowsHookEx(g_hHook);
}

// 設定変更
DLL_DECLSPEC void __stdcall HookSettingChange(int name, BOOL enable)
{
    if (name == IS_ENABLE_EXCEL) // Excel
    {
        isEnable_Excel = enable;
    }
    else if (name == IS_ENABLE_EXPLORER) // Explorer
    {
        isEnable_Explorer = enable;
        mode = 0;
    }
    else
    {
        ;
    }
}

BOOL APIENTRY DllMain(HMODULE hModule,
                      DWORD ul_reason_for_call,
                      LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = (HINSTANCE)hModule;
        hTimerQueue = CreateTimerQueue();
        break;
    case DLL_PROCESS_DETACH:
        if (hTimerQueue != NULL)
        {
            DeleteTimerQueueEx(hTimerQueue, NULL);
        }
        break;
    }
    return TRUE;
}
