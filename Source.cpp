#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <iostream>
#include <vector>
#include <thread>
#include <mutex>
#include <windows.h>
#include <Python.h>

TCHAR szClassName[] = TEXT("Window");

std::mutex cout_mutex;

void run_python_script(LPCSTR lpszSrc, HWND hWnd)
{
	// GILの状態を保存
	PyGILState_STATE gstate = PyGILState_Ensure();

	try {
		PyObject* sys, * io, * string_io, * output;

		// モジュールのインポート
		{
			std::lock_guard<std::mutex> lock(cout_mutex);
			sys = PyImport_ImportModule("sys");
			io = PyImport_ImportModule("io");
		}

		if (!sys || !io) throw std::runtime_error("Failed to import modules");

		// StringIOオブジェクトの作成
		string_io = PyObject_CallMethod(io, "StringIO", NULL);
		if (!string_io) throw std::runtime_error("Failed to create StringIO");

		// sys.stdoutを置き換え
		if (PyObject_SetAttrString(sys, "stdout", string_io) == -1) {
			throw std::runtime_error("Failed to redirect stdout");
		}

		// Pythonコードの実行
		if (PyRun_SimpleString(lpszSrc)) {
			PyErr_Print();
			throw std::runtime_error("Python script error");
		}

		// 出力の取得
		output = PyObject_CallMethod(string_io, "getvalue", NULL);
		if (!output) throw std::runtime_error("Failed to get output");

		// 結果の表示
		const char* result = PyUnicode_AsUTF8(output);
		{
			std::lock_guard<std::mutex> lock(cout_mutex);
			HWND hEdit = GetDlgItem(hWnd, 1001);
			if (hEdit) {
				DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, result, -1, 0, 0);
				WCHAR* szResult = new WCHAR[dwSize];
				MultiByteToWideChar(CP_UTF8, 0, result, -1, szResult, dwSize);
				SetWindowText(hEdit, szResult);
				delete[] szResult;
			}
		}

		// リソース解放
		Py_XDECREF(output);
		Py_XDECREF(string_io);
		Py_XDECREF(io);
		Py_XDECREF(sys);
	}
	catch (const std::exception& e) {
		std::lock_guard<std::mutex> lock(cout_mutex);
		HWND hEdit = GetDlgItem(hWnd, 1001);
		if (hEdit) {
			DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, e.what(), -1, 0, 0);
			WCHAR* szResult = new WCHAR[dwSize];
			MultiByteToWideChar(CP_UTF8, 0, e.what(), -1, szResult, dwSize);
			SetWindowText(hEdit, szResult);
			delete[] szResult;
		}
	}

	// GIL解放
	PyGILState_Release(gstate);
}

//void ThreadFunc(HWND hWnd)
DWORD WINAPI ThreadFunc(LPVOID hWnd)
{
	std::vector<std::thread> threads;

	// 各スレッドに異なるスクリプトを渡す
	HWND hEdit = GetDlgItem((HWND)hWnd, 1000);

	// エディットコントロールのテキストを取得
	DWORD dwSize = GetWindowTextLength(hEdit) + 1;
	WCHAR* szSrcW = new WCHAR[dwSize];
	GetWindowText(hEdit, szSrcW, dwSize);

	// ワイド文字列からUTF-8文字列に変換
	dwSize = WideCharToMultiByte(CP_UTF8, 0, szSrcW, -1, 0, 0, 0, 0);
	char* lpszSrc = new char[dwSize];
	WideCharToMultiByte(CP_UTF8, 0, szSrcW, -1, lpszSrc, dwSize, 0, 0);

	// メモリの解放
	delete[] szSrcW;

	// スレッド起動
	threads.emplace_back(run_python_script, lpszSrc, (HWND)hWnd);

	// スレッドの終了を待つ
	for (auto& t : threads) {
		t.join();
	}
	threads.clear();

	// メモリの解放
	delete[] lpszSrc;

	PostMessage((HWND)hWnd, WM_APP, 0, 0);
	ExitThread(0);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	static HWND hEdit1;
	static HWND hEdit2;
	static PyThreadState* main_state;
	static HANDLE hThread = 0;
	DWORD dwParam;

	switch (msg)
	{
	case WM_CREATE:
		Py_Initialize();
		PyEval_InitThreads();
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("実行(F5)"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, (HMENU)1000, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_READONLY, 0, 0, 0, 0, hWnd, (HMENU)1001, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_SETFOCUS:
		SetFocus(hEdit1);
		break;
	case WM_SIZE:
		MoveWindow(hButton, 0, 0, 128, 32, TRUE);
		MoveWindow(hEdit1, 0, 32, LOWORD(lParam), (HIWORD(lParam) - 32) / 2, TRUE);
		MoveWindow(hEdit2, 0, 32 + (HIWORD(lParam) - 32) / 2, LOWORD(lParam), (HIWORD(lParam) - 32) / 2, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			EnableWindow(hButton, FALSE);
			EnableWindow(hEdit1, FALSE);
			EnableWindow(hEdit2, FALSE);
			main_state = PyEval_SaveThread();
			hThread = CreateThread(0, 0, ThreadFunc, (LPVOID)hWnd, 0, &dwParam);
		}
		break;
	case WM_APP:
		WaitForSingleObject(hThread, INFINITE);
		CloseHandle(hThread);
		hThread = 0;
		PyEval_RestoreThread(main_state);
		EnableWindow(hButton, TRUE);
		EnableWindow(hEdit1, TRUE);
		EnableWindow(hEdit2, TRUE);
		break;
	case WM_DESTROY:
		Py_Finalize();
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Pythonコードを実行する"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	ACCEL Accel[] = { {FVIRTKEY,VK_F5,IDOK} };
	HACCEL hAccel = CreateAcceleratorTable(Accel, _countof(Accel));
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return (int)msg.wParam;
}
