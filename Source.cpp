#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <iostream>
#include <vector>
#include <thread>
#include <sstream>
#include <string>
#include <mutex>
#include <windows.h>
#include <Python.h>

TCHAR szClassName[] = TEXT("Window");

std::mutex cout_mutex;

LPWSTR handle_python_error(const std::string& script)
{
	PyObject* type = nullptr, * value = nullptr, * traceback = nullptr;
	PyErr_Fetch(&type, &value, &traceback);
	PyErr_NormalizeException(&type, &value, &traceback);

	if (!type) {
		std::cerr << "Unknown error occurred" << std::endl;
		return nullptr;
	}

	// 例外基本情報
	PyObject* type_name = PyObject_GetAttrString(type, "__name__");
	const char* error_type = PyUnicode_AsUTF8(type_name);
	PyObject* error_msg = PyObject_Str(value);
	const char* msg_str = PyUnicode_AsUTF8(error_msg);

	// 行番号とソースコード行の取得
	int error_line = -1;
	std::string error_source_line;
	std::string formatted_trace;  // トレースバック情報を格納するstring

	if (traceback) {
		PyObject* tb_module = PyImport_ImportModule("traceback");
		if (tb_module) {
			PyObject* format_func = PyObject_GetAttrString(tb_module, "format_exception");
			if (format_func && PyCallable_Check(format_func)) {
				PyObject* args = PyTuple_Pack(3, type, value, traceback);
				if (args) {
					PyObject* formatted = PyObject_CallObject(format_func, args);
					if (formatted) {
						PyObject* joined = PyUnicode_Join(PyUnicode_FromString(""), formatted);
						if (joined) {
							const char* trace_cstr = PyUnicode_AsUTF8(joined);
							if (trace_cstr) {
								formatted_trace = trace_cstr;  // stringにコピー

								// 行番号解析
								std::istringstream iss(formatted_trace);
								std::string line;
								while (std::getline(iss, line)) {
									size_t pos = line.find("File \"<string>\", line ");
									if (pos != std::string::npos) {
										pos += 22;
										error_line = std::stoi(line.substr(pos));

										// ソースコード行の抽出
										std::istringstream script_stream(script);
										for (int i = 0; i < error_line; ++i) {
											std::getline(script_stream, error_source_line);
										}
										break;
									}
								}
							}
							Py_XDECREF(joined);
						}
						Py_XDECREF(formatted);
					}
					Py_XDECREF(args);
				}
			}
			Py_XDECREF(format_func);
			Py_XDECREF(tb_module);
		}
	}

	// トレースバックの改行コード変換
	auto convert_newlines = [](std::string& s) {
		size_t pos = 0;
		while ((pos = s.find('\n', pos)) != std::string::npos) {
			if (pos == 0 || s[pos - 1] != '\r') {
				s.replace(pos, 1, "\r\n");
				pos += 2;
			}
			else {
				pos++;
			}
		}
	};

	convert_newlines(error_source_line);
	convert_newlines(formatted_trace);

	std::string str =
		std::string("=== Python Error ===\r\nType: ") + error_type
		+ "\r\nMessage: " + msg_str
		+ "\r\nLine: " + std::to_string(error_line)
		+ "\r\nSource: " + error_source_line
		+ "\r\nFull Traceback:\r\n" + formatted_trace;

	// リソース解放
	Py_XDECREF(type_name);
	Py_XDECREF(error_msg);
	Py_XDECREF(type);
	Py_XDECREF(value);
	Py_XDECREF(traceback);

	// str to LPWSTR
	DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, 0, 0);
	LPWSTR lpszResult = new WCHAR[dwSize];
	MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, lpszResult, dwSize);

	return lpszResult;
}

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
		{
			PyObject* main_dict = PyModule_GetDict(PyImport_AddModule("__main__"));
			PyCompilerFlags flags = { 0 };

			// コード実行（PyRun_StringFlags使用）
			PyObject* result = PyRun_StringFlags(
				lpszSrc,
				Py_file_input,  // 実行モード（複数文）
				main_dict,
				main_dict,
				&flags
			);

			if (!result) {
				throw std::runtime_error("Python script error");
			}
			else {
				Py_DECREF(result);
			}
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
			LPWSTR lpszResult = handle_python_error(lpszSrc);
			if (!lpszResult) {
				LPCSTR what = e.what();
				DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, what, -1, 0, 0);
				lpszResult = new WCHAR[dwSize];
				MultiByteToWideChar(CP_UTF8, 0, what, -1, lpszResult, dwSize);
			}
			SetWindowText(hEdit, lpszResult);
			delete[] lpszResult;
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
		hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, (HMENU)1000, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_READONLY, 0, 0, 0, 0, hWnd, (HMENU)1001, ((LPCREATESTRUCT)lParam)->hInstance, 0);
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
		SetFocus(hEdit1);
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
