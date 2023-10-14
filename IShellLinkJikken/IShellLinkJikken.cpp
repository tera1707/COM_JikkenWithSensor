#include <windows.h>
#include <tchar.h>	
#include <stdio.h>	
#include <string.h>	
#include <shlwapi.h>
#include <shlguid.h>
#include <shobjidl.h>

#pragma comment(lib, "shlwapi.lib")
#pragma warning(disable : 4996) 

int main()
{
	MessageBox(NULL, L"start", L"", MB_OK);

	// COMの初期化.
	CoInitialize(NULL);

	// ポインタの初期化.
	IShellLink* pShellLink = NULL;
	IPersistFile* pPersistFile = NULL;
	TCHAR tszSrcPath[MAX_PATH] = { 0 };
	TCHAR tszDirPath[MAX_PATH] = { 0 };

	// 実行ファイルのパスを取得.
	GetModuleFileName(NULL, tszSrcPath, MAX_PATH);

	// インスタンス作成.
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void**)&pShellLink);

	// ファイル実体のパスをセット.
	pShellLink->SetPath(tszSrcPath);

	TCHAR tszPath[MAX_PATH] = { 0 };	// tszPath(長さMAX_PATH)を{0}で初期化.
	pShellLink->GetPath(tszPath, MAX_PATH, NULL, SLGP_UNCPRIORITY);	// pShellLink->GetPathでパスをtszPathに格納.

	// IPersistFileインターフェースポインタを取得.
	hr = pShellLink->QueryInterface(IID_PPV_ARGS(&pPersistFile));	// pShellLink->QueryInterfaceでIPersistFileインターフェースポインタpPersistFileを取得.

	// 自分(exe)と同じフォルダに、自分へのショートカットを保存
	TCHAR tszShortCutPath[MAX_PATH] = { 0 };
	GetModuleFileName(NULL, tszShortCutPath, MAX_PATH);
	PathRemoveFileSpec(tszShortCutPath);
	_tcscat(tszShortCutPath, _T("\\ShortCut.lnk"));
	// 保存実施
	pPersistFile->Save(tszShortCutPath, TRUE);

	// 解放
	pPersistFile->Release();
	pShellLink->Release();

	CoUninitialize();
}


