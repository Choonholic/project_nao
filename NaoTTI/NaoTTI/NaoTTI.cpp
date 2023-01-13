//
// NaoTTI.cpp
// Version 1.0
//
// Copyright (c) 2007 InSoftware House.
// April 9, 2007
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

LPCTSTR TakeFileName(LPCTSTR lpszFullPath, LPTSTR lpszName)
{
	LPTSTR	lpszFound	= _tcsrchr(lpszFullPath, _T('\\'));
	int		nPos		= -1;

	if (lpszFound != NULL)
	{
		nPos	= lpszFound - lpszFullPath;
	}

	if (nPos >= 0)
	{
		_tcscpy(lpszName, lpszFound + 1);
	}

	return (LPCTSTR)lpszName;
}

LPCTSTR TakeTitle(LPCTSTR lpszFullPath, LPTSTR lpszTitle)
{
	TCHAR	szName[MAX_PATH + 1];

	if (TakeFileName(lpszFullPath, szName))
	{
		LPTSTR	lpszFound	= _tcsrchr(szName, _T('.'));
		int		nPos		= -1;

		if (lpszFound != NULL)
		{
			nPos	= lpszFound - szName;
		}

		if (nPos > 0)
		{
			_tcsncpy(lpszTitle, szName, nPos);

			lpszTitle[nPos]	= NULL;
		}
	}

	return (LPCTSTR)lpszTitle;
}

int CombineFileName(LPTSTR lpszFullPath, LPCTSTR lpszFolder, LPCTSTR lpszFileName)
{
	if (!_tcslen(lpszFolder))
	{
		_tcscpy(lpszFullPath, _T("\\"));
	}
	else if (lpszFolder[0] != '\\')
	{
		_tcscpy(lpszFullPath, _T("\\"));
		_tcscat(lpszFullPath, lpszFolder);
	}
	else
	{
		_tcscpy(lpszFullPath, lpszFolder);
	}

	if (lpszFullPath[_tcslen(lpszFullPath) - 1] != _T('\\'))
	{
		_tcscat(lpszFullPath, _T("\\"));
	}

	_tcscat(lpszFullPath, lpszFileName);

	return (_tcslen(lpszFullPath));
}

INT GetSpecialDirectoryEx(INT nFolderID, LPTSTR lpszDir)
{
	int				rc;
	LPITEMIDLIST	pidl;
	BOOL			fUseIMalloc	= TRUE;
	LPMALLOC		lpMalloc	= NULL;

	rc = SHGetMalloc(&lpMalloc);

	if (rc == E_NOTIMPL)
	{
		fUseIMalloc = FALSE;
	}
	else if (rc != NOERROR)
	{
		return rc;
	}

	rc = SHGetSpecialFolderLocation(NULL, nFolderID, &pidl);

	if (rc == NOERROR)
	{
		if (SHGetPathFromIDList(pidl, lpszDir))
		{
			rc = E_FAIL;
		}
		if (fUseIMalloc)
		{
			lpMalloc->Free(pidl);
		}
	}

	if (fUseIMalloc)
	{
		lpMalloc->Release();
	}

	return rc;
}

BOOL InstallTheme(LPCTSTR lpszFilename)
{
	HKEY	hKey;

	if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Today"), 0, 0, &hKey) == ERROR_SUCCESS)
	{
		TCHAR				szCmdLine[MAX_PATH + 1];
		PROCESS_INFORMATION	pi;

		RegDeleteValue(hKey, _T("UseStartImage"));

#if (_WIN32_WCE < 0x0500)
 		wsprintf(szCmdLine, _T("/safe /noui /nouninstall /delete 0 \"%s\""), lpszFilename);
#else
 		wsprintf(szCmdLine, _T("\"%s\" /silent /nodelete /safe"), lpszFilename);
#endif

		if (::CreateProcess(_T("\\Windows\\wceload.exe"), szCmdLine, NULL, NULL, FALSE, 0, NULL, NULL, NULL, &pi))
		{
			CloseHandle(pi.hThread);
			WaitForSingleObject(pi.hProcess, INFINITE);
			CloseHandle(pi.hProcess);
			RegSetValueEx(hKey, _T("Skin"), 0, REG_SZ, (LPBYTE)lpszFilename, (_tcslen(lpszFilename) + 1) * sizeof(TCHAR));
		}

		RegCloseKey(hKey);

		SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, 0);
		Sleep(1000);
		SendMessage(HWND_BROADCAST, WM_SYSCOLORCHANGE, 0, 0);
		Sleep(1000);
		SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0xF2, 0);
		Sleep(1000);
		return TRUE;
	}

	return FALSE;
}

#define THEMESET_TEMP	_T("'%s' 이(가) 오늘 화면으로 지정되었습니다.")
#define MESSAGE_TITLE	_T("알림")

int _tmain(int argc, _TCHAR* argv[])
{
	TCHAR	szThemeName[MAX_PATH];
	TCHAR	szMessage[256];
	HCURSOR	hCursor	= SetCursor(LoadCursor(NULL, IDC_WAIT));

	if (argc == 2)
	{
		ShowCursor(TRUE);

		TakeTitle(argv[1], szThemeName);

		BOOL	bSuccess	= InstallTheme(argv[1]);
		
		ShowCursor(FALSE);
		SetCursor(hCursor);

		if (bSuccess)
		{
			wsprintf(szMessage, THEMESET_TEMP, szThemeName);
			MessageBox(NULL, szMessage, MESSAGE_TITLE, MB_OK | MB_ICONINFORMATION | MB_TOPMOST);
		}
	}

	return 0;
}
