/// NaoRLB.cpp
// Version 1.0
//
// Copyright (c) 2006 InSoftware House.
// October 31, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

#define	REG_KEY_COLOR		_T("SOFTWARE\\Microsoft\\Color")
#define REG_NAME_BASEHUE	_T("BaseHue")
#define	MESSAGE_TITLE		_T("알림")
#define	FINISH_MESSAGE		_T("확장 리스트 컨트롤의 배경 색상이 제거되었습니다.")

BOOL SetUserShellColor(int nIndex, DWORD dwColor)
{
	BOOL		bResult	= FALSE;
	HKEY		hKey	= 0;
	COLORREF	cr		= 0;

	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, REG_KEY_COLOR, 0, 0, &hKey) == ERROR_SUCCESS)
	{
		TCHAR	szName[32];

		wsprintf(szName, _T("%d"), nIndex);

		if (RegSetValueEx(hKey, szName, 0, REG_BINARY, (LPBYTE)&dwColor, sizeof(DWORD)) == ERROR_SUCCESS)
		{
			bResult	= TRUE;
		}

		RegCloseKey(hKey);
	}

	return bResult;	
}

void RefreshSystem()
{
	SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0xF2, 0);
}

int _tmain(int argc, _TCHAR* argv[])
{
	SetUserShellColor(15, RGB(0xFF, 0xFF, 0xFF));
	SetUserShellColor(16, RGB(0xFF, 0xFF, 0xFF));
	SetUserShellColor(36, RGB(0xFF, 0xFF, 0xFF));
	SetUserShellColor(37, RGB(0xFF, 0xFF, 0xFF));

	RefreshSystem();

	MessageBox(NULL, FINISH_MESSAGE, MESSAGE_TITLE, MB_ICONINFORMATION | MB_OK | MB_TOPMOST);

	return 0;
}