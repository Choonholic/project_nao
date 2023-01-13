//
// NaoNFC.cpp
// Version 1.1
//
// Copyright (c) 2006 InSoftware House.
// November 13, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

#define	REG_KEY_POWER				_T("\\Services\\Power")
#define	REG_KEY_SHELL				_T("SOFTWARE\\Microsoft\\Shell")
#define	REG_KEY_TASKBAR				_T("SOFTWARE\\Microsoft\\Shell\\TaskBar")
#define	REG_VAL_SHOWICON			_T("ShowIcon")
#define	REG_VAL_SHOWTITLEBARCLOCK	_T("ShowTitleBarClock")
#define	REG_VAL_LIMITEDCLOCK		_T("LimitedClock")

#define	MESSAGE_TITLE	_T("알림")
#define	FINISH_MESSAGE	_T("시계 표시가 새로운 형식으로 변경되었습니다. 설정을 적용하기 위해 기기를 리셋합니다.")

int _tmain(int argc, _TCHAR* argv[])
{
	HKEY	hKey;
	DWORD	dwDisposition;
	DWORD	dwValue;

	if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, REG_KEY_POWER, 0, NULL, 0, NULL, NULL, &hKey, &dwDisposition) == ERROR_SUCCESS)
	{
		dwValue	= 1;

		RegSetValueEx(hKey, REG_VAL_SHOWICON, 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
		RegCloseKey(hKey);
	}

	if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, REG_KEY_SHELL, 0, NULL, 0, NULL, NULL, &hKey, &dwDisposition) == ERROR_SUCCESS)
	{
		dwValue	= 0;

		RegSetValueEx(hKey, REG_VAL_SHOWTITLEBARCLOCK, 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
		RegCloseKey(hKey);
	}

	if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, REG_KEY_TASKBAR, 0, NULL, 0, NULL, NULL, &hKey, &dwDisposition) == ERROR_SUCCESS)
	{
		dwValue	= 1;

		RegSetValueEx(hKey, REG_VAL_LIMITEDCLOCK, 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
		RegCloseKey(hKey);
	}

	MessageBox(NULL, FINISH_MESSAGE, MESSAGE_TITLE, MB_ICONINFORMATION | MB_OK | MB_TOPMOST);
	ExitWindowsEx(EWX_REBOOT, 0);

	return 0;
}