//
// NaoTBC.cpp
// Version 1.0
//
// Copyright (c) 2006 InSoftware House.
// October 26, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

#define	CHECK_INTERVAL	5000	// 5 seconds
#define	MIN_BRIGHTNESS	0x00
#define	MAX_BRIGHTNESS	0x0A

#define	NAOTBC_KEY		_T("Software\\Project Nao\\NaoTBC")
#define	NAOTBC_BASE		_T("Base")

int _tmain(int argc, _TCHAR* argv[])
{
	SYSTEMTIME	stNow;
	HKEY		hKey;
	DWORD		dwType;
	DWORD		dwSize;
	DWORD		dwValue;
	DWORD		dwBrightness;
	TCHAR		szHour[3];
	DWORD		dwDisposition;

	while (TRUE)
	{
		Sleep(CHECK_INTERVAL);

		GetLocalTime(&stNow);
		wsprintf(szHour, _T("%02d"), stNow.wHour);

		dwType	= REG_DWORD;
		dwSize	= sizeof(DWORD);

		RegCreateKeyEx(HKEY_CURRENT_USER, NAOTBC_KEY, 0, NULL, 0, NULL, NULL, &hKey, &dwDisposition);

		if (RegQueryValueEx(hKey, NAOTBC_BASE, 0, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
		{
			dwBrightness	= dwValue;
		}
		else
		{
			dwBrightness	= (MAX_BRIGHTNESS - MIN_BRIGHTNESS) / 2;

			RegSetValueEx(hKey, NAOTBC_BASE, 0, REG_DWORD, (LPBYTE)&dwBrightness, sizeof(DWORD));
		}

		RegCloseKey(hKey);

		if (RegOpenKeyEx(HKEY_CURRENT_USER, NAOTBC_KEY, 0, 0, &hKey) == ERROR_SUCCESS)
		{
			if (RegQueryValueEx(hKey, szHour, 0, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
			{
				dwBrightness	= dwValue;
			}

			RegCloseKey(hKey);
		}

		if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("ControlPanel\\Backlight"), 0, 0, &hKey) == ERROR_SUCCESS)
		{
			RegSetValueEx(hKey, _T("Brightness"), 0, REG_DWORD, (LPBYTE)&dwBrightness, sizeof(DWORD));
			RegCloseKey(hKey);
		}

		HANDLE	hEvent	= CreateEvent(NULL, FALSE, TRUE, _T("BackLightChangeEvent"));

		if (hEvent)
		{
			SetEvent(hEvent);
			CloseHandle(hEvent);
		}
	}

	return 0;
}