//
// NaoABC.cpp
// Version 1.0
//
// Copyright (c) 2006 InSoftware House.
// October 26, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

#define	CHECK_INTERVAL	5000	// 5 Seconds
#define	MIN_BRIGHTNESS	0x00
#define	MAX_BRIGHTNESS	0x0A

#define	NAOABC_KEY		_T("Software\\Project Nao\\NaoABC")
#define	NAOABC_MAXVALUE	_T("Maximized")

int _tmain(int argc, _TCHAR* argv[])
{
	BOOL					bActive		= FALSE;
	DWORD					dwValue;
	DWORD					dwNormal	= 0xFFFFFFFF;
	DWORD					dwMaximized	= MAX_BRIGHTNESS;
	HKEY					hKey;
	DWORD					dwType;
	DWORD					dwSize;
	SYSTEM_POWER_STATUS_EX2 spsex2;

	while (TRUE)
	{
		Sleep(CHECK_INTERVAL);

		if (dwNormal == 0xFFFFFFFF)
		{
			dwType	= REG_DWORD;
			dwSize	= sizeof(DWORD);

			if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("ControlPanel\\Backlight"), 0, 0, &hKey) == ERROR_SUCCESS)
			{
				RegQueryValueEx(hKey, _T("Brightness"), 0, &dwType, (LPBYTE)&dwNormal, &dwSize);
				RegCloseKey(hKey);
			}
		}

		GetSystemPowerStatusEx2(&spsex2, sizeof(SYSTEM_POWER_STATUS_EX2), TRUE);

		if (spsex2.ACLineStatus == AC_LINE_ONLINE)
		{
			if (bActive == FALSE)
			{
				dwMaximized	= MAX_BRIGHTNESS;
				bActive		= TRUE;
				dwType		= REG_DWORD;
				dwSize		= sizeof(DWORD);

				if (RegOpenKeyEx(HKEY_CURRENT_USER, NAOABC_KEY, 0, 0, &hKey) == ERROR_SUCCESS)
				{
					if (RegQueryValueEx(hKey, NAOABC_MAXVALUE, 0, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
					{
						dwMaximized	= dwValue;
					}

					RegCloseKey(hKey);
				}

				if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("ControlPanel\\Backlight"), 0, 0, &hKey) == ERROR_SUCCESS)
				{
					RegSetValueEx(hKey, _T("Brightness"), 0, REG_DWORD, (LPBYTE)&dwMaximized, sizeof(DWORD));
					RegCloseKey(hKey);
				}

				HANDLE	hEvent	= CreateEvent(NULL, FALSE, TRUE, _T("BackLightChangeEvent"));

				if (hEvent)
				{
					SetEvent(hEvent);
					CloseHandle(hEvent);
				}
			}
		}
		else
		{
			if (bActive == TRUE)
			{
				bActive	= FALSE;

				if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("ControlPanel\\Backlight"), 0, 0, &hKey) == ERROR_SUCCESS)
				{
					RegSetValueEx(hKey, _T("Brightness"), 0, REG_DWORD, (LPBYTE)&dwNormal, sizeof(DWORD));
					RegCloseKey(hKey);
				}

				HANDLE	hEvent	= CreateEvent(NULL, FALSE, TRUE, _T("BackLightChangeEvent"));

				if (hEvent)
				{
					SetEvent(hEvent);
					CloseHandle(hEvent);
				}
			}
			else
			{
				dwType	= REG_DWORD;
				dwSize	= sizeof(DWORD);

				if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("ControlPanel\\Backlight"), 0, 0, &hKey) == ERROR_SUCCESS)
				{
					RegQueryValueEx(hKey, _T("Brightness"), 0, &dwType, (LPBYTE)&dwValue, &dwSize);
					RegCloseKey(hKey);
				}

				if (dwNormal != dwValue)
				{
					dwNormal	= dwValue;
				}
			}
		}
	}

	return 0;
}