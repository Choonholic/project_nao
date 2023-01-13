//
// NaoRBN.cpp
// Version 1.0
//
// Copyright (c) 2006 InSoftware House.
// November 1, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>

#define BATTLOW_KEY		_T("ControlPanel\\Notifications\\{A877D663-239C-47a7-9304-0D347F580408}")
#define	OPTIONS_VALUE	_T("Options")
#define	MESSAGE_TITLE	_T("�˸�")
#define	FINISH_MESSAGE	_T("���͸� ���� ��� ��Ÿ���� �ʵ��� �����Ǿ����ϴ�.")

int _tmain(int argc, _TCHAR* argv[])
{
	HKEY	hKey;
	DWORD	dwDisposition;
	DWORD	dwValue	= 0;

	if (RegCreateKeyEx(HKEY_CURRENT_USER, BATTLOW_KEY, 0, NULL, 0, NULL, NULL, &hKey, &dwDisposition) == ERROR_SUCCESS)
	{
		RegSetValueEx(hKey, OPTIONS_VALUE, 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
		RegCloseKey(hKey);
	}

	MessageBox(NULL, FINISH_MESSAGE, MESSAGE_TITLE, MB_ICONINFORMATION | MB_OK | MB_TOPMOST);

	return 0;
}