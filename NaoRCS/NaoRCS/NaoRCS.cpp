//
// NaoRCS.cpp
// Version 1.1
//
// Copyright (c) 2006 InSoftware House.
// October 20, 2006
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>
#include "objbase.h"
#include "initguid.h" 
#include "pimstore.h"

#define	MESSAGE_TITLE	_T("알림")
#define	START_TEMP		_T("%d개의 연락처를 찾았습니다.\n\n변환하시겠습니까?")
#define	FINISH_TEMP		_T("%d개의 연락처에 대해 변환을 완료하였습니다.")

#if (_WIN32_WCE < 0x0500)
IPOutlookApp	*g_polApp	= NULL;
#else
IPOutlookApp2	*g_polApp	= NULL;
#endif

BOOL InitPOOM(HWND hWnd)
{
	BOOL	bSuccess	= FALSE;

	if (SUCCEEDED(CoInitializeEx(NULL, 0)))
	{
#if (_WIN32_WCE < 0x0500)
		if (SUCCEEDED(CoCreateInstance(CLSID_Application, NULL, CLSCTX_INPROC_SERVER, IID_IPOutlookApp, reinterpret_cast<void **>(&g_polApp))))
#else
		if (SUCCEEDED(CoCreateInstance(CLSID_Application, NULL, CLSCTX_INPROC_SERVER, IID_IPOutlookApp2, reinterpret_cast<void **>(&g_polApp))))
#endif
		{
			if (SUCCEEDED(g_polApp->Logon((long)hWnd)))
			{
				bSuccess	= TRUE;
			}
		}
	}

	return bSuccess;
}

BOOL UninitPOOM(void)
{
    HRESULT	hr	= E_FAIL;
    
    if (g_polApp != NULL)
	{
		hr	= g_polApp->Logoff();
		hr	= g_polApp->Release();

		CoUninitialize();
	}

    return (SUCCEEDED(hr));    
}

BOOL GetPOOMFolder(int nFolder, IFolder **ppFolder)
{
	return (BOOL)(SUCCEEDED(g_polApp->GetDefaultFolder(nFolder, ppFolder)));
}

int _tmain(int argc, _TCHAR* argv[])
{
	UINT					uDone			= FALSE;
	IFolder					*pFolder		= NULL;
	IPOutlookItemCollection	*pItems			= NULL;
	IContact				*pContact		= NULL;
	int						nItems			= 0;
	BSTR					bstrLastName	= NULL;
	BSTR					bstrFirstName	= NULL;
	BSTR					bstrFileAs		= NULL;
	LPTSTR					lpszFileAs;
	int						i;
	TCHAR					szMessage[256];
	HCURSOR					hCursor	= SetCursor(LoadCursor(NULL, IDC_WAIT));
	
	InitPOOM(NULL);

	if (GetPOOMFolder(olFolderContacts, &pFolder))
	{
		if (SUCCEEDED(pFolder->get_Items(&pItems)))
		{
			pItems->get_Count(&nItems);

			wsprintf(szMessage, START_TEMP, nItems);

			uDone	= MessageBox(NULL, szMessage, MESSAGE_TITLE, MB_ICONQUESTION | MB_YESNO | MB_TOPMOST);

			if (uDone == IDYES)
			{
				ShowCursor(TRUE);

				for (i = 1; i <= nItems; i++)
				{
					if (SUCCEEDED(pItems->Item(i, reinterpret_cast<IDispatch **>(&pContact))))
					{
						pContact->get_LastName(&bstrLastName);
						pContact->get_FirstName(&bstrFirstName);

						lpszFileAs	= new TCHAR[SysStringLen(bstrLastName) + SysStringLen(bstrFirstName) + 1];

						wsprintf(lpszFileAs, _T("%s%s"), bstrLastName, bstrFirstName);

						bstrFileAs	= SysAllocString(lpszFileAs);

						delete [] lpszFileAs;

						pContact->put_FileAs(bstrFileAs);
						pContact->Save();

						SysFreeString(bstrLastName);
						SysFreeString(bstrFirstName);
						SysFreeString(bstrFileAs);

						bstrLastName	= NULL;
						bstrFirstName	= NULL;
						bstrFileAs		= NULL;

						pContact->Release();
					}
				}

				ShowCursor(FALSE);
				SetCursor(hCursor);
			}

			pItems->Release();
		}

		pFolder->Release();
	}

	UninitPOOM();

	if (uDone == IDYES)
	{
		wsprintf(szMessage, FINISH_TEMP, nItems);
		MessageBox(NULL, szMessage, MESSAGE_TITLE, MB_ICONINFORMATION | MB_OK | MB_TOPMOST);
	}

	return 0;
}