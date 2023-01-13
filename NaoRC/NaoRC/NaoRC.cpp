//
// NaoRC.cpp
// Version 1.0
//
// Copyright (c) 2007 Infinite Revolution.
// September 30, 2007
//

#include "stdafx.h"
#include <windows.h>
#include <commctrl.h>
#include "objbase.h"
#include "initguid.h"
#include "pimstore.h"

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
	IFolder					*pFolder		= NULL;
	IPOutlookItemCollection	*pItems			= NULL;
	IContact				*pContact		= NULL;
	int						nItems			= 0;
	BSTR					bstrPhone		= NULL;
	int						i;
	HCURSOR					hCursor	= SetCursor(LoadCursor(NULL, IDC_WAIT));
	
	InitPOOM(NULL);

	if (GetPOOMFolder(olFolderContacts, &pFolder))
	{
		if (SUCCEEDED(pFolder->get_Items(&pItems)))
		{
			pItems->get_Count(&nItems);

			ShowCursor(TRUE);

			for (i = 1; i <= nItems; i++)
			{
				if (SUCCEEDED(pItems->Item(i, reinterpret_cast<IDispatch **>(&pContact))))
				{
					// Mobile Phone
					pContact->get_MobileTelephoneNumber(&bstrPhone);
					pContact->put_MobileTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Business Phone
					pContact->get_BusinessTelephoneNumber(&bstrPhone);
					pContact->put_BusinessTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Business Phone 2
					pContact->get_Business2TelephoneNumber(&bstrPhone);
					pContact->put_Business2TelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Home Phone
					pContact->get_HomeTelephoneNumber(&bstrPhone);
					pContact->put_HomeTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Home Phone 2
					pContact->get_Home2TelephoneNumber(&bstrPhone);
					pContact->put_Home2TelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Business Fax
					pContact->get_BusinessFaxNumber(&bstrPhone);
					pContact->put_BusinessFaxNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Home Fax
					pContact->get_HomeFaxNumber(&bstrPhone);
					pContact->put_HomeFaxNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Car Telephone Number
					pContact->get_CarTelephoneNumber(&bstrPhone);
					pContact->put_CarTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Radio Telephone Number
					pContact->get_RadioTelephoneNumber(&bstrPhone);
					pContact->put_RadioTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Assistant Telephone Number
					pContact->get_AssistantTelephoneNumber(&bstrPhone);
					pContact->put_AssistantTelephoneNumber(bstrPhone);
					SysFreeString(bstrPhone);

					bstrPhone	= NULL;

					// Save
					pContact->Save();
					pContact->Release();
				}

				ShowCursor(FALSE);
				SetCursor(hCursor);
			}

			pItems->Release();
		}

		pFolder->Release();
	}

	UninitPOOM();
	return 0;
}
