//
// NaoPND.cpp
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

#define	MESSAGE_TITLE	_T("�˸�")
#define	START_TEMP		_T("%d���� ����ó�� ã�ҽ��ϴ�.\n\n��ȯ�Ͻðڽ��ϱ�?")
#define	FINISH_TEMP		_T("%d���� ����ó�� ���� ��ȯ�� �Ϸ��Ͽ����ϴ�.")

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
	BSTR					bstrPhoneBefore	= NULL;
	BSTR					bstrPhone		= NULL;
	LPTSTR					lpszPhoneBefore;
	LPTSTR					lpszPhone;
	int						i;
	DWORD					j;
	int						nPos;
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
						// Mobile Phone

						pContact->get_MobileTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_MobileTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}
						
						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Business Phone

						pContact->get_BusinessTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_BusinessTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}
						
						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Business Phone 2

						pContact->get_Business2TelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_Business2TelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Home Phone

						pContact->get_HomeTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_HomeTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Home Phone 2

						pContact->get_Home2TelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_Home2TelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Business Fax

						pContact->get_BusinessFaxNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_BusinessFaxNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Home Fax

						pContact->get_HomeFaxNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_HomeFaxNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}
						
						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Car Telephone Number

						pContact->get_CarTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_CarTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Radio Telephone Number

						pContact->get_RadioTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_RadioTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Assistant Telephone Number

						pContact->get_AssistantTelephoneNumber(&bstrPhoneBefore);

						lpszPhoneBefore	= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						lpszPhone		= new TCHAR[SysStringLen(bstrPhoneBefore) + 1];
						nPos			= 0;

						memset(lpszPhone, 0, sizeof(TCHAR) * (SysStringLen(bstrPhoneBefore) + 1));

						_tcscpy(lpszPhoneBefore, bstrPhoneBefore);

						for (j = 0; j < _tcslen(lpszPhoneBefore); j++)
						{
							if ((lpszPhoneBefore[j] >= _T('0')) && (lpszPhoneBefore[j] <= _T('9')))
							{
								lpszPhone[nPos]	= lpszPhoneBefore[j];
								nPos++;
							}
						}

						if (_tcslen(lpszPhone) > 0)
						{
							bstrPhone	= SysAllocString(lpszPhone);

							pContact->put_AssistantTelephoneNumber(bstrPhone);

							SysFreeString(bstrPhone);

							bstrPhone	= NULL;
						}

						SysFreeString(bstrPhoneBefore);

						bstrPhoneBefore	= NULL;

						delete [] lpszPhoneBefore;
						delete [] lpszPhone;

						// Save

						pContact->Save();
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