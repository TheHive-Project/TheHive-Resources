// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include <atlstr.h>  
#include <Windows.h>
#include <Lmcons.h>
#include "ReportPhishing_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

//using namespace ATL;


// CConnect

typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>
IDTImpl;
typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>
CallbackImpl;
typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
RibbonImpl;

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_ReportPhishingLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public RibbonImpl,
	public CallbackImpl,
	public IDTImpl
{
public:
	CConnect()
	{
	}
	STDMETHOD(Invoke)(DISPID dispidMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexceptinfo, UINT *puArgErr)
	{
		HRESULT hr = DISP_E_MEMBERNOTFOUND;
		if (dispidMember == 42)
		{
			hr = CallbackImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);
		}

		if (DISP_E_MEMBERNOTFOUND == hr)
			hr = IDTImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

		return hr;
	}
	DECLARE_REGISTRY_RESOURCEID(IDR_CONNECT)


	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
		COM_INTERFACE_ENTRY(IConnect)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
		COM_INTERFACE_ENTRY(IRibbonCallback)
	END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:


	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		_outlookApplication = Application;
		return S_OK;
	}
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return S_OK;
	}

public:


	HRESULT HrGetResource(int nId,
		LPCTSTR lpType,
		LPVOID* ppvResourceData,
		DWORD* pdwSizeInBytes)
	{
		HMODULE hModule = _AtlBaseModule.GetModuleInstance();
		if (!hModule)
			return E_UNEXPECTED;
		HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
		if (!hRsrc)
			return HRESULT_FROM_WIN32(GetLastError());
		HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
		if (!hGlobal)
			return HRESULT_FROM_WIN32(GetLastError());
		*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
		*ppvResourceData = LockResource(hGlobal);
		return S_OK;
	}

	BSTR GetXMLResource(int nId)
	{
		LPVOID pResourceData = NULL;
		DWORD dwSizeInBytes = 0;
		HRESULT hr = HrGetResource(nId, TEXT("XML"),
			&pResourceData, &dwSizeInBytes);
		if (FAILED(hr))
			return NULL;
		// Assumes that the data is not stored in Unicode.
		CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
		return cbstr.Detach();
	}

	SAFEARRAY* GetOFSResource(int nId)
	{
		LPVOID pResourceData = NULL;
		DWORD dwSizeInBytes = 0;
		if (FAILED(HrGetResource(nId, TEXT("OFS"),
			&pResourceData, &dwSizeInBytes)))
			return NULL;
		SAFEARRAY* psa;
		SAFEARRAYBOUND dim = { dwSizeInBytes, 0 };
		psa = SafeArrayCreate(VT_UI1, 1, &dim);
		if (psa == NULL)
			return NULL;
		BYTE* pSafeArrayData;
		SafeArrayAccessData(psa, (void**)&pSafeArrayData);
		memcpy((void*)pSafeArrayData, pResourceData, dwSizeInBytes);
		SafeArrayUnaccessData(psa);
		return psa;
	}

	// IRibbonExtensibility Methods
public:

	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml)
	{
		if (!RibbonXml)
			return E_POINTER;
		*RibbonXml = GetXMLResource(IDR_XML1);
		return S_OK;
	}

	// IRibbonCallback Methods

	STDMETHOD(ButtonClicked)(IDispatch* ribbon)
	{

		// Get Current Timezone
		SYSTEMTIME st;
		GetLocalTime(&st);

		CString cstrMessage;
		cstrMessage.Format(_T("%d%02d%02d%02d%02d%02d%03d"),
			st.wYear,
			st.wMonth,
			st.wDay,
			st.wHour,
			st.wMinute,
			st.wSecond,
			st.wMilliseconds);

		INT _userSelection = MessageBoxW(NULL, L"The selected email will be reported as phishing email and will delete the email from your Inbox. Click OK to Continue", L"Report Phishing Emails", MB_OKCANCEL| MB_ICONWARNING | MB_APPLMODAL);
		
		if (_userSelection == 1)
		{
			reportPhishing(cstrMessage);
		}
		return S_OK;
	}


	// Initiate the Phishing Repoting Procedure

	void reportPhishing(CString cstrMessage)
	{

		CComQIPtr <Outlook::_Application> spApp(_outlookApplication);
		ATLASSERT(spApp);

		// Outlook User Details
		BSTR jobTitle, department, alias, extNumber, name, emailAddress;
		CComQIPtr <Outlook::_NameSpace> session_User;
		spApp->get_Session(&session_User);
		CComQIPtr <Outlook::Recipient> _userSession;
		session_User->get_CurrentUser(&_userSession);
		CComQIPtr <Outlook::AddressEntry> _userAddressEntry;
		_userSession->get_AddressEntry(&_userAddressEntry);
		CComQIPtr <Outlook::_ExchangeUser> _userExchangeUser;
		_userAddressEntry->GetExchangeUser(&_userExchangeUser);

		_userExchangeUser->get_Name(&name);
		_userExchangeUser->get_BusinessTelephoneNumber(&extNumber);
		_userExchangeUser->get_Alias(&alias);
		_userExchangeUser->get_PrimarySmtpAddress(&emailAddress);
		//_userExchangeUser->get_FirstName(&firstName);
		//_userExchangeUser->get_LastName(&lastName);
		_userExchangeUser->get_JobTitle(&jobTitle);
		_userExchangeUser->get_Department(&department);

		// Get Windows User and Hostname
		char username[UNLEN + 1];
		char hostname[UNLEN + 1];
		DWORD size = UNLEN + 1;
		DWORD hostsize = UNLEN + 1;
		GetUserName((TCHAR*)username, &size);
		GetComputerName((TCHAR*)hostname, &hostsize);


		//Get currently selected Mail item
		CComPtr<Outlook::_Explorer> Explorer;
		spApp->ActiveExplorer(&Explorer);
		CComPtr<Outlook::Selection> Selection;
		HRESULT hr = Explorer->get_Selection(&Selection);
		long lCount = 0;
		hr = Selection->get_Count(&lCount);
		for (long lTmp = 0; lTmp < lCount; ++lTmp)
		{
			CComVariant TmpVar(lTmp + 1);
			CComPtr<IDispatch> Item;
			hr = Selection->Item(TmpVar, &Item);
			CComQIPtr<Outlook::_MailItem> a_mailItem(Item);

			// Create Report Email
			CComQIPtr<Outlook::_MailItem> mailItem;
			spApp->CreateItem(olMailItem, (IDispatch**)&mailItem);

			// Body of the email when sending to the common inbox
			
			BSTR mailBody = _bstr_t("<html><body><table style=""margin: 0 0px;"" border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center""><tbody>"
				"<tr><td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" style=""font - size: large;""><strong>Phishing Email Notification</strong></td></tr>"
				"<tr><td style=""text-align: center;"">&nbsp;</td></tr>"
				"<tr><td style=""text-align: center;"">A possible phishing email concern has been raised by the user. The email has been attached for investigation.</td></tr>"
				"<tr><td style=""text-align: center;"">&nbsp;</td></tr>"
				"</tbody></table>"
				"<table style=""text-align: center; height: 180px; margin-left: auto; margin-right: auto;"" width=""600""><tbody>"
				"<tr>"
				"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;""><strong>Windows Username</strong></span></td>"
				"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;"">") + _bstr_t((TCHAR*)username) + _bstr_t("</span></td></tr>"
					"<tr>"
					"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;""><strong>Windows Hostname</strong></span></td>"
					"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;"">") + _bstr_t((TCHAR*)hostname) + _bstr_t("</span></td></tr>"
						"<tr>"
						"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;""><strong>Name</strong></span></td>"
						"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;"">") + _bstr_t(name) + _bstr_t("</span></td>"
							"</tr>"
							"<tr>"
							"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;""><strong>Alias</strong></span></td>"
							"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;"">") + _bstr_t(alias) + _bstr_t("</span></td>"
								"</tr>"
								"<tr>"
								"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;""><strong>Email</strong></span></td>"
								"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;"">") + _bstr_t(emailAddress) + _bstr_t("</span></td>"
									"</tr>"
									"<tr>"
									"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;""><strong>Title</strong></span></td>"
									"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;"">") + _bstr_t(jobTitle) + _bstr_t("</span></td>"
										"</tr>"
										"<tr>"
										"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;""><strong>Department</strong></span></td>"
										"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#e7e7e7""><span style=""font-size: small;"">") + _bstr_t(department) + _bstr_t("</span></td>"
											"</tr>"
											"<tr>"
											"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;""><strong>Extension</strong></span></td>"
											"<td style=""font-family: 'Helvetica Neue', Arial, Helvetica, Geneva, sans-serif;"" ""text-align: center;"" bgcolor=""#f4f4f4""><span style=""font-size: small;"">") + _bstr_t(extNumber) + _bstr_t("</span></td>"
											"</tr></tbody></table></body></html>");
			
			mailItem->put_Importance(olImportanceHigh);

			// Change this to your organization common mailbox to send the email to
			mailItem->put_To(_bstr_t("spam@yourorganization.com"));

			// Common email subject when received in the inbox for receiving reported email 
			mailItem->put_Subject(_bstr_t("WARNING! Possible Phishing Email Alert"));

			mailItem->put_HTMLBody(mailBody);
			//mailItem->put_Body(mailBody);
			
			// Attach Selected Email

			Outlook::OlAttachmentType::olEmbeddeditem;

			VARIANT emptyParam; VariantInit(&emptyParam); emptyParam.vt = VT_EMPTY;

			VARIANT _attachmentMail, _attachmentType, _attachmentshow; VariantInit(&_attachmentMail);
			_attachmentMail.vt = VT_DISPATCH;
			_attachmentMail.byref = a_mailItem;

			_attachmentType.vt = VT_UI4;
			_attachmentType.ulVal = Outlook::OlAttachmentType::olEmbeddeditem;

			_attachmentshow.vt = VT_UI4;
			_attachmentshow.ulVal = 1;
						
			CComQIPtr <Outlook::Attachment> attachment;
			CComQIPtr <Outlook::Attachments> attachments;
			mailItem->get_Attachments(&attachments);
			attachments->Add(_attachmentMail, _attachmentType, _attachmentshow, emptyParam, &attachment);

			// Send and Delete Report Email
			mailItem->put_DeleteAfterSubmit(True);
			mailItem->Send();

			
			// Delete Selected Email
			//a_mailItem->put_BodyFormat(olFormatPlain);
			//a_mailItem->put_Body(_bstr_t(" "));

			// Replace Subject with Timezone
			a_mailItem->put_Subject(_bstr_t(cstrMessage));


			// Delete All Attachments
			//CComQIPtr <Outlook::Attachment> a_attachment;
			//CComQIPtr <Outlook::Attachments> a_attachments;
			//a_mailItem->get_Attachments(&a_attachments);

			//LONG a_attach_Count;
			//a_attachments->get_Count(&a_attach_Count);

			//if (a_attachments != NULL)
			//{
			//	while (a_attach_Count > 0) {
			//		a_attachments->Remove(1);
			//		a_attach_Count--;
			//	}
			//}

			a_mailItem->Save();
			a_mailItem->Delete();

			// Inform user of email being reported

			//MessageBoxW(NULL, L"Thank you for reporting the Email", L"Report Phishing Emails", MB_OK);

			// Delete the deleted email from Deleted Items

			CComQIPtr <Outlook::_NameSpace> nameSpace;
			spApp->get_Session(&nameSpace);
			
			CComQIPtr <Outlook::MAPIFolder> mapiFolder;
			nameSpace->GetDefaultFolder(olFolderDeletedItems, &mapiFolder);

			CComQIPtr<Outlook::_Items> ad_Items;
			CComQIPtr<Outlook::_Items> ad_restrictItems;
			
			//BSTR _filter = _bstr_t("[Subject] =""") + _bstr_t(cstrMessage) + _bstr_t("""");
			BSTR _filter = _bstr_t("@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = '") + _bstr_t(cstrMessage) + _bstr_t("'");
			mapiFolder->get_Items(&ad_Items);
			ad_Items->Restrict(_filter,&ad_restrictItems);
			LONG _restrictCount;
			ad_restrictItems->get_Count(&_restrictCount);
			while (_restrictCount > 0) {
				CComPtr<IDispatch> ad_Item;
				ad_restrictItems->Item(_variant_t(_restrictCount),&ad_Item);
				CComQIPtr<Outlook::_MailItem> ad_mailItem(ad_Item);
				ad_mailItem->Delete();
				_restrictCount--;

			}

	
		}
	}


private:
	Outlook::_ApplicationPtr _outlookApplication;


};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
