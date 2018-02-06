// dllmain.h : Declaration of module class.

class CReportPhishingModule : public ATL::CAtlDllModuleT< CReportPhishingModule >
{
public :
	DECLARE_LIBID(LIBID_ReportPhishingLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_REPORTPHISHING, "{74F7E12D-35A0-4662-88DC-35C33FB283BB}")
};

extern class CReportPhishingModule _AtlModule;
