#include "stdafx.h"
#include "VirtualDesktopsTestBase.h"

IServiceProvider* VirtualDesktops::Tests::VirtualDesktopsTestBase::GetServiceProvider()
{
    IServiceProvider *pServiceProvider;

    HRESULT hr = ::CoCreateInstance(
        API::CLSID_ImmersiveShell,
        NULL,
        CLSCTX_LOCAL_SERVER,
        __uuidof(IServiceProvider),
        (LPVOID*)(&pServiceProvider)
        );

    if (FAILED(hr))
        throw std::exception("CoCreateInstance for IServiceProvider failed.");

    return pServiceProvider;
}

void VirtualDesktops::Tests::VirtualDesktopsTestBase::printGuid(const GUID & guid)
{
    std::wstring guidStr(40, L'\0');
    ::StringFromGUID2(guid, const_cast<LPOLESTR>(guidStr.c_str()), (int)guidStr.length());

    std::wcout << guidStr.c_str();
}
