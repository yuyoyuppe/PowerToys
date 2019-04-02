#include "stdafx.h"
#include "VirtualDesktopsManagerTest.h"

VirtualDesktops::Tests::VirtualDesktopsManagerTest::VirtualDesktopsManagerTest()
    : m_pServiceProvider(nullptr),
    m_pDesktopManager(nullptr)
{
    m_pServiceProvider = GetServiceProvider();

    HRESULT hr = m_pServiceProvider->QueryService(__uuidof(API::IVirtualDesktopManager), &m_pDesktopManager);

    if (FAILED(hr))
        throw std::exception("QueryService for IVirtualDesktopManager failed.");
}

VirtualDesktops::Tests::VirtualDesktopsManagerTest::~VirtualDesktopsManagerTest()
{
    if (m_pDesktopManager)
    {
        m_pDesktopManager->Release();
        m_pDesktopManager = nullptr;
    }

    if (m_pServiceProvider)
    {
        m_pServiceProvider->Release();
        m_pServiceProvider = nullptr;
    }
}

void VirtualDesktops::Tests::VirtualDesktopsManagerTest::GetWindowDesktopId(HWND hWnd)
{
    std::wcout << L"<<< GetWindowDesktopId >>>" << std::endl;

    GUID desktopId = { 0 };
    HRESULT hr = m_pDesktopManager->GetWindowDesktopId(hWnd, &desktopId);

    if (SUCCEEDED(hr))
    {
        std::wcout << L"\t";
        printGuid(desktopId);
        std::wcout << std::endl;
    }

    std::wcout << std::endl;
}
