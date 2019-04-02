#include "stdafx.h"

#include "VirtualDesktopsManagerInternalTest.h"
#include <Windows.h>

VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::VirtualDesktopsManagerInternalTest()
    : m_pServiceProvider(nullptr),
    m_pDesktopManager(nullptr)
{
    m_pServiceProvider = GetServiceProvider();

    HRESULT hr = m_pServiceProvider->QueryService(
        VirtualDesktops::API::CLSID_VirtualDesktopAPI_Unknown,
        &m_pDesktopManager);

    if (FAILED(hr))
        throw std::exception("QueryService for CLSID_VirtualDesktopAPI_Unknown failed.");
}

VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::~VirtualDesktopsManagerInternalTest()
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

void VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::EnumVirtualDesktops()
{
    std::wcout << L"<<< EnumDesktops >>>" << std::endl;

    IObjectArray *pObjectArray = nullptr;
    HRESULT hr = m_pDesktopManager->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr))
    {
        UINT count;
        hr = pObjectArray->GetCount(&count);

        if (SUCCEEDED(hr))
        {
            std::wcout << L"Count: " << count << std::endl;

            for (UINT i = 0; i < count; i++)
            {
                API::IVirtualDesktop *pDesktop = nullptr;

                if (FAILED(pObjectArray->GetAt(i, __uuidof(API::IVirtualDesktop), (void**)&pDesktop)))
                    continue;

                GUID id = { 0 };
                if (SUCCEEDED(pDesktop->GetID(&id)))
                {
                    std::wcout << L"\t #" << i << L": ";
                    printGuid(id);
                    std::wcout << std::endl;
                }

                pDesktop->Release();
            }
        }

        pObjectArray->Release();
    }

    std::wcout << std::endl;
}

void VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::CurrentVirtualDesktop()
{
    std::wcout << L"<<< GetCurrentVirtualDesktop >>>" << std::endl;

    API::IVirtualDesktop *pDesktop = nullptr;
    HRESULT hr = m_pDesktopManager->GetCurrentDesktop(&pDesktop);

    if (SUCCEEDED(hr))
    {
        GUID id = { 0 };

        if (SUCCEEDED(pDesktop->GetID(&id)))
        {
            std::wcout << L"Current desktop id: ";
            printGuid(id);
            std::wcout << std::endl;
        }

        pDesktop->Release();
    }

    std::wcout << std::endl;
}

void VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::EnumAdjacentDesktops()
{
    std::wcout << L"<<< EnumAdjacentDesktops >>>" << std::endl;

    API::IVirtualDesktop *pDesktop = nullptr;
    HRESULT hr = m_pDesktopManager->GetCurrentDesktop(&pDesktop);

    if (SUCCEEDED(hr))
    {
        GUID id = { 0 };
        API::IVirtualDesktop *pAdjacentDesktop = nullptr;
        hr = m_pDesktopManager->GetAdjacentDesktop(pDesktop, API::AdjacentDesktop::LeftDirection, &pAdjacentDesktop);

        std::wcout << L"At left direction: ";

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(pAdjacentDesktop->GetID(&id)))
                printGuid(id);

            pAdjacentDesktop->Release();
        }
        else
            std::wcout << L"NULL";
        std::wcout << std::endl;

        id = { 0 };
        pAdjacentDesktop = nullptr;
        hr = m_pDesktopManager->GetAdjacentDesktop(pDesktop, API::AdjacentDesktop::RightDirection, &pAdjacentDesktop);

        std::wcout << L"At right direction: ";

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(pAdjacentDesktop->GetID(&id)))
                printGuid(id);

            pAdjacentDesktop->Release();
        }
        else
            std::wcout << L"NULL";
        std::wcout << std::endl;

        pDesktop->Release();
    }

    std::wcout << std::endl;
}

void VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest::ManageVirtualDesktops(DWORD dwSleepMiliseconds)
{
    std::wcout << L"<<< ManageVirtualDesktops >>>" << std::endl;
    std::wcout << L"Sleep period: 2000 ms" << std::endl;

    ::Sleep(dwSleepMiliseconds);


    API::IVirtualDesktop *pDesktop = nullptr;
    HRESULT hr = m_pDesktopManager->GetCurrentDesktop(&pDesktop);

    if (FAILED(hr))
    {
        std::wcout << L"\tFAILED can't get current desktop" << std::endl;
        return;
    }

    std::wcout << L"Creating desktop..." << std::endl;

    API::IVirtualDesktop *pNewDesktop = nullptr;
    hr = m_pDesktopManager->CreateDesktopW(&pNewDesktop);

    if (SUCCEEDED(hr))
    {
        GUID id;
        hr = pNewDesktop->GetID(&id);

        if (FAILED(hr))
        {
            std::wcout << L"\tFAILED GetID" << std::endl;
            pNewDesktop->Release();
            return;
        }

        std::wcout << L"\t";
        printGuid(id);
        std::wcout << std::endl;

        std::wcout << L"Switching to desktop..." << std::endl;
        hr = m_pDesktopManager->SwitchDesktop(pNewDesktop);

        if (FAILED(hr))
        {
            std::wcout << L"\tFAILED SwitchDesktop" << std::endl;
            pNewDesktop->Release();
            return;
        }

        ::Sleep(dwSleepMiliseconds);

        std::wcout << L"Removing desktop..." << std::endl;

        if (SUCCEEDED(hr))
        {
            hr = m_pDesktopManager->RemoveDesktop(pNewDesktop, pDesktop);
            pDesktop->Release();

            if (FAILED(hr))
            {
                std::wcout << L"\tFAILED RemoveDesktop" << std::endl;
                pNewDesktop->Release();
                return;
            }
        }
    }

    std::wcout << std::endl;
}
