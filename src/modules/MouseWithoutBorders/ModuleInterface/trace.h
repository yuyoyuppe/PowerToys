#pragma once

class Trace
{
public:
    static void RegisterProvider() noexcept;
    static void UnregisterProvider() noexcept;

    class MouseWithoutBorders
    {
    public:
        static void Enable(bool enabled) noexcept;
    };
};
