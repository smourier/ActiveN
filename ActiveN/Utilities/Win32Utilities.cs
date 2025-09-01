namespace ActiveN.Utilities;

public static class Win32Utilities
{
    public static KEYMODIFIERS GetKEYMODIFIERS()
    {
        KEYMODIFIERS km = 0;
        if (Functions.GetKeyState((int)VIRTUAL_KEY.VK_SHIFT) < 0)
        {
            km |= KEYMODIFIERS.KEYMOD_SHIFT;
        }

        if (Functions.GetKeyState((int)VIRTUAL_KEY.VK_CONTROL) < 0)
        {
            km |= KEYMODIFIERS.KEYMOD_CONTROL;
        }

        if (Functions.GetKeyState((int)VIRTUAL_KEY.VK_MENU) < 0)
        {
            km |= KEYMODIFIERS.KEYMOD_ALT;
        }
        return km;
    }
}
