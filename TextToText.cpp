////////////////////////////////////////////////////////////////////////////
// TextToText.cpp -- a text converting utility for Win32
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#ifdef MZC_NO_INLINES
    #undef MZC_INLINE
    #define MZC_INLINE
    #include "TextToText_inl.hpp"
    #undef MZC_INLINE
    #define MZC_INLINE inline
#endif

///////////////////////////////////////////////////////////////////////////////

#ifdef UNITTEST
    int main(void)
    {
        #ifdef JAPAN
            assert(lstrcmpA(MWideToAnsi(MAnsiToWide("‚ ‚¢‚¤")), "‚ ‚¢‚¤") == 0);
            assert(lstrcmpW(MAnsiToWide(MWideToAnsi(L"‚ ‚¢‚¤")), L"‚ ‚¢‚¤") == 0);
            assert(lstrcmpA(MUtf8ToAnsi(MAnsiToUtf8("‚ ‚¢‚¤")), "‚ ‚¢‚¤") == 0);
            assert(lstrcmpW(MUtf8ToWide(MWideToUtf8(L"‚ ‚¢‚¤")), L"‚ ‚¢‚¤") == 0);
        #else
            assert(lstrcmpA(MWideToAnsi(MAnsiToWide("TEST123")), "TEST123") == 0);
            assert(lstrcmpW(MAnsiToWide(MWideToAnsi(L"TEST123")), L"TEST123") == 0);
            assert(lstrcmpA(MUtf8ToAnsi(MAnsiToUtf8("TEST123")), "TEST123") == 0);
            assert(lstrcmpW(MUtf8ToWide(MWideToUtf8(L"TEST123")), L"TEST123") == 0);
        #endif
        assert(lstrcmpA(MAnsiToUtf8(MUtf8ToAnsi("\xE3\x81\x82")), "\xE3\x81\x82") == 0);
        assert(lstrcmpA(MWideToUtf8(MUtf8ToWide("\xE3\x81\x82")), "\xE3\x81\x82") == 0);
        return 0;
    }
#endif

///////////////////////////////////////////////////////////////////////////////
