////////////////////////////////////////////////////////////////////////////
// TextToText_inl.hpp -- a text converting utility for Win32
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// MAnsiToWide

MZC_INLINE MAnsiToWide::MAnsiToWide() : m_wide(_wcsdup(L""))
{
}

MZC_INLINE MAnsiToWide::MAnsiToWide(const char *ansi) : m_wide(NULL)
{
    assert(ansi);
    #ifdef _WIN32
        const int len = ::MultiByteToWideChar(CP_ACP, 0, ansi, -1, NULL, 0);
        assert(len);
        const size_t siz = len * sizeof(wchar_t);
        wchar_t *psz = reinterpret_cast<wchar_t *>(malloc(siz));
        if (psz)
        {
            ::MultiByteToWideChar(CP_ACP, 0, ansi, -1, psz, len);
            m_wide = psz;
        }
    #else
        std::mbstate_t mb;
        const int len = 1 + std::mbsrtowcs(NULL, &ansi, 0, &mb);
        const size_t siz = len * sizeof(wchar_t);
        wchar_t *psz = reinterpret_cast<wchar_t *>(malloc(siz));
        if (psz)
        {
            std::mbsrtowcs(psz, &ansi, len, &mb);
            m_wide = psz;
        }
    #endif
}

MZC_INLINE MAnsiToWide::MAnsiToWide(const MAnsiToWide& a2w) :
    m_wide(_wcsdup(a2w.m_wide))
{
}

MZC_INLINE MAnsiToWide& MAnsiToWide::operator=(const MAnsiToWide& a2w)
{
    free(m_wide);
    m_wide = _wcsdup(a2w.m_wide);
    return *this;
}

#ifdef MODERN_CXX
    MZC_INLINE MAnsiToWide::MAnsiToWide(MAnsiToWide&& a2w) : m_wide(a2w.m_wide)
    {
        a2w.m_wide = NULL;
    }

    MZC_INLINE MAnsiToWide& MAnsiToWide::operator=(MAnsiToWide&& a2w)
    {
        m_wide = a2w.m_wide;
        a2w.m_wide = NULL;
        return *this;
    }
#endif

MZC_INLINE /*virtual*/ MAnsiToWide::~MAnsiToWide()
{
    free(m_wide);
}

MZC_INLINE bool MAnsiToWide::empty() const
{
    return m_wide[0] == 0;
}

MZC_INLINE size_t MAnsiToWide::size() const
{
    return wcslen(m_wide);
}

MZC_INLINE const wchar_t *MAnsiToWide::c_str() const
{
    return m_wide;
}

MZC_INLINE MAnsiToWide::operator const wchar_t *() const
{
    return m_wide;
}

///////////////////////////////////////////////////////////////////////////////
// MWideToAnsi

MZC_INLINE MWideToAnsi::MWideToAnsi() : m_ansi(strdup(""))
{
}

MZC_INLINE MWideToAnsi::MWideToAnsi(const wchar_t *wide) : m_ansi(NULL)
{
    assert(wide);
    #ifdef _WIN32
        const int len =
            ::WideCharToMultiByte(CP_ACP, 0, wide, -1, NULL, 0, NULL, NULL);
        assert(len);
        const size_t siz = len * sizeof(char);
        char *psz = reinterpret_cast<char *>(malloc(siz));
        if (psz)
        {
            ::WideCharToMultiByte(CP_ACP, 0, wide, -1, psz, len, NULL, NULL);
            m_ansi = psz;
        }
    #else
        std::mbstate_t mb;
        const int len = 1 + std::wcsrtombs(NULL, &wide, 0, &mb);
        const size_t siz = len * sizeof(char);
        char *psz = reinterpret_cast<char *>(malloc(siz));
        if (psz)
        {
            std::wcsrtombs(psz, &wide, len, &mb);
            m_ansi = psz;
        }
    #endif
}

MZC_INLINE MWideToAnsi::MWideToAnsi(const MWideToAnsi& w2a) :
    m_ansi(strdup(w2a.m_ansi))
{
}

MZC_INLINE MWideToAnsi& MWideToAnsi::operator=(const MWideToAnsi& w2a)
{
    free(m_ansi);
    m_ansi = strdup(w2a.m_ansi);
    return *this;
}

#ifdef MODERN_CXX
    MZC_INLINE MWideToAnsi::MWideToAnsi(MWideToAnsi&& w2a) : m_ansi(w2a.m_ansi)
    {
        w2a.m_ansi = NULL;
    }

    MZC_INLINE MWideToAnsi& MWideToAnsi::operator=(MWideToAnsi&& w2a)
    {
        m_ansi = w2a.m_ansi;
        w2a.m_ansi = NULL;
        return *this;
    }
#endif

MZC_INLINE /*virtual*/ MWideToAnsi::~MWideToAnsi()
{
    free(m_ansi);
}

MZC_INLINE bool MWideToAnsi::empty() const
{
    return m_ansi[0] == 0;
}

MZC_INLINE size_t MWideToAnsi::size() const
{
    return strlen(m_ansi);
}

MZC_INLINE const char *MWideToAnsi::c_str() const
{
    return m_ansi;
}

MZC_INLINE MWideToAnsi::operator const char *() const
{
    return m_ansi;
}

///////////////////////////////////////////////////////////////////////////////
// MUtf8ToWide

MZC_INLINE MUtf8ToWide::MUtf8ToWide() : m_wide(_wcsdup(L""))
{
}

MZC_INLINE MUtf8ToWide::MUtf8ToWide(const char *utf8) : m_wide(NULL)
{
    #ifndef _WIN32
        #error You lose.
    #endif
    assert(utf8);
    const int len = ::MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    assert(len);
    const size_t siz = len * sizeof(wchar_t);
    wchar_t *psz = reinterpret_cast<wchar_t *>(malloc(siz));
    if (psz)
    {
        ::MultiByteToWideChar(CP_UTF8, 0, utf8, -1, psz, len);
        m_wide = psz;
    }
}

MZC_INLINE MUtf8ToWide::MUtf8ToWide(const MUtf8ToWide& u2w) :
    m_wide(_wcsdup(u2w.m_wide))
{
}

MZC_INLINE MUtf8ToWide& MUtf8ToWide::operator=(const MUtf8ToWide& u2w)
{
    free(m_wide);
    m_wide = _wcsdup(u2w.m_wide);
    return *this;
}

#ifdef MODERN_CXX
    MZC_INLINE MUtf8ToWide::MUtf8ToWide(MUtf8ToWide&& u2w) : m_wide(u2w.m_wide)
    {
        u2w.m_wide = NULL;
    }

    MZC_INLINE MUtf8ToWide& MUtf8ToWide::operator=(MUtf8ToWide&& u2w)
    {
        m_wide = u2w.m_wide;
        u2w.m_wide = NULL;
        return *this;
    }
#endif

MZC_INLINE /*virtual*/ MUtf8ToWide::~MUtf8ToWide()
{
    free(m_wide);
}

MZC_INLINE bool MUtf8ToWide::empty() const
{
    return m_wide[0] == 0;
}

MZC_INLINE size_t MUtf8ToWide::size() const
{
    return wcslen(m_wide);
}

MZC_INLINE const wchar_t *MUtf8ToWide::c_str() const
{
    return m_wide;
}

MZC_INLINE MUtf8ToWide::operator const wchar_t *() const
{
    return m_wide;
}

///////////////////////////////////////////////////////////////////////////////
// MWideToUtf8

MZC_INLINE MWideToUtf8::MWideToUtf8() : m_utf8(strdup(""))
{
}

MZC_INLINE MWideToUtf8::MWideToUtf8(const wchar_t *wide) : m_utf8(NULL)
{
    #ifndef _WIN32
        #error You lose.
    #endif
    assert(wide);
    const int len =
        ::WideCharToMultiByte(CP_UTF8, 0, wide, -1, NULL, 0, NULL, NULL);
    assert(len);
    const size_t siz = len * sizeof(char);
    char *psz = reinterpret_cast<char *>(malloc(siz));
    if (psz)
    {
        ::WideCharToMultiByte(CP_UTF8, 0, wide, -1, psz, len, NULL, NULL);
        m_utf8 = psz;
    }
}

MZC_INLINE MWideToUtf8::MWideToUtf8(const MWideToUtf8& w2u) :
    m_utf8(strdup(w2u.m_utf8))
{
}

MZC_INLINE MWideToUtf8& MWideToUtf8::operator=(const MWideToUtf8& w2u)
{
    free(m_utf8);
    m_utf8 = strdup(w2u.m_utf8);
    return *this;
}

#ifdef MODERN_CXX
    MZC_INLINE MWideToUtf8::MWideToUtf8(MWideToUtf8&& w2u) : m_utf8(w2u.m_utf8)
    {
        w2u.m_utf8 = NULL;
    }

    MZC_INLINE MWideToUtf8& MWideToUtf8::operator=(MWideToUtf8&& w2u)
    {
        m_utf8 = w2u.m_utf8;
        w2u.m_utf8 = NULL;
        return *this;
    }
#endif

MZC_INLINE /*virtual*/ MWideToUtf8::~MWideToUtf8()
{
    free(m_utf8);
}

MZC_INLINE bool MWideToUtf8::empty() const
{
    return m_utf8[0] == 0;
}

MZC_INLINE size_t MWideToUtf8::size() const
{
    return strlen(m_utf8);
}

MZC_INLINE const char *MWideToUtf8::c_str() const
{
    return m_utf8;
}

MZC_INLINE MWideToUtf8::operator const char *() const
{
    return m_utf8;
}

///////////////////////////////////////////////////////////////////////////////
