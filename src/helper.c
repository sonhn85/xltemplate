#include "helper.h"

#include <stdlib.h>

void show_error(HWND hwnd, const wchar_t *msg)
{
    MessageBoxW(hwnd, msg, L"Error", MB_OK | MB_ICONERROR);
}

int xlstr_to_utf8(char **dest, const wchar_t *src, size_t *n)
{
    if (!dest)
        return 0;

    *dest = NULL;

    if (n)
        *n = 0;

    if (!src)
        return 1;

    char *utf8;

    int wchar_count = (unsigned short)src[0];
    /* Empty string */
    if (wchar_count == 0)
    {
        utf8 = malloc(1);
        if (!utf8)
            return 0;

        utf8[0] = '\0';

        *dest = utf8;

        return 1;
    }

    const wchar_t *start = src + 1;

    int utf8_size = WideCharToMultiByte(
        CP_UTF8,
        WC_ERR_INVALID_CHARS,
        start,
        wchar_count,
        NULL,
        0,
        NULL,
        NULL
    );
    if (utf8_size == 0)
        return 0;

    utf8 = malloc((size_t)utf8_size + 1);
    if (!utf8)
        return 0;

    int chars_written = WideCharToMultiByte(
        CP_UTF8,
        WC_ERR_INVALID_CHARS,
        start,
        wchar_count,
        utf8,
        utf8_size,
        NULL,
        NULL
    );
    if (chars_written == 0)
    {
        free(utf8);
        return 0;
    }

    utf8[chars_written] = '\0';

    if (n)
        *n = chars_written;

    *dest = utf8;

    return 1;
}

int utf8_to_xlstr(wchar_t **dest, const char *src, int n)
{

    if (!dest)
        return 0;

    *dest = NULL;

    if (n < -1)
        return 0;

    if (!src)
        return 1;

    wchar_t *xlstr;

    /* Empty Excel string */
    if (n == 0)
    {
        /* One wchar for length prefix, one for trailing L'\0' */
        xlstr = malloc(2 * sizeof(*xlstr));
        if (!xlstr)
            return 0;

        xlstr[0] = 0;
        xlstr[1] = L'\0';

        *dest = xlstr;

        return 1;
    }

    int wchar_count = MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        src,
        n,
        NULL,
        0
    );
    if (wchar_count <= 0)
        return 0;

    int char_only_len = (n == -1) ? wchar_count - 1 : wchar_count;
    if (char_only_len > XLSTR_MAX_LEN)
        return 0;

    /* Allocate space for UTF-16 characters plus Excel length prefix */
    xlstr = malloc(((size_t)wchar_count + 1) * sizeof(*xlstr));
    if (!xlstr)
        return 0;

    int chars_written = MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        src,
        n,
        xlstr+1,
        wchar_count
    );
    if (chars_written == 0)
    {
        free(xlstr);
        return 0;
    }

    /* Exclude terminating L'\0' from Excel string length */
    if (n == -1)
        --chars_written;

    xlstr[0] = (unsigned short)chars_written;

    *dest = xlstr;

    return 1;
}

LPXLOPER12 make_string_cell(const char *utf8str)
{
    if (!utf8str)
        return NULL;

    wchar_t *xlstr = NULL;
    if (utf8_to_xlstr(&xlstr, utf8str, -1) == 0 || !xlstr)
        return NULL;

    LPXLOPER12 result = malloc(sizeof(*result));
    if (!result)
    {
        free(xlstr);
        return NULL;
    }

    /* Excel releases this result through xlAutoFree12(). */
    result->xltype = xltypeStr | xlbitDLLFree; 
    result->val.str = xlstr;

    return result;
}
