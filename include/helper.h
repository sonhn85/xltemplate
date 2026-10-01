#ifndef HELPER_H
#define HELPER_H

#include <windows.h>
#include <stddef.h>
#include "XLCALL.H"
#include "config.h"

/* Preprocessor stringification and wide-string helpers */
#define TO_STR(x)               #x
#define TO_STR_DELAY(x)         TO_STR(x)
#define WIDEN(x)                L##x
#define WIDEN_DELAY(x)          WIDEN(x)
#define TO_WSTR(x)              WIDEN_DELAY(TO_STR_DELAY(x))

/* XLOPER12 type and ownership helpers. */
#define LPXLOPER12_TYPE(P)              ((P)->xltype & XLTYPEMASK)
#define LPXLOPER12_DLL_FREE(P)          (((P)->xltype & xlbitDLLFree) != 0)

/* Excel and UTF-8 string conversion. */

/* Convert a length-prefixed Excel string to allocated UTF-8. */
int xlstr_to_utf8(char **dest, const wchar_t *src, size_t *n);
/* Convert UTF-8 to an allocated length-prefixed Excel string. */
int utf8_to_xlstr(wchar_t **dest, const char *src, int n);

/* UI */
void show_error(HWND hwnd, const wchar_t *msg);

/* XLOPER12 creation and validation. */

/* Create an add-in-owned Excel string result. */
LPXLOPER12 make_string_cell(const char *utf8str);

#endif /* HELPER_H */
