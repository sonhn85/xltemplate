#ifndef XLTEMPLATE_H
#define XLTEMPLATE_H

#include <windows.h>
#include "XLCALL.H"
#include "FRAMEWRK.H"

#include "helper.h"
#include "config.h"

/*
 * Helper macros used to generate worksheet function parameter lists,
 * registration type strings, and help text from WORKSHEET_PARAMS.
 *
 * These macros ensure that worksheet function signatures and Excel
 * registration metadata remain consistent.
 */
#define KEY(X) key##X
#define KEY_SEPARATED(X) key##X,
#define VALUE(X) value##X
#define VALUE_SEPARATED(X) value##X,
/* Registration type code for an LPXLOPER12 argument. */
#define TYPE_SYMBOL(X) L"Q"
#define KEY_AND_TYPE_SEPARATED(X) LPXLOPER12 KEY(X),
#define VALUE_AND_TYPE(X) LPXLOPER12 VALUE(X)
#define VALUE_AND_TYPE_SEPARATED(X) LPXLOPER12 VALUE(X),
#define KEY_HELP_SEPARATED(X) L"key,"
#define VALUE_HELP(X) L"value"
#define VALUE_HELP_SEPARATED(X) L"value,"
/* List of parameters (comma-separated) */
#define WORKSHEET_PARAM_LIST WORKSHEET_PARAMS(KEY_SEPARATED, VALUE_SEPARATED, VALUE)
/* Type codes for parameters (not comma-separated) */
#define WORKSHEET_PARAM_STRING WORKSHEET_PARAMS(TYPE_SYMBOL, TYPE_SYMBOL, TYPE_SYMBOL)
/* Type and name pair of parameters (comma-separated) */
#define WORKSHEET_PARAM_AND_TYPE_LIST WORKSHEET_PARAMS(KEY_AND_TYPE_SEPARATED, VALUE_AND_TYPE_SEPARATED, VALUE_AND_TYPE)
/* Comma-separated parameter help text. */
#define HELP_TEXT L"template," WORKSHEET_PARAMS(KEY_HELP_SEPARATED, VALUE_HELP_SEPARATED, VALUE_HELP)
/* Complete Excel registration type string. */
#define TYPE_STRING(PREFIX, SUFFIX) PREFIX WORKSHEET_PARAM_STRING SUFFIX

/*
 * Worksheet function registration metadata.
 *
 * Each entry contains:
 *   (c_function,
 *    excel_function_name,
 *    registration_type_string,
 *    parameter_help_text,
 *    category)
 */
#define XLL_FUNCTIONS(X) \
X(render,     L"RENDER",      TYPE_STRING(L"QD%", L"$"), HELP_TEXT, FUNCTION_CATEGORY) \
X(addin_info, L"RENDER.INFO", L"Q",                      L"",       FUNCTION_CATEGORY)

#ifdef __cplusplus
extern "C" {
#endif

#define DLLEXPORT __declspec(dllexport)

/* Excel add-in initialization entry point. */
DLLEXPORT int WINAPI xlAutoOpen(void);

/* Excel add-in shutdown entry point. */
DLLEXPORT int WINAPI xlAutoClose(void);

/* Called when the add-in is removed from Excel. */
DLLEXPORT int WINAPI xlAutoRemove(void);

/* Release an add-in-owned value returned to Excel. */
DLLEXPORT void WINAPI xlAutoFree12(LPXLOPER12 pxFree);

/* Return add-in information to the Add-In Manager. */
DLLEXPORT LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction);

/*
 * 
 */
DLLEXPORT LPXLOPER12 WINAPI render(
    const wchar_t *template_str,
    WORKSHEET_PARAM_AND_TYPE_LIST
);

/*
 * Retrieve information about the add-in version.
 */
DLLEXPORT LPXLOPER12 WINAPI addin_info(void);

#ifdef __cplusplus
}
#endif

#endif /* XLTEMPLATE_H */
