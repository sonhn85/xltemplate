#include <stdbool.h>
#include <stddef.h>
#include <stdio.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>

#include "xltemplate.h"
#include "uthash.h"
#include "tinytemplate.h"

typedef struct data_dict_t {
    char           *key;
    LPXLOPER12     val;
    char           *utf8_str;   /* Owned UTF-8 representation of val when val is xltypeStr. */
    size_t         utf8_len;
    UT_hash_handle hh;
} data_dict_t;

typedef struct context_t {
    data_dict_t *data_dict;
    char        *output_buffer;
    size_t      output_capacity;
    size_t      output_len;
    bool        failed;
} context_t;

#define BLOCK_SIZE ((size_t)1024)

#define COUNT(X) ((int) (sizeof(X) / sizeof((X)[0])))

#define ERR_MSG_INFO "Error: Failed to get add-in info."

/* Look up each registered function ID and unregister it. */
static void unreg_funcs(const LPXLOPER12 xllPath)
{
    XLOPER12 xRegId;

    #define UNREGISTER_FUNCTION(XLLNAME, XLNAME, PARAMS, HELP_TEXT, CATEGORY) \
    do { \
        if (Excel12f( \
            xlfRegisterId, \
            &xRegId, \
            2, \
            xllPath, \
            TempStr12(XLNAME) \
        ) == xlretSuccess) { \
            Excel12f(xlfUnregister, 0, 1, &xRegId); \
            Excel12f(xlFree, 0, 1, &xRegId); \
        } \
    } while (0);

    XLL_FUNCTIONS(UNREGISTER_FUNCTION)
}

/* Unregister worksheet functions. */
static void xlUnload(void)
{
    XLOPER12 xllPath;

    if (Excel12f(xlGetName, &xllPath, 0) == xlretSuccess)
    {
        unreg_funcs(&xllPath);
        Excel12f(xlFree, NULL, 1, &xllPath);
    }
}

int WINAPI xlAutoOpen(void)
{
    XLOPER12 handle;
    XLOPER12 xllPath;

    if (Excel12f(xlGetName, &xllPath, 0) != xlretSuccess)
        return 0;

    if (Excel12f(xlGetHwnd, &handle, 0) != xlretSuccess)
    {   
        Excel12f(xlFree, 0, 1, &xllPath); 
        return 0;
    }

    HWND hwnd = (HWND)(INT_PTR)handle.val.w;
    Excel12f(xlFree, 0, 1, &handle); 

    /* Register worksheet functions */
    #define REGISTER_FUNCTION(XLLNAME, XLNAME, PARAMS, HELP_TEXT, CATEGORY) \
    do { \
        if(Excel12f(xlfRegister, 0, 7, \
            &xllPath, \
            TempStr12(TO_WSTR(XLLNAME)), \
            TempStr12(PARAMS), \
            TempStr12(XLNAME), \
            TempStr12(HELP_TEXT), \
            TempInt12(1), \
            TempStr12(CATEGORY) \
        ) != xlretSuccess) \
            goto register_failure; \
    } while (0);

    XLL_FUNCTIONS(REGISTER_FUNCTION)

    return 1;

register_failure:

    show_error(hwnd, L"Fail to register worksheet functions");
	xlUnload();

    Excel12f(xlFree, 0, 1, &xllPath); 

    return 0;
}

int WINAPI xlAutoRemove(void)
{
    xlUnload();
	
    return 1;
}

int WINAPI xlAutoClose(void)
{
    return xlAutoRemove();
}

void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
    if (!pxFree || !LPXLOPER12_DLL_FREE(pxFree))
        return;

    if (LPXLOPER12_TYPE(pxFree) == xltypeStr)
        free(pxFree->val.str);

    free(pxFree);
}

LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
    LPXLOPER12 xInfo = NULL;
    
    if (!pxAction)
        return NULL;

    XLOPER12 xIntAction;
    if(Excel12f(
        xlCoerce,
        &xIntAction,
        2, pxAction,
        TempInt12(xltypeInt)
       ) != xlretSuccess)
    {
        return NULL;
    }
    
    if (xIntAction.val.w == 1) 
    {
        xInfo = make_string_cell(ADDIN_MANAGER_TEXT);
    }
    else 
    {
        xInfo = malloc(sizeof(*xInfo));
        if (xInfo)
        {
            xInfo->xltype = xltypeErr | xlbitDLLFree;
            xInfo->val.err = xlerrValue;
        }
    }

    Excel12f(xlFree, NULL, 1, &xIntAction);

    return xInfo;
}

static void free_hash(data_dict_t *data_dict)
{
    data_dict_t *p, *tmp;
    HASH_ITER(hh, data_dict, p, tmp)
    {
        HASH_DEL(data_dict, p);
        free(p->key);
        free(p->utf8_str);
        free(p);
    }
}

/*
 * 
 */
static int parse_params(
    data_dict_t **out_data_dict,
    WORKSHEET_PARAM_AND_TYPE_LIST)
{
    LPXLOPER12 params[] = { WORKSHEET_PARAM_LIST };
    const size_t max_params = sizeof(params) / sizeof(params[0]);

    data_dict_t *data_dict = NULL;
    data_dict_t *pair = NULL;
    char *key_utf8 = NULL;
    char *val_utf8 = NULL;
    size_t utf8_len = 0;

    if (!out_data_dict || max_params % 2 != 0)
        return 0;

    *out_data_dict = NULL;

    size_t npairs = 0;
    size_t pos = max_params;
    while (pos > 0)
    {
        const size_t key_idx = pos - 2;
        LPXLOPER12 key = params[key_idx];
        LPXLOPER12 val = params[key_idx + 1];

        if (!key || !val)
            goto fail;

        if (LPXLOPER12_TYPE(key) != xltypeMissing)
        {
            if (LPXLOPER12_TYPE(key) != xltypeStr)
                goto fail;

            npairs = key_idx / 2 + 1;
            break;
        }

        pos -= 2;
    }

    for (size_t i = 0; i < npairs; i++)
    {
        val_utf8 = NULL;
        utf8_len = 0;

        LPXLOPER12 key = params[i * 2];
        LPXLOPER12 val = params[i * 2 + 1];

        if (!key 
            || !val
            || LPXLOPER12_TYPE(key) != xltypeStr)
        {
            goto fail;
        }

        switch (LPXLOPER12_TYPE(val)) {
            case xltypeInt:
            case xltypeNum:
            case xltypeBool:
                break;
            case xltypeStr:
            {
                const wchar_t *src = val->val.str;
                if (!src || xlstr_to_utf8(&val_utf8, src, &utf8_len) == 0 || !val_utf8)
                    goto fail;

                break;
            }
            default:
                goto fail;
        }

        const wchar_t *src = key->val.str;
        if (!src || xlstr_to_utf8(&key_utf8, src, NULL) == 0 || !key_utf8)
            goto fail;

        data_dict_t *existing = NULL;
        HASH_FIND_STR(data_dict, key_utf8, existing);
        if (existing)
            goto fail;

        pair = calloc(1, sizeof(*pair));
        if (!pair)
            goto fail;

        pair->key = key_utf8;
        key_utf8 = NULL;
        pair->val = val;
        pair->utf8_str = val_utf8;
        val_utf8 = NULL;
        pair->utf8_len = utf8_len;

        HASH_ADD_STR(data_dict, key, pair);
        pair = NULL;
    }

    *out_data_dict = data_dict;

    return 1;

fail:

    free(key_utf8);
    free(val_utf8);

    if (pair)
    {
        free(pair->key);
        free(pair->utf8_str);
        free(pair);
    }

    free_hash(data_dict);

    return 0;
}


/* Return add-in and loaded DuckDB version information. */
LPXLOPER12 WINAPI addin_info(void)
{
    char buf[XLSTR_MAX_LEN];
    int n = snprintf(
        buf,
        sizeof(buf),
        "Add-in version: %s\n"
        ADDIN_VERSION
    );

    if (n < 0 || (size_t)n >= sizeof(buf))
        return make_string_cell(ERR_MSG_INFO);

    return make_string_cell(buf);
}

static bool get_param(
    void *data,
    const char *key,
    size_t len, 
    tinytemplate_value_t *dst_val)
{
    if (!data || !key || !dst_val)
        return false;

    context_t *ctx = data;
    data_dict_t *data_dict = ctx->data_dict;
    data_dict_t *pair = NULL;
    HASH_FIND(hh, data_dict, key, len, pair);
    if (!pair || !pair->val)
        return false;

    LPXLOPER12 val = pair->val;
    switch (LPXLOPER12_TYPE(val)) {
        case xltypeInt:
            tinytemplate_set_int(dst_val, (int64_t)(val->val.w));
            return true;
        case xltypeNum:
            tinytemplate_set_double(dst_val, val->val.num);
            return true;
        case xltypeBool:
            if (val->val.xbool == true)
                tinytemplate_set_string(dst_val, "TRUE", 4);
            else
                tinytemplate_set_string(dst_val, "FALSE", 5);
            return true;
        case xltypeStr:
            if (!pair->utf8_str)
                return false;

            tinytemplate_set_string(dst_val, pair->utf8_str, pair->utf8_len);
            return true;

        default:
            return false;
    }
}

static void writer(
    void *data,
    const char *str,
    size_t len)
{
    if (!data || !str || len == 0)
        return;

    context_t *ctx = data;
    if (ctx->failed)
        return;

    if (len > SIZE_MAX - ctx->output_len - 1)
    {
        ctx->failed = true;
        return;
    }

    const size_t required = ctx->output_len + len + 1;

    if (required > ctx->output_capacity)
    {
        if (required > SIZE_MAX - (BLOCK_SIZE - 1))
        {
            ctx->failed = true;
            return;
        }

        const size_t new_capacity =
            ((required + BLOCK_SIZE - 1) / BLOCK_SIZE) *
            BLOCK_SIZE;

        char *buf = realloc(ctx->output_buffer, new_capacity);
        if (!buf)
        {
            ctx->failed = true;
            return;
        }
        ctx->output_buffer = buf;
        ctx->output_capacity = new_capacity;
    }

    memcpy(ctx->output_buffer + ctx->output_len, str, len);
    ctx->output_len += len;
    ctx->output_buffer[ctx->output_len] = '\0';
}

LPXLOPER12 WINAPI render(
    const wchar_t *template_str,
    WORKSHEET_PARAM_AND_TYPE_LIST)
{
    data_dict_t *data_dict = NULL;
    LPXLOPER12 result = NULL;
    char *template_str_utf8 = NULL;
    tinytemplate_instr_t renderer[32];
    context_t *ctx = NULL;

    size_t utf8_len = 0;
    if (!template_str
        || xlstr_to_utf8(&template_str_utf8, template_str, &utf8_len) == 0
        || !template_str_utf8)
    {
        goto cleanup;
    }

    size_t num_instr = 0;
    if(tinytemplate_compile(
        template_str_utf8,
        utf8_len,
        renderer,
        COUNT(renderer),
        &num_instr,
        NULL,
        0
    ) != TINYTEMPLATE_STATUS_DONE)
    {
        goto cleanup;
    } 

    ctx = calloc(1, sizeof(*ctx));
    if (!ctx)
        goto cleanup;

    if (parse_params(&data_dict, WORKSHEET_PARAM_LIST) == 0)
        goto cleanup;

    ctx->data_dict = data_dict;
    data_dict = NULL;

    if (tinytemplate_eval(
        template_str_utf8,
        renderer,
        ctx,
        get_param,
        writer,
        NULL,
        0
    ) != TINYTEMPLATE_STATUS_DONE || ctx->failed)
    {
        goto cleanup;
    }

    result = make_string_cell(ctx->output_buffer);

cleanup:

    free(template_str_utf8);

    free_hash(data_dict);

    if (ctx)
    {
        free(ctx->output_buffer);
        free_hash(ctx->data_dict);
        free(ctx);
    }

    return result;
}
