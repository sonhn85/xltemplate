#ifndef CONFIG_H
#define CONFIG_H

/* Name displayed in Excel Add-In Manager. */
#ifndef ADDIN_VERSION
#define ADDIN_MANAGER_TEXT          "xlTemplate(dev)"
#else
#define ADDIN_MANAGER_TEXT          "xlTemplate"
#endif

/* Add-in version information. */
#ifndef ADDIN_VERSION
#define ADDIN_VERSION               "dev"
#endif

/*
 * XLOPER12 constants defined by the Excel C API.
 *
 * Do not modify.
 */
#define XLTYPEMASK                  0x00000fff
#define XLSTR_MAX_LEN               0x00007fff

/*
 * Worksheet functions expose 128 optional arguments by default,
 * representing 64 key/value pairs.
 *
 * Extend WORKSHEET_PARAMS() to increase the maximum number of
 * key/value arguments supported by worksheet functions.
 */
#define WORKSHEET_PARAMS(X, Y, Z) \
    X(1)  Y(1)  X(2)  Y(2)  X(3)  Y(3)  X(4)  Y(4)   \
    X(5)  Y(5)  X(6)  Y(6)  X(7)  Y(7)  X(8)  Y(8)   \
    X(9)  Y(9)  X(10) Y(10) X(11) Y(11) X(12) Y(12)  \
    X(13) Y(13) X(14) Y(14) X(15) Y(15) X(16) Y(16)  \
    X(17) Y(17) X(18) Y(18) X(19) Y(19) X(20) Y(20)  \
    X(21) Y(21) X(22) Y(22) X(23) Y(23) X(24) Y(24)  \
    X(25) Y(25) X(26) Y(26) X(27) Y(27) X(28) Y(28)  \
    X(29) Y(29) X(30) Y(30) X(31) Y(31) X(32) Y(32)  \
    X(33) Y(33) X(34) Y(34) X(35) Y(35) X(36) Y(36)  \
    X(37) Y(37) X(38) Y(38) X(39) Y(39) X(40) Y(40)  \
    X(41) Y(41) X(42) Y(42) X(43) Y(43) X(44) Y(44)  \
    X(45) Y(45) X(46) Y(46) X(47) Y(47) X(48) Y(48)  \
    X(49) Y(49) X(50) Y(50) X(51) Y(51) X(52) Y(52)  \
    X(53) Y(53) X(54) Y(54) X(55) Y(55) X(56) Y(56)  \
    X(57) Y(57) X(58) Y(58) X(59) Y(59) X(60) Y(60)  \
    X(61) Y(61) X(62) Y(62) X(63) Y(63) X(64) Z(64)

/* Excel function category */
#define FUNCTION_CATEGORY           L"xlTemplate"

#endif /* CONFIG_H */
