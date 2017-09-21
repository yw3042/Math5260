// sort.cpp - std::sort from the STL
//!!! Write xll_sort that calls std::sort and hook it up to XLL.SORT in Excel.
AddIn xai_unique(
    Function(XLL_LPOPER, L"?xll_sort", L"XLL.SORT")
    .Arg(XLL_LPOPER, L"range", L"is a range.")
    .FunctionHelp(L"SORT entries from range.")
    .Category(CATEGORY)
);
LPOPER WINAPI xll_sort(LPOPER po)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        if (po->rows() == 1) {
            o.resize(1, std::distance(po->begin(), e));
        }
        else {
            o.resize(std::distance(po->begin(), e), 1);
        }
        std::copy(po->begin(), e, o.begin());
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
        o = OPER(xlerr::NA);
    }

    return &o;
}
