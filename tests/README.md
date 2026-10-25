# wxAutoExcel tests

The tests exercise wxAutoExcel's public API against a locally installed
Microsoft Excel. CMake uses an installed Catch2 3.15 or later when available;
otherwise it downloads a pinned Catch2 revision during the first configure.

Configure and run the tests with:

```console
cmake -S . -B <build-dir> -DwxAutoExcel_BUILD_TESTS=ON
cmake --build <build-dir> --config Debug --target wxAutoExcel_tests
ctest --test-dir <build-dir> -C Debug --output-on-failure
```

Catch2 tag expressions can be passed directly to `wxAutoExcel_tests` when only
one area is needed, for example `wxAutoExcel_tests "[excel][range][value]"`.

For an offline build, install Catch2 first or set
`FETCHCONTENT_SOURCE_DIR_CATCH2` to an existing Catch2 source checkout.

The test executable creates one hidden Excel instance and reuses it for the
entire run. Each test case receives a new temporary unsaved workbook, which is
closed when that test case finishes. The tests cover scalar and rectangular
range values, range acquisition and relative addressing, formulas, range
formatting and gradients, borders, data series, names, outlines, merged
ranges, worksheet creation, ordering, naming, copying, and deletion, chart
sheets, workbooks saved and opened again in a temporary folder, links between
workbooks, workbook and worksheet protection, ranges editable on protected
worksheets, document and custom properties, cell errors, tables and table
styles, sorting and filtering, validation, hyperlinks, comments, page setup,
conditional formatting, charts, error bars, trendlines, shapes of all kinds
and their formats, form controls, OLE objects, windows and their panes,
calculation, custom lists, transferring the values and formulas of all cells
at once, finding and replacing, and iterating over collections and ranges.
They are registered as serial tests to avoid concurrent Excel automation.

## Getter and setter smoke tests

*wxExcel_getters.cpp* creates objects of about 170 classes (charts and
their elements, sparklines, shapes, tables, conditional formats, ranges with
different values in their cells, windows, protection, page setup, and more) and calls every getter which takes no arguments,
checking only that the call succeeds. Then it sets every property which has
such a getter to the value the getter returned. This finds wrappers using a
wrong member name, result type, or argument, which make the getter or the
setter always fail.

The getters and setters are listed in *getter_lists.h*, generated from the
library headers by *tools/generate_getter_lists.py*. Run the script again after
adding getters, setters, or classes. The members which exist only with some of
the `WXAUTOEXCEL_USE_*` options in *wxAutoExcel_setup.h* are listed in separate
macros, empty without the options, and the tests of charts, shapes, and
conditional formats are compiled only with their options, so the tests can be
built with any options. Those which Excel fails to get or set in
the state the tests create (e.g. properties of chart sheets for an embedded
chart), as well as methods named like getters which show a dialog, are
excluded in the script, with the reason. The exclusions were checked against
what the same properties do in VBA.

When an assertion fails, the messages logged by wxAutoExcel and wxWidgets since
the previous assertion are printed before it, to explain why a call failed.
A failed wxWidgets assert is reported as a failure of the current test case.

## Package test

The *package* folder contains a small program checking that an installed
wxAutoExcel package can be used with `find_package()`. It does not need Excel
and is built separately, after installing wxAutoExcel:

```console
cmake -S . -B <build-dir> -DwxAutoExcel_BUILD_INSTALL=ON -DCMAKE_INSTALL_PREFIX=<install-dir>
cmake --build <build-dir> --config Release
cmake --install <build-dir> --config Release
cmake -S tests/package -B <package-build-dir> -DCMAKE_PREFIX_PATH=<install-dir>
cmake --build <package-build-dir> --config Release
```

wxWidgets must be found the same way as when building wxAutoExcel, e.g. with
`wxWidgets_ROOT_DIR` and `wxWidgets_LIB_DIR`.

## Declared functions check

*tools/check_definitions.py* checks that every function declared in the
classes of the library headers is defined. A declared but undefined function
does not make building the library fail, only linking code which calls it.
It needs only Python 3 and is run by the MSVS 2022 workflow:

```console
python tests/tools/check_definitions.py
```
