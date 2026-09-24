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
range values, formulas, range formatting, merged ranges, and worksheet
creation, ordering, naming, and deletion. They are registered as serial tests
to avoid concurrent Excel automation.
