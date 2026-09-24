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

For an offline build, install Catch2 first or set
`FETCHCONTENT_SOURCE_DIR_CATCH2` to an existing Catch2 source checkout.

The tests create hidden Excel instances and temporary unsaved workbooks. They
cover scalar and rectangular range values, formulas, range formatting, merged
ranges, and worksheet creation, ordering, naming, and deletion. They are
registered as serial tests to avoid concurrent Excel automation.
