# Changelog

Changes that can break building existing code are marked **[breaks build]**,
changes that alter the behavior of existing code are marked **[behavior change]**.
The changes in version 2.1, as well as this changelog, were made by Claude
Opus 5.5 (Anthropic) under the direction of PB.

## 2.1.0

### General

- **[breaks build]** The library names now contain version 2.1, e.g.
  *wxAutoExcel21* instead of *wxAutoExcel20*, so build scripts referring to
  the library by name must be updated.
- **[breaks build]** `GetNumberFormat()`, `SetNumberFormat()`,
  `GetNumberFormatLocal()` and `SetNumberFormatLocal()` were renamed to
  `GetNumberformat()`, `SetNumberformat()`, `GetNumberformatLocal()` and
  `SetNumberformatLocal()` in all classes having them (`wxExcelAboveAverage`,
  `wxExcelDataLabel`, `wxExcelDataLabels`, `wxExcelDisplayFormat`,
  `wxExcelFormatCondition`, `wxExcelRange`, `wxExcelStyle`,
  `wxExcelTickLabels`, `wxExcelTop10` and `wxExcelUniqueValues`). The Windows
  headers define `GetNumberFormat` as a macro expanding to `GetNumberFormatW`
  or `GetNumberFormatA`, which renamed the methods in code including them and
  caused compilation or link errors depending on the order of the includes.
  The Local variants were renamed for consistency.
- *wxAutoExcel_object.h* did not compile unless *wx/log.h* had been included
  before it.
- **[behavior change]** Calls of methods which Excel documents as not returning
  a value no longer report failure when Excel returns a value anyway, which it
  does for many of them (e.g. Workbook.ChangeFileAccess returns True).
- **[behavior change]** Getters returning `long` or `double` now accept both
  numeric types from Excel, which returns some numbers as a different type than
  documented (e.g. `wxExcelApplication::GetMaxIterations()` always failed
  because Excel returns the value as a double).
- **[behavior change]** `WXAUTOEXCEL_RELEASE_UNEXPECTED_IDISPATCH` is now 1 by
  default, so an object returned where a different type is expected is released
  instead of leaked. The code it enables did not compile before.
- **[behavior change]** A member name containing a dot is now reported as an
  error without making the call; previously the error was reported but the call
  was made anyway.
- Objects passed as arguments are now released when a call is not made because
  of an invalid object or a dotted member name, which leaked references and
  could prevent Excel from closing.
- Arguments of type `wxVariant` passed by name lost their name.
- Methods with a `double` result returned it truncated to a whole number, e.g.
  `wxExcelApplication::CentimetersToPoints()` and `InchesToPoints()`.

### CMake build

- **[behavior change]** With multi-configuration generators (Visual Studio,
  Ninja Multi-Config), static and import libraries are created in a subfolder
  of *lib* for each configuration, like the DLLs in *bin*. Previously, Debug
  and Release libraries were in *lib* itself and the others in subfolders.
- **[behavior change]** `find_package(wxAutoExcel)` accepts only the same
  minor version, as minor versions have different library names and may break
  the build; e.g. a request for version 2.0 is no longer satisfied by 2.1.
- **[behavior change]** `find_package(wxAutoExcel)` fails when the selected
  wxWidgets version is not binary compatible with the one wxAutoExcel was built
  with, i.e. differs in the minor version, or in the release number for the
  development versions such as 3.3.x.
- **[behavior change]** The names of the 32-bit DLLs no longer contain *_x32*;
  like the names of wxWidgets DLLs, they have no architecture suffix.
- With MSVC, configuring fails when `wxAutoExcel_BUILD_USE_STATIC_RUNTIME`
  does not match the runtime used by the static wxWidgets libraries; previously
  linking failed with error LNK2038.
- With MSVC, the debug information of the static library is stored in the
  library, instead of in a compiler PDB file which was not installed.
- The PDB file of the DLL is installed for all configurations creating one,
  not only for Debug.
- When the compiler does not report the target architecture, it is determined
  from the pointer size instead of the host processor.
- GCC and clang compile with `-Wall -Wextra`, and `/MP` is used only with the
  Visual Studio generators.
- The list of the source files is no longer maintained by hand:
  *build/CMake/make_files.cmake.bat* was removed, and the files added to
  *include/wx* and *src* are picked up automatically.
- A clear error is reported when the version cannot be read from
  *wxAutoExcel_version.h*.
- The `USE_FOLDERS` global property is not changed when wxAutoExcel is included
  in another project.

### Fixed wrappers

- `wxExcelApplication`: `DDEInitiate()` always failed,
  `GetHighQualityModeForGraphics()` always failed, `SetMaxIterations()` set
  a wrong property.
- **[breaks build]** **[behavior change]** `wxExcelWorksheet::Protect()` took
  the Structure and Windows arguments of Workbook.Protect and failed when they
  were used. It now takes the Worksheet.Protect arguments (DrawingObjects,
  Contents, Scenarios, UserInterfaceOnly and the Allow... options), so the
  second and third arguments of existing calls have a different meaning.
  A `Protect(const wxVariantVector& optionalArgs)` overload was added, like
  the one `SaveAs()` has.
- **[behavior change]** `wxExcelWorksheet::Unprotect()`,
  `wxExcelRange::ClearComments()` and `wxExcelRange::ClearHyperlinks()` reported
  failure even when they succeeded and now return true.
- **[behavior change]** `wxExcelRange::SetValue()` and `SetValue2()` with a null
  `wxVariant` now clear the cells; previously the call failed without changing
  them.
- **[breaks build]** `wxExcelIconCriterion::GetIcon()` and `SetIcon()` use
  `XlIcon` instead of `wxExcelIcon`; they always failed before.
- **[breaks build]** `wxExcelPage` header and footer getters
  (`GetCenterFooter()`, `GetCenterHeader()`, `GetLeftFooter()`,
  `GetLeftHeader()`, `GetRightFooter()`, `GetRightHeader()`) return
  `wxExcelHeaderFooter` instead of `wxString`; they always failed before.
- **[behavior change]** `wxExcelShape::Duplicate()` returns the new shape
  instead of nothing.
- `wxExcelErrors::GetItem()` always failed.
- **[behavior change]** `wxExcelLinearGradient::GetDegree()` and `SetDegree()`
  use `double` instead of `long`.
- **[behavior change]** `wxExcelShapeRange::GetHasChart()` returns
  `MsoTriState` instead of `bool`, as `wxExcelShape::GetHasChart()` does.
- `wxExcelChart::Protect()` never called Excel; its second argument now sets
  DrawingObjects. `wxExcelChart::SetHasAxis()` failed or set a wrong value.
- `wxExcelCharts::Add2()` and `wxExcelAllowEditRanges::Add()` had inverted
  argument checks and so always failed.
- `wxExcelDocumentProperties::Add()` never passed the name and used a wrong
  argument name, so it always failed.
- `wxExcelHyperlinks::Add()` did not pass the screen tip.
- `wxExcelListObjects::Add()` for an external data source did not pass the
  destination and so always failed.
- `wxExcelWorkbooks::OpenDatabase()` shifted the arguments when the command
  text was empty.
- `wxExcelWorkbook::ChangeFileAccess()` failed when an optional argument was
  given, `CheckInWithVersion()` passed the version type to a wrong parameter.
- `wxExcelTextRange2::InsertChartField()` always failed.
  **[breaks build]** `wxExcelTextRange2::RtlRun()` returns nothing instead of
  `wxExcelTextRange2`, as Excel does; it always failed before.
- `wxExcelSparkline::ModifyLocation()` and `SetLocation()`,
  `wxExcelSparklineGroup::Modify()` and `SetLocation()` always failed or were
  ignored by Excel.
- `wxExcelIconSetCondition::SetIconSet()` always failed.
- `wxExcelAboveAverage::GetNumStDev()` and `SetNumStDev()` always failed.
- `wxExcelInterior::GetLinearGradient()` and `GetRectangularGradient()` always
  failed.
- `wxExcelControlFormat::List()` always failed.
- All members of `wxExcelTextFrame2` were private and could not be used.

### Added

- `xlPatternLinearGradient` and `xlPatternRectangularGradient` values of the
  `XlPattern` enum.
- `wxExcelChartFormat::GetTextFrame2()`.
- `wxExcelControlFormat::SetList()` overloads for setting all the list entries
  or a single one.

### Removed

These members could never work, as the corresponding Excel properties or
methods do not exist or are read-only. **[breaks build]** for code using them.

- `wxExcelChartFormat::GetReflection()` and `GetTextFrame()`; use
  `GetTextFrame2()` instead of the latter.
- `wxExcelIconSets::Delete()`, `ModifyAppliesToRange()`, `SetFirstPriority()`
  and `SetLastPriority()`; these are methods of `wxExcelIconSetCondition`.
- Setters of read-only properties:
  `wxExcelAboveAverage::SetType()`,
  `wxExcelCalloutFormat::SetAutoLength()`,
  `wxExcelChartCategory::SetName()`,
  `wxExcelChartFillFormat::SetPattern()`,
  `wxExcelComment::SetAuthor()`,
  `wxExcelConditionValue::SetValue()`,
  `wxExcelDatabar::SetStopIfTrue()`,
  `wxExcelFillFormat::SetPattern()`,
  `wxExcelIconSetCondition::SetStopIfTrue()`,
  `wxExcelListObject::SetSort()`,
  `wxExcelRange::SetHeight()` (use `SetRowHeight()`),
  `wxExcelUserAccess::SetName()`,
  `wxExcelWindow::SetType()`.
