# Changelog

Changes that can break building existing code are marked **[breaks build]**,
changes that alter the behavior of existing code are marked **[behavior change]**.
The changes in versions 2.1, 2.2, 2.3, and 2.4, as well as this changelog, were made by
Claude Opus 5.5 (Anthropic) under the direction of PB.

## 2.4.0

### General

- **[breaks build]** The library names now contain version 2.4, e.g.
  *wxAutoExcel24* instead of *wxAutoExcel23*, so build scripts referring to
  the library by name must be updated.

### Added

- `wxExcelRange::AddCommentThreaded()`, as threaded comments could be read
  but not added.
- `wxExcelRange::GetDisplayFormat()`, as there was no way to obtain a
  `wxExcelDisplayFormat`, which returns the formatting of a range including
  that applied by conditional formatting.
- `wxExcelRange::SetMergeCells()`.
- `wxExcelShape::SetLeft()`, `SetHeight()`, and `SetName()`, as only the
  getters of these read-write properties were wrapped.
- The setters of more read-write properties which had only a getter:
  `wxExcelAdjustments::SetItem()`, `wxExcelApplication::SetCursorMovement()`,
  `wxExcelChartGroup::SetAxisGroup()` and `SetDoughnutHoleSize()`,
  `wxExcelChartObject::SetName()`, `wxExcelColorFormat::SetBrightness()`,
  `wxExcelFillFormat::SetTextureAlignment()`,
  `wxExcelParagraphFormat2::SetAlignment()`,
  `wxExcelPlotArea::SetInsideLeft()`, `SetInsideTop()`, `SetInsideWidth()`,
  and `SetInsideHeight()`, `wxExcelRange::SetReadingOrder()` and
  `SetShowDetail()`, `wxExcelShape::SetBackgroundStyle()`,
  `wxExcelShapeRange::SetAlternativeText()`, `SetBlackWhiteMode()`,
  `SetHeight()`, `SetRotation()`, `SetTop()`, `SetVisible()`, and
  `SetWidth()`, `wxExcelSort::SetSortMethod()`,
  `wxExcelSparkHorizontalAxis::SetRightToLeftPlotOrder()`,
  `wxExcelTextFrame2::SetAutoSize()`, and
  `wxExcelWorksheet::SetDisplayRightToLeft()` and `SetEnableSelection()`.
- `wxExcelRange::Group()`, as ranges could be ungrouped but not grouped.
- `wxExcelRangeData`, a block of cells indexed by row and column, and
  `wxExcelRange::GetValues_()`, `SetValues_()`, `GetFormulas_()`, and
  `SetFormulas_()`, which transfer the values or formulas of all cells of
  a range at once, without the SAFEARRAY handling the BulkData sample shows.
  The data must have the same size as the range, otherwise nothing is set and
  an error of the Err_InvalidArgument kind is reported.

### Fixed wrappers

- **[behavior change]** `wxExcelPageSetup::GetZoom()` failed when the Zoom
  property is False, i.e. when the worksheet is scaled with FitToPagesWide
  and FitToPagesTall. It now returns 0 then, and `SetZoom(0)` sets the
  property to False, which could not be done before.

## 2.3.0

### General

- **[breaks build]** The library names now contain version 2.3, e.g.
  *wxAutoExcel23* instead of *wxAutoExcel22*, so build scripts referring to
  the library by name must be updated.
- The library failed to compile with `WXAUTOEXCEL_USE_CHARTS` set to 0 in
  *wxAutoExcel_setup.h*.
- The library no longer uses `nullptr`, so it does not need C++11, as the readme
  says.
- The links in the documentation of the classes lead to the current
  Microsoft documentation instead of the archived pages for Excel 2007,
  where such a page exists.
- **[behavior change]** The `Err_AssertOnInvalidArgument`,
  `Err_AssertOnFailedInvoke`, `Err_AssertOnInvalidReturnType`, and
  `Err_AssertOnOtherError` error mode flags never asserted. They do now, and
  as they are in the default error mode, failed calls assert in debug
  builds by default, as documented. Use `wxExcelObject::SetErrorMode_()` or
  `wxAutoExcelObjectErrorModeOverrider` to change it.

### Fixed wrappers

- **[breaks build]** `wxExcelUserAccessList::DeleteAll()` was declared but
  not defined, so code calling it could not be linked. It now returns
  nothing instead of `wxExcelUserAccess`, like `UserAccessList.DeleteAll`
  in Excel.
- **[behavior change]** `wxExcelRange::ClearContents()` returned false although
  it succeeded, as Excel returns Empty instead of True. Methods returning
  a Boolean now return true when Excel returns Empty or Null for a
  successful call.
- **[breaks build]** The `toAbsolute` parameter of
  `wxExcelApplication::ConvertFormula()` is `XlReferenceType*` instead of
  `XlReferenceStyle*`, as `ToAbsolute` takes `XlReferenceType` values.

### Added

- `wxExcelApplication::GetFullBuild_()`, returning the full build of the
  running Excel, e.g. "16.0.20430.20182", while `GetVersion()` and
  `GetBuild()` return just "16.0" and 20430. The library now links with
  the Windows *version* library.

## 2.2.0

### General

- **[breaks build]** The library names now contain version 2.2, e.g.
  *wxAutoExcel22* instead of *wxAutoExcel21*, so build scripts referring to
  the library by name must be updated.
- **[behavior change]** Getters returning an object no longer report an error
  when Excel returns Nothing, which is not an error, e.g.
  `wxExcelWorksheet::GetAutoFilter()` for a worksheet without an AutoFilter
  or `wxExcelRange::GetComment()` for a cell without a comment. The call
  succeeds and the returned object is empty; previously the error was logged,
  asserted, or thrown, depending on the error mode.
- **[behavior change]** Getters no longer report an error when Excel returns
  Null (or Empty), which it does when a property has no single value, e.g.
  `wxExcelFont::GetBold()` for a range with both bold and regular cells, or
  when a property is not set, e.g. `GetNumberformat()` of a conditional format
  without a number format. The call succeeds and returns the default value of
  its type (0, false, an empty string, an invalid colour, or an empty object);
  the new `wxExcelObject::IsLastResultNull_()` tells that Excel returned Null.
- **[behavior change]** Results expected to be a Boolean or a number accept any
  of these types, as Excel returns numbers as long or double regardless of
  their declared type and some Boolean values as numbers. Previously, e.g.
  `wxExcelApplication::GetControlCharacters()` and `GetMailSystem()` always
  failed, as did any method returning a Boolean when Excel returned a number.

### Fixed wrappers

- `wxExcelApplication::GetClipboardFormats()` always failed.
- `wxExcelCharacters::SetCaption()` and `wxExcelApplication::SendKeys()` took
  a non-const reference to a string, so they could not be called with a
  temporary or a const string.
- **[behavior change]** `wxExcelDataLabel::GetSeparator()` and
  `wxExcelDataLabels::GetSeparator()` failed unless a separator had been set;
  they return an empty string when Excel selects the separator.
- `wxExcelDataLabel::GetCharacters()` always failed.
- **[behavior change]** `wxExcelTextRange2::GetMathZones()` with the default
  arguments failed for a text without a math zone; it returns the whole text
  range, as MathZones() in VBA does.
- `wxExcelWorksheet::GetConsolidationSources()` asserted for a worksheet
  without consolidation; it returns an empty array then.
- `wxExcelTab::GetTintAndShade()` failed for a tab without a color, for which
  Excel returns False; it returns 0 then.
- `wxExcelChart::GetCreator()`, `wxExcelTextFrame2::GetCreator()`,
  `wxExcelFormatCondition::GetApplication()` and `GetParent()`, and
  `wxExcelIconCriteria::GetCount()` were declared but not defined, so using
  them failed to link.
- All members of `wxExcelThreeDFormat` were private and could not be used.

### Added

- `wxExcelObject::IsLastResultNull_()`.
- `wxExcelApplication::GetListSeparator_()`, returning the list separator of
  the Windows regional settings. Excel uses it in the addresses of ranges with
  several areas, e.g. "A1:B2;D4:E5" instead of "A1:B2,D4:E5" in VBA,
  regardless of the automation LCID, contrary to what the documentation of
  `wxExcelObject::SetAutomationLCID_()` said.
- `wxExcelRange::GetFormula()`, `GetFormulaLocal()`, `GetFormulaR1C1()` and
  `GetFormulaR1C1Local()` overloads taking a `wxVariant&`, which return the
  formulas of all cells of a range like `GetValue()` returns their values.
  For more than one cell, Excel returns an array, so the getters returning
  a string fail. The corresponding setters now take a `wxVariant`, so they
  can also set a different formula for each cell; passing a string works
  as before.
- The enumerators of `XlSortOn`, `XlCredentialsMethod`, `XlFormatFilterTypes`,
  `XlGradientFillType`, `XlLookFor` and `XlStdColorScale` with the `xl` prefix
  Excel uses, e.g. `xlSortOnValues`.

### Deprecated

- The enumerators of the above enumerations without the `xl` prefix, e.g.
  `SortOnValues`, which Excel does not use; they remain as aliases.

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
