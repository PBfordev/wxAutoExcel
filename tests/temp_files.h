/////////////////////////////////////////////////////////////////////////////
// Purpose:     Temporary files used by the tests, e.g. for pictures
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef WXAUTOEXCEL_TESTS_TEMP_FILES_H
#define WXAUTOEXCEL_TESTS_TEMP_FILES_H

#include <catch2/catch_test_macros.hpp>

#include <wx/ffile.h>
#include <wx/filename.h>

// A temporary file with the given extension and content, deleted when the
// object is destroyed.
class TempFile
{
public:
    TempFile(const wxString& extension, const void* data, size_t size)
    {
        path = wxFileName::CreateTempFileName("wxae");
        wxRemoveFile(path);
        path += "." + extension;

        wxFFile file(path, "wb");
        REQUIRE(file.IsOpened());
        REQUIRE(file.Write(data, size) == size);
    }

    ~TempFile()
    {
        wxRemoveFile(path);
    }

    const wxString& GetPath() const { return path; }

private:
    wxString path;

    TempFile(const TempFile&);
    TempFile& operator=(const TempFile&);
};

// A temporary 2x2 pixel BMP picture.
class TempPicture : public TempFile
{
public:
    TempPicture() : TempFile("bmp", GetData(), GetSize()) { }

private:
    // the file and info headers, followed by two rows of 2 pixels,
    // each row padded to 4 bytes
    static const unsigned char* GetData()
    {
        static const unsigned char bmp[] =
        {
            'B', 'M', 70, 0, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0,
            40, 0, 0, 0, 2, 0, 0, 0, 2, 0, 0, 0, 1, 0, 24, 0,
            0, 0, 0, 0, 16, 0, 0, 0, 0x13, 0x0B, 0, 0, 0x13, 0x0B, 0, 0,
            0, 0, 0, 0, 0, 0, 0, 0,
            0, 0, 255, 0, 255, 0, 0, 0,
            255, 0, 0, 255, 255, 255, 0, 0
        };
        return bmp;
    }

    static size_t GetSize() { return 70; }
};

// A temporary text file, e.g. to embed in a worksheet as an OLE object.
class TempTextFile : public TempFile
{
public:
    TempTextFile() : TempFile("txt", "wxAutoExcel", 11) { }
};

#endif // WXAUTOEXCEL_TESTS_TEMP_FILES_H
