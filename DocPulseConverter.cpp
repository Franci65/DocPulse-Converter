#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <dwmapi.h>
#include <commdlg.h>
#include <shellapi.h>
#include <wincodec.h>
#include <vector>
#include <cstdint>
#include <array>
#include <string>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <cwctype>
#include <cstdio>
#include "XlsbReader.hpp"
#include "OfficeXmlReader.hpp"

constexpr COLORREF COL_BG    = RGB(24, 24, 27);
constexpr COLORREF COL_PANEL = RGB(36, 36, 40);
constexpr COLORREF COL_TEXT  = RGB(235, 235, 240);

enum ControlID {
    ID_BTN_SELECT = 101,
    ID_BTN_CONVERT,
    ID_BTN_CLEAR,
    ID_STATIC_FILE,
    ID_STATIC_STATUS
};

struct ImageJpegData {
    std::vector<uint8_t> jpegBytes;
    UINT width = 0;
    UINT height = 0;
};

HWND hMain = nullptr, hFileBox = nullptr, hStatus = nullptr;
HBRUSH hBrushBg = nullptr, hBrushPanel = nullptr;
HFONT hFontUI = nullptr;
std::wstring g_selectedFile;

std::wstring ToLower(const std::wstring& s) {
    std::wstring r = s;
    std::transform(r.begin(), r.end(), r.begin(), [](wchar_t c) { return std::towlower(c); });
    return r;
}

std::wstring GetFileExt(const std::wstring& path) {
    size_t dot = path.find_last_of(L'.');
    return (dot == std::wstring::npos) ? L"" : ToLower(path.substr(dot));
}

std::wstring GetFileName(const std::wstring& path) {
    size_t slash = path.find_last_of(L"\\/");
    return (slash == std::wstring::npos) ? path : path.substr(slash + 1);
}

bool IsImageExt(const std::wstring& ext) {
    return ext == L".jpg" || ext == L".jpeg" || ext == L".png" ||
           ext == L".bmp" || ext == L".gif"  || ext == L".tif" ||
           ext == L".tiff"|| ext == L".webp";
}

bool IsTextExt(const std::wstring& ext) {
    return ext == L".txt" || ext == L".md"   || ext == L".json" ||
           ext == L".xml" || ext == L".log"  || ext == L".csv"  ||
           ext == L".tsv" || ext == L".ini"  || ext == L".cpp"  ||
           ext == L".h"   || ext == L".py"   || ext == L".js";
}

bool IsSupportedFile(const std::wstring& path) {
    std::wstring ext = GetFileExt(path);
    return ext == L".xlsb" || ext == L".docx" || ext == L".rtf" ||
           ext == L".html" || ext == L".htm"  || IsImageExt(ext) || IsTextExt(ext);
}

bool LoadImageAsJpeg(const std::wstring& path, ImageJpegData& out) {
    IWICImagingFactory* factory = nullptr;
    if (FAILED(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory))))
        return false;

    IWICBitmapDecoder* decoder = nullptr;
    if (FAILED(factory->CreateDecoderFromFilename(path.c_str(), nullptr, GENERIC_READ, WICDecodeMetadataCacheOnDemand, &decoder))) {
        factory->Release();
        return false;
    }

    IWICBitmapFrameDecode* frame = nullptr;
    if (FAILED(decoder->GetFrame(0, &frame))) {
        decoder->Release();
        factory->Release();
        return false;
    }

    frame->GetSize(&out.width, &out.height);

    IWICFormatConverter* converter = nullptr;
    factory->CreateFormatConverter(&converter);
    converter->Initialize(frame, GUID_WICPixelFormat24bppBGR, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);

    IStream* memStream = nullptr;
    CreateStreamOnHGlobal(nullptr, TRUE, &memStream);

    IWICBitmapEncoder* encoder = nullptr;
    factory->CreateEncoder(GUID_ContainerFormatJpeg, nullptr, &encoder);
    encoder->Initialize(memStream, WICBitmapEncoderNoCache);

    IWICBitmapFrameEncode* encFrame = nullptr;
    encoder->CreateNewFrame(&encFrame, nullptr);
    encFrame->Initialize(nullptr);
    encFrame->SetSize(out.width, out.height);
    WICPixelFormatGUID fmt = GUID_WICPixelFormat24bppBGR;
    encFrame->SetPixelFormat(&fmt);
    encFrame->WriteSource(converter, nullptr);
    encFrame->Commit();
    encoder->Commit();

    STATSTG stat{};
    memStream->Stat(&stat, STATFLAG_NONAME);
    ULONG sz = static_cast<ULONG>(stat.cbSize.QuadPart);
    out.jpegBytes.resize(sz);

    LARGE_INTEGER zero{};
    memStream->Seek(zero, STREAM_SEEK_SET, nullptr);
    ULONG readBytes = 0;
    memStream->Read(out.jpegBytes.data(), sz, &readBytes);

    encFrame->Release();
    encoder->Release();
    memStream->Release();
    converter->Release();
    frame->Release();
    decoder->Release();
    factory->Release();
    return !out.jpegBytes.empty();
}
std::string EscapePdfText(const std::string& line) {
    std::string out;
    out.reserve(line.size() + 16);
    for (unsigned char c : line) {
        if (c == '(' || c == ')' || c == '\\') {
            out.push_back('\\');
            out.push_back(static_cast<char>(c));
        } else if (c >= 32 && c <= 126) {
            out.push_back(static_cast<char>(c));
        } else if (c == '\t') {
            out += "    ";
        } else {
            out.push_back(' ');
        }
    }
    return out;
}

std::vector<std::string> ReadTextLines(const std::wstring& path) {
    std::ifstream f(path.c_str(), std::ios::binary);
    std::vector<std::string> lines;
    if (!f) return lines;
    std::string content((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    size_t idx0 = 0, idx1 = 1, idx2 = 2;
    if (content.size() >= 3 &&
        static_cast<unsigned char>(content[idx0]) == 0xEF &&
        static_cast<unsigned char>(content[idx1]) == 0xBB &&
        static_cast<unsigned char>(content[idx2]) == 0xBF) {
        content.erase(0, 3);
    }
    std::istringstream iss(content);
    std::string line;
    while (std::getline(iss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        while (line.size() > 95) {
            lines.push_back(line.substr(0, 95));
            line.erase(0, 95);
        }
        lines.push_back(line);
    }
    return lines;
}

bool ConvertSingleFileToPdf(const std::wstring& file, const std::wstring& outPdfPath) {
    struct PdfObj { std::string data; };
    std::vector<PdfObj> objs;
    auto addObj = [&](std::string s) -> int {
        objs.push_back({ std::move(s) });
        return static_cast<int>(objs.size());
    };

    int catalogId = addObj("");
    int pagesId   = addObj("");
    int fontId    = addObj("<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>");
    std::vector<int> pageIds;

    auto appendTextPages = [&](const std::vector<std::string>& lines, double fontSize, double leftX, bool landscape = false) {
        double pageW = landscape ? 842.0 : 595.0;
        double pageH = landscape ? 595.0 : 842.0;
        double startY = landscape ? 560.0 : 805.0;
        const size_t linesPerPage = landscape ? 40 : 58;

        for (size_t i = 0; i < lines.size(); i += linesPerPage) {
            std::ostringstream cs;
            cs << "BT\n/F1 " << fontSize << " Tf\n" << leftX << " " << startY << " Td\n13 TL\n";
            size_t end = std::min(lines.size(), i + linesPerPage);
            for (size_t j = i; j < end; ++j) {
                cs << "(" << EscapePdfText(lines[j]) << ") Tj\nT*\n";
            }
            cs << "ET\n";
            std::string streamStr = cs.str();
            int contentId = addObj("<< /Length " + std::to_string(streamStr.size()) + " >>\nstream\n" + streamStr + "endstream");
            int pageId = addObj("<< /Type /Page /Parent " + std::to_string(pagesId) +
                                " 0 R /MediaBox [0 0 " + std::to_string(static_cast<int>(pageW)) + " " +
                                std::to_string(static_cast<int>(pageH)) + "] /Resources << /Font << /F1 " +
                                std::to_string(fontId) + " 0 R >> >> /Contents " +
                                std::to_string(contentId) + " 0 R >>");
            pageIds.push_back(pageId);
        }
    };


    std::wstring ext = GetFileExt(file);
    if (IsImageExt(ext)) {
        ImageJpegData img;
        if (!LoadImageAsJpeg(file, img) || img.width == 0 || img.height == 0) return false;
        double pageW = 595.0, pageH = 842.0;
        double scale = std::min(pageW / img.width, pageH / img.height);
        double drawW = img.width * scale, drawH = img.height * scale;
        double offX = (pageW - drawW) * 0.5, offY = (pageH - drawH) * 0.5;

        std::string imgDict = "<< /Type /XObject /Subtype /Image /Width " + std::to_string(img.width) +
                              " /Height " + std::to_string(img.height) +
                              " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " +
                              std::to_string(img.jpegBytes.size()) + " >>\nstream\n";
        imgDict.append(reinterpret_cast<const char*>(img.jpegBytes.data()), img.jpegBytes.size());
        imgDict += "\nendstream";
        int imgId = addObj(std::move(imgDict));

        std::ostringstream cs;
        cs << "q\n" << drawW << " 0 0 " << drawH << " " << offX << " " << offY << " cm\n/Im1 Do\nQ\n";
        std::string streamStr = cs.str();
        int contentId = addObj("<< /Length " + std::to_string(streamStr.size()) + " >>\nstream\n" + streamStr + "endstream");
        int pageId = addObj("<< /Type /Page /Parent " + std::to_string(pagesId) +
                            " 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im1 " +
                            std::to_string(imgId) + " 0 R >> >> /Contents " +
                            std::to_string(contentId) + " 0 R >>");
        pageIds.push_back(pageId);
    } else if (IsTextExt(ext)) {
        auto lines = ReadTextLines(file);
        if (lines.empty()) lines.push_back("");
        appendTextPages(lines, 9.5, 40.0);
    } else if (ext == L".docx") {
        auto lines = OfficeXml::ExtractDocxLines(file);
        if (lines.empty()) lines.push_back("(Empty DOCX document)");
        appendTextPages(lines, 9.5, 40.0);
    } else if (ext == L".html" || ext == L".htm" || ext == L".rtf") {
        auto lines = OfficeXml::ExtractHtmlOrRtfLines(file, ext != L".rtf");
        if (lines.empty()) lines.push_back("(Empty document)");
        appendTextPages(lines, 9.5, 40.0);
    } else if (ext == L".xlsb") {
        auto sheets = Xlsb::LoadXlsbWorkbook(file);
        for (const auto& sh : sheets) {
            std::vector<size_t> colWidths;
            for (const auto& row : sh.rows) {
                if (row.size() > colWidths.size()) colWidths.resize(row.size(), 0);
                for (size_t c = 0; c < row.size(); ++c)
                    colWidths[c] = std::max(colWidths[c], std::min<size_t>(row[c].size(), 42));
            }
            std::vector<std::string> tableLines = { "=== SHEET: " + sh.name + " ===", "" };
            for (size_t r = 0; r < sh.rows.size(); ++r) {
                const auto& row = sh.rows[r];
                std::string line;
                for (size_t c = 0; c < colWidths.size(); ++c) {
                    if (c > 0) line += " | ";
                    std::string cell = (c < row.size()) ? row[c] : "";
                    if (cell.size() > 42) cell = cell.substr(0, 39) + "...";
                    line += cell;
                    if (cell.size() < colWidths[c]) line.append(colWidths[c] - cell.size(), ' ');
                }
                while (line.size() > 155) {
                    tableLines.push_back(line.substr(0, 155));
                    line.erase(0, 155);
                }
                tableLines.push_back(line);
                if (r == 0 && !colWidths.empty()) {
                    std::string sep;
                    for (size_t c = 0; c < colWidths.size(); ++c) {
                        if (c > 0) sep += "-+-";
                        sep.append(colWidths[c], '-');
                    }
                    if (sep.size() > 155) sep.resize(155);
                    tableLines.push_back(sep);
                }
            }
            if (tableLines.size() <= 2) tableLines.push_back("(Empty sheet)");
            appendTextPages(tableLines, 8.0, 30.0, true);
        }
    }

    if (pageIds.empty()) return false;
    std::string kids = "[ ";
    for (int pid : pageIds) kids += std::to_string(pid) + " 0 R ";
    kids += "]";
    objs[catalogId - 1].data = "<< /Type /Catalog /Pages " + std::to_string(pagesId) + " 0 R >>";
    objs[pagesId - 1].data   = "<< /Type /Pages /Kids " + kids + " /Count " + std::to_string(pageIds.size()) + " >>";

    std::ofstream out(outPdfPath.c_str(), std::ios::binary);
    if (!out) return false;
    out << "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n";
    std::vector<size_t> offsets(objs.size() + 1, 0);
    for (size_t i = 0; i < objs.size(); ++i) {
        offsets[i + 1] = static_cast<size_t>(out.tellp());
        out << (i + 1) << " 0 obj\n" << objs[i].data << "\nendobj\n";
    }
    size_t xrefPos = static_cast<size_t>(out.tellp());
    out << "xref\n0 " << (objs.size() + 1) << "\n0000000000 65535 f \n";
    std::array<char, 32> xrefLine{};
    for (size_t i = 1; i <= objs.size(); ++i) {
        std::snprintf(xrefLine.data(), xrefLine.size(), "%010zu 00000 n \n", offsets[i]);
        out << xrefLine.data();
    }
    out << "trailer\n<< /Size " << (objs.size() + 1) << " /Root " << catalogId << " 0 R >>\nstartxref\n" << xrefPos << "\n%%EOF\n";
    return true;
}
void SetSelectedFile(const std::wstring& path) {
    if (!IsSupportedFile(path)) {
        MessageBoxW(hMain,
                    L"Unsupported file format.\nPrimary: .xlsb, .docx\nBonus: .jpg, .png, .webp, .bmp, .rtf, .html, .csv, .txt, .md, .json, .xml, .log",
                    L"DocPulse Converter", MB_ICONWARNING);
        return;
    }
    g_selectedFile = path;
    std::wstring info = L"Selected File:\r\n" + GetFileName(path) + L"\r\n\r\nFull Path:\r\n" + path;
    SetWindowTextW(hFileBox, info.c_str());
    SetWindowTextW(hStatus, L"Ready to convert. Click 'Convert to PDF'.");
}

void SelectFileDialog() {
    std::array<wchar_t, MAX_PATH> buf{};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMain;
    ofn.lpstrFilter =
        L"Primary Documents (*.xlsb;*.docx)\0*.xlsb;*.docx\0"
        L"All Supported Files\0*.xlsb;*.docx;*.rtf;*.html;*.htm;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.webp;*.txt;*.md;*.json;*.xml;*.log;*.csv;*.tsv;*.ini;*.cpp;*.h;*.py;*.js\0"
        L"Images (*.jpg;*.png;*.bmp;*.webp)\0*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.webp\0"
        L"Text & Web (*.csv;*.txt;*.md;*.json;*.html;*.rtf)\0*.txt;*.md;*.json;*.xml;*.log;*.csv;*.tsv;*.ini;*.rtf;*.html;*.htm;*.cpp;*.h;*.py;*.js\0"
        L"All Files (*.*)\0*.*\0";
    ofn.lpstrFile = buf.data();
    ofn.nMaxFile = static_cast<DWORD>(buf.size());
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (GetOpenFileNameW(&ofn)) {
        SetSelectedFile(buf.data());
    }
}

void SavePdfAction() {
    if (g_selectedFile.empty()) {
        MessageBoxW(hMain, L"Please select or drag & drop a file first.", L"DocPulse Converter", MB_ICONINFORMATION);
        return;
    }

    std::wstring firstName = GetFileName(g_selectedFile);
    size_t dotPos = firstName.find_last_of(L'.');
    std::wstring baseName = (dotPos != std::wstring::npos) ? firstName.substr(0, dotPos) : firstName;
    baseName += L".pdf";

    std::array<wchar_t, MAX_PATH> outPath{};
    wcsncpy_s(outPath.data(), outPath.size(), baseName.c_str(), _TRUNCATE);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMain;
    ofn.lpstrFilter = L"PDF Document (*.pdf)\0*.pdf\0";
    ofn.lpstrFile = outPath.data();
    ofn.nMaxFile = static_cast<DWORD>(outPath.size());
    ofn.lpstrDefExt = L"pdf";
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&ofn)) {
        SetWindowTextW(hStatus, L"Converting to PDF... Please wait.");
        if (ConvertSingleFileToPdf(g_selectedFile, outPath.data())) {
            std::wstring msg = L"PDF successfully generated:\n" + std::wstring(outPath.data());
            SetWindowTextW(hStatus, L"Done! PDF created successfully.");
            MessageBoxW(hMain, msg.c_str(), L"DocPulse Converter", MB_ICONINFORMATION);
        } else {
            SetWindowTextW(hStatus, L"Error: Could not convert file to PDF.");
            MessageBoxW(hMain, L"Failed to convert the selected file.", L"DocPulse Converter", MB_ICONERROR);
        }
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        BOOL dark = TRUE;
        DwmSetWindowAttribute(hwnd, 20, &dark, sizeof(dark));
        DragAcceptFiles(hwnd, TRUE);

        hBrushBg = CreateSolidBrush(COL_BG);
        hBrushPanel = CreateSolidBrush(COL_PANEL);
        hFontUI = CreateFontW(17, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                              OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");

        HWND bSelect  = CreateWindowW(L"BUTTON", L"Select File...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 15, 15, 130, 32, hwnd, (HMENU)ID_BTN_SELECT, nullptr, nullptr);
        HWND bClear   = CreateWindowW(L"BUTTON", L"Clear",          WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 155, 15, 85, 32, hwnd, (HMENU)ID_BTN_CLEAR, nullptr, nullptr);
        HWND bConvert = CreateWindowW(L"BUTTON", L"Convert to PDF", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 485, 15, 160, 32, hwnd, (HMENU)ID_BTN_CONVERT, nullptr, nullptr);

        hFileBox = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT",
                                   L"Drag & drop an .xlsb or .docx file (or image / text file) anywhere here,\r\nor click 'Select File...' above.",
                                   WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_READONLY,
                                   15, 60, 630, 150, hwnd, (HMENU)ID_STATIC_FILE, nullptr, nullptr);

        hStatus = CreateWindowW(L"STATIC", L"Primary: .xlsb, .docx  |  Bonus: Images, HTML, RTF, CSV, TXT, Code",
                                WS_CHILD | WS_VISIBLE, 15, 222, 630, 24, hwnd, (HMENU)ID_STATIC_STATUS, nullptr, nullptr);

        HWND ctrls[] = { bSelect, bClear, bConvert, hFileBox, hStatus };
        for (HWND c : ctrls) SendMessageW(c, WM_SETFONT, (WPARAM)hFontUI, TRUE);
        break;
    }
    case WM_DROPFILES: {
        HDROP hDrop = reinterpret_cast<HDROP>(wParam);
        std::array<wchar_t, MAX_PATH> path{};
        if (DragQueryFileW(hDrop, 0, path.data(), static_cast<UINT>(path.size()))) {
            SetSelectedFile(path.data());
        }
        DragFinish(hDrop);
        break;
    }
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_PANEL);
        return reinterpret_cast<LRESULT>(hBrushPanel);
    }
    case WM_ERASEBKGND: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(reinterpret_cast<HDC>(wParam), &rc, hBrushBg);
        return 1;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == ID_BTN_SELECT) SelectFileDialog();
        else if (id == ID_BTN_CLEAR) {
            g_selectedFile.clear();
            SetWindowTextW(hFileBox, L"Drag & drop an .xlsb or .docx file (or image / text file) anywhere here,\r\nor click 'Select File...' above.");
            SetWindowTextW(hStatus, L"Primary: .xlsb, .docx  |  Bonus: Images, HTML, RTF, CSV, TXT, Code");
        } else if (id == ID_BTN_CONVERT) {
            SavePdfAction();
        }
        break;
    }
    case WM_DESTROY:
        DeleteObject(hBrushBg);
        DeleteObject(hBrushPanel);
        DeleteObject(hFontUI);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int nCmdShow) {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.lpszClassName = L"DocPulseConverterWnd";
    RegisterClassExW(&wc);

    hMain = CreateWindowExW(0, wc.lpszClassName, L"DocPulse Converter v1.0 - Native XLSB & DOCX to PDF",
                            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                            CW_USEDEFAULT, CW_USEDEFAULT, 675, 295,
                            nullptr, nullptr, hInst, nullptr);

    ShowWindow(hMain, nCmdShow);
    UpdateWindow(hMain);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    CoUninitialize();
    return 0;
}
