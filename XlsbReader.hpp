#pragma once
#include <windows.h>
#include <shellapi.h>
#include <vector>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <cmath>
#include <cstdio>
#include <filesystem>

namespace Xlsb {

struct SheetData {
    std::string name;
    std::vector<std::vector<std::string>> rows;
};

inline std::string WideToUtf8(const wchar_t* wstr, size_t len) {
    if (len == 0) return "";
    int sz = WideCharToMultiByte(CP_UTF8, 0, wstr, static_cast<int>(len), nullptr, 0, nullptr, nullptr);
    if (sz <= 0) return "";
    std::string out(sz, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr, static_cast<int>(len), out.data(), sz, nullptr, nullptr);
    return out;
}

inline std::string Utf8ToAsciiClean(const std::string& s) {
    // Normalize common Romanian/European diacritics for standard PDF Type1 fonts
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        unsigned char c = static_cast<unsigned char>(s[i]);
        if (c < 128) {
            out.push_back(static_cast<char>(c));
        } else if (i + 1 < s.size()) {
            unsigned char c2 = static_cast<unsigned char>(s[i + 1]);
            uint16_t cp = (static_cast<uint16_t>(c) << 8) | c2;
            if (cp == 0xC482 || cp == 0xC382) { out.push_back('A'); ++i; }
            else if (cp == 0xC483 || cp == 0xC3A2) { out.push_back('a'); ++i; }
            else if (cp == 0xC38E) { out.push_back('I'); ++i; }
            else if (cp == 0xC3AE) { out.push_back('i'); ++i; }
            else if (cp == 0xC898 || cp == 0xC59E) { out.push_back('S'); ++i; }
            else if (cp == 0xC899 || cp == 0xC59F) { out.push_back('s'); ++i; }
            else if (cp == 0xC89A || cp == 0xC5A2) { out.push_back('T'); ++i; }
            else if (cp == 0xC89B || cp == 0xC5A3) { out.push_back('t'); ++i; }
            else { out.push_back(' '); ++i; }
        }
    }
    return out;
}

class BinReader {
    const uint8_t* ptr_;
    size_t size_;
    size_t pos_ = 0;
public:
    BinReader(const uint8_t* p, size_t s) : ptr_(p), size_(s) {}
    bool eof() const { return pos_ >= size_; }
    size_t pos() const { return pos_; }
    void skip(size_t n) { pos_ = (pos_ + n <= size_) ? (pos_ + n) : size_; }

    uint8_t u8() { return (pos_ < size_) ? ptr_[pos_++] : 0; }
    uint16_t u16() {
        if (pos_ + 2 > size_) { pos_ = size_; return 0; }
        uint16_t v = static_cast<uint16_t>(ptr_[pos_]) | (static_cast<uint16_t>(ptr_[pos_ + 1]) << 8);
        pos_ += 2;
        return v;
    }
    uint32_t u32() {
        if (pos_ + 4 > size_) { pos_ = size_; return 0; }
        uint32_t v = static_cast<uint32_t>(ptr_[pos_]) |
                     (static_cast<uint32_t>(ptr_[pos_ + 1]) << 8) |
                     (static_cast<uint32_t>(ptr_[pos_ + 2]) << 16) |
                     (static_cast<uint32_t>(ptr_[pos_ + 3]) << 24);
        pos_ += 4;
        return v;
    }
    double f64() {
        if (pos_ + 8 > size_) { pos_ = size_; return 0.0; }
        double d = 0.0;
        memcpy(&d, ptr_ + pos_, 8);
        pos_ += 8;
        return d;
    }
    uint32_t recId() {
        uint32_t id = 0;
        for (int i = 0; i < 4 && pos_ < size_; ++i) {
            uint8_t b = ptr_[pos_++];
            id |= static_cast<uint32_t>(b & 0x7F) << (7 * i);
            if ((b & 0x80) == 0) break;
        }
        return id;
    }
    uint32_t recLen() {
        uint32_t len = 0;
        for (int i = 0; i < 4 && pos_ < size_; ++i) {
            uint8_t b = ptr_[pos_++];
            len |= static_cast<uint32_t>(b & 0x7F) << (7 * i);
            if ((b & 0x80) == 0) break;
        }
        return len;
    }
    std::string xlWideString() {
        uint32_t chars = u32();
        if (chars == 0xFFFFFFFF || pos_ + static_cast<size_t>(chars) * 2 > size_) {
            pos_ = size_;
            return "";
        }
        std::string s = WideToUtf8(reinterpret_cast<const wchar_t*>(ptr_ + pos_), chars);
        pos_ += static_cast<size_t>(chars) * 2;
        return Utf8ToAsciiClean(s);
    }
};

inline std::vector<uint8_t> ReadFileBytes(const std::filesystem::path& p) {
    std::ifstream f(p, std::ios::binary);
    if (!f) return {};
    return std::vector<uint8_t>((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
}

inline std::vector<std::string> ParseSharedStrings(const std::filesystem::path& binPath) {
    std::vector<std::string> sst;
    auto data = ReadFileBytes(binPath);
    if (data.empty()) return sst;
    BinReader br(data.data(), data.size());
    while (!br.eof()) {
        uint32_t id = br.recId();
        uint32_t len = br.recLen();
        size_t nextPos = br.pos() + len;
        if (id == 0x0013 && len >= 5) { // BrtSSTItem
            br.u8(); // flags
            sst.push_back(br.xlWideString());
        }
        br.skip(nextPos - br.pos());
    }
    return sst;
}
inline std::string FormatNumber(double val) {
    if (std::fabs(val - std::round(val)) < 1e-9 && std::fabs(val) < 1e15) {
        long long iv = static_cast<long long>(std::llround(val));
        return std::to_string(iv);
    }
    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.4f", val);
    std::string s(buf);
    while (s.find('.') != std::string::npos && (s.back() == '0' || s.back() == '.')) {
        bool isDot = (s.back() == '.');
        s.pop_back();
        if (isDot) break;
    }
    return s;
}

inline double DecodeRk(uint32_t rk) {
    double val = 0.0;
    if (rk & 0x02) {
        int32_t iv = static_cast<int32_t>(rk) >> 2;
        val = static_cast<double>(iv);
    } else {
        uint64_t bits = static_cast<uint64_t>(rk & 0xFFFFFFFC) << 32;
        memcpy(&val, &bits, 8);
    }
    if (rk & 0x01) val /= 100.0;
    return val;
}

inline SheetData ParseWorksheet(const std::filesystem::path& binPath,
                                const std::vector<std::string>& sst,
                                const std::string& sheetName) {
    SheetData sheet;
    sheet.name = sheetName;
    auto data = ReadFileBytes(binPath);
    if (data.empty()) return sheet;

    BinReader br(data.data(), data.size());
    std::vector<std::string> curRow;
    bool inRow = false;

    auto setCell = [&](uint32_t col, const std::string& val) {
        if (col > 256) return;
        if (curRow.size() <= col) curRow.resize(col + 1);
        curRow[col] = val;
    };

    while (!br.eof()) {
        uint32_t id = br.recId();
        uint32_t len = br.recLen();
        size_t nextPos = br.pos() + len;

        if (id == 0x0000 && len >= 4) { // BrtRowHdr
            if (inRow && !curRow.empty()) sheet.rows.push_back(std::move(curRow));
            curRow.clear();
            inRow = true;
        } else if (id == 0x0002 && len >= 12) { // BrtCellRk
            uint32_t col = br.u32();
            br.u32(); // styleRef
            setCell(col, FormatNumber(DecodeRk(br.u32())));
        } else if (id == 0x0004 && len >= 9) { // BrtCellBool
            uint32_t col = br.u32();
            br.u32();
            setCell(col, br.u8() ? "TRUE" : "FALSE");
        } else if (id == 0x0005 && len >= 16) { // BrtCellReal
            uint32_t col = br.u32();
            br.u32();
            setCell(col, FormatNumber(br.f64()));
        } else if (id == 0x0006 && len >= 12) { // BrtCellSt (inline string)
            uint32_t col = br.u32();
            br.u32();
            setCell(col, br.xlWideString());
        } else if (id == 0x0007 && len >= 12) { // BrtCellIsst (shared string)
            uint32_t col = br.u32();
            br.u32();
            uint32_t idx = br.u32();
            if (idx < sst.size()) setCell(col, sst[idx]);
        } else if (id == 0x0008 && len >= 12) { // BrtFmlaString
            uint32_t col = br.u32();
            br.u32();
            setCell(col, br.xlWideString());
        } else if (id == 0x0009 && len >= 16) { // BrtFmlaNum
            uint32_t col = br.u32();
            br.u32();
            setCell(col, FormatNumber(br.f64()));
        }
        br.skip(nextPos - br.pos());
    }
    if (inRow && !curRow.empty()) sheet.rows.push_back(std::move(curRow));
    return sheet;
}

inline std::vector<SheetData> LoadXlsbWorkbook(const std::wstring& xlsbPath) {
    namespace fs = std::filesystem;
    std::vector<SheetData> sheets;

static uint32_t batchSeq = 0;
++batchSeq;
wchar_t tmpDir[MAX_PATH] = {};
GetTempPathW(MAX_PATH, tmpDir);
fs::path workDir = fs::path(tmpDir) / (L"DocPulse_Xlsb_" + std::to_wstring(GetCurrentProcessId()) + L"_" + std::to_wstring(batchSeq));

    std::error_code ec;
    fs::remove_all(workDir, ec);
    fs::create_directories(workDir, ec);

    // Extract .xlsb archive silently via built-in tar.exe (bsdtar on Win10/11)
    std::wstring cmd = L"tar.exe -xf \"" + xlsbPath + L"\" -C \"" + workDir.wstring() + L"\"";
    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi{};
    if (CreateProcessW(nullptr, cmd.data(), nullptr, nullptr, FALSE, CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        WaitForSingleObject(pi.hProcess, 15000);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }

    auto sst = ParseSharedStrings(workDir / L"xl" / L"sharedStrings.bin");

    // Parse workbook.bin for actual sheet names
    std::vector<std::string> sheetNames;
    auto wbData = ReadFileBytes(workDir / L"xl" / L"workbook.bin");
    if (!wbData.empty()) {
        BinReader br(wbData.data(), wbData.size());
        while (!br.eof()) {
            uint32_t id = br.recId();
            uint32_t len = br.recLen();
            size_t nextPos = br.pos() + len;
            if (id == 0x009C && len >= 12) { // BrtBundleSh
                br.u32(); // hsState
                br.u32(); // iTabID
                br.xlWideString(); // strRelID
                sheetNames.push_back(br.xlWideString());
            }
            br.skip(nextPos - br.pos());
        }
    }

    for (size_t i = 1; i <= 64; ++i) {
        fs::path shPath = workDir / L"xl" / L"worksheets" / (L"sheet" + std::to_wstring(i) + L".bin");
        if (!fs::exists(shPath, ec)) break;
        std::string title = (i - 1 < sheetNames.size() && !sheetNames[i - 1].empty())
                            ? sheetNames[i - 1]
                            : ("Sheet " + std::to_string(i));
        sheets.push_back(ParseWorksheet(shPath, sst, title));
    }

    fs::remove_all(workDir, ec);
    return sheets;
}

} // namespace Xlsb
