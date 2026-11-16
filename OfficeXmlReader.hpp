#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <filesystem>
#include "XlsbReader.hpp"

namespace OfficeXml {

inline std::string DecodeXmlEntities(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        if (s[i] == '&') {
            if (s.compare(i, 5, "&amp;") == 0)       { out.push_back('&');  i += 4; continue; }
            if (s.compare(i, 4, "&lt;") == 0)        { out.push_back('<');  i += 3; continue; }
            if (s.compare(i, 4, "&gt;") == 0)        { out.push_back('>');  i += 3; continue; }
            if (s.compare(i, 6, "&quot;") == 0)      { out.push_back('"');  i += 5; continue; }
            if (s.compare(i, 6, "&apos;") == 0)      { out.push_back('\''); i += 5; continue; }
            if (s.compare(i, 6, "&nbsp;") == 0)      { out.push_back(' ');  i += 5; continue; }
        }
        out.push_back(s[i]);
    }
    return Xlsb::Utf8ToAsciiClean(out);
}

inline std::vector<std::string> ExtractDocxLines(const std::wstring& docxPath) {
    namespace fs = std::filesystem;
    std::vector<std::string> result;

    wchar_t tmpDir[MAX_PATH] = {};
    GetTempPathW(MAX_PATH, tmpDir);
    fs::path workDir = fs::path(tmpDir) / (L"DocPulse_Docx_" + std::to_wstring(GetCurrentProcessId()));
    std::error_code ec;
    fs::remove_all(workDir, ec);
    fs::create_directories(workDir, ec);

    std::wstring cmd = L"tar.exe -xf \"" + docxPath + L"\" -C \"" + workDir.wstring() + L"\"";
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

    auto bytes = Xlsb::ReadFileBytes(workDir / L"word" / L"document.xml");
    fs::remove_all(workDir, ec);
    if (bytes.empty()) return result;

    std::string xml(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    std::string curPara;
    bool inTextTag = false;

    for (size_t i = 0; i < xml.size(); ++i) {
        if (xml[i] == '<') {
            size_t close = xml.find('>', i);
            if (close == std::string::npos) break;
            std::string tag = xml.substr(i + 1, close - i - 1);

            if (tag == "w:t" || tag.rfind("w:t ", 0) == 0) {
                inTextTag = true;
            } else if (tag == "/w:t") {
                inTextTag = false;
            } else if (tag == "w:tab/" || tag == "w:tab") {
                curPara += "    ";
            } else if (tag == "/w:tc") {
                curPara += " | ";
            } else if (tag == "w:br/" || tag == "w:br" || tag == "/w:p") {
                std::string clean = DecodeXmlEntities(curPara);
                while (clean.size() > 95) {
                    result.push_back(clean.substr(0, 95));
                    clean.erase(0, 95);
                }
                result.push_back(clean);
                curPara.clear();
            }
            i = close;
        } else if (inTextTag) {
            curPara.push_back(xml[i]);
        }
    }
    if (!curPara.empty()) result.push_back(DecodeXmlEntities(curPara));
    return result;
}

inline std::vector<std::string> ExtractHtmlOrRtfLines(const std::wstring& path, bool isHtml) {
    auto bytes = Xlsb::ReadFileBytes(path);
    std::vector<std::string> result;
    if (bytes.empty()) return result;

    std::string raw(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    std::string plain;
    plain.reserve(raw.size());

    if (isHtml) {
        bool inTag = false;
        for (size_t i = 0; i < raw.size(); ++i) {
            if (raw[i] == '<') {
                if (raw.compare(i, 4, "<br>") == 0 || raw.compare(i, 5, "<br/>") == 0 ||
                    raw.compare(i, 4, "</p>") == 0 || raw.compare(i, 6, "</div>") == 0 ||
                    raw.compare(i, 5, "</tr>") == 0) {
                    plain.push_back('\n');
                }
                inTag = true;
            } else if (raw[i] == '>') {
                inTag = false;
            } else if (!inTag) {
                if (raw[i] != '\r' && raw[i] != '\n') plain.push_back(raw[i]);
            }
        }
        plain = DecodeXmlEntities(plain);
    } else {
        // Basic RTF text stripper
        int braceDepth = 0;
        for (size_t i = 0; i < raw.size(); ++i) {
            char c = raw[i];
            if (c == '{') { braceDepth++; }
            else if (c == '}') { if (braceDepth > 0) braceDepth--; }
            else if (c == '\\' && i + 1 < raw.size()) {
                if (raw.compare(i, 4, "\\par") == 0) { plain.push_back('\n'); i += 3; }
                else if (raw.compare(i, 4, "\\tab") == 0) { plain += "    "; i += 3; }
                else {
                    while (i + 1 < raw.size() && (std::isalpha((unsigned char)raw[i + 1]) || std::isdigit((unsigned char)raw[i + 1]) || raw[i + 1] == '-')) ++i;
                    if (i + 1 < raw.size() && raw[i + 1] == ' ') ++i;
                }
            } else if (c != '\r' && c != '\n') {
                plain.push_back(c);
            }
        }
    }

    std::istringstream iss(plain);
    std::string line;
    while (std::getline(iss, line)) {
        while (line.size() > 95) {
            result.push_back(line.substr(0, 95));
            line.erase(0, 95);
        }
        result.push_back(line);
    }
    return result;
}

} // namespace OfficeXml
