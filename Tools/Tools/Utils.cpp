#include "Utils.h"
#include <windows.h>
#include <fstream>
#include <iostream>
#include <string>

std::string Utf8ToGbk(const std::string& utf8_str) {
    int wide_size = MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), (int)utf8_str.size(), NULL, 0);
    if (wide_size == 0) {
        return utf8_str;
    }
    
    std::vector<wchar_t> utf16_str(wide_size + 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), (int)utf8_str.size(), &utf16_str[0], wide_size);
    
    int gbk_size = WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], wide_size, NULL, 0, NULL, NULL);
    if (gbk_size == 0) {
        return utf8_str;
    }
    
    std::vector<char> gbk_str(gbk_size + 1, 0);
    WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], wide_size, &gbk_str[0], gbk_size, NULL, NULL);
    
    return std::string(&gbk_str[0]);
}

std::string GbkToUtf8(const std::string& gbk_str) {
	int wide_size = MultiByteToWideChar(CP_ACP, 0, gbk_str.c_str(), (int)gbk_str.size(), NULL, 0);
	if (wide_size == 0) {
		return gbk_str;
	}

	std::vector<wchar_t> utf16_str(wide_size + 1, 0);
	MultiByteToWideChar(CP_ACP, 0, gbk_str.c_str(), (int)gbk_str.size(), &utf16_str[0], wide_size);

	int utf8_size = WideCharToMultiByte(CP_UTF8, 0, &utf16_str[0], wide_size, NULL, 0, NULL, NULL);
	if (utf8_size == 0) {
		return gbk_str;
	}

	std::vector<char> utf8_str(utf8_size + 1, 0);
	WideCharToMultiByte(CP_UTF8, 0, &utf16_str[0], wide_size, &utf8_str[0], utf8_size, NULL, NULL);

	return std::string(&utf8_str[0]);
}

std::string WcharToChar(const std::wstring& wstr) {
	const wchar_t* wccstr = wstr.c_str();
	int len = WideCharToMultiByte(CP_UTF8, 0, wccstr, -1, nullptr, 0, nullptr, nullptr);
	if (len == 0) {
		STDERR << "WideCharToMultiByte failed" << STDEND;
		return "";
	}

	std::vector<char> mbstr(len);
	WideCharToMultiByte(CP_UTF8, 0, wccstr, -1, &mbstr[0], len, nullptr, nullptr);
	return std::string(mbstr.begin(), mbstr.end());
}

std::string MakeDesPath(const std::string& srcPath, const std::string& extension) {
	size_t dotPos = srcPath.rfind('.');
	if (dotPos != std::string::npos) {
		return srcPath.substr(0, dotPos) + extension;
	}
	return srcPath + extension;
}
// 获取文件名（含扩展名）
std::string GetFilename(const std::string& path) {
	size_t pos = path.find_last_of("\\/");
	return (pos == std::string::npos) ? path : path.substr(pos + 1);
}

// 获取文件名（不含扩展名）
std::string GetFilenameWithoutExt(const std::string& path) {
	std::string filename = GetFilename(path);
	size_t dot_pos = filename.find_last_of('.');
	return (dot_pos == std::string::npos) ? filename : filename.substr(0, dot_pos);
}

bool LoadFile(const std::string& filename, std::vector<uint8_t>& buffer) {
	std::ifstream file(filename, std::ios::binary | std::ios::ate);
	if (!file.is_open())
		return false;
	size_t size = file.tellg();
	file.seekg(0, std::ios::beg);
	buffer.resize(size);
	file.read(reinterpret_cast<char*>(buffer.data()), size);
	file.close();
	return true;
}
bool WriteFile(const std::string& filename, std::vector<uint8_t>& data) {
	std::ofstream file(filename, std::ios::binary);
	if (!file.is_open())
		return false;
	file.write(reinterpret_cast<const char*>(data.data()), data.size());
	file.close();
	return true;
}
std::string StrTrim(const std::string& str) {
	size_t start = str.find_first_not_of(" \t\n\r");
	if (start == std::string::npos) return "";
	size_t end = str.find_last_not_of(" \t\n\r");
	return str.substr(start, end - start + 1);
}
std::string StrJoin(const std::vector<std::string>& elements, const std::string& delimiter) {
	if (elements.empty()) {
		return "";
	}

	std::string result;
	result.reserve(elements.size() * 20); // 粗略预分配

	result += elements[0];
	for (size_t i = 1; i < elements.size(); ++i) {
		result += delimiter;
		result += elements[i];
	}

	return result;
}
void StrSplit(const std::string& str, const std::string& delimiters,
	std::function<void(const std::string&)> callback) {
	size_t start = 0;
	size_t end = 0;
	while ((end = str.find_first_of(delimiters, start)) != std::string::npos) {
		if (end != start) {  // 忽略空字符串
			callback(str.substr(start, end - start));
		}
		start = end + 1;
	}
	if (start < str.length()) {
		callback(str.substr(start));
	}
}