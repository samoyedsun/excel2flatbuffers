#include "Utils.h"
#include <windows.h>
#include <fstream>
#include <iostream>
#include <string>

#define MAX_ENCODING_BUFFER_SIZE 4096

char* UTF8ToGB2312(const char* utf8) {
	if (!utf8 || *utf8 == '\0') {
		thread_local char empty[] = "";
		return empty;
	}

	thread_local char buffer[MAX_ENCODING_BUFFER_SIZE];
	thread_local wchar_t wbuffer[MAX_ENCODING_BUFFER_SIZE / 2];

	int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
	if (wlen == 0 || wlen >= MAX_ENCODING_BUFFER_SIZE / 2) {
		return nullptr;  // 字符串太长
	}

	MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wbuffer, wlen);

	int len = WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, nullptr, 0, nullptr, nullptr);
	if (len == 0 || len >= MAX_ENCODING_BUFFER_SIZE) {
		return nullptr;
	}

	WideCharToMultiByte(CP_ACP, 0, wbuffer, -1, buffer, len, nullptr, nullptr);

	return buffer;
}
char* GB2312ToUTF8(const char* gb2312) {
	if (!gb2312 || *gb2312 == '\0') {
		thread_local char empty[] = "";
		return empty;
	}

	thread_local char buffer[MAX_ENCODING_BUFFER_SIZE];
	thread_local wchar_t wbuffer[MAX_ENCODING_BUFFER_SIZE / 2];

	int wlen = MultiByteToWideChar(CP_ACP, 0, gb2312, -1, nullptr, 0);
	if (wlen == 0 || wlen >= MAX_ENCODING_BUFFER_SIZE / 2) {
		return nullptr;
	}

	MultiByteToWideChar(CP_ACP, 0, gb2312, -1, wbuffer, wlen);

	int len = WideCharToMultiByte(CP_UTF8, 0, wbuffer, -1, nullptr, 0, nullptr, nullptr);
	if (len == 0 || len >= MAX_ENCODING_BUFFER_SIZE) {
		return nullptr;
	}

	WideCharToMultiByte(CP_UTF8, 0, wbuffer, -1, buffer, len, nullptr, nullptr);

	return buffer;
}

std::string WcharToChar(const std::wstring& wstr) {
	const wchar_t* wccstr = wstr.c_str();
	int len = WideCharToMultiByte(CP_UTF8, 0, wccstr, -1, nullptr, 0, nullptr, nullptr);
	if (len == 0) {
		std::cerr << "WideCharToMultiByte failed" << std::endl;
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

std::vector<uint8_t> LoadFile(const std::string& filename) {
	std::ifstream file(filename, std::ios::binary | std::ios::ate);
	if (!file.is_open()) {
		throw std::runtime_error("无法打开文件: " + filename);
	}

	size_t size = file.tellg();
	file.seekg(0, std::ios::beg);

	std::vector<uint8_t> buffer(size);
	file.read(reinterpret_cast<char*>(buffer.data()), size);
	file.close();

	std::cout << "加载文件: " << filename << " (" << size << " 字节)" << std::endl;
	return buffer;
}

// 辅助函数：写入文件
void WriteFile(const std::string& filename, std::vector<uint8_t>& data) {
	std::ofstream file(filename, std::ios::binary);
	if (!file.is_open()) {
		throw std::runtime_error("无法写入文件: " + filename);
	}
	file.write(reinterpret_cast<const char*>(data.data()), data.size());
	file.close();
	std::cout << "写入文件: " << filename << " (" << data.size() << " 字节)" << std::endl;
}

std::string StringTrim(const std::string& str) {
	size_t start = str.find_first_not_of(" \t\n\r");
	if (start == std::string::npos) return "";
	size_t end = str.find_last_not_of(" \t\n\r");
	return str.substr(start, end - start + 1);
}
