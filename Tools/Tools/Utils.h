#ifndef INCLUDE_UTILS_HPP_
#define INCLUDE_UTILS_HPP_

#include <vector>
#include <functional>

char* UTF8ToGB2312(const char* utf8);
char* GB2312ToUTF8(const char* gb2312);
std::string WcharToChar(const std::wstring& wstr);
std::string MakeDesPath(const std::string& srcPath, const std::string& extension);
// 获取文件名（不含扩展名）
std::string GetFilenameWithoutExt(const std::string& path);
std::vector<uint8_t> LoadFile(const std::string& filename);
void WriteFile(const std::string& filename, std::vector<uint8_t>& data);

std::string StrTrim(const std::string& str);
std::string StrJoin(const std::vector<std::string>& elements, const std::string& delimiter);
void StrSplit(const std::string& str, const std::string& delimiters,
	std::function<void(const std::string&)> callback);

#endif  // INCLUDE_UTILS_HPP_
