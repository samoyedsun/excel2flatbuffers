#ifndef INCLUDE_UTILS_HPP_
#define INCLUDE_UTILS_HPP_

#include <vector>

char* UTF8ToGB2312(const char* utf8);
char* GB2312ToUTF8(const char* gb2312);
std::string WcharToChar(const std::wstring& wstr);
std::string MakeDesPath(const std::string& srcPath, const std::string& extension);
// 获取文件名（不含扩展名）
std::string GetFilenameWithoutExt(const std::string& path);
std::vector<uint8_t> LoadFile(const std::string& filename);
void WriteFile(const std::string& filename, std::vector<uint8_t>& data);

std::string StringTrim(const std::string& str);



#endif  // INCLUDE_UTILS_HPP_
