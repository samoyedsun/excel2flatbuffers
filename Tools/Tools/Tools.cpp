#include "./ExcelToFlatBuffer.h"
#include <chrono>
#include <sstream>
#include <winsock2.h>
#include <iphlpapi.h>
#include "Utils.h"

#pragma comment(lib, "iphlpapi.lib")

struct AdapterInfo {
    std::string mac;
    std::vector<std::string> ips;
    
    std::string dump() const {
        return mac + "," + StrJoin(ips, ",");
    }
};

std::vector<AdapterInfo> GetAdaptersInfoList()
{
    std::vector<AdapterInfo> adapters;

    ULONG outBufLen = sizeof(IP_ADAPTER_INFO);
    PIP_ADAPTER_INFO pAdapterInfo = (IP_ADAPTER_INFO*)malloc(outBufLen);

    // 第一次调用获取缓冲区大小
    if (GetAdaptersInfo(pAdapterInfo, &outBufLen) == ERROR_BUFFER_OVERFLOW) {
        free(pAdapterInfo);
        pAdapterInfo = (IP_ADAPTER_INFO*)malloc(outBufLen);
    }

    if (GetAdaptersInfo(pAdapterInfo, &outBufLen) == NO_ERROR) {
        PIP_ADAPTER_INFO pAdapter = pAdapterInfo;
        while (pAdapter) {
            AdapterInfo info;

            // 获取MAC地址
            char macStr[18] = { 0 };
            sprintf_s(macStr, sizeof(macStr),
                "%02X:%02X:%02X:%02X:%02X:%02X",
                static_cast<unsigned char>(pAdapter->Address[0]),
                static_cast<unsigned char>(pAdapter->Address[1]),
                static_cast<unsigned char>(pAdapter->Address[2]),
                static_cast<unsigned char>(pAdapter->Address[3]),
                static_cast<unsigned char>(pAdapter->Address[4]),
                static_cast<unsigned char>(pAdapter->Address[5]));
            info.mac = macStr;

            // 获取IPv4地址（可选：过滤特定类型）
            PIP_ADDR_STRING pIpAddr = &pAdapter->IpAddressList;
            while (pIpAddr) {
                if (pIpAddr->IpAddress.String[0] != '\0') {
                    info.ips.emplace_back(pIpAddr->IpAddress.String);
                }
                pIpAddr = pIpAddr->Next;
            }

            adapters.emplace_back(info);
            pAdapter = pAdapter->Next;
        }
    }

    free(pAdapterInfo);
    return adapters;
}

std::string GetLocalAddress() {
    std::vector<std::string> vec;
    auto adapters = GetAdaptersInfoList();
    for (const auto& adapter : adapters) {
        vec.emplace_back(adapter.dump());
    }
    return StrJoin(vec, "|");
}

std::string GetHostInfo() {
    std::string hostName;
    do {
        DWORD size = 0;
        GetComputerNameW(nullptr, &size);
        std::vector<wchar_t> buffer(size);
        if (!GetComputerNameW(buffer.data(), &size))
            break;
        hostName = WcharToChar(buffer.data());
    } while (false);
    std::string userName;
    do {
        DWORD size = 0;
        GetUserNameW(nullptr, &size);
        std::vector<wchar_t> buffer(size);
        if (!GetUserNameW(buffer.data(), &size))
            break;
        userName = WcharToChar(buffer.data());
    } while (false);
    return StrJoin({ Utf8ToGbk(hostName), Utf8ToGbk(userName) }, "|");
}

std::string GetCurrentTimeString() {
    auto now = std::chrono::system_clock::now();
    auto nowTime = std::chrono::system_clock::to_time_t(now);
    std::tm localTm;
    localtime_s(&localTm, &nowTime);
    std::ostringstream oss;
    oss << std::put_time(&localTm, "%Y-%m-%d %H:%M:%S");
    return oss.str();
}

enum EOpenType
{
    EOT_InPathCodeConvert = 1 << 0,
    EOT_OutPathCodeConvert = 1 << 1,
    EOT_OpenSystemInfo = 1 << 2,
};

int main(int argc, char* argv[]) {
    if (argc < 5 || argc > 6) {
        STDERR << "用法: " << argv[0] << " <schema.bfbs> <excel.xlsx> <output.bin> <flag>" << STDEND;
        return -1;
    }
    std::string metadataFile;
    std::string bfbsFile;
    std::string excelFile;
    std::string outputFile;
    std::uint8_t openFlag = 0;
    if (argc >= 5) {
        metadataFile = argv[1];
        bfbsFile = argv[2];
        excelFile = argv[3];
        outputFile = argv[4];
    }
    if (argc == 6) {
        openFlag = std::stoul(argv[5]);
    }
    STDOUT << "metadataFile: " << metadataFile << STDEND;
    STDOUT << "bfbsFile: " << bfbsFile << STDEND;
    STDOUT << "excelFile: " << excelFile << STDEND;
    STDOUT << "outputFile: " << outputFile << STDEND;
    bool inpncc = openFlag & EOT_InPathCodeConvert;
    bool outpncc = openFlag & EOT_OutPathCodeConvert;
    std::string currentTime, hostInfo, localAddress;
    if (openFlag & EOT_OpenSystemInfo) {
        currentTime = GetCurrentTimeString();
        hostInfo = GetHostInfo();
        localAddress = GetLocalAddress();
    }
    ExcelToFlatBuffer converter;
    converter.SetSymbol(inpncc, outpncc, currentTime, hostInfo, localAddress);
    if (!converter.Convert(metadataFile, bfbsFile, excelFile, outputFile)) {
        STDERR << "转换失败: " << converter.GetLastError() << STDEND;
        return -2;
    }
    STDOUT << "转换成功!" << STDEND;
    return 0;
}
