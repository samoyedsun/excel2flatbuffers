#include "./ExcelToFlatBuffer.h"
#include <chrono>
#include <winsock2.h>
#include <iphlpapi.h>

#pragma comment(lib, "iphlpapi.lib")

struct AdapterInfo {
    std::string mac;
    std::vector<std::string> ips;

    static std::string join(const std::vector<std::string>& elements, const std::string& delimiter) {
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
    
    std::string dump() const {
        return mac + "," + AdapterInfo::join(ips, ",");
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
    return AdapterInfo::join(vec, "|");
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

int main(int argc, char* argv[]) {
    if (argc != 5) {
        std::cerr << "用法: " << argv[0] << " <schema.bfbs> <excel.xlsx> <output.bin>" << std::endl;
        return 1;
    }
    std::string metadataFile = argv[1];
    std::string bfbsFile = argv[2];
    std::string excelFile = argv[3];
    std::string outputFile = argv[4];
    std::cout << "metadataFile: " << metadataFile << std::endl;
    std::cout << "bfbsFile: " << bfbsFile << std::endl;
    std::cout << "excelFile: " << excelFile << std::endl;
    std::cout << "outputFile: " << outputFile << std::endl;
    ExcelToFlatBuffer converter;
    converter.SetSymbol(GetCurrentTimeString(), GetLocalAddress());
    if (!converter.Convert(metadataFile, bfbsFile, excelFile, outputFile)) {
        std::cerr << "转换失败: " << converter.GetLastError() << std::endl;
        system("pause");
        return 1;
    }
    std::cout << "转换成功!" << std::endl;
    system("pause");
    return 0;
}
