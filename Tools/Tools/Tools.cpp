#include "./ExcelToFlatBuffer.h"
#include "./FlatBufferToExcel.h"
#include <chrono>
#include <sstream>
#include <windows.h>
#include "Utils.h"

enum EOpenType
{
    EOT_InPathCodeConvert = 1 << 0,
    EOT_OutPathCodeConvert = 1 << 1,
    EOT_SendCommand = 1 << 2,
};

void printUsage(const char* programName) {
    STDERR << "用法:" << STDEND;
    STDERR << "  " << programName << " convert <metadata.json> <schema.bfbs> <excel.xlsx> <output.bytes> [flag]" << STDEND;
    STDERR << "  " << programName << " reverse  <metadata.json> <schema.bfbs> <input.bytes>  <output.xlsx> [flag]" << STDEND;
    STDERR << "flag: 1=输入路径GBK转UTF8, 2=输出路径UTF8转GBK, 4=发送进度命令" << STDEND;
}

int main(int argc, char* argv[]) {
    STDOUT << "参数数量:" << argc << ",编码:" << GetConsoleCP() << STDEND;
    if (argc < 2) {
        printUsage(argv[0]);
        return -1;
    }
    std::string command = argv[1];
    if (command != "convert" && command != "reverse") {
        STDERR << "未知命令: " << command << STDEND;
        printUsage(argv[0]);
        return -1;
    }
    if (argc != 6 && argc != 7) {
        printUsage(argv[0]);
        return -1;
    }
    std::string metadataFile = argv[2];
    std::string bfbsFile = argv[3];
    std::string inputPath = argv[4];
    std::string outputPath = argv[5];
    std::uint8_t openFlag = 0;
    if (argc == 7)
        openFlag = static_cast<uint8_t>(std::stoul(argv[6]));
    if (openFlag & EOT_InPathCodeConvert) {
        metadataFile = GbkToUtf8(metadataFile);
        bfbsFile = GbkToUtf8(bfbsFile);
        inputPath = GbkToUtf8(inputPath);
        outputPath = GbkToUtf8(outputPath);
    }
    STDOUT << "metadataFile: " << metadataFile << STDEND;
    STDOUT << "bfbsFile: " << bfbsFile << STDEND;
    STDOUT << "inputPath: " << inputPath << STDEND;
    STDOUT << "outputPath: " << outputPath << STDEND;
    bool outpncc = openFlag & EOT_OutPathCodeConvert;
    bool sendcmd = openFlag & EOT_SendCommand;
    if (command == "convert") {
        ExcelToFlatBuffer converter;
        converter.SetSymbol(sendcmd, outpncc);
        auto success = converter.Convert(metadataFile, bfbsFile, inputPath, outputPath);
        if (!success) {
            STDERR << "转换失败: " << converter.GetLastError() << STDEND;
            return -2;
        }
    }
    else if (command == "reverse") {
        FlatBufferToExcel converter;
        converter.SetSymbol(sendcmd, outpncc);
        auto success = converter.Convert(metadataFile, bfbsFile, inputPath, outputPath);
        if (!success) {
            STDERR << "逆向转换失败: " << converter.GetLastError() << STDEND;
            return -2;
        }
    }
    STDOUT << (command == "convert" ? "转换" : "逆向转换") << "成功!" << STDEND;
    return 0;
}
