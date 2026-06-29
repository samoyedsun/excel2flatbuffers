#include "./ExcelToFlatBuffer.h"
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

int main(int argc, char* argv[]) {
    STDOUT << "参数数量:" << argc << ",编码:" << GetConsoleCP() << STDEND;
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
    if (openFlag & EOT_InPathCodeConvert) {
        metadataFile = GbkToUtf8(metadataFile);
        bfbsFile = GbkToUtf8(bfbsFile);
        excelFile = GbkToUtf8(excelFile);
        outputFile = GbkToUtf8(outputFile);
    }
    STDOUT << "metadataFile: " << metadataFile << STDEND;
    STDOUT << "bfbsFile: " << bfbsFile << STDEND;
    STDOUT << "excelFile: " << excelFile << STDEND;
    STDOUT << "outputFile: " << outputFile << STDEND;
    bool outpncc = openFlag & EOT_OutPathCodeConvert;
    bool sendcmd = openFlag & EOT_SendCommand;
    ExcelToFlatBuffer converter;
    converter.SetSymbol(sendcmd, outpncc);
    if (!converter.Convert(metadataFile, bfbsFile, excelFile, outputFile)) {
        STDERR << "转换失败: " << converter.GetLastError() << STDEND;
        return -2;
    }
    STDOUT << "转换成功!" << STDEND;
    return 0;
}
