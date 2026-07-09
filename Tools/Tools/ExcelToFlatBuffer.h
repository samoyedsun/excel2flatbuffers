#ifndef INCLUDE_EXCELTOFLATBUFFER_HPP_
#define INCLUDE_EXCELTOFLATBUFFER_HPP_

#define _CRT_SECURE_NO_WARNINGS

#include <iostream>
#include <fstream>
#include <map>
#include <string>
#include <vector>
#include <functional>
#include <OpenXLSX/OpenXLSX.hpp>
#include "flatbuffers/reflection.h"
#include "./nlohmann/json.hpp"
#include "./Utils.h"

class ExcelToFlatBuffer {
public:
    // 构造函数
    ExcelToFlatBuffer();

    void SetSymbol(bool sendcmd);

    // 执行转换
    bool Convert(const std::string& metadataPath, const std::string& bfbsPath,
        const std::string& excelPath, const std::string& outputPath);

    // 获取错误信息
    std::string GetLastError() const { return m_lastError; }

private:
    // 内部类型定义
    using InfoOffsetsType = std::vector<flatbuffers::Offset<flatbuffers::Table>>;

    // 核心处理方法
    bool LoadSchema(const std::string& bfbsPath);
    bool LoadMetadata(const std::string& metadataPath);
    void ParseField(flatbuffers::FlatBufferBuilder& builder, const reflection::Field* pField, const std::string& value);
    void ReadExcelSheet(OpenXLSX::XLWorksheet& ws,
        flatbuffers::FlatBufferBuilder& builder,
        InfoOffsetsType& infoOffsets,
        const reflection::Object* pObject,
        nlohmann::json& objMetadata);
    bool ParseExcel(const std::string& excelPath, const std::string& outputPath);

    // 成员变量
    std::string m_lastError;
    const reflection::Schema* m_pSchema = nullptr;
    std::vector<uint8_t> m_schemaData;
    nlohmann::json m_metadataRoot;
    std::string m_excelFileName;
    std::map<std::string, InfoOffsetsType> m_tblOffsets;
    std::vector<uint8_t> m_outputData;

    bool m_sendCommand = false;
};

#endif  // INCLUDE_EXCELTOFLATBUFFER_HPP_
