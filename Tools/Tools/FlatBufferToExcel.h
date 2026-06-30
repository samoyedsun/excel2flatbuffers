#ifndef INCLUDE_FLATBUFFERTOEXCEL_HPP_
#define INCLUDE_FLATBUFFERTOEXCEL_HPP_

#define _CRT_SECURE_NO_WARNINGS

#include <iostream>
#include <fstream>
#include <map>
#include <string>
#include <vector>
#include <functional>
#include <sstream>
#include <OpenXLSX/OpenXLSX.hpp>
#include "flatbuffers/reflection.h"
#include "./nlohmann/json.hpp"
#include "./Utils.h"

class FlatBufferToExcel {
public:
    FlatBufferToExcel();

    void SetSymbol(bool sendcmd, bool outpncc);

    // bytes -> Excel 转换
    bool Convert(const std::string& metadataPath,
                 const std::string& bfbsPath,
                 const std::string& bytesPath,
                 const std::string& outputPath);

    std::string GetLastError() const { return m_lastError; }

private:
    bool LoadSchema(const std::string& bfbsPath);
    bool LoadMetadata(const std::string& metadataPath);
    bool ParseFlatBuffers(const std::string& bytesPath, const std::string& outputPath);

    std::string ReadFieldValue(const flatbuffers::Table* pTable, const reflection::Field* pField);

    // Template for reading vector values
    template<typename T>
    std::string ReadVectorValueT(const flatbuffers::Vector<T>* pVector) {
        std::stringstream ss;
        for (flatbuffers::uoffset_t i = 0; i < pVector->size(); ++i) {
            if (i > 0) ss << ",";
            ss << pVector->Get(i);
        }
        return ss.str();
    }

    std::string m_lastError;
    const reflection::Schema* m_pSchema = nullptr;
    std::vector<uint8_t> m_schemaData;
    nlohmann::json m_metadataRoot;
    std::string m_excelFileName;

    bool m_outPathNeedCodeConversion = false;
    bool m_sendCommand = false;
};

#endif
