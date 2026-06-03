#include "ExcelToFlatBuffer.h"
#include <sstream>

std::string toLower(const std::string& str) {
    std::string result;
    result.reserve(str.size());
    for (unsigned char c : str) {
        result += std::tolower(c);
    }
    return result;
}

ExcelToFlatBuffer::ExcelToFlatBuffer() {
    m_pSchema = nullptr;
}

void ExcelToFlatBuffer::SetSymbol(bool outpncc
    , const std::string& dataTime
    , const std::string& hostInfo
    , const std::string& macAddress) {
    m_outPathNeedCodeConversion = outpncc;
    m_dateTime = dataTime;
    m_hostInfo = hostInfo;
    m_macAddress = macAddress;
}

bool ExcelToFlatBuffer::Convert(
    const std::string& metadataPath,
    const std::string& bfbsPath,
    const std::string& excelPath,
    const std::string& outputPath) {
    // 提取文件名用于查找元数据
    m_excelFileName = GetFilenameWithoutExt(excelPath);
    // 加载 Schema
    if (!LoadSchema(bfbsPath)) {
        return false;
    }
    // 加载 Metadata
    if (!LoadMetadata(metadataPath)) {
        return false;
    }
    // 解析 Excel
    if (!ParseExcel(excelPath, outputPath)) {
        return false;
    }
    return true;
}

bool ExcelToFlatBuffer::LoadSchema(const std::string& bfbsPath) {
    if (!LoadFile(bfbsPath, m_schemaData)) {
        m_lastError = "无法打开文件: " + bfbsPath;
        return false;
    }
    flatbuffers::Verifier schemaVerifier(m_schemaData.data(), m_schemaData.size());
    if (!reflection::VerifySchemaBuffer(schemaVerifier)) {
        m_lastError = ".bfbs 文件无效: " + bfbsPath;
        return false;
    }
    STDOUT << "加载文件(" << m_schemaData.size() << " 字节): " << bfbsPath << STDEND;
    m_pSchema = reflection::GetSchema(m_schemaData.data());
    STDOUT << "=== Schema 信息 ===" << STDEND;
    STDOUT << "根表: " << m_pSchema->root_table()->name()->str() << STDEND;
    STDOUT << "对象数量: " << m_pSchema->objects()->size() << STDEND;
    return true;
}

bool ExcelToFlatBuffer::LoadMetadata(const std::string& metadataPath) {
    std::vector<uint8_t> buffer;
    if (!LoadFile(metadataPath, buffer)) {
        m_lastError = "无法打开文件: " + metadataPath;
        return false;
    }
    STDOUT << "加载文件(" << buffer.size() << " 字节): " << metadataPath << STDEND;
    std::string jsonStr(buffer.begin(), buffer.end());
    if (jsonStr.empty()) {
        m_lastError = "元数据文件为空: " + metadataPath;
        return false;
    }
    m_metadataRoot = nlohmann::json::parse(jsonStr);
    return true;
}

void ExcelToFlatBuffer::ParseField(flatbuffers::FlatBufferBuilder& builder,
    const reflection::Field* pField,
    const std::string& key,
    const std::string& value) {
    switch (pField->type()->base_type()) {
    case reflection::Bool:
    {
        bool boVal = (value == "true" || value == "1" || value == "True" || value == "TRUE");
        builder.AddElement<bool>(pField->offset(), boVal, pField->default_integer() != 0);
    }
    break;
    case reflection::Byte:
    {
        char* endptr = nullptr;
        long val = std::strtol(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val >= INT8_MIN && val <= INT8_MAX)
            builder.AddElement<int8_t>(pField->offset(), static_cast<int8_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to int8 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::UByte:
    {
        char* endptr = nullptr;
        unsigned long val = std::strtoul(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val <= UINT8_MAX)
            builder.AddElement<uint8_t>(pField->offset(), static_cast<uint8_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to uint8 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Short:
    {
        char* endptr = nullptr;
        long val = std::strtol(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val >= INT16_MIN && val <= INT16_MAX)
            builder.AddElement<int16_t>(pField->offset(), static_cast<int16_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to int16 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::UShort:
    {
        char* endptr = nullptr;
        unsigned long val = std::strtoul(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val <= UINT16_MAX)
            builder.AddElement<uint16_t>(pField->offset(), static_cast<uint16_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to uint16 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Int:
    {
        char* endptr = nullptr;
        long val = std::strtol(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val >= INT32_MIN && val <= INT32_MAX)
            builder.AddElement<int32_t>(pField->offset(), static_cast<int32_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to int32 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::UInt:
    {
        char* endptr = nullptr;
        unsigned long val = std::strtoul(value.c_str(), &endptr, 10);
        if (endptr != value.c_str() && val <= UINT32_MAX)
            builder.AddElement<uint32_t>(pField->offset(), static_cast<uint32_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to uint32 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Long:
    {
        char* endptr = nullptr;
        long long val = std::strtoll(value.c_str(), &endptr, 10);
        if (endptr != value.c_str())
            builder.AddElement<int64_t>(pField->offset(), static_cast<int64_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to int64 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::ULong:
    {
        char* endptr = nullptr;
        unsigned long long val = std::strtoull(value.c_str(), &endptr, 10);
        if (endptr != value.c_str())
            builder.AddElement<uint64_t>(pField->offset(), static_cast<uint64_t>(val), pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to uint64 for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Float:
    {
        char* endptr = nullptr;
        float val = std::strtof(value.c_str(), &endptr);
        if (endptr != value.c_str())
            builder.AddElement<float>(pField->offset(), val, pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to float for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Double:
    {
        char* endptr = nullptr;
        double val = std::strtod(value.c_str(), &endptr);
        if (endptr != value.c_str())
            builder.AddElement<double>(pField->offset(), val, pField->default_integer());
        else
            STDERR << "Failed to convert '" << value << "' to double for field " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::Array:
    {
        STDERR << "Error: array type is not supported yet: " << pField->name()->str() << STDEND;
    }
    break;
    case reflection::String:
    {
        auto strOffset = builder.CreateString(value);
        builder.AddOffset(pField->offset(), strOffset);
    }
    break;
    case reflection::Vector:
    {
        auto elementType = pField->type()->element();
        if (elementType == reflection::UByte) {
            std::vector<uint8_t> bytes(value.begin(), value.end());
            auto bytesOffset = builder.CreateVector(bytes);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::Int) {
            std::vector<int32_t> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                long val = std::strtol(token.c_str(), &endptr, 10);
                if (endptr != token.c_str() && val >= INT32_MIN && val <= INT32_MAX) {
                    tokens.emplace_back(static_cast<int32_t>(val));
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::UInt) {
            std::vector<uint32_t> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                unsigned long val = std::strtoul(token.c_str(), &endptr, 10);
                if (endptr != token.c_str() && val <= UINT32_MAX) {
                    tokens.emplace_back(static_cast<uint32_t>(val));
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::Long) {
            std::vector<int64_t> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                long long val = std::strtoll(token.c_str(), &endptr, 10);
                if (endptr != token.c_str()) {
                    tokens.emplace_back(static_cast<int64_t>(val));
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::ULong) {
            std::vector<uint64_t> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                unsigned long long val = std::strtoull(token.c_str(), &endptr, 10);
                if (endptr != token.c_str()) {
                    tokens.emplace_back(static_cast<uint64_t>(val));
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::Float) {
            std::vector<float> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                float val = std::strtof(token.c_str(), &endptr);
                if (endptr != token.c_str()) {
                    tokens.emplace_back(val);
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::Double) {
            std::vector<double> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                char* endptr = nullptr;
                double val = std::strtod(token.c_str(), &endptr);
                if (endptr != token.c_str()) {
                    tokens.emplace_back(val);
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else {
            STDERR << "Unsupported vector element type: " << elementType
                << " for field " << pField->name()->str() << STDEND;
        }
    }
    break;
    default:
        STDERR << "Unsupported field type: " << pField->type()->base_type()
            << " for field " << pField->name()->str() << STDEND;
    }
}

void ExcelToFlatBuffer::ReadExcelLine(size_t maxColumn, std::function<void(int32_t colIndex)> process) {
    for (int32_t colIndex = 1; colIndex <= maxColumn; ++colIndex) {
        process(colIndex);
    }
}

void ExcelToFlatBuffer::ReadExcelSheet(OpenXLSX::XLWorksheet& ws,
    flatbuffers::FlatBufferBuilder& builder,
    InfoOffsetsType& infoOffsets,
    const reflection::Object* pObject,
    nlohmann::json& infoMetadataObj) {
    size_t maxRow = ws.rowCount();
    size_t maxColumn = ws.columnCount();
    std::vector<std::string> keys;

    // 读取第一行（表头）
    int32_t rowIndex = 1;
    ReadExcelLine(maxColumn, [rowIndex, &ws, &keys](int32_t colIndex) {
        auto cell = ws.cell(rowIndex, colIndex);
        keys.emplace_back(cell.getString());
        });

    // 读取数据行
    for (rowIndex += 1; rowIndex <= maxRow; ++rowIndex) {
        auto tableStart = builder.StartTable();

        ReadExcelLine(maxColumn, [this, rowIndex, &ws, &keys, &builder, pObject, &infoMetadataObj](int32_t colIndex) {
            auto cell = ws.cell(rowIndex, colIndex);
            auto& key = keys[colIndex - 1];
            auto& val = cell.value();
            if (val.type() == OpenXLSX::XLValueType::Empty) {
                // 未配的不需要导出 因为字段都是可选的
                return;
            }
            if (!infoMetadataObj.contains(key)) {
                STDERR << "未找到对应字段的元数据：" << colIndex << "-" << rowIndex << "-" << ws.name() << ":" << key << STDEND;
                return;
            }
            auto jsonValue = infoMetadataObj[key];
            if (jsonValue.is_null()) {
                // 元数据配空代表不需要导出
                return;
            }
            auto pField = pObject->fields()->LookupByKey(jsonValue);
            if (!pField) {
                STDERR << "反射字段未定义：" << colIndex << "-" << rowIndex << "-" << ws.name() << ":" << key << "-" << jsonValue << STDEND;
                return;
            }
            std::string value = Utf8ToGbk(cell.getString());
            ParseField(builder, pField, key, StrTrim(value));
            });

        auto infoOffset = builder.EndTable(tableStart);
        infoOffsets.push_back(infoOffset);
    }
}

bool ExcelToFlatBuffer::ParseExcel(const std::string& excelPath, const std::string& outputPath) {
    try {
        OpenXLSX::XLDocument doc(excelPath);;
        STDOUT << "=== Excel 信息 ===" << STDEND;
        std::map<std::string, OpenXLSX::XLWorksheet> sheets;
        auto workbook = doc.workbook();
        int sheetCount = workbook.sheetCount();
        STDOUT << "总共有 " << sheetCount << " 个工作表" << STDEND;
        std::vector<std::string> sheetNames = workbook.sheetNames();
        for (const auto& sheeName : sheetNames) {
            auto ws = workbook.worksheet(sheeName);
            sheets.emplace(toLower(sheeName), ws);
        }

        //============================================================
        flatbuffers::FlatBufferBuilder builder(1024);
        m_tblOffsets.clear();

        auto pRootTable = m_pSchema->root_table();
        if (!pRootTable) {
            m_lastError = "未找到根表定义";
            return false;
        }

        // 检查元数据中是否有当前 Excel 文件的配置
        if (!m_metadataRoot.contains(m_excelFileName)) {
            m_lastError = "未找到对应的元数据: " + m_excelFileName;
            return false;
        }

        auto tblMetadata = m_metadataRoot[m_excelFileName];

        // 遍历根表的所有字段
        for (auto pField : *pRootTable->fields()) {
            if (!pField) continue;

            if (pField->type()->base_type() == reflection::Vector) {
                auto typeElement = pField->type()->element();
                std::stringstream os;
                os << "sheet:" << pField->name()->str();
                os << "\ttypeElement:" << typeElement;

                if (typeElement == reflection::Obj) {
                    auto typeIndex = pField->type()->index();
                    os << " typeIndex:" << typeIndex;

                    if (typeIndex >= 0) {
                        auto pObject = m_pSchema->objects()->Get(typeIndex);
                        os << " typeName:" << pObject->name()->str();

                        // 检查工作表是否存在
                        if (sheets.find(pField->name()->str()) == sheets.end()) {
                            STDERR << "错误: 未找到对应的数据表 "
                                << m_excelFileName << ":" << pField->name()->str() << STDEND;
                            continue;
                        }

                        // 检查元数据是否存在
                        if (!tblMetadata.contains(pObject->name()->str())) {
                            STDERR << "错误: 未找到对应表字段的元数据 "
                                << m_excelFileName << ":" << pObject->name()->str() << STDEND;
                            continue;
                        }

                        auto infoMetadata = tblMetadata[pObject->name()->str()];
                        InfoOffsetsType infoOffsets;
                        OpenXLSX::XLWorksheet ws = sheets[pField->name()->str()];
                        size_t maxRow = ws.rowCount();
                        size_t maxColumn = ws.columnCount();
                        STDOUT << "sheet:" << pField->name()->str() << "\t行数：" << maxRow << "\t列数：" << maxColumn << STDEND;
                        ReadExcelSheet(ws, builder, infoOffsets, pObject, infoMetadata);
                        m_tblOffsets.emplace(pField->name()->str(), infoOffsets);
                    }
                }
                STDOUT << os.str() << STDEND;
            }
        }

        // 构建最终输出
        if (!m_tblOffsets.empty()) {
            auto tableStart = builder.StartTable();

            for (auto pField : *pRootTable->fields()) {
                if (!pField)
                    continue;
                if (pField->type()->base_type() == reflection::Vector) {
                    auto elementType = pField->type()->element();
                    if (elementType == reflection::Obj && pField->type()->index() >= 0) {
                        auto it = m_tblOffsets.find(pField->name()->str());
                        if (it != m_tblOffsets.end()) {
                            auto infosVector = builder.CreateVector(it->second);
                            builder.AddOffset(pField->offset(), infosVector);
                        }
                    }
                }
                else {
                    if (pField->type()->base_type() == reflection::String) {
                        if (pField->name()->str() == "__date_time") {
                            auto strOffset = builder.CreateString(m_dateTime);
                            builder.AddOffset(pField->offset(), strOffset);
                        }
                        else if (pField->name()->str() == "__host_info") {
                            auto strOffset = builder.CreateString(m_hostInfo);
                            builder.AddOffset(pField->offset(), strOffset);
                        }
                        else if (pField->name()->str() == "__mac_address") {
                            auto strOffset = builder.CreateString(m_macAddress);
                            builder.AddOffset(pField->offset(), strOffset);
                        }
                    }
                }
            }

            auto tblOffset = builder.EndTable(tableStart);
            builder.Finish<reflection::Object>(tblOffset);

            m_outputData.assign(builder.GetBufferPointer(),
                builder.GetBufferPointer() + builder.GetSize());

            std::string validFilePath = outputPath;
            if (m_outPathNeedCodeConversion)
                validFilePath = Utf8ToGbk(outputPath);
            if (!WriteFile(validFilePath, m_outputData)) {
                m_lastError = "无法写入文件: " + outputPath;
                return false;
            }
            STDOUT << "写入文件(" << m_outputData.size() << " 字节): " << outputPath << STDEND;
            return true;
        }

        m_lastError = "没有找到任何有效的数据表";
        return false;

    }
    catch (const std::exception& e) {
        m_lastError = "构建输出失败: " + std::string(e.what());
        return false;
    }
}
