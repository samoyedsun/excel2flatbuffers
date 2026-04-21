#include "ExcelToFlatBuffer.h"

ExcelToFlatBuffer::ExcelToFlatBuffer() {
    m_pSchema = nullptr;
}

void ExcelToFlatBuffer::SetSymbol(const std::string& dataTime
    , const std::string& hostInfo
    , const std::string& macAddress) {
    m_dateTime = dataTime;
    m_hostInfo = hostInfo;
    m_macAddress = macAddress;
    std::cout << "时间：" << m_dateTime << std::endl;
    std::cout << "主机：" << m_hostInfo << std::endl;
    std::cout << "地址：" << m_macAddress << std::endl;
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
    // 加载 Excel 文件
    if (!LoadExcel(excelPath)) {
        return false;
    }
    // 构建输出
    if (!BuildOutput(outputPath)) {
        return false;
    }
    return true;
}

bool ExcelToFlatBuffer::LoadSchema(const std::string& bfbsPath) {
    try {
        m_schemaData = LoadFile(bfbsPath);

        flatbuffers::Verifier schemaVerifier(m_schemaData.data(), m_schemaData.size());
        if (!reflection::VerifySchemaBuffer(schemaVerifier)) {
            m_lastError = ".bfbs 文件无效: " + bfbsPath;
            std::cerr << "错误: " << m_lastError << std::endl;
            return false;
        }

        m_pSchema = reflection::GetSchema(m_schemaData.data());

        std::cout << "=== Schema 信息 ===" << std::endl;
        std::cout << "根表: " << m_pSchema->root_table()->name()->str() << std::endl;
        std::cout << "对象数量: " << m_pSchema->objects()->size() << std::endl;

        return true;
    }
    catch (const std::exception& e) {
        m_lastError = "加载 Schema 失败: " + std::string(e.what());
        std::cerr << "错误: " << m_lastError << std::endl;
        return false;
    }
}

bool ExcelToFlatBuffer::LoadMetadata(const std::string& metadataPath) {
    try {
        auto jsonBuffer = LoadFile(metadataPath);
        std::string jsonStr(jsonBuffer.begin(), jsonBuffer.end());

        if (jsonStr.empty()) {
            m_lastError = "元数据文件为空: " + metadataPath;
            std::cerr << "错误: " << m_lastError << std::endl;
            return false;
        }

        m_metadataRoot = nlohmann::json::parse(jsonStr);

        return true;
    }
    catch (const std::exception& e) {
        m_lastError = "加载元数据失败: " + std::string(e.what());
        std::cerr << "错误: " << m_lastError << std::endl;
        return false;
    }
}

bool ExcelToFlatBuffer::LoadExcel(const std::string& excelPath) {
    try {
        // 转换路径编码（支持中文路径）
        std::string utf8Path = GB2312ToUTF8(excelPath.c_str());
        m_workbook.load(utf8Path);

        std::cout << "=== Excel 信息 ===" << std::endl;

        m_sheets.clear();
        int sheetCount = m_workbook.sheet_count();
        std::cout << "总共有 " << sheetCount << " 个工作表" << std::endl;

        for (int i = 0; i < sheetCount; ++i) {
            xlnt::worksheet ws = m_workbook.sheet_by_index(i);
            m_sheets.emplace(ws.title(), ws);
        }

        return true;
    }
    catch (const std::exception& e) {
        m_lastError = "加载 Excel 失败: " + std::string(e.what());
        std::cerr << "错误: " << m_lastError << std::endl;
        return false;
    }
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
        int8_t outVal;
        if (sscanf(value.c_str(), "%hhd", &outVal) == 1)
            builder.AddElement<int8_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to int8 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::UByte:
    {
        uint8_t outVal;
        if (sscanf(value.c_str(), "%hhu", &outVal) == 1)
            builder.AddElement<uint8_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to uint8 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Short:
    {
        int16_t outVal;
        if (sscanf(value.c_str(), "%hd", &outVal) == 1)
            builder.AddElement<int16_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to int16 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::UShort:
    {
        uint16_t outVal;
        if (sscanf(value.c_str(), "%hu", &outVal) == 1)
            builder.AddElement<uint16_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to uint16 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Int:
    {
        int32_t outVal;
        if (std::sscanf(value.c_str(), "%d", &outVal) == 1)
            builder.AddElement<int32_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to int32 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::UInt:
    {
        uint32_t outVal;
        if (std::sscanf(value.c_str(), "%u", &outVal) == 1) {
            builder.AddElement<uint32_t>(pField->offset(), outVal, pField->default_integer());
        }
        else
            std::cerr << "Failed to convert '" << value << "' to uint32 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Long:
    {
        int64_t outVal;
        if (std::sscanf(value.c_str(), "%lld", &outVal) == 1)
            builder.AddElement<int64_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to int64 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::ULong:
    {
        uint64_t outVal;
        if (std::sscanf(value.c_str(), "%llu", &outVal) == 1)
            builder.AddElement<uint64_t>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to uint64 for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Float:
    {
        float outVal;
        if (std::sscanf(value.c_str(), "%f", &outVal) == 1)
            builder.AddElement<float>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to float for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Double:
    {
        double outVal;
        if (std::sscanf(value.c_str(), "%lf", &outVal) == 1)
            builder.AddElement<double>(pField->offset(), outVal, pField->default_integer());
        else
            std::cerr << "Failed to convert '" << value << "' to double for field " << pField->name()->str() << std::endl;
    }
    break;
    case reflection::Array:
    {
        std::cerr << "Error: array type is not supported yet: " << pField->name()->str() << std::endl;
    }
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
                int32_t outVal;
                if (std::sscanf(token.c_str(), "%d", &outVal) == 1) {
                    tokens.emplace_back(outVal);
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
        else if (elementType == reflection::UInt) {
            std::vector<uint32_t> tokens;
            StrSplit(value, ",", [&tokens](const std::string& token) {
                uint32_t outVal;
                if (std::sscanf(token.c_str(), "%u", &outVal) == 1) {
                    tokens.emplace_back(outVal);
                }
                });
            auto bytesOffset = builder.CreateVector(tokens);
            builder.AddOffset(pField->offset(), bytesOffset);
        }
    }
    break;
    default:
        std::cerr << "Unsupported field type: " << pField->type()->base_type()
            << " for field " << pField->name()->str() << std::endl;
    }
}

void ExcelToFlatBuffer::ReadExcelLine(size_t maxColumn, std::function<void(int32_t colIndex)> process) {
    for (int32_t colIndex = 1; colIndex <= maxColumn; ++colIndex) {
        process(colIndex);
    }
}

void ExcelToFlatBuffer::ReadExcelSheet(xlnt::worksheet& ws,
    flatbuffers::FlatBufferBuilder& builder,
    InfoOffsetsType& infoOffsets,
    const reflection::Object* pObject,
    nlohmann::json& infoMetadataObj) {
    auto dim = ws.calculate_dimension();
    size_t maxRow = dim.height();
    size_t maxColumn = dim.width();
    std::cout << "sheet:" << ws.title() << "\t行数：" << maxRow << "\t列数：" << maxColumn << std::endl;

    std::vector<std::string> keys;

    // 读取第一行（表头）
    int32_t rowIndex = 1;
    ReadExcelLine(maxColumn, [rowIndex, &ws, &keys](int32_t colIndex) {
        xlnt::cell cell = ws.cell(colIndex, rowIndex);
        keys.emplace_back(cell.to_string());
        });

    // 读取数据行
    for (rowIndex += 1; rowIndex <= maxRow; ++rowIndex) {
        auto tableStart = builder.StartTable();

        ReadExcelLine(maxColumn, [this, rowIndex, &ws, &keys, &builder, pObject, &infoMetadataObj](int32_t colIndex) {
            xlnt::cell cell = ws.cell(colIndex, rowIndex);
            auto& key = keys[colIndex - 1];
            if (!cell.has_value()) {
                return;
            }
            if (!infoMetadataObj.contains(key)) {
                //std::cerr << "错误: 未找到对应字段的元数据 " << ws.title()
                //    << ":" << std::string(UTF8ToGB2312(key.c_str())) << std::endl;
                // 元数据未配 说明不需要导出
                return;
            }
            auto jsonValue = infoMetadataObj.at(key);
            auto pField = pObject->fields()->LookupByKey(jsonValue);
            if (!pField) {
                std::cerr << "错误: 找不到字段 " << key << " 定义" << std::endl;
                return;
            }
            std::string value = UTF8ToGB2312(cell.to_string().c_str());
            ParseField(builder, pField, key, StrTrim(value));
            });

        auto infoOffset = builder.EndTable(tableStart);
        infoOffsets.push_back(infoOffset);
    }
}

bool ExcelToFlatBuffer::BuildOutput(const std::string& outputPath) {
    try {
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
            std::cerr << "错误: " << m_lastError << std::endl;
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
                        if (m_sheets.find(pField->name()->str()) == m_sheets.end()) {
                            std::cerr << "错误: 未找到对应的数据表 "
                                << m_excelFileName << ":" << pField->name()->str() << std::endl;
                            continue;
                        }

                        // 检查元数据是否存在
                        if (!tblMetadata.contains(pObject->name()->str())) {
                            std::cerr << "错误: 未找到对应表字段的元数据 "
                                << m_excelFileName << ":" << pObject->name()->str() << std::endl;
                            continue;
                        }

                        auto infoMetadata = tblMetadata[pObject->name()->str()];
                        InfoOffsetsType infoOffsets;
                        xlnt::worksheet ws = m_sheets[pField->name()->str()];

                        ReadExcelSheet(ws, builder, infoOffsets, pObject, infoMetadata);
                        m_tblOffsets.emplace(pField->name()->str(), infoOffsets);
                    }
                }
                std::cout << os.str() << std::endl;
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

            std::cout << "输出数据大小: " << m_outputData.size() << " 字节" << std::endl;
            WriteFile(outputPath, m_outputData);

            return true;
        }

        m_lastError = "没有找到任何有效的数据表";
        return false;

    }
    catch (const std::exception& e) {
        m_lastError = "构建输出失败: " + std::string(e.what());
        std::cerr << "错误: " << m_lastError << std::endl;
        return false;
    }
}
