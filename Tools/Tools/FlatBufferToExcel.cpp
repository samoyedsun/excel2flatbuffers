#include "FlatBufferToExcel.h"
#include <sstream>

FlatBufferToExcel::FlatBufferToExcel() {
    m_pSchema = nullptr;
}

void FlatBufferToExcel::SetSymbol(bool sendcmd) {
    m_sendCommand = sendcmd;
}

bool FlatBufferToExcel::Convert(const std::string& metadataPath,
                                const std::string& bfbsPath,
                                const std::string& bytesPath,
                                const std::string& outputPath) {
    if (m_sendCommand) {
        auto processId = GetProcessId();
        STDCMD << "@processid(" << processId << ")" << STDEND;
    }
    try {
        m_excelFileName = GetFilenameWithoutExt(bytesPath);
        if (!LoadSchema(bfbsPath)) {
            return false;
        }
        if (!LoadMetadata(metadataPath)) {
            return false;
        }
        if (!ParseFlatBuffers(bytesPath, outputPath)) {
            return false;
        }
    }
    catch (const std::exception& e) {
        m_lastError = std::string(e.what());
        return false;
    }
    return true;
}

bool FlatBufferToExcel::LoadSchema(const std::string& bfbsPath) {
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

bool FlatBufferToExcel::LoadMetadata(const std::string& metadataPath) {
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

std::string FlatBufferToExcel::ReadFieldValue(const flatbuffers::Table* pTable, const reflection::Field* pField) {
    auto baseType = pField->type()->base_type();
    auto offset = pField->offset();

    switch (baseType) {
    case reflection::Bool:
        return std::to_string(pTable->GetField<uint8_t>(offset, 0) != 0);
    case reflection::Byte:
        return std::to_string(pTable->GetField<int8_t>(offset, 0));
    case reflection::UByte:
        return std::to_string(pTable->GetField<uint8_t>(offset, 0));
    case reflection::Short:
        return std::to_string(pTable->GetField<int16_t>(offset, 0));
    case reflection::UShort:
        return std::to_string(pTable->GetField<uint16_t>(offset, 0));
    case reflection::Int:
        return std::to_string(pTable->GetField<int32_t>(offset, 0));
    case reflection::UInt:
        return std::to_string(pTable->GetField<uint32_t>(offset, 0));
    case reflection::Long:
        return std::to_string(pTable->GetField<int64_t>(offset, 0));
    case reflection::ULong:
        return std::to_string(pTable->GetField<uint64_t>(offset, 0));
    case reflection::Float:
        return std::to_string(pTable->GetField<float>(offset, 0.0f));
    case reflection::Double:
        return std::to_string(pTable->GetField<double>(offset, 0.0));
    case reflection::String: {
        auto pStr = pTable->GetPointer<const flatbuffers::String*>(offset);
        return pStr ? pStr->str() : "";
    }
    case reflection::Vector: {
        // 使用具体类型的 Vector
        auto elementType = pField->type()->element();
        switch (elementType) {
        case reflection::UByte: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<uint8_t>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::Int: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<int32_t>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::UInt: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<uint32_t>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::Long: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<int64_t>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::ULong: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<uint64_t>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::Float: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<float>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        case reflection::Double: {
            auto pVector = pTable->GetPointer<const flatbuffers::Vector<double>*>(offset);
            return pVector ? ReadVectorValueT(pVector) : "";
        }
        default:
            STDERR << "Unsupported vector element type: " << elementType << STDEND;
            return "";
        }
    }
    case reflection::Array:
        return "";
    default:
        STDERR << "Unsupported field type: " << baseType
            << " for field " << pField->name()->str() << STDEND;
    }
    return "";
}

bool FlatBufferToExcel::ParseFlatBuffers(const std::string& bytesPath, const std::string& outputPath) {
    std::vector<uint8_t> bytesData;
    if (!LoadFile(bytesPath, bytesData)) {
        m_lastError = "无法打开文件: " + bytesPath;
        return false;
    }
    STDOUT << "加载文件(" << bytesData.size() << " 字节): " << bytesPath << STDEND;

    // Get root table
    auto pRootTable = m_pSchema->root_table();
    auto pRoot = flatbuffers::GetAnyRoot(bytesData.data());
    auto pTable = reinterpret_cast<const flatbuffers::Table*>(pRoot);

    // Create Excel document
    OpenXLSX::XLDocument doc;
    doc.create(outputPath, OpenXLSX::XLForceOverwrite);
    auto workbook = doc.workbook();

    bool hasCreatedSheet = false;

    // Iterate through root table fields (Vector<Obj> fields)
    for (auto pField : *pRootTable->fields()) {
        if (!pField) continue;

        if (pField->type()->base_type() == reflection::Vector) {
            auto elementType = pField->type()->element();

            if (elementType == reflection::Obj) {
                auto typeIndex = pField->type()->index();
                if (typeIndex < 0) continue;

                auto pObject = m_pSchema->objects()->Get(typeIndex);
                auto sheetName = pField->name()->str();

                // Get metadata for this object first
                std::string objName = pObject->name()->str();
                if (!m_metadataRoot.contains(m_excelFileName) ||
                    !m_metadataRoot[m_excelFileName].contains(objName)) {
                    STDERR << "未找到对应表字段的元数据: " << objName << STDEND;
                    continue;
                }

                auto objMetadata = m_metadataRoot[m_excelFileName][objName];

                // Get the vector (Vector of Offset<Table>)
                auto pVector = pTable->GetPointer<const flatbuffers::Vector<flatbuffers::Offset<flatbuffers::Table>>*>(pField->offset());
                if (!pVector || pVector->size() == 0) {
                    STDOUT << "空向量: " << sheetName << STDEND;
                    continue;
                }

                // Create worksheet
                workbook.addWorksheet(sheetName);
                auto ws = workbook.worksheet(sheetName);
                hasCreatedSheet = true;

                // Build field order from metadata
                std::vector<const reflection::Field*> fieldOrder;
                int32_t colIndex = 1;

                // First pass: write header row
                for (auto childField : *pObject->fields()) {
                    if (!childField) continue;
                    auto fbName = childField->name()->str();
                    for (auto& [excelName, fbFieldName] : objMetadata.items()) {
                        if (fbFieldName == fbName && !fbFieldName.is_null()) {
                            ws.cell(1, colIndex).value() = excelName;
                            fieldOrder.push_back(childField);
                            colIndex++;
                            break;
                        }
                    }
                }

                // Write data rows
                size_t maxRow = pVector->size();
                for (size_t rowIndex = 0; rowIndex < maxRow; ++rowIndex) {
                    int32_t excelRow = static_cast<int32_t>(rowIndex) + 2;
                    auto pChildTable = pVector->Get(static_cast<flatbuffers::uoffset_t>(rowIndex));

                    int32_t writeCol = 1;
                    for (auto pChildField : fieldOrder) {
                        if (!pChildField) continue;
                        std::string value = ReadFieldValue(pChildTable, pChildField);
                        ws.cell(excelRow, writeCol).value() = GbkToUtf8(value);
                        writeCol++;
                    }
                    if (m_sendCommand) {
                        STDCMD << "@progress(" << rowIndex << "," << maxRow << ")" << STDEND;
                    }
                }

                STDOUT << "创建工作表: " << sheetName
                       << " (" << pVector->size() << " 行)" << STDEND;
            }
        }
    }

    // Delete default Sheet1 after creating all sheets
    if (hasCreatedSheet) {
        workbook.deleteSheet("Sheet1");
    }

    doc.save();
    doc.close();
    STDOUT << "写入文件: " << outputPath << STDEND;
    return true;
}