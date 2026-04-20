#define _CRT_SECURE_NO_WARNINGS

#include <iostream>
#include <xlnt/xlnt.hpp>

#include "flatbuffers/flatbuffers.h"
#include "flatbuffers/reflection.h"
#include "flatbuffers/reflection_generated.h"

#include "./nlohmann/json.hpp"
#include "./Utils.hpp"

void ParseField(flatbuffers::FlatBufferBuilder& builder,
	const reflection::Field* pField,
	std::string& key, const std::string& value) {
	/*
	std::cout << pField->name()->str()
		<< "-" << pField->id()
		<< "-" << pField->offset()
		<< "-" << pField->type()->base_type()
		<< std::endl;
	*/
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
		if (std::sscanf(value.c_str(), "%u", &outVal) == 1)
			builder.AddElement<uint32_t>(pField->offset(), outVal, pField->default_integer());
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
	}
	break;
	default:
		std::cerr << "Unsupported field type: " << pField->type()->base_type() << " for field " << pField->name()->str() << std::endl;
	}
}

void ReadExcelLine(size_t maxColumn, std::function<void(int32_t colIndex)> process) {
	for (int32_t colIndex = 1; colIndex <= maxColumn; ++colIndex) {
		process(colIndex);
	}
}

using InfoOffsetsType = std::vector<flatbuffers::Offset<flatbuffers::Table>>;

void ReadExcelSheet(xlnt::worksheet& ws,
	flatbuffers::FlatBufferBuilder& builder,
	InfoOffsetsType& infoOffsets,
	const reflection::Object* pObject,
	nlohmann::json& infoMetadataObj) {
	auto dim = ws.calculate_dimension();
	size_t maxRow = dim.height();
	size_t maxColumn = dim.width();
	std::cout << "sheet:" << ws.title() << "\t行数：" << maxRow << "\t列数：" << maxColumn << std::endl;

	std::vector<std::string> keys;
	// 读取第一行
	int32_t rowIndex = 1;
	ReadExcelLine(maxColumn, [rowIndex, &ws, &keys](int32_t colIndex) {
		xlnt::cell cell = ws.cell(colIndex, rowIndex);
		keys.emplace_back(cell.to_string());
		});
	// 读取其他行
	for (rowIndex += 1; rowIndex <= maxRow; ++rowIndex) {
		auto tableStart = builder.StartTable();
		ReadExcelLine(maxColumn, [rowIndex, &ws, &keys, &builder, pObject, &infoMetadataObj](int32_t colIndex) {
			xlnt::cell cell = ws.cell(colIndex, rowIndex);
			auto& key = keys[colIndex - 1];
			if (!cell.has_value())
				return;
			const auto& value = cell.to_string();
			if (!infoMetadataObj.contains(key)) {
				std::cerr << "错误: 未找到对应的元数据" << ws.title()
					<< ":" << key << std::endl;
				return;
			}
			auto jsonValue = infoMetadataObj.at(key);

			auto pField = pObject->fields()->LookupByKey(jsonValue);
			if (!pField) {
				std::cerr << "错误: 找不到 MydbMiscInfo " << key << " 定义" << std::endl;
				return;
			}
			ParseField(builder, pField, key, value);
			});
		auto infoOffset = builder.EndTable(tableStart);
		infoOffsets.push_back(infoOffset);
	}
}

int main111(int argc, char* argv[]) {
	if (argc != 4) {
		std::cerr << "参数错误:" << argc << std::endl;
		return 1;
	}
	const char* bfbsFile = argv[1];
	const char* inputFile = argv[2];
	const char* outputFile = argv[3];
	std::cout << "bfbsFile:" << bfbsFile << std::endl;
	std::cout << "inputFile:" << inputFile << std::endl;
	std::cout << "outputFile:" << outputFile << std::endl;

	auto schemaData = LoadFile(bfbsFile);
	// 验证 Schema
	flatbuffers::Verifier schemaVerifier(schemaData.data(), schemaData.size());
	if (!reflection::VerifySchemaBuffer(schemaVerifier)) {
		std::cerr << "错误: .bfbs 文件无效!" << std::endl;
		return 2;
	}
	const reflection::Schema* pSchema = reflection::GetSchema(schemaData.data());
	std::cout << "=== Schema 信息 ===" << std::endl;
	std::cout << "根表: " << pSchema->root_table()->name()->str() << std::endl;
	std::cout << "对象数量: " << pSchema->objects()->size() << std::endl;

	xlnt::workbook wb;
	wb.load(GB2312ToUTF8(inputFile));
	std::cout << "=== Excel 信息 ===" << std::endl;
	std::map<std::string, xlnt::worksheet> sheets;
	int sheetCount = wb.sheet_count();
	std::cout << "总共有 " << sheetCount << " 个工作表" << std::endl;
	for (int i = 0; i < sheetCount; ++i) {
		// 通过索引获取工作表 (索引从 0 开始)
		xlnt::worksheet ws = wb.sheet_by_index(i);
		std::string sheetName = ws.title();
		std::cout << i << "-sheetName:" << sheetName << std::endl;
		sheets.emplace(sheetName, ws);
	}
	//=========================================
	std::string fileName = GetFilenameWithoutExt(inputFile);

	auto jsonBuffer = LoadFile("./metadata.json");
	std::string jsonStr(jsonBuffer.begin(), jsonBuffer.end());
	if (jsonStr.empty()) {
		std::cerr << "错误: 读取元数据文件失败" << fileName << std::endl;
		return 3;
	}
	nlohmann::json jsonRoot = nlohmann::json::parse(jsonStr);
	if (!jsonRoot.contains(fileName)) {
		std::cerr << "错误: 未找到对应的元数据" << fileName << std::endl;
		return 3;
	}
	auto tblMetadata = jsonRoot[fileName];
	//=========================================
	flatbuffers::FlatBufferBuilder builder(1024);
	std::map<std::string, InfoOffsetsType> tblOffsets;
	auto pRootTable = pSchema->root_table();
	if (!pRootTable) {
		std::cerr << "错误: 未找到元数据" << fileName << std::endl;
		return 3;
	}
	for (auto pField : *pRootTable->fields()) {
		if (pField) {
			if (pField->type()->base_type() == reflection::Vector) {
				auto typeElement = pField->type()->element();
				std::stringstream os;
				os << "sheet:" << pField->name()->str();
				os << "\ttypeElement:" << typeElement;
				if (typeElement == reflection::Obj) {
					auto typeIndex = pField->type()->index();
					os << " typeIndex:" << typeIndex;
					if (typeIndex >= 0) {
						auto pObject = pSchema->objects()->Get(typeIndex);
						os << " typeName:" << pObject->name()->str();
						if (sheets.find(pField->name()->str()) == sheets.end()) {
							std::cerr << "错误: 未找到对应的数据表" << fileName
								<< ":" << pField->name()->str() << std::endl;
							continue;
						}
						if (!tblMetadata.contains(pObject->name()->str())) {
							std::cerr << "错误: 未找到对应的元数据" << fileName
								<< ":" << pObject->name()->str() << std::endl;
							continue;
						}
						auto infoMetadata = tblMetadata[pObject->name()->str()];;
						InfoOffsetsType infoOffsets;
						xlnt::worksheet ws = wb.sheet_by_title(pField->name()->str());
						ReadExcelSheet(ws, builder, infoOffsets, pObject, infoMetadata);
						tblOffsets.emplace(pField->name()->str(), infoOffsets);
					}
				}
				std::cout << os.str() << std::endl;
			}
		}
	}
	if (!tblOffsets.empty()) {
		auto tableStart = builder.StartTable();
		for (auto pField : *pRootTable->fields()) {
			if (pField) {
				if (pField->type()->base_type() == reflection::Vector) {
					auto elementType = pField->type()->element();
					if (elementType == reflection::Obj) {
						if (pField->type()->index() >= 0) {
							auto& infoOffsets = tblOffsets[pField->name()->str()];
							auto infosVector = builder.CreateVector(infoOffsets);
							builder.AddOffset(pField->offset(), infosVector);
						}
					}
				}
			}
		}
		auto tblOffset = builder.EndTable(tableStart);
		builder.Finish<reflection::Object>(tblOffset);
		std::vector<uint8_t> data(builder.GetBufferPointer(), builder.GetBufferPointer() + builder.GetSize());
		std::cout << data.size() << std::endl;

		WriteFile(outputFile, data);
	}

	system("pause");
}

/*
int main(int argc, char* argv[]) {
	if (argc != 4) {
		std::cerr << "参数错误:" << argc << std::endl;
		return 1;
	}
	const char* bfbsFile = argv[1];
	const char* inputFile = argv[2];
	const char* outputFile = argv[3];
	std::cout << "bfbsFile:" << bfbsFile << std::endl;
	std::cout << "inputFile:" << inputFile << std::endl;
	std::cout << "outputFile:" << outputFile << std::endl;

	auto schemaData = LoadFile(bfbsFile);
	// 验证 Schema
	flatbuffers::Verifier schemaVerifier(schemaData.data(), schemaData.size());
	if (!reflection::VerifySchemaBuffer(schemaVerifier)) {
		std::cerr << "错误: .bfbs 文件无效!" << std::endl;
		return 1;
	}
	const reflection::Schema* pSchema = reflection::GetSchema(schemaData.data());
	std::cout << "=== Schema 信息 ===" << std::endl;
	std::cout << "根表: " << pSchema->root_table()->name()->str() << std::endl;
	std::cout << "对象数量: " << pSchema->objects()->size() << std::endl;

	xlnt::workbook wb;
	wb.load(inputFile);
	std::cout << "=== Excel 信息 ===" << std::endl;
	int sheetCount = wb.sheet_count();
	std::cout << "总共有 " << sheetCount << " 个工作表" << std::endl;
	for (int i = 0; i < sheetCount; ++i) {
		// 通过索引获取工作表 (索引从 0 开始)
		xlnt::worksheet ws = wb.sheet_by_index(i);
		std::string sheetName = ws.title();
		std::cout << i << "-sheetName:" << sheetName << std::endl;
	}

	auto ws = wb.active_sheet();
	auto dim = ws.calculate_dimension();
	size_t maxRow = dim.height();
	size_t maxColumn = dim.width();
	std::cout << "行数：" << maxRow << "\t列数：" << maxColumn << std::endl;

	// 动态构建 FlatBuffer
	flatbuffers::FlatBufferBuilder builder(1024);
	std::vector<flatbuffers::Offset<flatbuffers::Table>> infoOffsets;
	auto pObject = pSchema->objects()->LookupByKey("MydbMiscInfo");
	if (!pObject) {
		std::cerr << "错误: 找不到 MydbMiscInfo 定义" << std::endl;
		return 2;
	}
	auto pObjectTbl = pSchema->objects()->LookupByKey("MydbMiscTbl");
	if (!pObjectTbl) {
		std::cerr << "错误: 找不到 MydbMiscTbl 定义" << std::endl;
		return 3;
	}
	std::vector<std::string> keys;
	// 读取第一行
	int32_t rowIndex = 1;
	ReadExcelLine(maxColumn, [rowIndex, &ws, &keys](int32_t colIndex) {
		xlnt::cell cell = ws.cell(colIndex, rowIndex);
		keys.emplace_back(UTF8ToGB2312(cell.to_string().c_str()));
		});
	// 读取其他行
	for (rowIndex += 1; rowIndex <= maxRow; ++rowIndex) {
		auto tableStart = builder.StartTable();
		ReadExcelLine(maxColumn, [rowIndex, &ws, &keys, &builder, pObject](int32_t colIndex) {
			xlnt::cell cell = ws.cell(colIndex, rowIndex);
			auto& key = keys[colIndex - 1];
			if (!cell.has_value())
				return;
			const auto& value = cell.to_string();
			auto pField = pObject->fields()->LookupByKey(tableKeyMap[key]);
			if (!pField) {
				std::cerr << "错误: 找不到 MydbMiscInfo " << key << " 定义" << std::endl;
				return;
			}
			ParseField(builder, pField, key, value);
			});
		auto infoOffset = builder.EndTable(tableStart);
		infoOffsets.push_back(infoOffset);
	}
	auto infosVector = builder.CreateVector(infoOffsets);

	auto tableStart = builder.StartTable();
	auto pField = pObjectTbl->fields()->LookupByKey("misc_infos");
	if (!pField) {
		std::cerr << "错误: 找不到 Root " << "misc_infos" << " 定义" << std::endl;
		return 4;
	}
	builder.AddOffset(pField->offset(), infosVector);
	auto tblOffset = builder.EndTable(tableStart);

	builder.Finish<reflection::Object>(tblOffset);

	std::vector<uint8_t> data(builder.GetBufferPointer(), builder.GetBufferPointer() + builder.GetSize());
	std::cout << data.size() << std::endl;

	WriteFile(outputFile, data);
	system("pause");
}
*/