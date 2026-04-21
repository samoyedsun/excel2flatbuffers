// Test.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "../config/reader/mydb_misc_tbl_generated.h"
#include <iostream>
#include <fstream>
#include <vector>

std::vector<uint8_t> LoadFile(const std::string& filename) {
	std::ifstream file(filename, std::ios::binary | std::ios::ate);
	if (!file.is_open()) {
		throw std::runtime_error("无法打开文件: " + filename);
	}

	size_t size = file.tellg();
	file.seekg(0, std::ios::beg);

	std::vector<uint8_t> buffer(size);
	file.read(reinterpret_cast<char*>(buffer.data()), size);
	file.close();

	std::cout << "加载文件: " << filename << " (" << size << " 字节)" << std::endl;
	return buffer;
}

int main()
{
	auto buffer = LoadFile("./mydb_misc_tbl.bytes");

	// 2. 验证 FlatBuffers 数据完整性（可选但推荐）
	flatbuffers::Verifier verifier(buffer.data(), buffer.size());
	if (!VerifyMydbMiscTblBuffer(verifier)) {
		std::cerr << "FlatBuffers 数据验证失败" << std::endl;
		return 1;
	}

	auto pTbl = GetMydbMiscTbl(buffer.data());
	std::cout << "misc_infos条数：" << pTbl->misc_infos()->size() << std::endl;
	for (size_t i = 0; i < pTbl->misc_infos()->size(); ++i) {
		auto pInfo = pTbl->misc_infos()->Get(i);
		if (pInfo) {
			std::cout << "id:" << pInfo->id()
				<< ",type:" << pInfo->type()
				<< ",num:" << pInfo->num()
				<< ",size:" << pInfo->num_list()->size()
				<< ",str:" << pInfo->str()->str()
				<< std::endl;
			if (pInfo->num_list()) {
				for (const auto num : *(pInfo->num_list())) {
					std::cout << "num:" << num << std::endl;
				}
			}
		}
		else {
			std::cout << "什么情况：" << i << std::endl;
		}
	}
	std::cout << "type1_datas条数：" << pTbl->type1_datas()->size() << std::endl;
	for (size_t i = 0; i < pTbl->type1_datas()->size(); ++i) {
		auto pInfo = pTbl->type1_datas()->Get(i);
		if (pInfo) {
			std::cout << "id:" << pInfo->id()
				<< ",field:" << pInfo->field()->str()
				<< std::endl;
		}
		else {
			std::cout << "什么情况：" << i << std::endl;
		}
	}
	std::cout << "type2_datas条数：" << pTbl->type2_datas()->size() << std::endl;
	for (size_t i = 0; i < pTbl->type2_datas()->size(); ++i) {
		auto pInfo = pTbl->type2_datas()->Get(i);
		if (pInfo) {
			std::cout << "id:" << pInfo->id()
				<< ",field:" << pInfo->field()
				<< std::endl;
		}
		else {
			std::cout << "什么情况：" << i << std::endl;
		}
	}
	std::cout << "type3_datas条数：" << pTbl->type3_datas()->size() << std::endl;
	for (size_t i = 0; i < pTbl->type3_datas()->size(); ++i) {
		auto pInfo = pTbl->type3_datas()->Get(i);
		if (pInfo) {
			std::cout << "id:" << pInfo->id()
				<< ",field1:" << pInfo->field1()
				<< ",field2:" << pInfo->field2()
				<< std::endl;
		}
		else {
			std::cout << "什么情况：" << i << std::endl;
		}
	}
	std::cout << "type4_datas条数：" << pTbl->type4_datas()->size() << std::endl;
	for (size_t i = 0; i < pTbl->type4_datas()->size(); ++i) {
		auto pInfo = pTbl->type4_datas()->Get(i);
		if (pInfo) {
			std::cout << "id:" << pInfo->id()
				<< ",field1:" << pInfo->field1()
				<< ",field2:" << pInfo->field2()
				<< ",field3:" << pInfo->field3()
				<< ",field4:" << pInfo->field4()
				<< std::endl;
		}
		else {
			std::cout << "什么情况：" << i << std::endl;
		}
	}
	std::cout << "导表时间：" << pTbl->__date_time()->str() << std::endl;
	std::cout << "导表主机：" << pTbl->__host_info()->str() << std::endl;
	std::cout << "导表地址：" << pTbl->__mac_address()->str() << std::endl;
	system("pause");
}

// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件
