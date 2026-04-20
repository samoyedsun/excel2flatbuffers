- 使用说明
    1. 设计 ConfigDirectory 文件夹下的Excel表格文件
    2. 设计 ProjectDirectory\config\schema 文件夹下的fbs配置结构文件
        - excel文件名字要与fbs文件名字对应
        - fbs中root_type指定的表格里列表类型的字段名对应Excel文件中的表格（Sheet）名
    3. 配置 ConfigDirectory\Tools 文件夹下metadata.json文件 将Excel表格列名与fbs文件中的字段名映射起来
    4. 进入 ProjectDirectory 文件夹 执行CompileSchema.bat生成bfbs序列化映射文件以及 配置读取文件
        - 同时将生成的bfbs映射文件拷贝到 ConfigDirectory\Tools 文件夹下供Tools.exe读取
        - 配置读取文件是在 ProjectDirectory\config\reader 文件夹下的 xxx_generated.h文件，可加入Test项目中使用
    5. 进入 ConfigDirectory 文件夹 执行ExcelToBytes.bat生成 bytes文件 到 Temp 目录
        - 同时将生成的bytes配置文件拷贝到 ProjectDirectory\Bin\Release\Test 文件夹下供Test.exe程序读取使用
		- 或者执行 .\ExcelToBytes.bat .\mydb_misc_tbl.xlsx 可以指定单个导出