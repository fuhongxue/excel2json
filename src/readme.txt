conf.json文件需要和excel2json.exe放在同一个目录下

conf.json的配置说明

{
	"excel_dir":"指定excel源文件目录,支持相对路径和绝对路径",
	"json_dir":"指定json文件生成目录,支持相对路径和绝对路径"
}


excel目录源的excel文件说明:
	excel文件只支持.xlsx后缀格式
	excel文件命名标准:中文名称-英文名称.xlsx 如 用户等级表-UserLevel.xlsx
	其中英文名称部分将用于生成json文件名，即 用户等级表-UserLevel.xlsx => UserLevel.json

	注意excel文件只读取第一张sheet1，其它sheet会自动忽略
	excel文件里的内容,前几行为备注，前几行备注之后为英文字段名行【用于充当json的字段名称】
	之后各行为数据域，如果出现空行会自动过滤。
	注意每行的第一行第一列单元格不能为空，否则会忽略整行数据

	excel2json.exe会有检测，成功，失败，警告四种状态，失败为严重错误需要检查配置，该文件不会生成。
	
	警告状态工具会指出excel配置中可能出问题的行列，需要策划自行找excel确认。警告不影响文件生成(但数据不能保证正常,一般是策划配置有出入)


-----------------------------
conf.json and excel2json.exe need in the same directory

confi.json:
{
	"excel_dir":"set the input excel files source directory,your can set it with relative directory or absolute directory"
	"json_dir":"set the output files directory,your can set it with relative directory or absolute directory"
}

something about the tool:
	excel file only support .xlsx

	excel file only read sheet1,other sheets will be ignored
	the beginning row of excel content will be ignored if the content is utf8,until find the first english row(this row will be titled for json)
	notice:every beginning col of the row can't be blank or it will be ignored be program
	excel2json.exe has check,success,fail,warnning state.fail state is the worst error,the json file wouldn't created,you'd better to check the excel file config
	through the warnning state show the col and row about something wrong,it will create json file,but you'd better check the excel if the config is ok. 