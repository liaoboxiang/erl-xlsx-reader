%% @author box
%% @doc xlsx 文件读取


-module(ex_reader).
-include_lib("xmerl/include/xmerl.hrl").


%% ====================================================================
%% API functions
%% ====================================================================
-export([
		 read/1,
		 read/2
		]).

%% 读数据
-spec read(FileName::string()) -> SheetList::list() | {error, Reason::term()}.
%% SheetList::[SheetData|...]
%% SheetData::{SheetName, RowDataList}
%% RowData::[CellVal()::term()|...]
%% 例子：read("../xlsx/t1.xlsx") -> [].
read(FileName) ->
	%% 解压到内存
	case zip:zip_open(FileName, [memory]) of
		{error, Reason} -> {error, Reason};
		{ok, ZipHandle} ->
			%% 获得sharedStrings.xml中的字符串的映射关系
			ShareStringList = ex_share_strings:parse(ZipHandle),
			%% 读取每个sheet
			Result = ex_workbook:parse(ZipHandle, ShareStringList),
			zip:zip_close(ZipHandle),
			Result
	end.

%% SheetNameList::[SheetName|...]
%% SheetName使用unicode编码，如"我" == [25105]
%% 例子：read("../xlsx/t1.xlsx", ["sheet1"]) -> SheetList
%% read("../xlsx/t1.xlsx", [[25105]]) -> SheetList
read(FileName, SheetNameList) ->
	%% 解压到内存
	case zip:zip_open(FileName, [memory]) of
		{error, Reason} -> {error, Reason};
		{ok, ZipHandle} ->
			%% 获得sharedStrings.xml中的字符串的映射关系
			ShareStringList = ex_share_strings:parse(ZipHandle),
			%% 读取sheet
			Result = ex_workbook:parse_sheet(ZipHandle, ShareStringList, SheetNameList),
			zip:zip_close(ZipHandle),
			Result
	end.

%% ====================================================================
%% Internal functions
%% ====================================================================


