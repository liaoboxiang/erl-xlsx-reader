%% @author box
%% @doc sheet.xml读取


-module(ex_workbook).
-include_lib("xmerl/include/xmerl.hrl").

%% ====================================================================
%% API functions
%% ====================================================================
-export([parse/2,
		 parse_sheet/3]).

%% return -> SheetList | {error, Reason}
%% SheetList::[SheetData|...]
%% SheetData::{SheetName, RowDataList}
%% RowData::[CellVal()::term()|...]
%% 解析所有的sheet
parse(ZipHandle, ShareStringList) ->
	%% 所有的sheet
	case get_all_sheet(ZipHandle) of
		{error, Reason} -> {error, Reason};
		SheetInfoList -> 
			Rels = ex_workbook_rels:get_rels(ZipHandle),
			%% Rels::[{SheetId, SheetPath}|...]
			lists:map(fun({SheetId, SheetName}) -> 
							  SheetPath = math_path_from_rels(SheetId, Rels),
							  ex_sheet:parse(ZipHandle, ShareStringList, SheetName, SheetPath)
					  end, SheetInfoList)
	end.

parse_sheet(ZipHandle, ShareStringList, SheetNameList) ->
	%% 所有的sheet
	case get_all_sheet(ZipHandle) of
		{error, Reason} -> {error, Reason};
		SheetInfoList -> 
			SheetInfoList1 = lists:filter(fun({_SheetId, SheetName}) -> 
												  lists:member(SheetName, SheetNameList) 
										  end, SheetInfoList),
			Rels = ex_workbook_rels:get_rels(ZipHandle),
			%% Rels::[{SheetId, SheetPath}|...]
			lists:map(fun({SheetId, SheetName}) -> 
							  SheetPath = math_path_from_rels(SheetId, Rels),
							  ex_sheet:parse(ZipHandle, ShareStringList, SheetName, SheetPath)
					  end, SheetInfoList1)
	end.
	

math_path_from_rels(SheetId, Rels) ->
	{_, SheetPath} = lists:keyfind(SheetId, 1, Rels),
	SheetPath.	
%% ====================================================================
%% Internal functions
%% ====================================================================
%% return -> [{Id::string(), SheetName::string(})|...]
get_all_sheet(ZipHandle) ->
	case zip:zip_get("xl/workbook.xml", ZipHandle) of
		{error, Reason} -> {error, Reason};
		{ok, Res} ->
			{_File, Binary} = Res,
			{XmlElement, _Rest} = xmerl_scan:string(binary_to_list(Binary)),
			SheetList = xmerl_xpath:string("/workbook/sheets/sheet", XmlElement),
			lists:map(fun(E) -> 
							  AttrList = E#xmlElement.attributes,
							  SheetName = get_name_from_attrs(AttrList),
							  Id = get_id_from_attrs(AttrList),
							  {Id, SheetName}
					  end, SheetList)
	end.

get_name_from_attrs(XmlAttrList) ->
	case lists:keyfind(name, #xmlAttribute.name, XmlAttrList) of
		false -> "";
		Attr ->
			Attr#xmlAttribute.value
	end.

get_id_from_attrs(XmlAttrList) ->
	case lists:keyfind('r:id', #xmlAttribute.name, XmlAttrList) of
		false -> "";
		Attr ->
			Attr#xmlAttribute.value
	end.





