%% @author box
%% @doc sheet.xml读取


-module(ex_sheet).
-include_lib("xmerl/include/xmerl.hrl").

%% ====================================================================
%% API functions
%% ====================================================================
-export([
		 parse/4
		]).

%% return -> {SheetName, RowDataList}
%% RowData::[CellVal()::term()|...]
parse(ZipHandle, ShareStringList, SheetName, SheetPath) ->
	FileName = filename:join("xl", SheetPath),
	case zip:zip_get(FileName, ZipHandle) of
		{error, Reason} ->
			io:format("parse sheet error, file:~p, reason:~p~n", [SheetName,Reason]),
			{SheetName, []};
		{ok, Res} ->
			Res,
			{_File, Binary} = Res,
			{XmlElement, _Rest} = xmerl_scan:string(binary_to_list(Binary)),
			RowDataList = parse_content(XmlElement, ShareStringList),
			{SheetName, RowDataList}
	end.

%% ====================================================================
%% Internal functions
%% ====================================================================
%% return -> RowDataList
%% RowData::[{CellName, CellVal()::term()}|...]
%% <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
%% 	<dimension ref="A1:C2"/>
%% 	<sheetViews> <sheetView workbookViewId="0"><selection activeCell="D6" sqref="D6"/></sheetView></sheetViews>
%% 	<sheetFormatPr defaultRowHeight="14.25" x14ac:dyDescent="0.2"/>
%% 	<sheetData>
%% 		<row r="1" spans="1:3" x14ac:dyDescent="0.2">
%% 			<c r="A1"><v>1</v></c>
%% 			<c r="B1"><v>2</v></c>
%% 			<c r="C1"><v>3</v></c>
%% 		</row>
%% 		<row r="2" spans="1:3" x14ac:dyDescent="0.2">
%% 			<c r="A2" t="s"><v>0</v></c>
%% 			<c r="B2" t="s"><v>1</v></c>
%% 			<c r="C2" t="s"><v>2</v></c>
%% 		</row>
%% 	</sheetData>
%% 	<phoneticPr fontId="1" type="noConversion"/>
%% 	<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
%% 	<pageSetup paperSize="9" orientation="portrait" r:id="rId1"/>
%% </worksheet>
parse_content(XmlElement, ShareStringList) ->
	RowList = xmerl_xpath:string("/worksheet/sheetData/row", XmlElement),
	MergeList = get_merge_cell_info(XmlElement),
	{_, RowDataList} = 
		lists:foldl(fun(E, {MergeL, RowDataL}) ->
							CellList = xmerl_xpath:string("c", E),
							{NewMergeL, AddRowDataL} = lists:foldl(fun(CellE, {MergeL1, L}) -> 
												{CellName, Val, MergeL2} = get_cell_val(CellE, ShareStringList, MergeL1),
												{MergeL2, [{CellName, Val} |L]}
										end, {MergeL, []}, CellList),
							 {NewMergeL, [lists:reverse(AddRowDataL)|RowDataL]}
					end, {MergeList, []}, RowList),
	fix(lists:reverse(RowDataList)).


%% 获取合并单元格的信息
%% return -> [{CellName, MergeToCellName}|...]
%% 如 A1和A2合并， return -> [{"A1", "A1"}, {"A2","A1"}]
%% <mergeCells count="2">
%% 	<mergeCell ref="A1:A2"/>
%% 	<mergeCell ref="B1:D1"/>
%% </mergeCells>
get_merge_cell_info(XmlElement) ->
	MergeList = xmerl_xpath:string("/worksheet/mergeCells/mergeCell", XmlElement),
	lists:foldl(fun(E, L) -> 
					  XmlAttr = lists:keyfind('ref', #xmlAttribute.name, E#xmlElement.attributes),
					  [FromCellName, ToCellName] = string:tokens(XmlAttr#xmlAttribute.value, ":"),
					  AllCell = spread_cell(FromCellName, ToCellName),
					  [{Cell, FromCellName} ||Cell<-AllCell] ++ L
			  end, [], MergeList).

%% 铺开
%% spread_cell("A1", "A3") -> ["A1", "A2", "A3"].
%% spread_cell("A1", "B2") -> ["A1", "A2", "B1", "B2"].
spread_cell(FromCellName, ToCellName) ->
	Column1 = get_column_by_cell_name(FromCellName),
	Column2 = get_column_by_cell_name(ToCellName),
	Row1 = get_row_by_cell_name(FromCellName),
	Row2 = get_row_by_cell_name(ToCellName),
	ColList = lists:seq(Column1, Column2),
	RowList = lists:seq(Row1, Row2),
	lists:foldl(fun(Col, L) -> 
						ColStr = column_num_to_string(Col),
						AddL = lists:map(fun(Row) -> 
												 RowStr = integer_to_list(Row),
												 ColStr ++ RowStr
										 end, RowList),
						AddL ++ L
				end, [], ColList).


%% 'r'标签
get_cell_name(CellE) ->
	AttrList = CellE#xmlElement.attributes,
	Attr = lists:keyfind('r', #xmlAttribute.name, AttrList),
	Attr#xmlAttribute.value.

%% return -> {CellName, Val, NewMergeList}
get_cell_val(CellE, ShareStringList, MergeList) ->
	Name = get_cell_name(CellE),
	{Val, NewMergeList} = get_cell_value2(CellE, ShareStringList, MergeList),
	{Name, Val, NewMergeList}.

%% 'v'节点
%% return -> {Val::string()|number(), MergeList}
%% 注意合并单元格的情况，是mergecell的有's'属性，有"v"标签的需要更新MergeList， 没有"v"的从MergeList中取值
%% A1,B1,C1合并
%% <c r="A1" s="1"><v>2</v></c>
%% <c r="B1" s="1"/>
%% <c r="C1" s="1"/>
get_cell_value2(CellE, ShareStringList, MergeList) ->
	%% 是否有's'属性
	IsMergeCell = has_attr('s', CellE),
	case xmerl_xpath:string("v", CellE) of
		[] -> 
			case IsMergeCell of
				true ->
					CellName = get_cell_name(CellE),
					Val = get_value_from_merge_list(CellName, MergeList),
					{Val, MergeList};
				false ->
					{"", MergeList}
			end;
		[VE|_] -> 
			[XmlText|_] = VE#xmlElement.content,
			Val = XmlText#xmlText.value,
			NewVal = 
				case is_string_value(CellE) of
					false -> 
						string_to_number(Val);
					true ->
						%% 需要从ShareStringList找出对应的值
						ex_share_strings:get_share_string(erlang:list_to_integer(Val), ShareStringList)
				end,
			case IsMergeCell of
				true -> 
					CellName = get_cell_name(CellE),
					NewMergeList = save_merge_value_to_merge_list(CellName,NewVal, MergeList),
					{NewVal, NewMergeList};
				false ->
					{NewVal, MergeList}
			end
	end.

get_value_from_merge_list(CellName, MergeList) ->
	case lists:keyfind(CellName, 1, MergeList) of
		false -> "";
		{_, MergeToCellName} -> 
			case lists:keyfind({merge_key, MergeToCellName}, 1, MergeList) of
				false -> "";
				{_, Val} -> Val
			end
	end.

save_merge_value_to_merge_list(CellName,Val, MergeList) ->
	[{{merge_key, CellName}, Val} |MergeList].

%% "1" -> 1, "0.4" -> 0.4
string_to_number(String) ->
	case erl_scan:string(String ++ ".") of
        {ok, Tokens, _} ->
            case erl_parse:parse_term(Tokens) of
                {ok, Term} -> Term;
                _Err -> _Err
            end;
        _Error ->
            undefined
    end.

%% 是否有't'属性
%% return -> true | false
is_string_value(CellE) ->
	has_attr('t', CellE).

%% return -> true | false
has_attr(AttrName, XmlEle) ->
	AttrList = XmlEle#xmlElement.attributes,
	case lists:keyfind(AttrName, #xmlAttribute.name, AttrList) of
		false -> false;
		_ -> true
	end.

%% 补全缺失的行和列
fix(RowList) ->
	{_, NewRowList} = 
		lists:foldl(fun(CellList, {RowN, RList}) -> 
							case CellList of
								[] ->
									{RowN, RList};
								[{CellName, _Val} | _] ->
									RowNum = get_row_by_cell_name(CellName),
									NewCellList = fix_cell(CellList),
									case RowNum == RowN of
										true ->  {RowN+1, [NewCellList|RList]};
										false -> 
											{RowNum+1, [NewCellList, lists:duplicate(RowNum - RowN - 1,[])|RList]}
									end
							end
					end, {1, []}, RowList),
	lists:reverse(NewRowList).

fix_cell(CellList) ->
	fix_cell2(CellList, 1, []).

fix_cell2([], _N, ValList) ->
	lists:reverse(ValList);
fix_cell2([{CellName, Val} | CellList], N, ValList) ->
	Column = get_column_by_cell_name(CellName),
	case Column == N of
		true -> 
			fix_cell2(CellList, N+1, [Val|ValList]);
		false ->
			fix_cell2([{CellName, Val} | CellList], N+1, [""|ValList])
	end.

%% 获取行号
%% CellName::string() , 如"A1" 
%% return -> integer()
get_row_by_cell_name(CellName) ->
	[ColumnStr] = string:tokens(CellName, "0123456789"),
	RowNumStr = string:right(CellName, length(CellName) - length(ColumnStr)),
	list_to_integer(RowNumStr).

%% 获取列号
%% 编号顺序：A,B,...,Y,Z,AA,AB,... (26进制1~26)
get_column_by_cell_name(CellName) ->
	[ColumnStr] = string:tokens(CellName, "0123456789"),
	get_column_num(ColumnStr, 0).

%% "A" == [65]
get_column_num([], Res) -> Res;
get_column_num([H|ColumnStr], Res) ->
	Num = H - 64,
	NewRes = Res*26 + Num,
	get_column_num(ColumnStr, NewRes).

%% 1 -> "A", 2 -> "B"
column_num_to_string(N) ->
	[N+64].


