%% @author box
%% @doc workbook中的 id与sheet.xml文件的 对应关系


-module(ex_workbook_rels).
-include_lib("xmerl/include/xmerl.hrl").

%% ====================================================================
%% API functions
%% ====================================================================
-export([get_rels/1]).

%% return -> [{SheetId, SheetPath}].
get_rels(ZipHandle) ->
	case zip:zip_get("xl/_rels/workbook.xml.rels", ZipHandle) of
		{error, Reason} -> {error, Reason};
		{ok, Res} ->
			{_File, Binary} = Res,
			{XmlElement, _Rest} = xmerl_scan:string(binary_to_list(Binary)),
			RelationshipList = xmerl_xpath:string("/Relationships/Relationship", XmlElement),
			lists:map(fun(E) -> 
							  AttrList = E#xmlElement.attributes,
							  Id = get_id_from_attrs(AttrList),
							  SheetPath = get_sheet_path_from_attrs(AttrList),
							  {Id, SheetPath}
					  end, RelationshipList)
	end.

%% ====================================================================
%% Internal functions
%% ====================================================================
get_id_from_attrs(XmlAttrList) ->
	case lists:keyfind('Id', #xmlAttribute.name, XmlAttrList) of
		false -> "";
		Attr ->
			Attr#xmlAttribute.value
	end.

get_sheet_path_from_attrs(XmlAttrList) ->
	case lists:keyfind('Target', #xmlAttribute.name, XmlAttrList) of
		false -> "";
		Attr ->
			Attr#xmlAttribute.value
	end.





