%% @author box
%% @doc xlsx中的sharedStrings.xml信息提取


-module(ex_share_strings).
-include_lib("xmerl/include/xmerl.hrl").

%% ====================================================================
%% API functions
%% ====================================================================
-export([
		 parse/1,
		 get_share_string/2
		]).

%% return -> ShareStringList | {error, Reason}.
%% ShareStringList::[Val|...]
parse(ZipHandle) ->
	ShareStringFile = "xl/sharedStrings.xml",
	case zip:zip_get(ShareStringFile, ZipHandle) of
		{error, Reason} ->
			{error, Reason};
		{ok, ShareStrings} ->
			{_File, Binary} = ShareStrings,
			parse_content(Binary)
	end.

%% 解析sharedStrings.xml的内容
%% sharedStrings.xml的内容如下：
%% <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
%% <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="10" uniqueCount="10">
%% <si>
%% 		<t>a</t>
%% 		<phoneticPr fontId="1" type="noConversion"/>
%% 	</si>
%% 	<si>
%% 		<t>b</t>
%% 		<phoneticPr fontId="1" type="noConversion"/>
%% 	</si>
%% 	<si>
%% 		<t>c</t>
%% 		<phoneticPr fontId="1" type="noConversion"/>
%% 	</si>
%% </sst>
%% 获取所有<si>标签中的<t>
%% 按顺序放入ShareStringList
%% return -> ShareStringList.
%% ShareStringList:: [string() |...] unicode编码[21517,23383]("我")
parse_content(Binary) ->
	{XmlElement, _Rest} = xmerl_scan:string(binary_to_list(Binary)),
	TList = xmerl_xpath:string("/sst/si/t", XmlElement),
	lists:map(fun(E) -> 
						  [XmlText|_] = E#xmlElement.content,
						  XmlText#xmlText.value
				  end, TList).
		


%% N::integer() >= 0
get_share_string(N, ShareStringList) ->
	lists:nth(N+1, ShareStringList).

%% ====================================================================
%% Internal functions
%% ====================================================================


