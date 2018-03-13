%% @author box
%% @doc 测试

-module(ex_test).

%% ====================================================================
%% API functions
%% ====================================================================
-export([
		 t1/0,
		 t2/0,
		 unzip/1
		]).

t1() ->
	FileName = "../xlsx/t1.xlsx",
	ex_reader:read(FileName).

t2() ->
	FileName = "../xlsx/t2.xlsx",
	ex_reader:read(FileName).

%% 把xlsx文件解压到文件
%% ex_test:unzip("t2").
unzip(FileName) ->
	FileName1 = if 
					is_atom(FileName) -> atom_to_list(FileName);
					is_integer(FileName) -> integer_to_list(FileName);
					true -> FileName
				end,
	FullFileName = filename:join(["../xlsx/", FileName1++".xlsx"]),
	zip:unzip(FullFileName).

%% ====================================================================
%% Internal functions
%% ====================================================================


