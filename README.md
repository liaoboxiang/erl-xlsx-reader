# erl-xlsx-reader
读取xlsx文件

## USAGE:
```
windows下
编译文件
>git clone https://github.com/liaoboxiang/erl-xlsx-reader.git
>rebar compile
启动shell
cd 到scripts/
>./start.bat
```

```
使用
ex_reader:read(FileName).

返回结果
[
{Sheet1, [[A1,B1|...],
		  [A2,B2|...]]},
{Sheet2, [[A1,B1|...],
		  [A2,B2|...]]}
|...		 
]
```

```
例子
ex_test:t1().
ex_test:t2().
```

```
注意：
1、单元格中没有数据时，获得的数据为""，例如["A1", "", "C1"]
2、合并的单元格，获取的值为左上角的单元格的值，例如 A1,B1,C1合并，return -> ["A1","A1","A1"]
3、erlang版本：R20,未测试其他版本，请注意编码
```
