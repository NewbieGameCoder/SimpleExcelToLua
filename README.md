本工具旨在提供一个简单的Excel转lua方式。Excel配置文件特定的格式如下：

1、需要转成Lua文件的Excel Sheet表需命名为LuaTable。

2、Excel Sheet表头4行需有特定作用：第一行是数据块的注释说明，第二行是数据块的命名，第三行是数据块的数据对应的类型，第四行是组织一堆数据块的方式。

3、本工具组织Excel配置数据块成LuaTable数据的方式有一下三种：[]、{}、|。

    1>“[d]”代表这一列和接下来的d-1列都会被组织成一个数组形式的LuaTable表（即不存在Key值（{{1, 2, 3}, {2, 3, 4}}）。d可以为任意数字或者不填任何数字。
  
    2>“{d}”代表这一列和接下来的d-1列都会被组织成一个键值对形式的LuaTable表（即不存在Key值（[1] = {1, 2, 3}））。d可以为任意数字或者不填任何数字。
  
    3>“|”在数据块中代表本数据块其实是一个数组，比如“20012|20013|20014”这个字符串，会在转成lua文件的时候变成{20012, 20013, 20014};当“|”在第四行（数据块组织方式）的时候，则代表“分表”的概念，用于处理在一张Excel配置里面配置2个不同格式的Lua表（具体可以见工具路径下的TestCase文件夹里面的示例文件）。

本工具缺少很多Excel已经集成的数据检测，只提供一些较简单的数据检测，比如Key是否重复等。具体使用可以参考工具路径下的TestCase文件夹里面的示例文件。
本工具还欠缺一个数据坍缩的功能（即直接将Excel文件转成[1] = 2013,的形式），尚不影响大部分情况下的使用，日后补全。

简单示例：
Excel :
![image](https://github.com/NewbieGameCoder/SimpleExcelToLua/raw/master/Snapshot/excelCSS.bmp)
Lua :
![image](https://github.com/NewbieGameCoder/SimpleExcelToLua/raw/master/Snapshot/resualt.bmp)