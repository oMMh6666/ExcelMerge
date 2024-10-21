Windows下控制台程序，已测试Win10 64位，其他未测试

不需要目标主机安装有Office，使用NPOI组件进行

合并原则如下：

1、不同SheetName的Sheet页直接置入

2、相同SheetName的Sheet页以处理的第一个文件的第一行为标题列，其他文件不会再合并第一行，其他行若列数多于第一个处理的文件，依然会写入合并文件


几种自动化处理合并的方式：

1、直接拖拽xls xlsx文件到exe可执行文件

2、将xls xslx文件和exe文件放置于同一目录，直接Enter键处理

3、直接Enter键若目录下没有excel文件，则按config.yaml中配置的内容进行合并动作
