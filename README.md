# ExcelUnion
enviroment：python 2.7 xlwt xlrd


接受两个参数，第一个参数输入要合并的文件所在文件夹的名字，第二个参数说明将所有文件的第N行合并到同一个文件中
**注意：日期格式写为:
年.月.日 或 年/月/日 或 月/日 、 年/月
不要 写成 月.日、 年.月**
<!--因为学号和班号这类在读取每个表时存到list中的是float，若直接输入会表示为科学计数法，采取的解决方案是先将float化为整形再化为字符串类型。-->
文件名不要包含破折号
统一放在文件的同一行，即表头正下面一行，不要空行，不要两人填一张表，要是原来有示范内容删掉再填

*也可以每个文件有多行，但是先就这样把*
