
<center><h1><b> EasyGoods 谷子排肾表工具 </b></h1></center>

一个让谷子的排表能转变为肾表的工具，支持合并入原始排表或新输出肾表。

---

实际上就是一个统计工具，灵感来源于两年前做计算机二级题库的时候看到的水果统计题（给定一系列输入，计算每个输入出现的次数）。排表转肾表的本质就是带有权重的水果统计题，而且还更格式化，因此就诞生了这个谷子排肾表工具！*也有可能是因为肾表一做就是一个小时，补药再做搬运数据的猴子工作了......*

说明书随Release一起发布，目录下有build所使用的命令。很多地方都是o3帮忙写的，数据科学做久了导致这种界面设计和基于openpyxl的数据处理完全苦手，因此最开始写了一版Pandas版的，后面发现打包体积太大了就让o3帮忙转成了全openpyxl。也算是一次小小尝试吧！这个工具就花了2天写出来，估计Bug挺多。

Contact me: Vanishark@163.com

* [ ] 角色名问题：只支持单字角色名
* [ ] 重构代码：完全不考虑设计模式的一锅炖，实在太丑陋，可维护性太低
* [ ] 兼容性提高：价格数字的兼容性处理
* [ ] 转单识别
* [ ] 更美丽的界面，能不能做上线啊（）
* [ ] 做个教程页面，录个视频
