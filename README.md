# BetterOrigin

我之前都是用 matlab 比较多，处理完数据直接绘图输出，易于编程便于修改。但实验室常用 origin 处理数据和绘图， 尝试了一下， origin 的内置函数比如数据拟合函数的种类和功能确实要比 matlab 的 CurveFit 工具箱要多得多好用得多，而且matlab默认画出来的图确实丑. 而对于批量的数据处理和绘图，用origin想来是要疯掉的。正好最近接到一个类似任务，小小调研了一下，寻得一神器，或许对其他人也有用，分享之。

doc_example.m 通过 Origin 的 Automation Server 功能，
首先创建了一个 Origin COM Server object
“originObj” ，然后通过 invoke 语句操纵 “originObj” ，
“originObj” 有很多可操纵的参数
详见 Origin 帮助文档 
Programming > Automation Server > ApplicationSI 

“originObj” 的 'Execute'
参数意味着可以传递origin自带的labtalk语言，
而labtalk语言理论上可以实现origin所有用点击能完成的功能，
嘿嘿，这样matlab就相当于完全控制origin啦。
详细信息同样参考origin帮助文档labtalk语言部分，
更多功能任你实现 
