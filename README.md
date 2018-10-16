POI之PPT-元素操纵

POI操作PPT提供了HSLF和XSLF两套API，下面的调研结果也是围绕着这两套API对比显示的。“读取”指的是API是否能够解析PPT中的已有元素，并获取相关属性信息。“写出”指的是API是否提供了新建PPT，并在PPT中新建所需元素的构建功能。

- **文本**

	读取

    HSLF: yes; XSLF: yes
	
	写出

	HSLF: yes; XSLF: yes


- **自选图形**

    读取

	HSLF: yes; XSLF: yes

    写出

    HSLF: yes; XSLF: yes

    **ps:** 

	在Git@OSC中的代码示例，读写自选图形方面主要展示了获取各种对应的图形类并输出了一些简单信息。至于图形类具体能够做哪些操作，怎么做，可以参照官方示例代码及API手册。
	
	在组合图形方面，HSLF和XSLF并没有提供官方直接的示例，google的结果也不尽人意，所以笔者认为调研结果不够充分，暂时没有好的办法去处理组合图形。

- **图片**

	读取

	HSLF: yes; XSLF: yes

    写出

    HSLF: yes; XSLF: yes

	**ps：**

	XSLF API要比HSLF支持更多的图片类型。具体可参见API文档，Picture中的图片类型常量用于HSLF；XSLFPictureData中的图片类型常量用于XSLF。

- **表格**

	读取
	
	HSLF: yes; XSLF: yes;

	写出

	HSLF: yes; XSLF: yes;

	**ps：**

	XSLF API操作表格元素的代码，笔者认为功能更丰富些，处理得更细腻，也更合乎逻辑习惯。

- **图表**

	读取
	
	HSLF: yes; XSLF: yes;

	写出

	HSLF: no; XSLF: no;

	**ps：**

	HSLF API并没有给出直接的，关于操作图表的示例，读取PPT时图表被解析为了图片。至于如何将图表写出到新建的PPT中，也没有方便的API可调用。但官网示例中提到了可以使用实现的PowerPoint 2D图形驱动直接在幻灯片中绘制图表，操作略繁琐，而且并不是完全兼容Java的java.awt.Graphics2D，所以有些特性并不支持。

	XSLF API操作图表，官方给出了一个关于饼状图的示例。示例中是通过读取既有的PPT图表和修改绑定的excel数据源，重新写出到新的PPT中。整个示例代码操作较为复杂（有点乱），关键一点，示例中用到的XSLFChart类，笔者在API文档发现了@Beta字样。所以笔者推断XSLF API在操作图表方面还处于测试阶段，官方也没有系统地给出相关教程。
	

- **超链接**

	读取

	HSLF: yes; XSLF: no

	写出

	HSLF: no; XSLF: yes

	**ps：**

	关于XSLF API操作超链接，官方示例中只演示了如何添加超链接到PPT，并没有读取的示例。笔者使用逆推的方式只能勉强获取到文本框的超链接地址，很不理想。添加超链接到PPT目前只针对文本框起作用。

	关于HSLF API，官方示例仅给出了如何读取PPT中的超链接。笔者查阅了相关API，目前只发现可以在图形（文本框也算一种图形）上绑定超链接生成PPT，未发现生成超链接文本的方法。而且最终生成的ppt，打开后虽然图形显示有超链接，但打不开指定的链接，目前无解。。。

- **音频**

	读取

	HSLF: yes; XSLF: no

    写出

    HSLF: no; XSLF: no

	**ps:** 

	HSLF API提供了可以读取PPT中音频文件的类，但笔者测试貌似只对WAV格式文件起作用；无法将音频文件写出到新的PPT中。

- **视频**

	读取

	HSLF: yes; XSLF: no;

	写出

	HSLF: no; XSLF: no;

	**ps：**

	HSLF API中提供了MovieShape类，可以读取到幻灯片中嵌入的视频绝对路径，可以使用IO流将视频文件写入到目标地址。但未发现将视频文件嵌入到新建PPT中的方法。

	XSLF API则未发现有关视频操作的类。


- **SmartArt**

	读取

	HSLF: yes; XSLF: no

	写出

	HSLF: yes; XSLF: no

	**ps:** 

	HSLF API解析SmartArt为图片，所以可以采用操作图片的方式将PPT中的SmartArt存储为png图片。但XSLF API解析SmartArt为XSLFGraphicFrame，暂时无解。。。


- **公式**

	读取

	HSLF: yes; XSLF: no

	写出

	HSLF: no; XSLF: no

	**ps:** 

	HSLF API解析公式为图片，所以可以采用操作图片的方式将PPT中的公式存储为png图片。但XSLF API解析不到公式，暂时无解。。。


- **艺术字**

	读取

	HSLF: yes; XSLF: no

	写出

	HSLF: no; XSLF: no

	**ps:** 

	HSLF API解析艺术字为图片，所以可以采用操作图片的方式将PPT中的艺术字存储为png图片。但XSLF API解析艺术字为XSLFAutoShape，暂时无解。。。

	
- **声明**

	笔者的调研结果，自我感觉有不足和遗漏之处，主要原因有以下几点：

	
	1 以上所列元素是笔者根据自己的业务所更加关注的，但肯定不是最全面的。

	2 POI官方给出的关于PPT的教程示例，其实挺混乱，过于零散，缺乏系统的规划与整理。从上述的对比和Git@OSC代码示例中不难看出，HSLF与XSLF看似类似，但其实真正用起来会发现差异很大，有些甚至是逻辑上的本质转变。有些元素的代码示例是笔者根据API文档自己推出来的，官方并没有明确举例，所以若存在错误或误解之处，欢迎各位朋友积极指正。

	3 关于XSLF API，官方有说明还在不断的发展和改善中，难免笔者的代码示例，在将来会发生变化，还望包涵。

官方HSLF示例：[https://poi.apache.org/slideshow/how-to-shapes.html](https://poi.apache.org/slideshow/how-to-shapes.html "官方HSLF示例")

官方XSLF示例：[https://poi.apache.org/slideshow/xslf-cookbook.html](https://poi.apache.org/slideshow/xslf-cookbook.html "官方XSLF示例")  

[http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xslf/usermodel/](http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xslf/usermodel/ "官方XSLF示例")