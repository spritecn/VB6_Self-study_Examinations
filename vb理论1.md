# 基础理论
- vb是面向对像的程序设计语言
- vb构成对象3要素:属性，方法和事件
- vb应用程序是通过事件来驱动的
- 对象的Name属性用于唯一标识 一个对象，是代码中的引用属性，所有对象都有Nmae属性
- Caption 属性，即标题 属性
- Enabled 属性，用于设置对象的有效性
- Visible 属性 用于设置对象的可见性
- FontName,FontSize,FontBold, 用于设置字体，字体大小，加粗
- ForeColor,BackColor 设置用景色、背影色
- Option Explicit 表示使用严格格式
- Option Base 1 设置arr(5)这样没有To的数组下标从１开始
- 所有的除变量定义之外的代码，必须放在函数内部执行，也就是说全局变量的赋值也必须在函数中，一般在Load中不然会报无交效的外部过程
- 三个工作模式
    - 设计模式，界面设计和更改代码
    - 运行模式
    - 中断模式，运行时暂时中断，可以编辑代码，但不能编辑界面，按F5继续到运行模式
- 相关文件类型
    - .vbp 工程文件
    - .frm 窗体文件
    - .bas标准模块组件
    - .cls类模块文件
- vb中将错误分为3类
    - 编译错误,语法错误,换行vb可以检查到
    - 运行时错误，可以运行但运行时会报错，比如除数为0
    - 逻辑错误，虽然能正常运行但得不到预期结果 
- 调试:单步调试,设置断点调试
- true = -1 ,false = 0,对于逻辑表达式非0即真
- dim
- const
- private,public,static
    - private 局部的,仅本模块中的其他过程可调用
    - public 本模块和其他模块都可以调用,一个框体是一个模块
    - 不能级别的域内出现同名变量时,从里到外查找
## 运算符/表达式
- + - X / \ ^ Mod
- / 为除, \为整除, Mod为取余
- \整除和Mod时,如果操作数带有小数,先按 四舍六入五凑偶(也叫五成双) 原则转化成整数再进行运算
- ^为乘方运算,可以解决数学上的平方,开方等(建议用Sqr,^逻辑有时候是有问题的)
    - -3^2  = 9 '负3的平方
    - 27^(1/3) = 3  '27的3分之一次,即27的立方根
    - 2^(-2) = (1/4)  '2的负2次,即2的平的的倒数 
    - 2^2^3 = 64  '从左向右运算
    - -27^(1/3) = -3  ' -27开3次方
    - -4^(1/2) 报错,-4不能开平方根(课本上说会报错,但实际写到程序里并不报...这课本真的是给不学生编的)
    - ***注^前不能用括号  (-9)^2  会报错

## 数组
- 声明
    - dim a() '声明一个没界限制的变体类型数组
    - dim b(4) '声明一个有5个元素的变体型数组,即dim (0 TO 4)
    - dim c() as Integer '无界的整型数组
    - dim d(4) as Integer '5个元素的整型数组
    - ()里不能写变量,仅对arr()进行ReDim时括号里才能写变量  
- 取上界和下界
    - LBound(arr) 取arr的上界
    - Ubound(arr) 下界
- 输入,输入
    - 循环输入如arr(4) 用 For i = LBound(arr) to UBound(arr)   \ arr(i) = xxx \ Next i 这样输入
    - Array函数仅用于给变体类型或仅有括号没有维界的变体数组输入比如dim arr()用 Array(1,2,3,4)输入
    - dim arr() as String这样的动态数组,只能用ReDim arr(ex)重定义之后才可以使用,ex可以是表达式
    - 使用Redim 原值会丢失,需要用ReDim preserve才能保留原值 ,但如果用preserve只能更改上界,不能更改下界(**这样只能用来切断数组了**)
    - 输出和循环输入一样用元素下标,arr(1)
- 二维数组
    - 定义dim　arr(1 to 3,1 to 3),即定义了一个 3*3的二维数组
    - 使用时用两个下标arr(1,3) 表示二维第一行第三列
    - 日历可以使用一个6*7的二维数组来表示，第一行星期几，第二行开始是日期对应的星期
- 三维及以上的数组移称为多维数组


### 控件数组
- 先创建控件手动设置index为0，通过Load添加控件元素，下标>原始控件的index，下标不能重复
- 可以使用Unload将load的控件元素删除，不删除index为0的控制元素
- 原始控件的index必须要手动设置
- 控制元素load之后除index,tabIndex,Visible之外其他的属性是复制了index最小的那个元素的属性所以，需要手动的设置top或left属性使之可以正常显示，option1(i).top = option1(i-1).top +300

## 数据类型

### 自定义数据类型
- 由若干个基本数据类型 组成的一个新的数据类型即自定义数据类型
- 定义语法 
    ```vb 
    private Type 类型名 
        元素名1 as 数据类型
        元素名2 as 数据类型 
        ...
    end Type
- 窗体等私有模块里只能使用private声明,在标准模块里public和private均可以使用
- 类型定义好后可以象普通变量 一样声明此类型的变量
- 为了区分自定义数据类型名,建议首字母大写
- 用.访问类型成元
- 为了方便吏用vb引入了with语句(***这高级的特性,居然vb的时候就有了,到现在java都没有**)
    ```vb
    private type Dog '定义dog
        color as String
        age as Integer
    end Type
    function makeDog(String color) as Dog '创建dog
        dim dog as Dog
        with dog
            dog.color = color
            dog.age = 1
        end with
        makeDog = dog
    end function
    ```
- 自定义数据类型娄组,和数组没区别,dim dog() as Dog

## 常用函数
- Val  将输入转换为数组值
- inputBOx
- msgBox
### 字符串
    - inStr(1,str1,str2),返回从str1的第一个字符开始算,str2出现的位置
    - 
### 数学函数
- Int,Fix,Round的区别
    - Int为左取整,即对于大于0的数小数点后直接舍弃,对于小于0的数,如果小数点后不为0,向左-1,即int(2.9)=2,int(-2.1) = -3
    - Fix为截断,不管小数点后是什么,都舍弃即 fix(2.x) = 2
    - Round 算法为***四舍六入五凑偶*** 即Round(2.4) =2, Round(2.6) = 3,Round(1.5) = 2,Rount(2.5) = 2

## 控制结构
### 循环
- vb 没有continue
- for 循环
    - for i=1 to 10 step 步长  ... exit For ...  next i
    - exit For 用于退出整个循环的
    - for循环执行次数公式 (终于 - 初值)/步长 + 1
    - For循环变量 不建议在循环外使用,for i = 1 to 10,表示for int i  = 1;i<=10;i++,这个i一轮循环完后结果为11，不是10
    - 在严格格式下 for变量的i也是需要 dim声明的
- Do [<..>]... Loop [<...>]
    - while表示表达式为真时执行,until相反
    - while/until条件可以写在do后面先判断条件再执行,也可以写在Loop后面表示先执行一次然后判断条件
    - exit do 退出循环
- While <...> ... wend
    - 先检查条件再执行
    - while没有exit语句退出循环
    - 通常使用flag对while进行退出如
    ```vb
    dim exitFlag
    exitFlag = 0
    while exitFlag <>1
        ...
        if xxx then exitFlag = 1
    wend
    ```
- exit for,exit Do 仅退出本层循环


### 分支
- If,ElseIf Else End if
- 一种是 Select Case语句,类似于 java的switch,比switch设计的好
    - 如下
    ``` vb
    Select Case a
        Case 1
            out = "一"  '不需要break
        Case 2
            out = "二"
        Case Else
            out = "不在范围中"
    End Select
    - case 后面可以写 1 TO 10这样的范围
    - Case 后面可用用逗号分割多个条件,表达或的关系
    - Case 后面可以用 IS 来代替测试表达式,比如 Case Is < 100
    - Case 后面不能写对测试表达式的判断语句,比如上面不能Case a = 0,但可以用 Case is = 0
- IIF 函数,来实现三元操作
    - max = IFF(a>b, a, b)
- Choose 函数,用来实现根据数值取值
    - session =　Chose(3,"春","夏","秋"，"冬")  '结果是秋

##  过程
- 函数过种 funciton,可返回值,在end function之前以 函数过程名 = 返回结果的形式返回
- 子过程  sub
- 函数过程 可以用staitc标识,表示该过程中所有的局部变量 都是表态变量 
- static 静态变量,仅用在过程中对变量进行声明,声明后此变量在下次执行函数时延用上次执行的值(***居然实现了闭包**)
- 函数过程必须有返回值
- 函数过程调用时,如果没有参数,可以省略括号
- vb有两种传参型式
    - 值传递(ByVal),使用时需要在形参定义时加上byVal,过程执行时会对实参copy一份使用
    - 地址传递(byRef),传地址,形参直接使用实形的地址进行处理,双向传递
    - *** vb默认传递方式为传地址 byRef ***
    - 数组,自定义数据类型,及控件 只能以地址传递的方式传递
- vb没有return
- 递归,即自己调用 自己 的过程 
- exit sub/exit function 可以用于提前 退出
## 文件
- TODO


## 控件
### 定时器
### 文件选择框
### 
