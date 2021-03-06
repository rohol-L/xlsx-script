# xlsx-script
一个通过指令控制、可编程的xlsx模板类库。

+ 支持多sheet页模板
+ 支持行级子模板
+ 支持语法拓展
+ 保留原始样式

## 快速开始
``` js
import xlsx_script from './xlsx-script'

let testData=[
    {day: "1", count: 10},
    {day: "2", count: 6},
    {day: "3", count: 8},
]

xlsx_script.loadUrl('/templet/模板.xlsx')
    .then((xs)=>{
        xs.render(testData)
        xs.export("输出")
    })
/*
==== 模板.xlsx ====
|日期          |数量    |
|{$.fill}{day}|{count}|

==== 输出.xlsx ====
|日期 |数量 |
|1   |10  |
|2   |6   |
|3   |8   |
*/
```


## 指令语法
```
// 完整指令结构
{$column.function1(arg1).function2(arg1,arg2)}

// 简写示例(只调用一个方法，并省略字段名称和参数)
{$.fill}
```


| token     | 含义                                                         |
| --------- | ------------------------------------------------------------ |
| $         | 指令类型，$:普通指令，@:后处理指令*。不指定时默认为填充指令  |
| column    | 字段名称，区分大小写，同一个指令块中的方法均可获取到此值。可省略 |
| function1 | 方法名称，参考“内置方法”节点                                 |
| arg1      | 参数列表，无参数可连带括号一起省略，可用的语法参考“参数类型”节点 |

后处理指令说明：

后处理指令中只能处理excel表格本身，无法读取模板绑定的数据，也无法控制渲染指针位置。后处理指令可用的方法会在“内置方法”节点中标记。



## 参数类型
不在下表中的表示法均作为 String 处理。
| 表示法 | 说明                                                         |
| ------ | ------------------------------------------------------------ |
| true   | 表示 Boolean 值 `true`                                       |
| false  | 表示 Boolean 值 `false`                                      |
| 1      | 表示 Number 值 `1`。（判定：`!Number.isNaN(xxx)`）           |
| "1"    | 表示 String 值 `1`。内容中的 `"` 和 `\` 需要使用 `\` 转义（表示为 `\"` 或 `\\`）。 |



## 内置方法 - 模板操作类别

模板操作类别中的方法会改变excel单元格结构（例如新增行、合并单元格），并控制在模板范围内的上下文数据集。



### fill - 行填充

| 参数名称 | 类型 | 含义 | 默认值 |
| -------- | ---- | ---- | ------ |
|mode|String|填充模式。可选填：group（填充时去除重复）||

读取该行的无类型指令（形如`{column_name}`），将上下文数据集按行向下填充。

示例：

```js
// 数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

/* 模板
|uid          |day  |count  |
|{$.fill}{uid}|{day}|{count}|
*/

/* 输出
|uid  |day |count |
|1000 |1   |10    |
|1000 |2   |16    |
|1001 |1   |9     |
|1001 |2   |18    |
*/
```



### for - 子模板循环

| 参数名称 | 类型 | 含义 | 默认值 |
| -------- | ---- | ---- | ------ |
| _column | String | 字段名称（公共参数） | 必填 |
|start|Number|模板开始行，最好是指令所在行|指令所在行|
|end|Number|模板结束行，开始到结束之间的行将会作为多行模板|指令所在行|

指定 `start` ~ `end` 行号范围，生成一个子模板，将上下文数据集按 `_column` 分组，循环向下填充到子模板中。

示例：

```js
// 数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

/* 模板
|{$uid.for(1,4)}{$uid.first} |
|day          |count         |
|{$.fill}{day}|{count}       |
*/

/* 输出
|1000       |
|day |count |
|1   |10    |
|2   |16    |
|    |      |
|1001       |
|day |count |
|1   |9     |
|2   |18    |
*/
```



### merge 合并单元格

| 参数名称  | 类型   | 含义                 | 默认值 |
| --------- | ------ | -------------------- | ------ |
| extWidth  | Number | 横向合并的单元格数量 | 必填   |
| extHeight | Number | 纵向合并的单元格数量 | 必填   |

向右合并`extWidth`个单元格，向下合并`extHeight`个单元格，支持后处理模式。

示例：

```
// 模板
|test{$.merge(1,1)}|  |
+------------------+--+
|                  |  |

// 输出
|test                 |
|                     |
```



## 内置方法 - 数据处理类别

数据处理类别中的方法会改变当前渲染指针的上下文数据集。



### filterMax 最大值过滤

| 参数名称 | 类型 | 含义 | 默认值 |
| -------- | ---- | ---- | ------ |
| _column | String | 字段名称（公共参数） | 可空 |
|columnName|String|字段名称，此参数会覆盖公共参数中的字段名称|可空|

求出数据集中 `_column` 字段的最大值，然后过滤出 `_column` 等于该值的数据。 

示例：

```js
// 指令: {$day.filterMax}
// 等价: {$.filterMax(day)}

// 数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

// 输出
[{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

```



### filterMin 最小值过滤

| 参数名称 | 类型 | 含义 | 默认值 |
| -------- | ---- | ---- | ------ |
| _column | String | 字段名称（公共参数） | 可空 |
|columnName|String|字段名称，此参数会覆盖公共参数中的字段名称|可空|

求出数据集中 `_column` 字段的最小值，然后过滤出 `_column` 等于该值的数据。 

示例：

```js
// 指令: {$day.filterMin} 
// 等价: {$.filterMin(day)}

// 数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

// 输出
[{uid:1000,day:1,count:10},
{uid:1000,day:2,count:16}]

```



### select - 选择数据集

| 参数名称 | 类型   | 含义                            | 默认值 |
| -------- | ------ | ------------------------------- | ------ |
| path     | String | 节点路径，支持绝对路径/相对路径 | 必填   |

修改当前渲染指针的上下文数据集。

示例：

```js
// 数据
{
  modeA:[{data:"dataA"}]
  modeB:[{data:"dataB"}]
}

/* 模板
|{$.select(modeA)}{$data.first}|
|{$.select(modeB)}{$data.first}|
*/

/* 输出
|dataA |
|dataB |
*/
```



### cancelSelect - 撤销选择数据集

撤回到上一次调用 `select` 函数之前的上下文数据集。

注意：跨子模板配对使用 `select` 和 `cancelSelect` 可能导致意外的渲染结果。





## 内置方法 - 数值输出类别

数据处理类别中的方法会直接在指令所在位置输出结果

### first - 取第一项

| 参数名称   | 类型   | 含义                                       | 默认值 |
| ---------- | ------ | ------------------------------------------ | ------ |
| _column    | String | 字段名称（公共参数）                       | 可空   |
| columnName | String | 字段名称，此参数会覆盖公共参数中的字段名称 | 可空   |

取出上下文数据集中的第一个对象的 `_column` 字段值，并填入到指令所在的位置。

示例：

```js
// 指令: {$count.first} 
// 等价: {$.first($count)}

//数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

//输出
10
```



### max - 取最大值

|参数名称|类型|含义|默认值|
| -------- | ---- | ---- | ------ |
| _column    | String | 字段名称（公共参数）                       | 可空   |
| columnName | String | 字段名称，此参数会覆盖公共参数中的字段名称 | 可空   |

取出上下文数据集中 `_column` 字段的最大值，并填入到指令所在的位置。

示例：

```js
// 指令: {$count.max} 
// 等价: {$.max($count)}

//数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

//输出
18
```



### min - 取最小值

| 参数名称   | 类型   | 含义                                       | 默认值 |
| ---------- | ------ | ------------------------------------------ | ------ |
| _column    | String | 字段名称（公共参数）                       | 可空   |
| columnName | String | 字段名称，此参数会覆盖公共参数中的字段名称 | 可空   |

取出上下文数据集中 `_column` 字段的最大值，并填入到指令所在的位置。

示例：

```js
// 指令: {$min.max} 
// 等价: {$.min($count)}

//数据
[{uid:1000,day:1,count:10}
{uid:1000,day:2,count:16},
{uid:1001,day:1,count:9},
{uid:1001,day:2,count:18}]

//输出
9
```



### print - 打印输出

| 参数名称 | 类型   | 含义           | 默认值 |
| -------- | ------ | -------------- | ------ |
| text     | String | 需要输出的文本 | 可空   |

输出`text`到指令所在位置，支持 `\` 转义

示例：

```js
// 指令：{$.print("{}")}
// 输出：{}
```

