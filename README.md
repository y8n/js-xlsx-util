# js-xlsx-util
A simple utility collection for [js-xlsx](https://github.com/SheetJS/js-xlsx).

## Installation
With [npm](https://www.npmjs.com/package/js-xlsx-util):

```
npm install js-xlsx-util --save
```

## Usage

```
var _ = require('js-xlsx-util');
```
## Methods
### Parse Methods
对`xlsx`文件进行读取和解析的方法

- `_.readFile(xlsxPath, noCache)`：读取文件，自动执行`_.formatCell`方法，对workbook进行格式化。默认进行缓存，第二个参数设置为`true`的话表示不进行缓存

	当文件过大的时候，每次多去都会比较耗时，如果源文件不会经常变动的话，建议使用缓存。
- `_.formatCell(workbookOrWorksheet)`：对`workbook`或`worksheet`进行格式化，格式化之后的单元格对象会新增`col`列和`row`行属性，同时也会增加`top(gap)`,`right(gap)`,`bottom(gap)`,`left(gap)`四个**方法**用于获取间隔`gap`的单元格，默认是相邻的。

	```
	cellA5.top() === cellA4
	cellA5.bottom(2) === cellA7
	cellC3.left() === cellB3
	cellD4.right(4) === cellI4.left()
	```
- `_.getPrevCol(colOrCell)`：获取某一列的上一列序号，也可以传入一个**格式化之后**的单元格

	```
	_.getPrevCol('DF') // DE
	var formatCell = {
		t:'n',
		v:1234,
		col:'SA' // 格式化之后才有这个属性
	}
	_.getPrevCol(formatCell) // 'QZ'
	```
- `_.getNextCol(colOrCell)`：获取某一列的下一列序号，用法同`_.getPrevCol(colOrCell)`
- `_.each(worksheet, iterator)`：遍历`sheet`中的所有单元格，迭代器的参数是单元格位置，属性和`sheet`

	```
	_.each(worksheet,function(key,value,sheet){
		// key = 'SA12'
		// value = {t:'n',v:123 ... }
	});
	```
- `_.filter(worksheet, iterator)`：遍历`sheet`并进行过滤，根据`iterator`函数的执行结果确定是否符合条件，返回符合条件的单元格的数组

	```
	var result = _.filter(worksheet,function(key,value,sheet){
		return value.t === 'n'; // 过滤所有的数字单元格
	});
	// result [{
		t:'n',
		...
	},{
		t:'n',
		...
	}]
	```
- `_.formatKey(key)`：对单元格的位置（键）进行格式化，返回一个包含行和列的对象

	```
	_.formatKey('BS23') // {row:23,col:'BS'}
	```

### Writing Methods
输出文件的时候相关工具方法

- `_.writeFile(filepath, workbook)`：输出文件，和`xlsx.writeFile`类似，参数位置不太一样。
- `_.buildRef(worksheet)`：为`sheet`生成`!ref`属性值，因为在使用`xlsx`输出文件的时候如果没有该属性是无法生成表格的。*如果已经有了`!ref`属性，则不再生成*
- `_.cell(value)`：根据`value`的类型生成单元格对象

	```
	_.cell(234)    // {t:'n', v: 234}
	_.cell('str')  // {t:'s', v: 'str'} 
 	```
- `_.addRow(worksheet, rowObj)`：为`sheet`中添加一行，会自动根据`!ref`（如果没有的话使用`_.buildRef()`生成）自动在最后一行之后追加新的一行，可以指定列及其对应的值

	```
	var worksheet = {
		A1:...
		B1:...
	};
	_.addRow(worksheet, {
		A:'str',
		B:123,
		C:{t:'n',v:890}
	});
	console.log(worksheet); 
	/*
	{
		A1:...
		B1:...
		A2:{t:'s',v:'str'},
		B2:{t:'n',v:123},
		C2:{t:'n',v:890}
	}
	*/
	```
- `_.addWorkSheet(workbook, sheetName, worksheet)`：将`sheet`添加到`workbook`中

	```
	var workbook = {};
	_.addWorkSheet(workbook,'Sheet1', worksheet);
	console.log(workbook);
	/*
	{
		Sheets:{
			'Sheet1':worksheet
		},
		SheetNames:['Sheet1']
	}
	*/
	```
### Others
其他方法

- `_.isEqual(valA,valB)`：判断两个值是否一样，经过转化字符串，然后去除首尾空格，因为有些单元格中首尾是有空格的。
- `_.isUndefined(value)`：判断一个值是否是`undefined`
- `_.isDefined(value)`：判断一个值是否是定义的

## License
MIT