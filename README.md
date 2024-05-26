# VBA-UFunction
VBA函数库
```VB	
数组-------------------------------------------------------------------------------------------------------------------------------------
*适用于所有数组函数*：索引Index参数可以使用@修饰符 表示从头数第n个行列 例如ArrGetRegion(Array(1, 2, 3), 1, 1)->[2]   ArrGetRegion(Array(1, 2, 3), 1@, 1)->[1]
Let Titles(ParamArray TitleNames(), ByRef TitleIndexs As Variant) 缓存标题，将标题字段转成数字输出 例子：Titles("a", "b", "c") = Array(1, 2, 3)
Get Titles(ParamArray TitleNames()) As Variant 取出缓存标题 返回数组  T = Titles("a", "b", "c")->[1, 2, 3]
Get Title() As Object 返回缓存标题字典 利用这个取单个标题  Title!a -> 1  Title!b -> 2
ArrCache(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) 缓存数组属性，可以对其赋值取值操作，支持一维和二维
   ArrCache = arr 赋值整个数组
   ArrCache(RowIndex) = arr 修改二维数组的RowIndex行1列开始的值  或 修改一维数组从RowIndex开始的值
   ArrCache(, ColumnIndex) = arr 修改二维数组的1行ColumnIndex列开始的值
   ArrCache(RowIndex, ColumnIndex) = arr 修改RowIndex行ColumnIndex列开始的值 arr一维则竖着写入
   arr = ArrCache 取整个数组
   arr = ArrCache(RowIndex) 取二维数组一行 返回一维数组 或 取一维数组一个值
   arr = ArrCache(RowIndex数组) 取二维数组多行 返回二维数组 或 取一维数组多个值的数组 返回一维数组
   arr = ArrCache(, ColumnIndex) 取二维数组一列 返回一维数组
   arr = ArrCache(, ColumnIndex数组) 取二维数组多列 返回二维数组
   arr = ArrCache(RowIndex, ColumnIndex) 取二维数组一个值
   arr = ArrCache(RowIndex数组, ColumnIndex) 取ColumnIndex列里的RowIndex索引的多个值 返回一维数组
   arr = ArrCache(RowIndex, ColumnIndex数组) 取RowIndex行里的ColumnIndex索引的多个值 返回一维数组
   arr = ArrCache(RowIndex数组, ColumnIndex数组) 取RowIndex行ColumnIndex列索引相交的值 返回二维数组
ArrCache1 , ArrCache2 , ArrCache3 多个缓存数组
ArrBlend(ByRef arrC, Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) 数组区域复合操作 参照ArrCache

ArrGetValue(arr, ByVal RowCount, Optional ByVal ColumnCount, Optional EmptyContent = "") As Variant 数组取值操作，按元素第RowCount,ColumnCount个取,超出界限返回EmptyContent
不是数组时永远返回arr,数组元素数量为1时永远返回这个元素，数组为一行数组时只计算ColumnCount RowCount永=1，数组为一列或一维数组时只计算RowCount ColumnCount永=1

ArrGetValueCache(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
数组取值操作 同ArrGetValue 不同的是arr,EmptyContent写入函数缓存中 减少计算加快读取速度
WriteArr=True时写入arr缓存 WriteArr=False时传入RowCount,ColumnCount读取缓存数组
设置缓存数组示例：ArrGetValueCache WriteArr:=True, arr:=arr, EmptyContent:=""
读取缓存数组示例：v = ArrGetValueCache(i, j)
ArrGetValueCache1 , ArrGetValueCache2 , ArrGetValueCache3 , ArrGetValueCache4 , ArrGetValueCache5

ArrayDynamic(Optional ByRef v) As Variant 一维动态数组 传参则添加，不传参则取值或初始化
ArrayDynamic1 , ArrayDynamic2, ArrayDynamic3 多个ArrayDynamic
ArrayDynamic2D(ParamArray v()) As Variant 二维动态数组 传多个参数添加一行，不传参则取值或初始化
ArrayDynamic2D1 , ArrayDynamic2D2 , ArrayDynamic2D3  多个ArrayDynamic2D
ArrTranspose(ByRef arr) As Variant 数组转置
ArrFlip(arr) As Variant  数组翻转
ArrTo2D(ByRef arr1D, ByVal DCount As Long) As Variant 一维数组转二维数组
Arr2DTo1D(ByRef arr2D, Optional RowFirst As Boolean = True) As Variant 二维数组转一维数组
ArrF_T(ByRef arr, Optional ColumnCount = 0) As Variant 假数组变真数组  ColumnCount =0取最大列 >0使用ColumnCount作为列数超出被截去 <0按第一个元素的数量为列数
ArrF_T_LIndexToUIndex(ByRef arr) As Variant 假数组变真数组 保留数组上下标 *数组上标必须一致*
ArrFlatten_Single(ParamArray arr()) As Variant  展平数组(一维化) 单层
ArrFlatten(ParamArray arr()) As Variant  展平数组(一维化) 递归
Arr2DFlatten(ByRef arr2D, ByVal ColumnIndex) As Variant 二维数组内含有数组的情况,将对应的列复制多行展开
ArrMergeRow(ByVal arr) As Variant  合并数组，上下合并
ArrMergeRowParam(ParamArray arr()) As Variant 合并数组，上下合并(多参数)
ArrMergeColumn(ByVal arr) As Variant 合并数组，左右合并
ArrMergeColumnParam(ParamArray arr()) As Variant 合并数组，左右合并(多参数)

ArrCopyElement(ByRef arr, ParamArray ArrEleCount()) As Variant 一维数组 复制元素 ArrEleCount为对应arr大小的数量数组 ArrCopyElement([1,2,3],[2,3])->[1,1,2,2,2,3]
ArrCopyElement2(ByRef arr, ArrCopyIndex, ArrCopyCount) As Variant 一维数组 复制元素 ArrCopyIndex位置对应的复制ArrCopyCount个 ArrCopyElement2([1,2,3],[2,3],[2,3])->[1,2,2,3,3,3]
ArrCopyColumn(ByRef arr2D, ParamArray ArrEleCount()) As Variant 复制整列 ArrEleCount为对应arr2D列数量的数量数组
ArrCopyColumn2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant 复制整列 ArrCopyIndex位置对应的复制ArrCopyCount个
ArrCopyRow(ByRef arr2D, ParamArray ArrEleCount()) As Variant 复制整行 ArrEleCount为对应arr2D行数量的数量数组
ArrCopyRow2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant 复制整行 ArrCopyIndex位置对应的复制ArrCopyCount个

ArrInsert(ByRef arr, Optional ByVal Index, Optional ByVal EleCount As Long = 1, Optional EleCopy As Boolean = False) As Variant 一维数组 插入一个空值或多个空值 EleCopy=True复制插入
ArrInsertColumn(ByRef arr2D, Optional ByVal ColumnIndex, Optional ByVal ColumnCount As Long = 1, Optional EleCopy As Boolean = False) As Variant 数组 插入一列或多列 EleCopy=True复制插入
ArrInsertRow(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal RowCount As Long = 1, Optional EleCopy As Boolean = False) As Variant 数组 插入一行或多行 EleCopy=True复制插入
ArrGetIndex(ByRef arr, Optional GetRowIndex As Boolean = True) As Variant() 数组 取索引
ArrRemoveRegion(ByRef arr, ByRef Index, Optional ByVal Count = 1) As Variant 一维数组 删除一个元素或多个元素
ArrRemoveColumn(ByRef arr2D, ByRef Index, Optional ByVal ColumnCount = 1) As Variant 数组 删除一列或多列
ArrRemoveColumns(ByRef arr2D, ParamArray arrIndex()) As Variant 数组 删除一列或多列 多参数
ArrRemoveRow(ByRef arr2D, ByRef Index, Optional ByVal RowCount = 1) As Variant 数组 删除一行或多行
ArrRemoveRows(ByRef arr2D, ParamArray arrIndex()) As Variant 数组 删除一行或多行 多参数
ArrGetRow(ByRef arr2D, ByRef Index, Optional ByVal RowCount = 1, Optional Expansion As Boolean = False) As Variant 数组取整行 一行为一维数组 RowCount=0取到最后行
ArrGetRows(ByRef arr2D, ByVal arrIndex) As Variant  数组取多行到二维数组
ArrGetColumn(ByRef arr2D, ByRef Index, Optional ByVal ColumnCount = 1, Optional Expansion As Boolean = False) As Variant 数组取整列 一列为一维数组 ColumnCount=0取到最后列
ArrGetColumns(ByRef arr2D, ByVal arrIndex) As Variant  数组取多列到二维数组
ArrGetRegion2D(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
     Optional ByVal Height = 0, Optional ByVal Width = 0, Optional Expansion As Boolean = False) As Variant  数组取区域 索引加大小 二维数组
ArrGetRegion2D_To(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
        Optional ByVal RowIndex2, Optional ByVal ColumnIndex2, Optional Expansion As Boolean = False) As Variant  数组取区域 索引到索引 二维数组
ArrGetRegion(ByRef arr, Optional ByVal Index, Optional ByVal Count = 0, Optional Expansion As Boolean = False) As Variant 数组取区域 一维数组
ArrGetRegion_To(ByRef arr, Optional ByVal Index, Optional ByVal IndexTo, Optional Expansion As Boolean = False) As Variant 数组取区域 索引到索引 一维数组
ArrSizeExpansion(ByRef arr, ByRef RowCount, Optional ByRef ColumnCount, Optional FillValue = Empty) 数组扩充大小  **数组下标变1**

ArrSizeExpansionEx(ByRef arr, ByRef RowCount, ByRef ColumnCount, Optional FillValue = Empty)数组扩充大小 满足矩阵运算扩充  **数组下标变1**
不是数组时填充所有元素,数组元素数量为1时填充所有元素，数组为一行数组时填充所有列，数组为一列或一维数组时填充所有行

ArrSizeExpansion2(ByRef arr, ByRef ArrSizeCount, Optional FillValue = Empty) 数组扩充大小 所有数组都变成一维  **数组下标变1**  复杂数组计算用
ArrIndexExpansion(ByRef arr, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional FillValue = Empty) 数组扩充索引，当索引超出数组时会被扩充
Arr2DSetArr2D(ByRef arrL, ByRef arrR, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False)  数组赋值到数组 二维
Arr2DSetValues(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values()) 多个值按RowIndexArr与ColumnIndexArr交叉位置依次赋值到数组  从上到下一行一行写入 二维
Arr2DSetValues_LtoR(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values()) 多个值按RowIndexArr与ColumnIndexArr交叉位置依次赋值到数组  从左到右一列一列写入 二维
ArrSetValues(ByRef arr1D, ByRef IndexArr, ParamArray Values()) 多个值按IndexArr位置依次赋值到数组 一维
ArrSetEntireColumnValues(ByRef arr2D, ByRef ColumnIndexArr, ParamArray Values()) 赋值到数组一整列 多值对应多列 二维
ArrSetEntireRowValues(ByRef arr2D, ByRef RowIndexArr, ParamArray Values()) 赋值到数组一整行 多值对应多行 二维
ArrSetArr(ByRef arrL, ByRef arrR, Optional ByVal Index, Optional Expansion As Boolean = False)  数组赋值到数组 一维
ArrSetColumn(ByRef arrL2D, ByRef arrR, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False) 数组赋值到数组一列
ArrSetRow(ByRef arrL2D, ByRef arrR, Optional ByVal RowIndex, Optional Expansion As Boolean = False) 数组赋值到数组一行
ArrFromIndex(arr, arrIndex) As Variant   按索引数组顺序取回数组值，用来还原排序结果
ArrFromBoolea(arr, arrBoolea) As Variant 按布尔数组条件=True取回数组值，用来筛选数组
ArrRandSort(ByVal arr) As Variant  数组随机排序
ArrSort2D(arr, Index, Optional Order As Boolean = True) As Variant  二维数组稳定排序
ArrSort2Ds(arr, Indexs, Optional Orders = True) As Variant 二维数组多列稳定排序
ArrSort1D(arr, Optional Order As Boolean = True) As Variant 一维数组稳定排序
ArrSort(arr, Optional Order As Boolean = True) As Variant 一维数组稳定排序 返回索引，Order=True 升序排序
例子：排序arr二维数组
ArrColumns = ArrGetColumn(arr, 1)  取得arr排序列
arrIndex = ArrSort(ArrColumns)  对排序列排序返回排序索引
arrOrder = ArrFromIndex(arr, arrIndex) 根据索引数组取出有序数组
ArrSortNext(arr, Indexs, Optional Order As Boolean = True) As Variant  对数组多次排序
例子：按1,2列排序
arrIndex = ArrSort(ArrGetColumn(arr, 1)) 第一次排序
arrIndex = ArrSortNext(ArrGetColumn(arr, 2), arrIndex) 第2次排序
arrorder = ArrFromIndex(arr, arrIndex) 返回结果
ArrCustomSort2D(arrValue, arrKey, Index, Optional IsLike As Boolean = False) As Variant  二维数组自定义排序
ArrCustomSort(arrValue, arrKey, Optional IsLike As Boolean = False)  自定义排序  CustomSort(排序数组, 自定义序列, Like匹配) 返回索引数组
ArrInInterval(ByVal arrInterval, Number) As Long 查找Number在arrInterval里的区间位置 位置索引从LBound(arrInterval)到UBound(arr)+1 arrInterval必须升序顺序
ArrInIntervalEqual(ByVal arrInterval, Number) As Long 查找Number在arrInterval里的区间位置 含等于 位置索引从LBound(arrInterval)到UBound(arr)+1 arrInterval必须升序顺序
ArrFindLessIndex(arr_Small, V_Large, Optional ByVal Start) As Long 查找小于v的索引
ArrFindLessIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long 查找小于v的索引 反向
ArrFindLessEqualIndex(arr_Small, V_Large, Optional ByVal Start) As Long 查找小于等于v的索引
ArrFindLessEqualIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long 查找小于等于v的索引 反向
ArrFindGreaterIndex(arr_Large, V_Small, Optional ByVal Start) As Long 查找大于v的索引
ArrFindGreaterIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long 查找大于v的索引 反向
ArrFindGreaterEqualIndex(arr_Large, V_Small, Optional ByVal Start) As Long 查找大于等于v的索引
ArrFindGreaterEqualIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long 查找大于等于v的索引 反向
ArrFindLikeIndex(arr, v, Optional ByVal Start) As Long  查找对应值索引 Like
ArrFindLikeIndexRev(arr, v, Optional ByVal Start) As Long  查找对应值索引反向 Like
ArrFindNotLikeIndex(arr, v, Optional ByVal Start) As Long 查找对应值索引 Not Like
ArrFindNotLikeIndexRev(arr, v, Optional ByVal Start) As Long 查找对应值索引反向 Not Like
ArrFindIndex(arr, v, Optional ByVal Start) As Long  查找对应值索引
ArrFindIndexRev(arr, v, Optional ByVal Start) As Long  查找对应值索引反向
ArrFindNotIndex(arr, v, Optional ByVal Start) As Long 查找不等于的值索引
ArrFindNotIndexRev(arr, v, Optional ByVal Start) As Long 查找不等于的值索引反向
ArrFindRegIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long  查找对应值索引 正则
ArrFindRegIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long  查找对应值索引 正则 反向
ArrFindRegNotIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long 查找对应值索引 不满足正则
ArrFindRegNotIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long 查找对应值索引 不满足正则 反向
ArrFindIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant 二维数组查找索引 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrFindNotIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant 二维数组查找索引 不等于 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrFindLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant 二维数组查找索引 Like查找 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrFindNotLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant 二维数组查找索引 Not Like查找 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrFindRegIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant 二维数组查找索引 正则 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrFindRegNotIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant 二维数组查找索引 不满足正则 找到返回Array(RowIndex, ColumnIndex) 找不到返回空数组
ArrValid_InError(arr) As Boolean    数组数据效验 有错误返回True
ArrValid_NumericAll(arr, Optional InEmpty As Boolean = True, Optional IsStr As Boolean = True) As Boolean  数组数据效验 全部是数字返回True
ArrValid_DateAll(arr, Optional IsStr As Boolean = True) As Boolean  数组数据效验 全部是日期返回True
ArrValid_Reg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean 数组数据效验满足一个 正则 匹配返回True
ArrValid_RegAll(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean  数组数据效验满足全部 正则 全部匹配返回True
ArrValid_Repeat(arr) As Boolean 数组数据效验是否有重复 重复返回True
ArrValid_Incremental(ParamArray arr()) As Boolean 数组数据效验是否递增序列
ArrValid_IncrementalEqual(ParamArray arr()) As Boolean 数组数据效验是否递增序列包含相等
ArrValid_Decrement(ParamArray arr()) As Boolean 数组数据效验是否递减序列
ArrValid_DecrementEqual(ParamArray arr()) As Boolean 数组数据效验是否递减序列包含相等
ArrFilterRepeatCount(arr, Optional CountSmall = 0, Optional CountLarge = 1.79769313486231E+308, Optional CompareMode As CompareMethod = BinaryCompare) As Variant 筛选 重复次数  ,*返回筛选索引*
ArrFilterRangeInside(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 筛选 区间 内部 ,*返回筛选索引*
ArrFilterRangeExternal(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 筛选 区间 外部 ,*返回筛选索引*
ArrFilterGreater(arr_Large, V_Small) As Variant 筛选 大于V_Small的值 ,*返回筛选索引*
ArrFilterGreaterEqual(arr_Large, V_Small) As Variant 筛选 大于等于V_Small的值 ,*返回筛选索引*
ArrFilterLess(arr_Small, V_Large) As Variant 筛选 小于V_Large的值 ,*返回筛选索引*
ArrFilterLessEqual(arr_Small, V_Large) As Variant 筛选 小于V_Large的值 ,*返回筛选索引*
ArrFilter(arr, ByVal arrv) As Variant   筛选 ,**返回筛选索引**
ArrFilterLike(arr, ByVal arrv) As Variant  筛选like匹配 ,**返回筛选索引**
ArrFilterReg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant  筛选正则匹配 ,**返回筛选索引**
ArrFilterRemove(arr, ByVal arrv) As Variant  筛选排除 ,**返回筛选索引**
ArrFilterLikeRemove(arr, ByVal arrv) As Variant  筛选like排除 ,**返回筛选索引**
ArrFilterRegRemove(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant  筛选正则排除 ,**返回筛选索引**
ArrDistinct(arr) As Variant 去重 保留开头一个值
ArrDistinctIndex(arr) As Variant 去重，返回索引 保留开头索引
ArrDistinctIndexRev(arr) As Variant 去重，返回索引 保留最后索引
ArrLBoundToN_1D(arr, Optional StartLBound = 1) As Variant 数组下标变StartLBound 一维数组
ArrLBoundToN_2D(arr, Optional StartLBound1 = 1, Optional StartLBound2 = 1) As Variant 数组下标变StartLBound1,StartLBound2 二维数组
ArrMap(ByVal arr, EvaluateStr) As Variant  Evaluate修改数组 $表示当前值
ArrReplace(ByRef arr, FindValueArr, ReplaceValue) As Variant 数组替换数组所有完整元素 FindValueArr支持单值或数组
ArrErrorClear(ByRef arr, Optional EmptyValue = Empty) As Variant 清除数组错误值
ArrIsValid(ByRef arr) As Boolean  数组是否有效
ArrDimension(ByRef arr) As Long  数组维度
ArrCount(ByRef arr) As Long  数组元素个数
ArrCountRow(ByRef arr) As Long  数组行数
ArrCountColumn(ByRef arr) As Long 数组列数
ArrCountRowAndColumn(arr, ByRef RowCount, ByRef ColumnCount) 同时计算行列数量用变量RowCount,ColumnCount接收返回值，一维数组ColumnCount=1，不是数组RowCount=ColumnCount=1
ArrCountElement(ByVal arr) As Variant 数组标记元素个数，返回总数数组
ArrCountMergeElement(ByRef arr, Optional EmptyContent = "") As Variant 数组标记合并单元格形式元素个数，返回个数数组
ArrBetween(l, u) As Variant()  创建范围整数数组
ArrCreate(Number, Optional Number2 = 0, Optional FillValue = Empty) As Variant() 创建数组
ArrCreateRand(l, r, RowCount, Optional ColumnCount = 0) As Variant() 创建随机数数组
ArrCreateRandDic(l, r, RowCount, Optional ColumnCount = 0) As Variant() 创建随机数数组 不重复随机数
ArrFillDown(ByRef arr, Optional index = 1, Optional EmptyContent = "") As Variant 空值向下填充  arr一维或二维数组 index二维数组列索引  EmptyContent当做空值的内容
ArrFillUp(ByRef arr, Optional index = 1, Optional EmptyContent = "") As Variant  空值向上填充  arr一维或二维数组 index二维数组列索引  EmptyContent当做空值的内容
ArrPerspectiveRev(ByRef arrH, ByRef arrV, Optional ByRef arrRegion2D = "") As Variant
  逆透视 arrH竖标题(可以是多列)  arrV横标题(只能一行) arrRegion2D数据区域(行大小必须是arrH行数 列大小必须是arrV数量)
ArrPerspective(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant 透视 行列交叉有重复数据时取最后一值 arr2D二维表  VIndex变横标题的列  DataIndex变数据区域的列
ArrPerspective_Repeating(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant 透视 行列交叉有重复数据时写多行 arr2D二维表  VIndex变横标题的列  DataIndex变数据区域的列
ArrGroupSum(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrSumIndex) As Variant 分类求和 arr2D二维表 ArrGroupIndex分组列索引支持数组 ArrSumIndex求和列索引支持数组
ArrGroupCount(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrCountIndex, Optional NoEmpty As Boolean = True) As Variant 分类计数 arr2D二维表 ArrGroupIndex分组列索引支持数组 ArrCountIndex计数列索引支持数组 NoEmpty = True计算非空值数量
ArrGroupJoin(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrJoinIndex, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
    分类拼接字符串 arr2D二维表 ArrGroupIndex分组列索引支持数组 ArrJoinIndex求和列索引支持数组 Delimiter分隔符 OmittedEmpty忽略空字符串
ArrGroup_Class(ByRef arr2D, ByVal ArrClassIndex) As Variant 数组分组 按类别 ArrClassIndex分类索引支持数组  返回数组套数组的分组
ArrGroup_Find_First(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant 数组分组 按查找内容为分组界限 界限放在分组*首行*  FindIndex索引列 FindValue查找内容  返回数组套数组的分组
ArrGroup_Find_Last(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant  数组分组 按查找内容为分组界限 界限放在分组*末尾*  FindIndex索引列 FindValue查找内容  返回数组套数组的分组
ArrGroup_Differ(ByRef arr2D, ByVal ArrDifferIndex) As Variant 数组分组 按列上下内容不用为分组界限  ArrDifferIndex不同的列索引支持数组  返回数组套数组的分组
ArrGroup_Number_Column(ByRef arr2D, ByVal Number) As Variant 数组分组 按列数量  number数量  返回数组套数组的分组
ArrGroup_Number(ByRef arr2D, ByVal number) As Variant数组分组 按数量  number数量  返回数组套数组的分组
ArrGroup_Row_First(ByRef arr2D, ByVal ArrRowIndex) As Variant  数组分组 按行索引为界限分组  界限放在分组*首行* ArrRowIndex行索引支持数组  返回数组套数组的分组
ArrGroup_Row_Last(ByRef arr2D, ByVal ArrRowIndex) As Variant  数组分组 按行索引为界限分组  界限放在分组*末尾* ArrRowIndex行索引支持数组  返回数组套数组的分组
ArrGroup_Interval(ByVal arr2D, ByVal ColumnIndex, ParamArray ArrInterval()) As Variant 数组分组 按数值区间分组分组  小于不等于被放一组 ArrInterval区间数组  返回数组套数组的分组
ArrGroup_Interval_Equal(ByVal arr2D, ByVal ColumnIndex, ParamArray ArrInterval()) As Variant 数组分组 按数值区间分组分组  小于等于被放一组 ArrInterval区间数组  返回数组套数组的分组
ArrGroup_CustomClass(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant 数组分组 按自定义分类 不匹配的放最后一组 arrCustomValue匹配数组  返回数组套数组的分组
ArrGroup_CustomClass_Like(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant 数组分组 按自定义分类Like匹配  不匹配的放最后一组 arrCustomValue匹配数组  返回数组套数组的分组
ArrGroup_CustomClass_Reg(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomPattern()) As Variant 数组分组 按自定义分类 正则匹配  不匹配的放最后一组 arrCustomValue匹配数组  返回数组套数组的分组
ArrGroupAgg(ByRef ArrGroup, Optional OmittedNoneArg As Boolean = True, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, _
        Optional ByRef C1 As GroupAggregateMethod = Group_None, C2, C3,... C46 ) As Variant
       分组聚合函数  ArrGroup分组函数返回的数组套数组  OmittedNoneArg没有写Cn参数的列是否省略 Delimiter拼接字符分隔符 OmittedEmpty拼接字符串是否忽略空值
       C1-C46代表数组的1-46列 采用C1:=Group_Sum方式传入聚合模式GroupAggregateMethod  C1-C46传入正数取第N行传入负数取倒数第N行

ArrGroupAgg2(ByRef ArrGroup, ArrGroupIndex, ArrAggregateMethod, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
分组聚合函数 支持一列多种聚合 ArrGroup分组函数返回的数组套数组 ArrGroupIndex聚合列 ArrAggregateMethod对应的聚合模式 Delimiter拼接字符分隔符 OmittedEmpty拼接字符串是否忽略空值

ArrUnions(ParamArray arr()) As Variant 并集多个  取多个数组元素
ArrUnions_Distinct(ParamArray arr()) As Variant 并集多个  去重
ArrUnions_Sort(ParamArray arr()) As Variant 并集多个  排序
ArrUnions_DistinctSort(ParamArray arr()) As Variant 并集多个  去重排序
ArrUnion(ByRef arr1, ByRef arr2) As Variant 并集 取两个数组元素
ArrUnion_Distinct(ByRef arr1, ByRef arr2) As Variant 并集 去重
ArrUnion_Sort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant 并集 排序
ArrUnion_DistinctSort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant 并集 去重排序
ArrIntersects(ParamArray arr()) As Variant  交集多个  取多个数组元素
ArrIntersects_Distinct(ParamArray arr()) As Variant 交集多个  去重
ArrIntersects_arr1(ParamArray arr()) As Variant 交集多个 取第一个数组元素
ArrIntersects_arr1_Index(ParamArray arr()) As Variant 交集多个 取第一个数组元素
ArrIntersect(ByRef arr1, ByRef arr2) As Variant 交集 取两个数组元素
ArrIntersect_Distinct(ByRef arr1, ByRef arr2) As Variant 交集 去重
ArrIntersect_arr1(ByRef arr1, ByRef arr2) As Variant 交集 取arr1元素
ArrIntersect_arr2(ByRef arr1, ByRef arr2) As Variant 交集 取arr2元素
ArrIntersect_arr1_Index(ByRef arr1, ByRef arr2) As Variant 交集 取arr1索引
ArrIntersect_arr2_Index(ByRef arr1, ByRef arr2) As Variant 交集 取arr2索引
ArrExcepts_Single(ParamArray arr()) As Variant 差集多个  取多个数组元素(保留数组中其他数组没有的元素)[1,2,3,4,5,5][1,2,3][2,3,4,6]->[5,5,6]
ArrExcepts_RemoveAllIntersect(ParamArray arr()) As Variant 差集多个  取多个数组元素(去除所有数组都包含的元素)[1,2,3,4,5,5][1,2,3][2,3,4,6]->去除共有元素2,3得到[1,4,5,5,1,4,6]
ArrExcepts_arr1(ParamArray arr()) As Variant 差集多个  取第一个元素
ArrExcepts_arr1_Index(ParamArray arr()) As Variant 差集多个 取第一个数组元素索引
ArrExcept(ByRef arr1, ByRef arr2) As Variant 差集 取两个数组元素
ArrExcept_arr1(ByRef arr1, ByRef arr2) As Variant 差集 取arr1元素
ArrExcept_arr2(ByRef arr1, ByRef arr2) As Variant 差集 取arr2元素
ArrExcept_arr1_Index(ByRef arr1, ByRef arr2) As Variant 差集 取arr1索引
ArrExcept_arr2_Index(ByRef arr1, ByRef arr2) As Variant 差集 取arr2索引
ArrTitleToIndex(ByRef arrTitle, ByRef arrOrder) As Variant  arrTitle(一维)按arrOrder(一维)返回对应的顺序的标题索引数组,返回的数组为arrTitle索引不匹配位置返回(LBound-1),返回的数组大小与arrOrder相同
ArrIFs(ParamArray Calculates()) As Variant 数组IFs判断计算 ArrIFs(条件,值,条件,值,否则值)
ArrBoolea_And(ParamArray Calculates()) As Variant 数组布尔且计算
ArrBoolea_Or(ParamArray Calculates()) As Variant 数组布尔或计算
ArrBoolea_Not(ByVal arr) As Variant 数组布尔非计算
ArrComp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 数组区间比较计算 内部
ArrComp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 数组区间比较计算 外部
ArrComp_Like(ByVal arr, ByVal arr2) As Variant 数组比较Like计算
ArrComp_NotLike(ByVal arr, ByVal arr2) As Variant 数组比较Not Like计算
ArrComp_Equal(ByVal arr, ByVal arr2) As Variant 数组比较等于计算
ArrComp_NotEqual(ByVal arr, ByVal arr2) As Variant 数组比较不等于计算
ArrComp_Size(ByVal arr_Large, ByVal arr_Small) As Variant 数组比较大小计算
ArrComp_SizeEqual(ByVal arr_Large, ByVal arr_Small) As Variant 数组比较大小包含等于计算
ArrMath_Add(ParamArray Calculates()) As Variant 数组加法计算
ArrMath_Sub(ParamArray Calculates()) As Variant 数组减法计算
ArrMath_Multipli(ParamArray Calculates()) As Variant 数组乘法计算
ArrMath_Division(ParamArray Calculates()) As Variant 数组除法计算
ArrMath_Power(ParamArray Calculates()) As Variant 数组乘方计算
ArrMath_Join(ParamArray Calculates()) As Variant 数组连接计算
ArrMath_Round(ByVal arr, number, Optional ColumnIndex = 1) As Variant 数组四舍五入
ArrMath_Val(ByVal arr, Optional ColumnIndexArr = 1) As Variant
ArrMath_Abs(ByVal arr, Optional ColumnIndexArr = 1) As Variant 数组绝对值Abs
ArrMath_Format(ByVal arr, Pormat, Optional ColumnIndex = 1) As Variant 数组Format
ArrStr_Ucase(ByVal arr, Optional ColumnIndexArr = 1) As Variant 数组转大写
ArrStr_Lcase(ByVal arr, Optional ColumnIndexArr = 1) As Variant 数组转小写

ArrStr_Split(ByVal arr, Delimiter, Optional ColumnIndexArr = 1) As Variant  数组循环拆分字符串 返回数组套数组
ArrStr_Replace(ByVal arr, FindStr, ReplaceStr, Optional ColumnIndex = 1) As Variant 数组替换
ArrStr_ReplaceAll(ByVal arr, FindStr, ReplaceStr) As Variant 数组替换数组所有内容
ArrStr_RegexSearch(ByVal arr, Pattern, Optional RegIndex = 0, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant 数组正则取值
 
ArrStr_RegexSearchs(ByVal arr, Pattern, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant 数组正则取所有值返回数组套数组
        
ArrStr_RegexCount(ByVal arr, Pattern, Optional ByVal ColumnIndexArr = 1, Optional ByVal NumberAdd = 0, _
         Optional ByRef ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant 数组正则返回匹配数量
         
ArrStr_RegexReplace(ByVal arr, Pattern, ReplaceStr, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant 数组正则替换
 
ArrStr_Mid(ByVal arr, start, Optional length, Optional ColumnIndex = 1) As Variant 数组MID
ArrDate_DateSub(Interval, Date1, Date2) As Variant 数组日期差值 参照DateDiff
ArrDate_Year(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取年
ArrDate_Month(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取月
ArrDate_Day(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取天
ArrDate_Weekday(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取星期
ArrTime_Hour(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取小时
ArrTime_Minute(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取分钟
ArrTime_Second(ByVal arr, Optional ColumnIndex = 1) As Variant 数组取秒
ArrSerialNumber(ByVal arr, Optional ColumnIndex = 1, Optional StartNumber = 1) As Variant 加序号 传入数组返回1++序号
ArrSerialNumberCalssSelf(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant 加序号 按数组不同内容 相同内容1++ 返回1++序号
ArrSerialNumberCalss(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant 加序号 按数组不同内容1++ 返回1++序号
ArrMaxIndex(ByRef arr, Optional ColumnIndex = 1, Optional Front As Boolean = True) As Long 数组取最大值索引 ColumnIndex 二维数组列索引  Front = True 最前的索引
ArrMinIndex(ByRef arr, Optional ColumnIndex = 1, Optional Front As Boolean = True) As Long 数组取最小值索引 ColumnIndex 二维数组列索引  Front = True 最前的索引
ArrSum(ByRef arr) As Double  数组求和
ArrMax(ByRef arr) As Double  数组求最大值
ArrMin(ByRef arr) As Double  数组求最小值
ArrCountNoEmpty(ByRef arr) As Double 数组计算非空值数量
ArrSumColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant 数组按列求和
ArrSumRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant 数组按行求和
ArrMaxColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant 数组按列求最大值
ArrMaxRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant 数组按行求最大值
ArrMinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant 数组按列求最小值
ArrMinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant 数组按行求最小值
ArrJoinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant 数组按列拼接字符串
ArrJoinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant 数组按行拼接字符串
ArrCountNoEmptyColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant 数组按列计算非空值数量
ArrCountNoEmptyRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant 数组按行计算非空值数量
ArrCountClassColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant 数组按列计算种类数量
ArrCountClassRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant 数组按行计算种类数量
ArrAverageColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant 数组按列计算平均值  NumDigitsAfterDecimal舍入小数位数
ArrAverageRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant 数组按行计算平均值  NumDigitsAfterDecimal舍入小数位数
ArrAverage(ByRef arr, Optional NumDigitsAfterDecimal As Long = 2) As Double 数组计算求平均值  NumDigitsAfterDecimal舍入小数位数
ArrMoveUp(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant 空值移动 向上
ArrMoveDown(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant 空值移动 向下
ArrMoveLeft(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant 空值移动 向左
ArrMoveRight(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant 空值移动 向右
ArrMove(ByRef arr1D, Optional EmptyContent = "") As Variant 空值移动 一维数组 正向
ArrMoveRev(ByRef arr1D, Optional EmptyContent = "") As Variant 空值移动 一维数组 反向
ArrMove_Index(ByRef arr1D, Optional EmptyContent = "") As Variant 空值移动 一维数组 正向 返回索引
ArrMoveRev_Index(ByRef arr1D, Optional EmptyContent = "") As Variant 空值移动 一维数组 反向 返回索引
ArrScroll(ByRef arr, Index) As Variant 数组滚动 正向 Index索引滚动到开头
ArrScrollRev(ByRef arr, Index) As Variant 数组滚动 反向 Index索引滚动到末尾
ArrScroll_Index(ByRef arr, Index) As Variant 数组滚动 正向 Index索引滚动到开头 返回索引
ArrScrollRev_Index(ByRef arr, Index) As Variant 数组滚动 反向 Index索引滚动到末尾 返回索引
ArrScrollColumn(ByRef arr2D, Index) As Variant 二维数组列滚动 正向 Index索引滚动到开头
ArrScrollColumnRev(ByRef arr2D, Index) As Variant 二维数组列滚动 反向 Index索引滚动到末尾
ArrScrollColumn_Index(ByRef arr2D, Index) As Variant 二维数组列滚动  正向 Index索引滚动到开头 返回索引
ArrScrollColumnRev_Index(ByRef arr2D, Index) As Variant 二维数组列滚动 反向 Index索引滚动到末尾 返回索引
ArrCombinCon(arr, r) 组合  arr 一维数组 r抽取数量
ArrPermutCon(arr, r) 排列  arr 一维数组 r抽取数量


矩阵-------------------------------------------------------------------------------------------------------------------------------------
Matrix_Add(ParamArray Calculates()) As Variant 矩阵加法计算
Matrix_Sub(ParamArray Calculates()) As Variant 矩阵减法计算
Matrix_Multipli(ParamArray Calculates()) As Variant 矩阵乘法计算
Matrix_Division(ParamArray Calculates()) As Variant 矩阵除法计算
Matrix_Power(ParamArray Calculates()) As Variant 矩阵乘方计算
Matrix_Join(ParamArray Calculates()) As Variant 矩阵连接计算
Matrix_Comp_Equal(ByRef arr, ByRef arr2) As Variant 矩阵比较等于
Matrix_Comp_NotEqual(ByRef arr, ByRef arr2) As Variant 矩阵比较不等于
Matrix_Comp_Size(ByRef arr_Large, ByRef arr_Small) As Variant 矩阵比较大小
Matrix_Comp_SizeEqual(ByRef arr_Large, ByRef arr_Small) As Variant 矩阵比较大小包含等于
Matrix_Comp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 矩阵区间比较计算 内部
Matrix_Comp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant 矩阵区间比较计算 外部
Matrix_Comp_Like(ByRef arr, ByRef arr2) As Variant 矩阵比较Like
Matrix_Comp_NotLike(ByRef arr, ByRef arr2) As Variant 矩阵比较Not Like
Matrix_Boolea_And(ParamArray Calculates()) As Variant 矩阵布尔且计算
Matrix_Boolea_Or(ParamArray Calculates()) As Variant 矩阵布尔或计算
Matrix_Boolea_Not(ByRef arr) As Variant 矩阵布尔非计算
Matrix_IF(Expression, TruePart, FalsePart) As Variant 矩阵IF
Matrix_IFs(ParamArray Calculates()) As Variant 矩阵IFs
Matrix_Str_Mid(String1, Start, Optional Length) As Variant 矩阵Mid 矩阵参数：String1, Start, Length
Matrix_Str_Left(String1, Length) As Variant 矩阵Left 矩阵参数：String1, Length
Matrix_Str_Right(String1, Length) As Variant 矩阵Right 矩阵参数：String1, Length
Matrix_Str_InStr(StringLarge, StringSmall, Optional Start = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant 矩阵InStr 矩阵参数：StringLarge, StringSmall, Start
Matrix_Str_InStrRev(StringLarge, StringSmall, Optional Start = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant 矩阵InStr 矩阵参数：StringLarge, StringSmall, Start
Matrix_Str_Len(ByRef String1) As Variant 矩阵Len 矩阵参数：String1
Matrix_Str_Replace(Expression, Find, Replace, Optional Start = 1, Optional Count = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant 矩阵替换 矩阵参数：Expression, Find, Replace
Matrix_DateSub(Interval, Date1, Date2) As Variant 矩阵日期间隔 参照DateDiff 矩阵参数：Interval, Date1, Date2







字符串-----------------------------------------------------------------------------------------------------------------------------------
StringBuilder(Optional ByRef s) As Variant  传参则添加，不传参则取值或初始化
StringBuilder1 , StringBuilder2, StringBuilder3 多个StringBuilder
StrJoinArr2D(ByRef arr2D, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, Optional RowFirst As Boolean = True) As String 二维数组拼接
StrJoin_ArrDelimiter(ByRef arr, ParamArray ArrDelimiter()) As String 数组交错拼接
StrStrLike(str1, LikeStr) As Boolean  Like匹配
StrLeft(String1, Length) As String 支持负Length的Left
StrRight(String1, Length) As String 支持负Length的Right
StrMid(String1, ByVal Start, ByVal Length) As String 支持负Start负Length的Mid
StrMidBetween(String1, ByVal Start, Optional ByVal EndIndex = 0) As String 起始结束取字符串
StrGetLeft(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  取str左边内容，从左查找
StrGetLeftRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  取str左边内容，从右查找
StrGetRight(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  取str右边内容，从左查找
StrGetRightRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  取str右边内容，从右查找
StrGetCentre(String1, str1, str2, Optional SearchType As SearchDirection = LeftLeft) As String 取两个str中间内容
StrTrimChr(String1, Optional Chrs = " ") As String 按Chrs里的字符去除两端字符串
StrLTrimChr(String1, Optional Chrs = " ") As String 按Chrs里的字符去除左端字符串
StrRTrimChr(String1, Optional Chrs = " ") As String 按Chrs里的字符去除右端字符串
StrRepeat(ByVal string1, ByVal numberOfRepeats As Long) As String   重复字符串
StrReplaces(Expression, Finds, Replaces, Optional Counts = -1, _
      Optional Compare As VbCompareMethod = vbBinaryCompare) As String 批量替换 Finds,Replaces,Counts支持数组 StrReplaces("aabca",{"aa","a"},{"a","e"})->abce
StrReplaceChr(ByVal String1, StrKey, StrItem) As String 按StrKey里的字符 替换对应位置的StrItem  StrReplaceChr("aabbccdd","abc","123")->112233dd
StrReplacePlaceholder(ByVal String1, placeholder, ParamArray ValueStrs()) As String 替换占位符placeholder    StrReplacePlaceholder("a%b%c", "%", 1, 2) "a1b2c"
StrReplaceIndex(String1, ReplaceStr, ByVal Start, ByVal Length) As String 按索引位置替换
Str_Split(ByVal Expression, Optional Delimitre = "", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String()
    拆分字符串 支持多个分割符
Str_SplitMatch(String1, ParamArray Delimitre()) As Variant 处理 "序号=1,名称=abc,数量=1" 类型的数据，Str_SplitMatch("序号=1,名称=abc,数量=1", "序号=",",名称=",",数量=")返回数组，数组(0)是"序号="左边内容
Str_Split2D(ByVal string1, DelimitreRow, DelimitreColumn, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant 字符串拆分二维数组
StrReg_Split(ByVal Expression, ByVal Pattern As Variant, Optional ByVal ignoreCase As Boolean = True) As Variant 正则拆分
PinYin(Txt As Variant, Optional Delimiter = " ") As String  简单拼音，可以用来写拼音搜索 注：多音字和生僻字，可能不准
PinYinInitial(Txt As Variant) As String  拼音开头
StrFindSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long  编辑距离相似度算法 包含字符串顺序 查找FindStr在arr位置 Similarity为最小相似度
StrFindCosineSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long  余弦相似度算法 忽略字符串顺序 查找FindStr在arr位置 Similarity为最小相似度
StrSimilar(s1, s2) As Double  编辑距离相似度算法 判断字符串S1、S2的相似度,包含字符串顺序,相似度区间为0-100,100为完全一致
StrCosineSimilar(strA, strB) As Double  余弦相似度算法 判断字符串S1、S2的相似度,忽略字符串顺序,相似度区间为0-100,100为完全一致
StrRegexSearch( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef Index = 0, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant正则取单个值
 
StrRegexSearchs( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant()  正则取所有匹配，返回数组
 
StrRegexSearchOne( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As String  正则取第一个值
 
RegexInStr( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long  正则查找位置
 
StrRegexInStrRev( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long  正则查找位置 反向
 
StrRegexSearchSub( _
        ByRef string1, _
        ByRef Pattern, _
        Optional ByRef All As Boolean = True, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Variant() 正则取所有组匹配，返回正则里的()假二维数组
 
RegexCount( _
        ByRef string1, _
        ByRef Pattern, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Long  正则计数
 
StrRegexTest( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Boolean 正则验证
 
StrRegexReplace( _
    ByRef string1, _
    ByRef Pattern, _
    ByRef replacementString As String, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As String  正则替换
 
StrFormatter(ByVal formatString, ParamArray textArray() As Variant) As String  模版字符串 Formatter("姓名：{1},年龄：{2}","UFO",18)  返回"姓名：UFO,年龄：18"
ByteToStr(arrByte, strCharset As String) As String 流数据转成指定编码的文本 "Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
StrToByte(strText As String, strCharset As String) 文本按指定编码转为流数据 "Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
StrencodeURI(strText) As String  URL转码
StrdecodeURI(strText) As String  URL解码
StrConvert(ByVal strText As String) As String unicode字符转换成中文
StrencodeBase64(String1, Optional Charset = "") As String 字符串编码Base64
StrdecodeBase64(String1, Optional Charset = "") As String 字符串解码Base64



系统-------------------------------------------------------------------------------------------------------------------------------------
Clipboard_GetData() As String  剪贴板读取
Clipboard_SetData(strData) As Boolean  剪贴板写入
Clipboard_ClearData() As Boolean  剪贴板清空
UserName() As String  用户名
UserDomain() As String  用户的域名
ComputerName() As String  计算机名


文件-------------------------------------------------------------------------------------------------------------------------------------
TextRead(TextPath) As String  读取txt文件(ANSI编码)
TextWrite(TextPath, str) As Boolean  写入txt文件(ANSI编码)
TextAppend(TextPath, str) As Boolean 追加txt文件(ANSI编码)
TextRead2(TextPath, strCharset As String) As String  读取txt文件(自定义编码) "Unicode", "GB2312", "UTF-7", "UTF-8", "ASCII", "GBK", "Big5", "unicodeFEFF", "unicodeFFFE"
TextWrite2(TextPath, str, strCharset As String) As Boolean  写入txt文件(自定义编码)
TextAppend2(TextPath, str, strCharset As String) As Boolean  追加txt文件(自定义编码)
FileToByte(strFileName As String) As Byte() 读文件为字节数组
ByteToFile(arrByte, strFileName As String)  字节数组转文件
FolderExists(Path) As Boolean  文件夹是否存在
FolderDelete(Path) As Boolean  删除文件夹
FolderCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean  复制文件夹
FolderCreate(Path) As Boolean  创建文件夹，可以创建上级不存在的文件夹，创建多级
FolderSearch(pPath) As Variant  遍历文件夹里文件夹
FolderSearchSub(pPath) As Variant 遍历文件夹(含子文件夹)
FileExists(Path) As Boolean  文件是否存在
FileDelete(Path) As Boolean  删除文件
FileCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean 复制文件
FileSearch(pPath) As Variant 遍历文件夹里文件
FileSearchSub(pPath, Optional pMask As String = "") As Variant 遍历文件夹里文件(含子文件夹) pPath搜索起始路径，pMask如果要必填写,那得这样填写"*.xlsx",加星号


路径-------------------------------------------------------------------------------------------------------------------------------------
PathGetTemp() As String  返回临时路径
PathGetMyDocuments() As String  返回文档路径
PathGetDesktop() As String  返回桌面路径
PathBaseName(Path) As String  返回文件名，不含扩展名
PathFileName(Path) As String  返回文件名，包含扩展名
PathExtensionName(Path) As String  返回扩展名，不带.
PathParentFolderName(Path) As String  返回路径,末尾不带\
PathIsFolder(Path) As Boolean 判断是否是文件夹
PathTempName() As String  随机文件名
PathNameSerialNumber(Name, Optional DelimiterLeft = "(", Optional DelimiterRight = ")") As String 名称重复时给名称加序号 Name当前名称 DelimiterLeft序号左侧分隔符 DelimiterRight序号右侧分隔符

单元格-----------------------------------------------------------------------------------------------------------------------------------
ColumnChr(ByVal v) As String  数字转字母
ColumnChrArr(ParamArray arr()) As Variant  数字转字母Arr
ColumnI(ByVal s) As Long  字母转数字
ColumnIArr(ParamArray arr()) As Variant  字母转数字Arr
UnionEx(ByRef Rngs) As Range  单元格并集扩展,传入单元格数组或集合的Range对象，合并成Range
UnionEx_Str(ByRef Rngs, sh) As Range  单元格并集扩展,传入单元格数组或集合的字符串地址，合并成Range
SheetNew(wb As Workbook, Optional Name As String = "") As Worksheet  末尾新增工作表
SheetCopyAfter(sh, Optional Name As String = "") As Worksheet  复制工作表到末尾
SheetCopyNow(sh, Optional Name As String = "") As Worksheet  复制工作表到新工作簿
SheetIsName(wb As Workbook, ByVal Name As String) As Boolean  检查工作表是否存在
WorkbookIsName(ByVal Name As String) As Boolean  检查工作簿是否存在，Name不包含后缀
ArrToRange(ByRef arr, ByVal rng)  数组写入工作表
ArrToRangeUndo(ByRef arr, ByVal rng)  数组写入工作表带撤销
RangAddUndo(ByVal rng)  添加撤销数据
RangStartUndo()  启动撤销 先添加后启动
RngResizeDownRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range 单元格行向下扩展区域
RngResizeRightColumn(ByRef rng) As Range 单元格行向右扩展区域
RngResizeEndRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range 单元格行最后一行扩展区域
RngResizeEndColumn(ByRef rng) As Range 单元格行最后一列扩展区域
RngDownRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long 单元格向下一行
RngRightColumn(ByRef rng As Range) As Long 单元格向右一列
RngEndRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long 单元格最后一行
RngEndColumn(ByRef rng As Range) As Long 单元格最后一列
RangeToArr(rng As Range) As Variant 单元格值到数组,保证一个单元格也是数组
RngMerge_Empty(MergeRng As Range) 向下合并空值单元格
RngMerge_Repeat(MergeRng As Range) 重复值合并单元格
RngAddBorders(rng As Range) 加框线
RngAlignmentCenter(rng As Range) 单元格居中对齐
SheetsSummary(Optional SelectName = "*", Optional RemoveName = "", Optional RngAddress = "", Optional wb As Workbook = Nothing) As Variant 汇总工作表
    汇总工作表 SelectName包含的工作表名 RemoveName排除的工作表名 RngAddress单元格区域默认UsedRange  wb工作簿默认当前
UCreatePivotTable(SourceData As Range, TableDestination As Range, TableName) As PivotTable创建数据透视表 SourceData数据源单元格 TableDestination放置单元格 TableName透视表名字
USetPivotField(PTable As PivotTable, FieldName As String, Orientation As XlPivotFieldOrientation, _
        Position As Long, Optional Caption As String = "", Optional Fun As XlConsolidationFunction = xlCount)
    设置透视表字段 PTable透视表对象UCreatePivotTable返回值  FieldName表格标题
    Orientation 字段位置类型 xlRowField(行标签) xlColumnField(列标签) xlDataField(数据)
    Position 字段顺序
    Caption  字段标题
    Fun   Orientation=xlDataField(数据)时 设置汇总方式：xlSum  xlCount  xlMin  xlMax

FormatConditionAdd(Rng As Range, Formula, Color) As FormatCondition 新增条件格式  Rng条件格式范围  Formula公式  Color颜色RGB值
FormatConditionAdd_Pattern(Rng As Range, Formula, PatternColor, Optional Pattern As XlPattern = xlPatternGray50) As FormatCondition 新增条件格式图案  Rng条件格式范围  Formula公式  PatternColor颜色RGB值
FormatConditionFind(Rng As Range, ByVal Formula) As FormatCondition 按公式查找条件格式
FormatConditionFind_Color(Rng As Range, Color) As FormatCondition 按颜色查找条件格式
FormatConditionFind_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) As FormatCondition 按图案查找条件格式
FormatConditionFindCount(Rng As Range, ByVal Formula) As Long 按公式查找条件格式数量  注意Formula:="=ROW($A1)=*"是错误写法 剪贴后A1可能是A65536 所以Formula:="=ROW($A*)=*"
FormatConditionFindCount_Color(Rng As Range, Color) As Long 按颜色查找条件格式数量
FormatConditionFindCount_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) As Long 按图案查找条件格式数量
FormatConditionModify_Formula(FC As FormatCondition, Formula) 条件格式修改公式
FormatConditionModify_Color(FC As FormatCondition, Color) 条件格式修改颜色
FormatConditionModify_Pattern(FC As FormatCondition, Pattern As XlPattern, PatternColor) 条件格式修改图案颜色
FormatConditionModify_ClearColor(FC As FormatCondition) 条件格式清除颜色
FormatConditionDelete(Rng As Range, ByVal Formula) 按公式删除条件格式 注意Formula:="=ROW($A1)=*"是错误写法 剪贴后A1可能是A65536 所以Formula:="=ROW($A*)=*"
FormatConditionDelete_Color(Rng As Range, Color) 按颜色删除条件格式
FormatConditionDelete_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) 按图案删除条件格式
Rng_Validation(rng As Range, Formula, Optional ShowError As Boolean = True, Optional AlertStyle As XlDVAlertStyle = xlValidAlertStop) 数据有效性 rng单元格 Formula序列"a,b,c" ShowError 显示错误提示并且禁止输入 AlertStyle错误提示样式
RngAddComment(rng As Range, CommentText, Optional Visible As Boolean = False) As Comment 添加批注
RngAddPicture(PicturePath, rng As Range, Optional LowerWidth = 0, Optional LowerHeight = 0, Optional OriginalSizeRatio As Boolean = False) As Shape 添加图片 PicturePath本地路径 rng单元格 LowerWidth宽度缩进量 LowerHeight高度缩进量 OriginalSizeRatio是否按原大小比例


数学-------------------------------------------------------------------------------------------------------------------------------------
SumParams(ParamArray arr()) As Double 参数求和
MaxParams(ParamArray arr()) As Double  参数求最大值
MinParams(ParamArray arr()) As Double  参数求最小值
MaxParams2(Number1, Number2) As Double 两数取最大值 效率高
MinParams2(Number1, Number2) As Double 两数取最小值 效率高
MultiplesUp(Number, Multiples) As Double 向上舍入基数的倍数
MultiplesDown(Number, Multiples) As Double 向下舍入基数的倍数
IntUp(Number) As Long 向上舍入取整
IntDown(Number) As Long 向下舍入取整
RoundUp(Number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double 向上舍入
RoundDown(Number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double 向下舍入
MultipleUp(Number, Significance) As Double 向上舍入指定基数的倍数
MultipleDown(Number, Significance) As Double 向下舍入指定基数的倍数
MultipleRound(Number, Significance) As Double 四舍五入指定基数的倍数
Float_Clear(Number) 清除浮点数运算导致的精度缺失
RoundEX(number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double 真的四舍五入
RandAddSub(Optional Number As Double = 1) As Double 随机 +Number 或 -Number
ModNumber(Number1, Number2) As Double 求余  十亿大数求余不报错
RandBetween(l, r) As Long 按范围随机数
NumberSplit(Number, interval) As Variant  拆分数组 Number被拆分数组 interval拆分大小 NumberSplit(5, 2)->[2,2,1]
NumberLCase(NumberStr) As Double 数字大写转小写
NumberUCase(Number) As String 数字转大写
RMBLCase(NumberStr) As Currency 人民币小写
RMBUCase(curmoney) As String 人民币大写
NumberRangeInside(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean 区间比较 内部
NumberRangeExternal(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean 区间比较 外部
IsEven(Number) As Boolean 判断偶数
IsOdd(Number) As Boolean  判断奇数
Number_Cycle(ByRef Number, ByRef CycleCount) As Long 循环序号 (i,3)->1,2,3,1,2,3,1,2,3
Number_Repeat(ByRef Number, ByRef RepeatCount) As Long 重复序号 (i,3)->1,1,1,2,2,2,3,3,3
Number_Separated(ByRef Number, ByRef SeparatedCount) As Long 相隔序号 (i,3)->1,4,7,10,13,16,19,22,25
vbMaxNumber 常熟 最大值
vbMinNumber 常熟 最小值
vbPi() As Double Pi的值
AngleToRadian(Angle) As Double 角度转弧度
RadianToAngle(Radian, Optional ByVal NumDigitsAfterDecimal = 3) As Double 弧度转角度




功能-------------------------------------------------------------------------------------------------------------------------------------
Deconstruc(ParamArray DValue() As Variant, ByRef Value As Variant) 解构 Deconstruc(变量1, 变量2, 变量3) = Array(1, 2, 3)
Cover(iValue, jValue) 赋值  iValue = jValue
Exchange(iValue, jValue) 交换
ColToArr(ByRef col) As Variant   Col集合转数组
DictionaryCreate(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 item为数组索引 重复值索引取最前
DictionaryCreateRev(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 item为数组索引 反向 重复值索引取最后
DictionaryCreateIndex_ItemIsCol(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 重复值添加到集合索引
DictionaryCreate_DicIndex(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 item为字典自身索引
DictionaryCreate_Items(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 双数组到字典
DictionaryCreate_ItemsRev(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 双数组到字典 反向
DictionaryCreate_ItemsIsCol(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object 创建字典 双数组到字典 重复值添加到集合
DictionaryToArr2D(dic) As Variant 字典到二维数组 1列是Key 2列是Item
DictionaryGetValues(dic, ByVal arrKey, Optional NoExistsValue = Empty) As Variant 字典取多个值  arrKey可以是一维二维数组返回对应大小的Item值数组 NoExistsValue不存填充的值
DictionaryGetValuesParam(dic, ParamArray Keys()) As Variant 字典取多个值 多参数Key
DictionaryExists(dic, ByVal arrKey) As Variant 字典判断多个值 arrKey可以是一维二维数组返回对应大小的True/False数组
DictionaryAdds(Dic, arrKeys, arrItems) As Object 字典批量添加 重复不会修改原来值
DictionaryAddsRev(Dic, arrKeys, arrItems) As Object 字典批量添加 重复则覆盖原来值
DictionaryMerge(ParamArray Dics()) As Object 字典合并
DictionaryMergeRev(ParamArray Dics()) As Object 字典合并 反向 有重复后面替换前面
Application_Attribute(bol As Boolean) Application_Attribute(False)关闭一系列影响效率属性  **注意程序结束后必须 Application_Attribute(True)**
Sleep(PauseTime)  不挂起的不占CPU延迟,单位毫秒
GetTimer() 返回开机时间 单位毫秒
PrintEx(ByRef arg, Optional RowCount = 0, Optional DividerLine As Boolean = True) 打印函数 arg打印内容 RowCount打印行数，负数倒数  DividerLine是否有分隔线*普通类型默认不打印为False时才打印分割线，复杂类型默认打印为False时不打印*
encodeBase64(Bytes) As String 编码Base64
decodeBase64(String1) As Byte() 解码Base64
ImageSize(ImagePath) As Variant 图片像素宽长大小  返回Array(Width, Height)
LoadPictureEx(filename) As IPictureDisp 类似LoadPicture 支持多种图片格式
CLngEx(Expression) As Variant 扩展CLng 支持数组转换
CDateEx(Expression) As Variant 扩展CDate 支持数组转换
CDblEx(Expression) As Variant 扩展CDbl 支持数组转换
CCurEx(Expression) As Variant 扩展CCur 支持数组转换
CStrEx(Expression) As Variant 扩展CStr 支持数组转换
CVarEx(Expression) As Variant 扩展CVar 支持数组转换
CBoolEx(Expression) As Variant 扩展CBool 支持数组转换


Http-------------------------------------------------------------------------------------------------------------------------------------
HttpGet(Url, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Get请求
HttpDownload(Url, DownloadFileName, Optional RequestHeaderDic = Nothing) Get下载文件
HttpPost(Url, Optional SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post请求
HttpPost_Form(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post请求 发送表单数据
HttpPost_Json(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post请求 发送Json数据
HttpReadJson(Jsonstr As String, Routestr As String) As Variant 读取JSON属性
```
