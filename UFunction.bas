Attribute VB_Name = "UFunction"
'ʵ�ù��ܺ���

'����-------------------------------------------------------------------------------------------------------------------------------------
'*�������������麯��*������Index��������ʹ��@���η� ��ʾ��ͷ����n������ ����ArrGetRegion(Array(1, 2, 3), 1, 1)->[2]   ArrGetRegion(Array(1, 2, 3), 1@, 1)->[1]
'Let Titles(ParamArray TitleNames(), ByRef TitleIndexs As Variant) ������⣬�������ֶ�ת��������� ���ӣ�Titles("a", "b", "c") = Array(1, 2, 3)
'Get Titles(ParamArray TitleNames()) As Variant ȡ��������� ��������  T = Titles("a", "b", "c")->[1, 2, 3]
'Get Title() As Object ���ػ�������ֵ� �������ȡ��������  Title!a -> 1  Title!b -> 2
'ArrCache(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) �����������ԣ����Զ��丳ֵȡֵ������֧��һά�Ͷ�ά
'   ArrCache = arr ��ֵ��������
'   ArrCache(RowIndex) = arr �޸Ķ�ά�����RowIndex��1�п�ʼ��ֵ  �� �޸�һά�����RowIndex��ʼ��ֵ
'   ArrCache(, ColumnIndex) = arr �޸Ķ�ά�����1��ColumnIndex�п�ʼ��ֵ
'   ArrCache(RowIndex, ColumnIndex) = arr �޸�RowIndex��ColumnIndex�п�ʼ��ֵ arrһά������д��
'   arr = ArrCache ȡ��������
'   arr = ArrCache(RowIndex) ȡ��ά����һ�� ����һά���� �� ȡһά����һ��ֵ
'   arr = ArrCache(RowIndex����) ȡ��ά������� ���ض�ά���� �� ȡһά������ֵ������ ����һά����
'   arr = ArrCache(, ColumnIndex) ȡ��ά����һ�� ����һά����
'   arr = ArrCache(, ColumnIndex����) ȡ��ά������� ���ض�ά����
'   arr = ArrCache(RowIndex, ColumnIndex) ȡ��ά����һ��ֵ
'   arr = ArrCache(RowIndex����, ColumnIndex) ȡColumnIndex�����RowIndex�����Ķ��ֵ ����һά����
'   arr = ArrCache(RowIndex, ColumnIndex����) ȡRowIndex�����ColumnIndex�����Ķ��ֵ ����һά����
'   arr = ArrCache(RowIndex����, ColumnIndex����) ȡRowIndex��ColumnIndex�������ཻ��ֵ ���ض�ά����
'ArrCache1 , ArrCache2 , ArrCache3 �����������
'ArrBlend(ByRef arrC, Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) �������򸴺ϲ��� ����ArrCache

'ArrGetValue(arr, ByVal RowCount, Optional ByVal ColumnCount, Optional EmptyContent = "") As Variant ����ȡֵ��������Ԫ�ص�RowCount,ColumnCount��ȡ,�������޷���EmptyContent
'��������ʱ��Զ����arr,����Ԫ������Ϊ1ʱ��Զ�������Ԫ�أ�����Ϊһ������ʱֻ����ColumnCount RowCount��=1������Ϊһ�л�һά����ʱֻ����RowCount ColumnCount��=1

'ArrGetValueCache(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
'����ȡֵ���� ͬArrGetValue ��ͬ����arr,EmptyContentд�뺯�������� ���ټ���ӿ��ȡ�ٶ�
'WriteArr=Trueʱд��arr���� WriteArr=Falseʱ����RowCount,ColumnCount��ȡ��������
'���û�������ʾ����ArrGetValueCache WriteArr:=True, arr:=arr, EmptyContent:=""
'��ȡ��������ʾ����v = ArrGetValueCache(i, j)
'ArrGetValueCache1 , ArrGetValueCache2 , ArrGetValueCache3 , ArrGetValueCache4 , ArrGetValueCache5

'ArrayDynamic(Optional ByRef v) As Variant һά��̬���� ��������ӣ���������ȡֵ���ʼ��
'ArrayDynamic1 , ArrayDynamic2, ArrayDynamic3 ���ArrayDynamic
'ArrayDynamic2D(ParamArray v()) As Variant ��ά��̬���� ������������һ�У���������ȡֵ���ʼ��
'ArrayDynamic2D1 , ArrayDynamic2D2 , ArrayDynamic2D3  ���ArrayDynamic2D
'ArrTranspose(ByRef arr) As Variant ����ת��
'ArrFlip(arr) As Variant  ���鷭ת
'ArrTo2D(ByRef arr1D, ByVal DCount As Long) As Variant һά����ת��ά����
'Arr2DTo1D(ByRef arr2D, Optional RowFirst As Boolean = True) As Variant ��ά����תһά����
'ArrF_T(ByRef arr, Optional ColumnCount = 0) As Variant �������������  ColumnCount =0ȡ����� >0ʹ��ColumnCount��Ϊ������������ȥ <0����һ��Ԫ�ص�����Ϊ����
'ArrF_T_LIndexToUIndex(ByRef arr) As Variant ������������� �����������±� *�����ϱ����һ��*
'ArrFlatten_Single(ParamArray arr()) As Variant  չƽ����(һά��) ����
'ArrFlatten(ParamArray arr()) As Variant  չƽ����(һά��) �ݹ�
'Arr2DFlatten(ByRef arr2D, ByVal ColumnIndex) As Variant ��ά�����ں�����������,����Ӧ���и��ƶ���չ��
'ArrMergeRow(ByVal arr) As Variant  �ϲ����飬���ºϲ�
'ArrMergeRowParam(ParamArray arr()) As Variant �ϲ����飬���ºϲ�(�����)
'ArrMergeColumn(ByVal arr) As Variant �ϲ����飬���Һϲ�
'ArrMergeColumnParam(ParamArray arr()) As Variant �ϲ����飬���Һϲ�(�����)

'ArrCopyElement(ByRef arr, ParamArray ArrEleCount()) As Variant һά���� ����Ԫ�� ArrEleCountΪ��Ӧarr��С���������� ArrCopyElement([1,2,3],[2,3])->[1,1,2,2,2,3]
'ArrCopyElement2(ByRef arr, ArrCopyIndex, ArrCopyCount) As Variant һά���� ����Ԫ�� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount�� ArrCopyElement2([1,2,3],[2,3],[2,3])->[1,2,2,3,3,3]
'ArrCopyColumn(ByRef arr2D, ParamArray ArrEleCount()) As Variant �������� ArrEleCountΪ��Ӧarr2D����������������
'ArrCopyColumn2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant �������� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount��
'ArrCopyRow(ByRef arr2D, ParamArray ArrEleCount()) As Variant �������� ArrEleCountΪ��Ӧarr2D����������������
'ArrCopyRow2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant �������� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount��

'ArrInsert(ByRef arr, Optional ByVal Index, Optional ByVal EleCount As Long = 1, Optional EleCopy As Boolean = False) As Variant һά���� ����һ����ֵ������ֵ EleCopy=True���Ʋ���
'ArrInsertColumn(ByRef arr2D, Optional ByVal ColumnIndex, Optional ByVal ColumnCount As Long = 1, Optional EleCopy As Boolean = False) As Variant ���� ����һ�л���� EleCopy=True���Ʋ���
'ArrInsertRow(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal RowCount As Long = 1, Optional EleCopy As Boolean = False) As Variant ���� ����һ�л���� EleCopy=True���Ʋ���
'ArrGetIndex(ByRef arr, Optional GetRowIndex As Boolean = True) As Variant() ���� ȡ����
'ArrRemoveRegion(ByRef arr, ByRef Index, Optional ByVal Count = 1) As Variant һά���� ɾ��һ��Ԫ�ػ���Ԫ��
'ArrRemoveColumn(ByRef arr2D, ByRef Index, Optional ByVal ColumnCount = 1) As Variant ���� ɾ��һ�л����
'ArrRemoveColumns(ByRef arr2D, ParamArray arrIndex()) As Variant ���� ɾ��һ�л���� �����
'ArrRemoveRow(ByRef arr2D, ByRef Index, Optional ByVal RowCount = 1) As Variant ���� ɾ��һ�л����
'ArrRemoveRows(ByRef arr2D, ParamArray arrIndex()) As Variant ���� ɾ��һ�л���� �����
'ArrGetRow(ByRef arr2D, ByRef Index, Optional ByVal RowCount = 1, Optional Expansion As Boolean = False) As Variant ����ȡ���� һ��Ϊһά���� RowCount=0ȡ�������
'ArrGetRows(ByRef arr2D, ByVal arrIndex) As Variant  ����ȡ���е���ά����
'ArrGetColumn(ByRef arr2D, ByRef Index, Optional ByVal ColumnCount = 1, Optional Expansion As Boolean = False) As Variant ����ȡ���� һ��Ϊһά���� ColumnCount=0ȡ�������
'ArrGetColumns(ByRef arr2D, ByVal arrIndex) As Variant  ����ȡ���е���ά����
'ArrGetRegion2D(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
     Optional ByVal Height = 0, Optional ByVal Width = 0, Optional Expansion As Boolean = False) As Variant  ����ȡ���� �����Ӵ�С ��ά����
'ArrGetRegion2D_To(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
        Optional ByVal RowIndex2, Optional ByVal ColumnIndex2, Optional Expansion As Boolean = False) As Variant  ����ȡ���� ���������� ��ά����
'ArrGetRegion(ByRef arr, Optional ByVal Index, Optional ByVal Count = 0, Optional Expansion As Boolean = False) As Variant ����ȡ���� һά����
'ArrGetRegion_To(ByRef arr, Optional ByVal Index, Optional ByVal IndexTo, Optional Expansion As Boolean = False) As Variant ����ȡ���� ���������� һά����
'ArrSizeExpansion(ByRef arr, ByRef RowCount, Optional ByRef ColumnCount, Optional FillValue = Empty) ���������С  **�����±��1**

'ArrSizeExpansionEx(ByRef arr, ByRef RowCount, ByRef ColumnCount, Optional FillValue = Empty)���������С ���������������  **�����±��1**
'��������ʱ�������Ԫ��,����Ԫ������Ϊ1ʱ�������Ԫ�أ�����Ϊһ������ʱ��������У�����Ϊһ�л�һά����ʱ���������

'ArrSizeExpansion2(ByRef arr, ByRef ArrSizeCount, Optional FillValue = Empty) ���������С �������鶼���һά  **�����±��1**  �������������
'ArrIndexExpansion(ByRef arr, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional FillValue = Empty) ����������������������������ʱ�ᱻ����
'Arr2DSetArr2D(ByRef arrL, ByRef arrR, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False)  ���鸳ֵ������ ��ά
'Arr2DSetValues(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values()) ���ֵ��RowIndexArr��ColumnIndexArr����λ�����θ�ֵ������  ���ϵ���һ��һ��д�� ��ά
'Arr2DSetValues_LtoR(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values()) ���ֵ��RowIndexArr��ColumnIndexArr����λ�����θ�ֵ������  ������һ��һ��д�� ��ά
'ArrSetValues(ByRef arr1D, ByRef IndexArr, ParamArray Values()) ���ֵ��IndexArrλ�����θ�ֵ������ һά
'ArrSetEntireColumnValues(ByRef arr2D, ByRef ColumnIndexArr, ParamArray Values()) ��ֵ������һ���� ��ֵ��Ӧ���� ��ά
'ArrSetEntireRowValues(ByRef arr2D, ByRef RowIndexArr, ParamArray Values()) ��ֵ������һ���� ��ֵ��Ӧ���� ��ά
'ArrSetArr(ByRef arrL, ByRef arrR, Optional ByVal Index, Optional Expansion As Boolean = False)  ���鸳ֵ������ һά
'ArrSetColumn(ByRef arrL2D, ByRef arrR, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False) ���鸳ֵ������һ��
'ArrSetRow(ByRef arrL2D, ByRef arrR, Optional ByVal RowIndex, Optional Expansion As Boolean = False) ���鸳ֵ������һ��
'ArrFromIndex(arr, arrIndex) As Variant   ����������˳��ȡ������ֵ��������ԭ������
'ArrFromBoolea(arr, arrBoolea) As Variant ��������������=Trueȡ������ֵ������ɸѡ����
'ArrRandSort(ByVal arr) As Variant  �����������
'ArrSort2D(arr, Index, Optional Order As Boolean = True) As Variant  ��ά�����ȶ�����
'ArrSort2Ds(arr, Indexs, Optional Orders = True) As Variant ��ά��������ȶ�����
'ArrSort1D(arr, Optional Order As Boolean = True) As Variant һά�����ȶ�����
'ArrSort(arr, Optional Order As Boolean = True) As Variant һά�����ȶ����� ����������Order=True ��������
'���ӣ�����arr��ά����
'ArrColumns = ArrGetColumn(arr, 1)  'ȡ��arr������
'arrIndex = ArrSort(ArrColumns)  '�����������򷵻���������
'arrOrder = ArrFromIndex(arr, arrIndex) '������������ȡ����������
'ArrSortNext(arr, Indexs, Optional Order As Boolean = True) As Variant  ������������
'���ӣ���1,2������
'arrIndex = ArrSort(ArrGetColumn(arr, 1)) '��һ������
'arrIndex = ArrSortNext(ArrGetColumn(arr, 2), arrIndex) '��2������
'arrorder = ArrFromIndex(arr, arrIndex) '���ؽ��
'ArrCustomSort2D(arrValue, arrKey, Index, Optional IsLike As Boolean = False) As Variant  ��ά�����Զ�������
'ArrCustomSort(arrValue, arrKey, Optional IsLike As Boolean = False)  �Զ�������  CustomSort(��������, �Զ�������, Likeƥ��) ������������
'ArrInInterval(ByVal arrInterval, Number) As Long ����Number��arrInterval�������λ�� λ��������LBound(arrInterval)��UBound(arr)+1 arrInterval��������˳��
'ArrInIntervalEqual(ByVal arrInterval, Number) As Long ����Number��arrInterval�������λ�� ������ λ��������LBound(arrInterval)��UBound(arr)+1 arrInterval��������˳��
'ArrFindLessIndex(arr_Small, V_Large, Optional ByVal Start) As Long ����С��v������
'ArrFindLessIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long ����С��v������ ����
'ArrFindLessEqualIndex(arr_Small, V_Large, Optional ByVal Start) As Long ����С�ڵ���v������
'ArrFindLessEqualIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long ����С�ڵ���v������ ����
'ArrFindGreaterIndex(arr_Large, V_Small, Optional ByVal Start) As Long ���Ҵ���v������
'ArrFindGreaterIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long ���Ҵ���v������ ����
'ArrFindGreaterEqualIndex(arr_Large, V_Small, Optional ByVal Start) As Long ���Ҵ��ڵ���v������
'ArrFindGreaterEqualIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long ���Ҵ��ڵ���v������ ����
'ArrFindLikeIndex(arr, v, Optional ByVal Start) As Long  ���Ҷ�Ӧֵ���� Like
'ArrFindLikeIndexRev(arr, v, Optional ByVal Start) As Long  ���Ҷ�Ӧֵ�������� Like
'ArrFindNotLikeIndex(arr, v, Optional ByVal Start) As Long ���Ҷ�Ӧֵ���� Not Like
'ArrFindNotLikeIndexRev(arr, v, Optional ByVal Start) As Long ���Ҷ�Ӧֵ�������� Not Like
'ArrFindIndex(arr, v, Optional ByVal Start) As Long  ���Ҷ�Ӧֵ����
'ArrFindIndexRev(arr, v, Optional ByVal Start) As Long  ���Ҷ�Ӧֵ��������
'ArrFindNotIndex(arr, v, Optional ByVal Start) As Long ���Ҳ����ڵ�ֵ����
'ArrFindNotIndexRev(arr, v, Optional ByVal Start) As Long ���Ҳ����ڵ�ֵ��������
'ArrFindRegIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long  ���Ҷ�Ӧֵ���� ����
'ArrFindRegIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long  ���Ҷ�Ӧֵ���� ���� ����
'ArrFindRegNotIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long ���Ҷ�Ӧֵ���� ����������
'ArrFindRegNotIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long ���Ҷ�Ӧֵ���� ���������� ����
'ArrFindIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant ��ά����������� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrFindNotIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant ��ά����������� ������ �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrFindLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant ��ά����������� Like���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrFindNotLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant ��ά����������� Not Like���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrFindRegIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant ��ά����������� ���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrFindRegNotIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant ��ά����������� ���������� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
'ArrValid_InError(arr) As Boolean    ��������Ч�� �д��󷵻�True
'ArrValid_NumericAll(arr, Optional InEmpty As Boolean = True, Optional IsStr As Boolean = True) As Boolean  ��������Ч�� ȫ�������ַ���True
'ArrValid_DateAll(arr, Optional IsStr As Boolean = True) As Boolean  ��������Ч�� ȫ�������ڷ���True
'ArrValid_Reg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean ��������Ч������һ�� ���� ƥ�䷵��True
'ArrValid_RegAll(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean  ��������Ч������ȫ�� ���� ȫ��ƥ�䷵��True
'ArrValid_Repeat(arr) As Boolean ��������Ч���Ƿ����ظ� �ظ�����True
'ArrValid_Incremental(ParamArray arr()) As Boolean ��������Ч���Ƿ��������
'ArrValid_IncrementalEqual(ParamArray arr()) As Boolean ��������Ч���Ƿ�������а������
'ArrValid_Decrement(ParamArray arr()) As Boolean ��������Ч���Ƿ�ݼ�����
'ArrValid_DecrementEqual(ParamArray arr()) As Boolean ��������Ч���Ƿ�ݼ����а������
'ArrFilterRepeatCount(arr, Optional CountSmall = 0, Optional CountLarge = 1.79769313486231E+308, Optional CompareMode As CompareMethod = BinaryCompare) As Variant ɸѡ �ظ�����  ,*����ɸѡ����*
'ArrFilterRangeInside(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ɸѡ ���� �ڲ� ,*����ɸѡ����*
'ArrFilterRangeExternal(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ɸѡ ���� �ⲿ ,*����ɸѡ����*
'ArrFilterGreater(arr_Large, V_Small) As Variant ɸѡ ����V_Small��ֵ ,*����ɸѡ����*
'ArrFilterGreaterEqual(arr_Large, V_Small) As Variant ɸѡ ���ڵ���V_Small��ֵ ,*����ɸѡ����*
'ArrFilterLess(arr_Small, V_Large) As Variant ɸѡ С��V_Large��ֵ ,*����ɸѡ����*
'ArrFilterLessEqual(arr_Small, V_Large) As Variant ɸѡ С��V_Large��ֵ ,*����ɸѡ����*
'ArrFilter(arr, ByVal arrv) As Variant   ɸѡ ,**����ɸѡ����**
'ArrFilterLike(arr, ByVal arrv) As Variant  ɸѡlikeƥ�� ,**����ɸѡ����**
'ArrFilterReg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant  ɸѡ����ƥ�� ,**����ɸѡ����**
'ArrFilterRemove(arr, ByVal arrv) As Variant  ɸѡ�ų� ,**����ɸѡ����**
'ArrFilterLikeRemove(arr, ByVal arrv) As Variant  ɸѡlike�ų� ,**����ɸѡ����**
'ArrFilterRegRemove(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant  ɸѡ�����ų� ,**����ɸѡ����**
'ArrDistinct(arr) As Variant ȥ�� ������ͷһ��ֵ
'ArrDistinctIndex(arr) As Variant ȥ�أ��������� ������ͷ����
'ArrDistinctIndexRev(arr) As Variant ȥ�أ��������� �����������
'ArrLBoundToN_1D(arr, Optional StartLBound = 1) As Variant �����±��StartLBound һά����
'ArrLBoundToN_2D(arr, Optional StartLBound1 = 1, Optional StartLBound2 = 1) As Variant �����±��StartLBound1,StartLBound2 ��ά����
'ArrMap(ByVal arr, EvaluateStr) As Variant  Evaluate�޸����� $��ʾ��ǰֵ
'ArrReplace(ByRef arr, FindValueArr, ReplaceValue) As Variant �����滻������������Ԫ�� FindValueArr֧�ֵ�ֵ������
'ArrErrorClear(ByRef arr, Optional EmptyValue = Empty) As Variant ����������ֵ
'ArrIsValid(ByRef arr) As Boolean  �����Ƿ���Ч
'ArrDimension(ByRef arr) As Long  ����ά��
'ArrCount(ByRef arr) As Long  ����Ԫ�ظ���
'ArrCountRow(ByRef arr) As Long  ��������
'ArrCountColumn(ByRef arr) As Long ��������
'ArrCountRowAndColumn(arr, ByRef RowCount, ByRef ColumnCount) ͬʱ�������������ñ���RowCount,ColumnCount���շ���ֵ��һά����ColumnCount=1����������RowCount=ColumnCount=1
'ArrCountElement(ByVal arr) As Variant ������Ԫ�ظ�����������������
'ArrCountMergeElement(ByRef arr, Optional EmptyContent = "") As Variant �����Ǻϲ���Ԫ����ʽԪ�ظ��������ظ�������
'ArrBetween(l, u) As Variant()  ������Χ��������
'ArrCreate(Number, Optional Number2 = 0, Optional FillValue = Empty) As Variant() ��������
'ArrCreateRand(l, r, RowCount, Optional ColumnCount = 0) As Variant() �������������
'ArrCreateRandDic(l, r, RowCount, Optional ColumnCount = 0) As Variant() ������������� ���ظ������
'ArrFillDown(ByRef arr, Optional index = 1, Optional EmptyContent = "") As Variant ��ֵ�������  arrһά���ά���� index��ά����������  EmptyContent������ֵ������
'ArrFillUp(ByRef arr, Optional index = 1, Optional EmptyContent = "") As Variant  ��ֵ�������  arrһά���ά���� index��ά����������  EmptyContent������ֵ������
'ArrPerspectiveRev(ByRef arrH, ByRef arrV, Optional ByRef arrRegion2D = "") As Variant
'  ��͸�� arrH������(�����Ƕ���)  arrV�����(ֻ��һ��) arrRegion2D��������(�д�С������arrH���� �д�С������arrV����)
'ArrPerspective(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant ͸�� ���н������ظ�����ʱȡ���һֵ arr2D��ά��  VIndex���������  DataIndex�������������
'ArrPerspective_Repeating(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant ͸�� ���н������ظ�����ʱд���� arr2D��ά��  VIndex���������  DataIndex�������������
'ArrGroupSum(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrSumIndex) As Variant ������� arr2D��ά�� ArrGroupIndex����������֧������ ArrSumIndex���������֧������
'ArrGroupCount(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrCountIndex, Optional NoEmpty As Boolean = True) As Variant ������� arr2D��ά�� ArrGroupIndex����������֧������ ArrCountIndex����������֧������ NoEmpty = True����ǿ�ֵ����
'ArrGroupJoin(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrJoinIndex, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
'    ����ƴ���ַ��� arr2D��ά�� ArrGroupIndex����������֧������ ArrJoinIndex���������֧������ Delimiter�ָ��� OmittedEmpty���Կ��ַ���
'ArrGroup_Class(ByRef arr2D, ByVal ArrClassIndex) As Variant ������� ����� ArrClassIndex��������֧������  ��������������ķ���
'ArrGroup_Find_First(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant ������� ����������Ϊ������� ���޷��ڷ���*����*  FindIndex������ FindValue��������  ��������������ķ���
'ArrGroup_Find_Last(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant  ������� ����������Ϊ������� ���޷��ڷ���*ĩβ*  FindIndex������ FindValue��������  ��������������ķ���
'ArrGroup_Differ(ByRef arr2D, ByVal ArrDifferIndex) As Variant ������� �����������ݲ���Ϊ�������  ArrDifferIndex��ͬ��������֧������  ��������������ķ���
'ArrGroup_Number_Column(ByRef arr2D, ByVal Number) As Variant ������� ��������  number����  ��������������ķ���
'ArrGroup_Number(ByRef arr2D, ByVal number) As Variant������� ������  number����  ��������������ķ���
'ArrGroup_Row_First(ByRef arr2D, ByVal ArrRowIndex) As Variant  ������� ��������Ϊ���޷���  ���޷��ڷ���*����* ArrRowIndex������֧������  ��������������ķ���
'ArrGroup_Row_Last(ByRef arr2D, ByVal ArrRowIndex) As Variant  ������� ��������Ϊ���޷���  ���޷��ڷ���*ĩβ* ArrRowIndex������֧������  ��������������ķ���
'ArrGroup_Interval(ByVal arr2D, ByVal ColumnIndex, ParamArray ArrInterval()) As Variant ������� ����ֵ����������  С�ڲ����ڱ���һ�� ArrInterval��������  ��������������ķ���
'ArrGroup_Interval_Equal(ByVal arr2D, ByVal ColumnIndex, ParamArray ArrInterval()) As Variant ������� ����ֵ����������  С�ڵ��ڱ���һ�� ArrInterval��������  ��������������ķ���
'ArrGroup_CustomClass(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant ������� ���Զ������ ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
'ArrGroup_CustomClass_Like(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant ������� ���Զ������Likeƥ��  ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
'ArrGroup_CustomClass_Reg(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomPattern()) As Variant ������� ���Զ������ ����ƥ��  ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
'ArrGroupAgg(ByRef ArrGroup, Optional OmittedNoneArg As Boolean = True, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, _
        Optional ByRef C1 As GroupAggregateMethod = Group_None, C2, C3,... C46 ) As Variant
'       ����ۺϺ���  ArrGroup���麯�����ص�����������  OmittedNoneArgû��дCn���������Ƿ�ʡ�� Delimiterƴ���ַ��ָ��� OmittedEmptyƴ���ַ����Ƿ���Կ�ֵ
'       C1-C46���������1-46�� ����C1:=Group_Sum��ʽ����ۺ�ģʽGroupAggregateMethod  C1-C46��������ȡ��N�д��븺��ȡ������N��

'ArrGroupAgg2(ByRef ArrGroup, ArrGroupIndex, ArrAggregateMethod, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
'����ۺϺ��� ֧��һ�ж��־ۺ� ArrGroup���麯�����ص����������� ArrGroupIndex�ۺ��� ArrAggregateMethod��Ӧ�ľۺ�ģʽ Delimiterƴ���ַ��ָ��� OmittedEmptyƴ���ַ����Ƿ���Կ�ֵ

'ArrUnions(ParamArray arr()) As Variant �������  ȡ�������Ԫ��
'ArrUnions_Distinct(ParamArray arr()) As Variant �������  ȥ��
'ArrUnions_Sort(ParamArray arr()) As Variant �������  ����
'ArrUnions_DistinctSort(ParamArray arr()) As Variant �������  ȥ������
'ArrUnion(ByRef arr1, ByRef arr2) As Variant ���� ȡ��������Ԫ��
'ArrUnion_Distinct(ByRef arr1, ByRef arr2) As Variant ���� ȥ��
'ArrUnion_Sort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant ���� ����
'ArrUnion_DistinctSort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant ���� ȥ������
'ArrIntersects(ParamArray arr()) As Variant  �������  ȡ�������Ԫ��
'ArrIntersects_Distinct(ParamArray arr()) As Variant �������  ȥ��
'ArrIntersects_arr1(ParamArray arr()) As Variant ������� ȡ��һ������Ԫ��
'ArrIntersects_arr1_Index(ParamArray arr()) As Variant ������� ȡ��һ������Ԫ��
'ArrIntersect(ByRef arr1, ByRef arr2) As Variant ���� ȡ��������Ԫ��
'ArrIntersect_Distinct(ByRef arr1, ByRef arr2) As Variant ���� ȥ��
'ArrIntersect_arr1(ByRef arr1, ByRef arr2) As Variant ���� ȡarr1Ԫ��
'ArrIntersect_arr2(ByRef arr1, ByRef arr2) As Variant ���� ȡarr2Ԫ��
'ArrIntersect_arr1_Index(ByRef arr1, ByRef arr2) As Variant ���� ȡarr1����
'ArrIntersect_arr2_Index(ByRef arr1, ByRef arr2) As Variant ���� ȡarr2����
'ArrExcepts_Single(ParamArray arr()) As Variant ����  ȡ�������Ԫ��(������������������û�е�Ԫ��)[1,2,3,4,5,5][1,2,3][2,3,4,6]->[5,5,6]
'ArrExcepts_RemoveAllIntersect(ParamArray arr()) As Variant ����  ȡ�������Ԫ��(ȥ���������鶼������Ԫ��)[1,2,3,4,5,5][1,2,3][2,3,4,6]->ȥ������Ԫ��2,3�õ�[1,4,5,5,1,4,6]
'ArrExcepts_arr1(ParamArray arr()) As Variant ����  ȡ��һ��Ԫ��
'ArrExcepts_arr1_Index(ParamArray arr()) As Variant ���� ȡ��һ������Ԫ������
'ArrExcept(ByRef arr1, ByRef arr2) As Variant � ȡ��������Ԫ��
'ArrExcept_arr1(ByRef arr1, ByRef arr2) As Variant � ȡarr1Ԫ��
'ArrExcept_arr2(ByRef arr1, ByRef arr2) As Variant � ȡarr2Ԫ��
'ArrExcept_arr1_Index(ByRef arr1, ByRef arr2) As Variant � ȡarr1����
'ArrExcept_arr2_Index(ByRef arr1, ByRef arr2) As Variant � ȡarr2����
'ArrTitleToIndex(ByRef arrTitle, ByRef arrOrder) As Variant  arrTitle(һά)��arrOrder(һά)���ض�Ӧ��˳��ı�����������,���ص�����ΪarrTitle������ƥ��λ�÷���(LBound-1),���ص������С��arrOrder��ͬ
'ArrIFs(ParamArray Calculates()) As Variant ����IFs�жϼ��� ArrIFs(����,ֵ,����,ֵ,����ֵ)
'ArrBoolea_And(ParamArray Calculates()) As Variant ���鲼���Ҽ���
'ArrBoolea_Or(ParamArray Calculates()) As Variant ���鲼�������
'ArrBoolea_Not(ByVal arr) As Variant ���鲼���Ǽ���
'ArrComp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ��������Ƚϼ��� �ڲ�
'ArrComp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ��������Ƚϼ��� �ⲿ
'ArrComp_Like(ByVal arr, ByVal arr2) As Variant ����Ƚ�Like����
'ArrComp_NotLike(ByVal arr, ByVal arr2) As Variant ����Ƚ�Not Like����
'ArrComp_Equal(ByVal arr, ByVal arr2) As Variant ����Ƚϵ��ڼ���
'ArrComp_NotEqual(ByVal arr, ByVal arr2) As Variant ����Ƚϲ����ڼ���
'ArrComp_Size(ByVal arr_Large, ByVal arr_Small) As Variant ����Ƚϴ�С����
'ArrComp_SizeEqual(ByVal arr_Large, ByVal arr_Small) As Variant ����Ƚϴ�С�������ڼ���
'ArrMath_Add(ParamArray Calculates()) As Variant ����ӷ�����
'ArrMath_Sub(ParamArray Calculates()) As Variant �����������
'ArrMath_Multipli(ParamArray Calculates()) As Variant ����˷�����
'ArrMath_Division(ParamArray Calculates()) As Variant �����������
'ArrMath_Power(ParamArray Calculates()) As Variant ����˷�����
'ArrMath_Join(ParamArray Calculates()) As Variant �������Ӽ���
'ArrMath_Round(ByVal arr, number, Optional ColumnIndex = 1) As Variant ������������
'ArrMath_Val(ByVal arr, Optional ColumnIndexArr = 1) As Variant
'ArrMath_Abs(ByVal arr, Optional ColumnIndexArr = 1) As Variant �������ֵAbs
'ArrMath_Format(ByVal arr, Pormat, Optional ColumnIndex = 1) As Variant ����Format
'ArrStr_Ucase(ByVal arr, Optional ColumnIndexArr = 1) As Variant ����ת��д
'ArrStr_Lcase(ByVal arr, Optional ColumnIndexArr = 1) As Variant ����תСд

'ArrStr_Split(ByVal arr, Delimiter, Optional ColumnIndexArr = 1) As Variant  ����ѭ������ַ��� ��������������
'ArrStr_Replace(ByVal arr, FindStr, ReplaceStr, Optional ColumnIndex = 1) As Variant �����滻
'ArrStr_ReplaceAll(ByVal arr, FindStr, ReplaceStr) As Variant �����滻������������
'ArrStr_RegexSearch(ByVal arr, Pattern, Optional RegIndex = 0, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant ��������ȡֵ
 
'ArrStr_RegexSearchs(ByVal arr, Pattern, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant ��������ȡ����ֵ��������������
        
'ArrStr_RegexCount(ByVal arr, Pattern, Optional ByVal ColumnIndexArr = 1, Optional ByVal NumberAdd = 0, _
         Optional ByRef ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant �������򷵻�ƥ������
         
'ArrStr_RegexReplace(ByVal arr, Pattern, ReplaceStr, Optional ColumnIndex = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant ���������滻
 
'ArrStr_Mid(ByVal arr, start, Optional length, Optional ColumnIndex = 1) As Variant ����MID
'ArrDate_DateSub(Interval, Date1, Date2) As Variant �������ڲ�ֵ ����DateDiff
'ArrDate_Year(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ��
'ArrDate_Month(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ��
'ArrDate_Day(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ��
'ArrDate_Weekday(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ����
'ArrTime_Hour(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡСʱ
'ArrTime_Minute(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ����
'ArrTime_Second(ByVal arr, Optional ColumnIndex = 1) As Variant ����ȡ��
'ArrSerialNumber(ByVal arr, Optional ColumnIndex = 1, Optional StartNumber = 1) As Variant ����� �������鷵��1++���
'ArrSerialNumberCalssSelf(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant ����� �����鲻ͬ���� ��ͬ����1++ ����1++���
'ArrSerialNumberCalss(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant ����� �����鲻ͬ����1++ ����1++���
'ArrMaxIndex(ByRef arr, Optional ColumnIndex = 1, Optional Front As Boolean = True) As Long ����ȡ���ֵ���� ColumnIndex ��ά����������  Front = True ��ǰ������
'ArrMinIndex(ByRef arr, Optional ColumnIndex = 1, Optional Front As Boolean = True) As Long ����ȡ��Сֵ���� ColumnIndex ��ά����������  Front = True ��ǰ������
'ArrSum(ByRef arr) As Double  �������
'ArrMax(ByRef arr) As Double  ���������ֵ
'ArrMin(ByRef arr) As Double  ��������Сֵ
'ArrCountNoEmpty(ByRef arr) As Double �������ǿ�ֵ����
'ArrSumColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant ���鰴�����
'ArrSumRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant ���鰴�����
'ArrMaxColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant ���鰴�������ֵ
'ArrMaxRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant ���鰴�������ֵ
'ArrMinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant ���鰴������Сֵ
'ArrMinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant ���鰴������Сֵ
'ArrJoinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant ���鰴��ƴ���ַ���
'ArrJoinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant ���鰴��ƴ���ַ���
'ArrCountNoEmptyColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant ���鰴�м���ǿ�ֵ����
'ArrCountNoEmptyRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant ���鰴�м���ǿ�ֵ����
'ArrCountClassColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant ���鰴�м�����������
'ArrCountClassRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant ���鰴�м�����������
'ArrAverageColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant ���鰴�м���ƽ��ֵ  NumDigitsAfterDecimal����С��λ��
'ArrAverageRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant ���鰴�м���ƽ��ֵ  NumDigitsAfterDecimal����С��λ��
'ArrAverage(ByRef arr, Optional NumDigitsAfterDecimal As Long = 2) As Double ���������ƽ��ֵ  NumDigitsAfterDecimal����С��λ��
'ArrMoveUp(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant ��ֵ�ƶ� ����
'ArrMoveDown(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant ��ֵ�ƶ� ����
'ArrMoveLeft(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant ��ֵ�ƶ� ����
'ArrMoveRight(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant ��ֵ�ƶ� ����
'ArrMove(ByRef arr1D, Optional EmptyContent = "") As Variant ��ֵ�ƶ� һά���� ����
'ArrMoveRev(ByRef arr1D, Optional EmptyContent = "") As Variant ��ֵ�ƶ� һά���� ����
'ArrMove_Index(ByRef arr1D, Optional EmptyContent = "") As Variant ��ֵ�ƶ� һά���� ���� ��������
'ArrMoveRev_Index(ByRef arr1D, Optional EmptyContent = "") As Variant ��ֵ�ƶ� һά���� ���� ��������
'ArrScroll(ByRef arr, Index) As Variant ������� ���� Index������������ͷ
'ArrScrollRev(ByRef arr, Index) As Variant ������� ���� Index����������ĩβ
'ArrScroll_Index(ByRef arr, Index) As Variant ������� ���� Index������������ͷ ��������
'ArrScrollRev_Index(ByRef arr, Index) As Variant ������� ���� Index����������ĩβ ��������
'ArrScrollColumn(ByRef arr2D, Index) As Variant ��ά�����й��� ���� Index������������ͷ
'ArrScrollColumnRev(ByRef arr2D, Index) As Variant ��ά�����й��� ���� Index����������ĩβ
'ArrScrollColumn_Index(ByRef arr2D, Index) As Variant ��ά�����й���  ���� Index������������ͷ ��������
'ArrScrollColumnRev_Index(ByRef arr2D, Index) As Variant ��ά�����й��� ���� Index����������ĩβ ��������
'ArrCombinCon(arr, r) ���  arr һά���� r��ȡ����
'ArrPermutCon(arr, r) ����  arr һά���� r��ȡ����


'����-------------------------------------------------------------------------------------------------------------------------------------
'Matrix_Add(ParamArray Calculates()) As Variant ����ӷ�����
'Matrix_Sub(ParamArray Calculates()) As Variant �����������
'Matrix_Multipli(ParamArray Calculates()) As Variant ����˷�����
'Matrix_Division(ParamArray Calculates()) As Variant �����������
'Matrix_Power(ParamArray Calculates()) As Variant ����˷�����
'Matrix_Join(ParamArray Calculates()) As Variant �������Ӽ���
'Matrix_Comp_Equal(ByRef arr, ByRef arr2) As Variant ����Ƚϵ���
'Matrix_Comp_NotEqual(ByRef arr, ByRef arr2) As Variant ����Ƚϲ�����
'Matrix_Comp_Size(ByRef arr_Large, ByRef arr_Small) As Variant ����Ƚϴ�С
'Matrix_Comp_SizeEqual(ByRef arr_Large, ByRef arr_Small) As Variant ����Ƚϴ�С��������
'Matrix_Comp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ��������Ƚϼ��� �ڲ�
'Matrix_Comp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant ��������Ƚϼ��� �ⲿ
'Matrix_Comp_Like(ByRef arr, ByRef arr2) As Variant ����Ƚ�Like
'Matrix_Comp_NotLike(ByRef arr, ByRef arr2) As Variant ����Ƚ�Not Like
'Matrix_Boolea_And(ParamArray Calculates()) As Variant ���󲼶��Ҽ���
'Matrix_Boolea_Or(ParamArray Calculates()) As Variant ���󲼶������
'Matrix_Boolea_Not(ByRef arr) As Variant ���󲼶��Ǽ���
'Matrix_IF(Expression, TruePart, FalsePart) As Variant ����IF
'Matrix_IFs(ParamArray Calculates()) As Variant ����IFs
'Matrix_Str_Mid(String1, Start, Optional Length) As Variant ����Mid ���������String1, Start, Length
'Matrix_Str_Left(String1, Length) As Variant ����Left ���������String1, Length
'Matrix_Str_Right(String1, Length) As Variant ����Right ���������String1, Length
'Matrix_Str_InStr(StringLarge, StringSmall, Optional Start = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant ����InStr ���������StringLarge, StringSmall, Start
'Matrix_Str_InStrRev(StringLarge, StringSmall, Optional Start = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant ����InStr ���������StringLarge, StringSmall, Start
'Matrix_Str_Len(ByRef String1) As Variant ����Len ���������String1
'Matrix_Str_Replace(Expression, Find, Replace, Optional Start = 1, Optional Count = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant �����滻 ���������Expression, Find, Replace
'Matrix_DateSub(Interval, Date1, Date2) As Variant �������ڼ�� ����DateDiff ���������Interval, Date1, Date2







'�ַ���-----------------------------------------------------------------------------------------------------------------------------------
'StringBuilder(Optional ByRef s) As Variant  ��������ӣ���������ȡֵ���ʼ��
'StringBuilder1 , StringBuilder2, StringBuilder3 ���StringBuilder
'StrJoinArr2D(ByRef arr2D, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, Optional RowFirst As Boolean = True) As String ��ά����ƴ��
'StrJoin_ArrDelimiter(ByRef arr, ParamArray ArrDelimiter()) As String ���齻��ƴ��
'StrStrLike(str1, LikeStr) As Boolean  Likeƥ��
'StrLeft(String1, Length) As String ֧�ָ�Length��Left
'StrRight(String1, Length) As String ֧�ָ�Length��Right
'StrMid(String1, ByVal Start, ByVal Length) As String ֧�ָ�Start��Length��Mid
'StrMidBetween(String1, ByVal Start, Optional ByVal EndIndex = 0) As String ��ʼ����ȡ�ַ���
'StrGetLeft(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  ȡstr������ݣ��������
'StrGetLeftRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  ȡstr������ݣ����Ҳ���
'StrGetRight(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  ȡstr�ұ����ݣ��������
'StrGetRightRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String  ȡstr�ұ����ݣ����Ҳ���
'StrGetCentre(String1, str1, str2, Optional SearchType As SearchDirection = LeftLeft) As String ȡ����str�м�����
'StrTrimChr(String1, Optional Chrs = " ") As String ��Chrs����ַ�ȥ�������ַ���
'StrLTrimChr(String1, Optional Chrs = " ") As String ��Chrs����ַ�ȥ������ַ���
'StrRTrimChr(String1, Optional Chrs = " ") As String ��Chrs����ַ�ȥ���Ҷ��ַ���
'StrRepeat(ByVal string1, ByVal numberOfRepeats As Long) As String   �ظ��ַ���
'StrReplaces(Expression, Finds, Replaces, Optional Counts = -1, _
      Optional Compare As VbCompareMethod = vbBinaryCompare) As String �����滻 Finds,Replaces,Counts֧������ StrReplaces("aabca",{"aa","a"},{"a","e"})->abce
'StrReplaceChr(ByVal String1, StrKey, StrItem) As String ��StrKey����ַ� �滻��Ӧλ�õ�StrItem  StrReplaceChr("aabbccdd","abc","123")->112233dd
'StrReplacePlaceholder(ByVal String1, placeholder, ParamArray ValueStrs()) As String �滻ռλ��placeholder    StrReplacePlaceholder("a%b%c", "%", 1, 2) '"a1b2c"
'StrReplaceIndex(String1, ReplaceStr, ByVal Start, ByVal Length) As String ������λ���滻
'Str_Split(ByVal Expression, Optional Delimitre = "", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String()
'    ����ַ��� ֧�ֶ���ָ��
'Str_SplitMatch(String1, ParamArray Delimitre()) As Variant ���� "���=1,����=abc,����=1" ���͵����ݣ�Str_SplitMatch("���=1,����=abc,����=1", "���=",",����=",",����=")�������飬����(0)��"���="�������
'Str_Split2D(ByVal string1, DelimitreRow, DelimitreColumn, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant �ַ�����ֶ�ά����
'StrReg_Split(ByVal Expression, ByVal Pattern As Variant, Optional ByVal ignoreCase As Boolean = True) As Variant ������
'PinYin(Txt As Variant, Optional Delimiter = " ") As String  ��ƴ������������дƴ������ ע�������ֺ���Ƨ�֣����ܲ�׼
'PinYinInitial(Txt As Variant) As String  ƴ����ͷ
'StrFindSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long  �༭�������ƶ��㷨 �����ַ���˳�� ����FindStr��arrλ�� SimilarityΪ��С���ƶ�
'StrFindCosineSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long  �������ƶ��㷨 �����ַ���˳�� ����FindStr��arrλ�� SimilarityΪ��С���ƶ�
'StrSimilar(s1, s2) As Double  �༭�������ƶ��㷨 �ж��ַ���S1��S2�����ƶ�,�����ַ���˳��,���ƶ�����Ϊ0-100,100Ϊ��ȫһ��
'StrCosineSimilar(strA, strB) As Double  �������ƶ��㷨 �ж��ַ���S1��S2�����ƶ�,�����ַ���˳��,���ƶ�����Ϊ0-100,100Ϊ��ȫһ��
'StrRegexSearch( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef Index = 0, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant����ȡ����ֵ
 
'StrRegexSearchs( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant()  ����ȡ����ƥ�䣬��������
 
'StrRegexSearchOne( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As String  ����ȡ��һ��ֵ
 
'RegexInStr( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long  �������λ��
 
'StrRegexInStrRev( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long  �������λ�� ����
 
'StrRegexSearchSub( _
        ByRef string1, _
        ByRef Pattern, _
        Optional ByRef All As Boolean = True, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Variant() ����ȡ������ƥ�䣬�����������()�ٶ�ά����
 
'RegexCount( _
        ByRef string1, _
        ByRef Pattern, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Long  �������
 
'StrRegexTest( _
    ByRef string1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Boolean ������֤
 
'StrRegexReplace( _
    ByRef string1, _
    ByRef Pattern, _
    ByRef replacementString As String, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As String  �����滻
 
'StrFormatter(ByVal formatString, ParamArray textArray() As Variant) As String  ģ���ַ��� Formatter("������{1},���䣺{2}","UFO",18)  ����"������UFO,���䣺18"
'ByteToStr(arrByte, strCharset As String) As String ������ת��ָ��������ı� "Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
'StrToByte(strText As String, strCharset As String) �ı���ָ������תΪ������ "Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
'StrencodeURI(strText) As String  URLת��
'StrdecodeURI(strText) As String  URL����
'StrConvert(ByVal strText As String) As String unicode�ַ�ת��������
'StrencodeBase64(String1, Optional Charset = "") As String �ַ�������Base64
'StrdecodeBase64(String1, Optional Charset = "") As String �ַ�������Base64



'ϵͳ-------------------------------------------------------------------------------------------------------------------------------------
'Clipboard_GetData() As String  �������ȡ
'Clipboard_SetData(strData) As Boolean  ������д��
'Clipboard_ClearData() As Boolean  ���������
'UserName() As String  �û���
'UserDomain() As String  �û�������
'ComputerName() As String  �������


'�ļ�-------------------------------------------------------------------------------------------------------------------------------------
'TextRead(TextPath) As String  ��ȡtxt�ļ�(ANSI����)
'TextWrite(TextPath, str) As Boolean  д��txt�ļ�(ANSI����)
'TextAppend(TextPath, str) As Boolean ׷��txt�ļ�(ANSI����)
'TextRead2(TextPath, strCharset As String) As String  ��ȡtxt�ļ�(�Զ������) "Unicode", "GB2312", "UTF-7", "UTF-8", "ASCII", "GBK", "Big5", "unicodeFEFF", "unicodeFFFE"
'TextWrite2(TextPath, str, strCharset As String) As Boolean  д��txt�ļ�(�Զ������)
'TextAppend2(TextPath, str, strCharset As String) As Boolean  ׷��txt�ļ�(�Զ������)
'FileToByte(strFileName As String) As Byte() ���ļ�Ϊ�ֽ�����
'ByteToFile(arrByte, strFileName As String)  �ֽ�����ת�ļ�
'FolderExists(Path) As Boolean  �ļ����Ƿ����
'FolderDelete(Path) As Boolean  ɾ���ļ���
'FolderCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean  �����ļ���
'FolderCreate(Path) As Boolean  �����ļ��У����Դ����ϼ������ڵ��ļ��У������༶
'FolderSearch(pPath) As Variant  �����ļ������ļ���
'FolderSearchSub(pPath) As Variant �����ļ���(�����ļ���)
'FileExists(Path) As Boolean  �ļ��Ƿ����
'FileDelete(Path) As Boolean  ɾ���ļ�
'FileCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean �����ļ�
'FileSearch(pPath) As Variant �����ļ������ļ�
'FileSearchSub(pPath, Optional pMask As String = "") As Variant �����ļ������ļ�(�����ļ���) pPath������ʼ·����pMask���Ҫ����д,�ǵ�������д"*.xlsx",���Ǻ�


'·��-------------------------------------------------------------------------------------------------------------------------------------
'PathGetTemp() As String  ������ʱ·��
'PathGetMyDocuments() As String  �����ĵ�·��
'PathGetDesktop() As String  ��������·��
'PathBaseName(Path) As String  �����ļ�����������չ��
'PathFileName(Path) As String  �����ļ�����������չ��
'PathExtensionName(Path) As String  ������չ��������.
'PathParentFolderName(Path) As String  ����·��,ĩβ����\
'PathIsFolder(Path) As Boolean �ж��Ƿ����ļ���
'PathTempName() As String  ����ļ���
'PathNameSerialNumber(Name, Optional DelimiterLeft = "(", Optional DelimiterRight = ")") As String �����ظ�ʱ�����Ƽ���� Name��ǰ���� DelimiterLeft������ָ��� DelimiterRight����Ҳ�ָ���

'��Ԫ��-----------------------------------------------------------------------------------------------------------------------------------
'ColumnChr(ByVal v) As String  ����ת��ĸ
'ColumnChrArr(ParamArray arr()) As Variant  ����ת��ĸArr
'ColumnI(ByVal s) As Long  ��ĸת����
'ColumnIArr(ParamArray arr()) As Variant  ��ĸת����Arr
'UnionEx(ByRef Rngs) As Range  ��Ԫ�񲢼���չ,���뵥Ԫ������򼯺ϵ�Range���󣬺ϲ���Range
'UnionEx_Str(ByRef Rngs, sh) As Range  ��Ԫ�񲢼���չ,���뵥Ԫ������򼯺ϵ��ַ�����ַ���ϲ���Range
'SheetNew(wb As Workbook, Optional Name As String = "") As Worksheet  ĩβ����������
'SheetCopyAfter(sh, Optional Name As String = "") As Worksheet  ���ƹ�����ĩβ
'SheetCopyNow(sh, Optional Name As String = "") As Worksheet  ���ƹ������¹�����
'SheetIsName(wb As Workbook, ByVal Name As String) As Boolean  ��鹤�����Ƿ����
'WorkbookIsName(ByVal Name As String) As Boolean  ��鹤�����Ƿ���ڣ�Name��������׺
'ArrToRange(ByRef arr, ByVal rng)  ����д�빤����
'ArrToRangeUndo(ByRef arr, ByVal rng)  ����д�빤���������
'RangAddUndo(ByVal rng)  ��ӳ�������
'RangStartUndo()  �������� ����Ӻ�����
'RngResizeDownRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range ��Ԫ����������չ����
'RngResizeRightColumn(ByRef rng) As Range ��Ԫ����������չ����
'RngResizeEndRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range ��Ԫ�������һ����չ����
'RngResizeEndColumn(ByRef rng) As Range ��Ԫ�������һ����չ����
'RngDownRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long ��Ԫ������һ��
'RngRightColumn(ByRef rng As Range) As Long ��Ԫ������һ��
'RngEndRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long ��Ԫ�����һ��
'RngEndColumn(ByRef rng As Range) As Long ��Ԫ�����һ��
'RangeToArr(rng As Range) As Variant ��Ԫ��ֵ������,��֤һ����Ԫ��Ҳ������
'RngMerge_Empty(MergeRng As Range) ���ºϲ���ֵ��Ԫ��
'RngMerge_Repeat(MergeRng As Range) �ظ�ֵ�ϲ���Ԫ��
'RngAddBorders(rng As Range) �ӿ���
'RngAlignmentCenter(rng As Range) ��Ԫ����ж���
'SheetsSummary(Optional SelectName = "*", Optional RemoveName = "", Optional RngAddress = "", Optional wb As Workbook = Nothing) As Variant ���ܹ�����
'    ���ܹ����� SelectName�����Ĺ������� RemoveName�ų��Ĺ������� RngAddress��Ԫ������Ĭ��UsedRange  wb������Ĭ�ϵ�ǰ
'UCreatePivotTable(SourceData As Range, TableDestination As Range, TableName) As PivotTable��������͸�ӱ� SourceData����Դ��Ԫ�� TableDestination���õ�Ԫ�� TableName͸�ӱ�����
'USetPivotField(PTable As PivotTable, FieldName As String, Orientation As XlPivotFieldOrientation, _
        Position As Long, Optional Caption As String = "", Optional Fun As XlConsolidationFunction = xlCount)
'    ����͸�ӱ��ֶ� PTable͸�ӱ����UCreatePivotTable����ֵ  FieldName������
'    Orientation �ֶ�λ������ xlRowField(�б�ǩ) xlColumnField(�б�ǩ) xlDataField(����)
'    Position �ֶ�˳��
'    Caption  �ֶα���
'    Fun   Orientation=xlDataField(����)ʱ ���û��ܷ�ʽ��xlSum  xlCount  xlMin  xlMax

'FormatConditionAdd(Rng As Range, Formula, Color) As FormatCondition ����������ʽ  Rng������ʽ��Χ  Formula��ʽ  Color��ɫRGBֵ
'FormatConditionAdd_Pattern(Rng As Range, Formula, PatternColor, Optional Pattern As XlPattern = xlPatternGray50) As FormatCondition ����������ʽͼ��  Rng������ʽ��Χ  Formula��ʽ  PatternColor��ɫRGBֵ
'FormatConditionFind(Rng As Range, ByVal Formula) As FormatCondition ����ʽ����������ʽ
'FormatConditionFind_Color(Rng As Range, Color) As FormatCondition ����ɫ����������ʽ
'FormatConditionFind_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) As FormatCondition ��ͼ������������ʽ
'FormatConditionFindCount(Rng As Range, ByVal Formula) As Long ����ʽ����������ʽ����  ע��Formula:="=ROW($A1)=*"�Ǵ���д�� ������A1������A65536 ����Formula:="=ROW($A*)=*"
'FormatConditionFindCount_Color(Rng As Range, Color) As Long ����ɫ����������ʽ����
'FormatConditionFindCount_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) As Long ��ͼ������������ʽ����
'FormatConditionModify_Formula(FC As FormatCondition, Formula) ������ʽ�޸Ĺ�ʽ
'FormatConditionModify_Color(FC As FormatCondition, Color) ������ʽ�޸���ɫ
'FormatConditionModify_Pattern(FC As FormatCondition, Pattern As XlPattern, PatternColor) ������ʽ�޸�ͼ����ɫ
'FormatConditionModify_ClearColor(FC As FormatCondition) ������ʽ�����ɫ
'FormatConditionDelete(Rng As Range, ByVal Formula) ����ʽɾ��������ʽ ע��Formula:="=ROW($A1)=*"�Ǵ���д�� ������A1������A65536 ����Formula:="=ROW($A*)=*"
'FormatConditionDelete_Color(Rng As Range, Color) ����ɫɾ��������ʽ
'FormatConditionDelete_Pattern(Rng As Range, Pattern As XlPattern, PatternColor) ��ͼ��ɾ��������ʽ
'Rng_Validation(rng As Range, Formula, Optional ShowError As Boolean = True, Optional AlertStyle As XlDVAlertStyle = xlValidAlertStop) ������Ч�� rng��Ԫ�� Formula����"a,b,c" ShowError ��ʾ������ʾ���ҽ�ֹ���� AlertStyle������ʾ��ʽ
'RngAddComment(rng As Range, CommentText, Optional Visible As Boolean = False) As Comment �����ע
'RngAddPicture(PicturePath, rng As Range, Optional LowerWidth = 0, Optional LowerHeight = 0, Optional OriginalSizeRatio As Boolean = False) As Shape ���ͼƬ PicturePath����·�� rng��Ԫ�� LowerWidth��������� LowerHeight�߶������� OriginalSizeRatio�Ƿ�ԭ��С����


'��ѧ-------------------------------------------------------------------------------------------------------------------------------------
'SumParams(ParamArray arr()) As Double �������
'MaxParams(ParamArray arr()) As Double  ���������ֵ
'MinParams(ParamArray arr()) As Double  ��������Сֵ
'MaxParams2(Number1, Number2) As Double ����ȡ���ֵ Ч�ʸ�
'MinParams2(Number1, Number2) As Double ����ȡ��Сֵ Ч�ʸ�
'MultiplesUp(Number, Multiples) As Double ������������ı���
'MultiplesDown(Number, Multiples) As Double ������������ı���
'IntUp(Number) As Long ��������ȡ��
'IntDown(Number) As Long ��������ȡ��
'RoundUp(Number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double ��������
'RoundDown(Number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double ��������
'MultipleUp(Number, Significance) As Double ��������ָ�������ı���
'MultipleDown(Number, Significance) As Double ��������ָ�������ı���
'MultipleRound(Number, Significance) As Double ��������ָ�������ı���
'Float_Clear(Number) ������������㵼�µľ���ȱʧ
'RoundEX(number, Optional ByVal NumDigitsAfterDecimal As Long = 0) As Double �����������
'RandAddSub(Optional Number As Double = 1) As Double ��� +Number �� -Number
'ModNumber(Number1, Number2) As Double ����  ʮ�ڴ������಻����
'RandBetween(l, r) As Long ����Χ�����
'NumberSplit(Number, interval) As Variant  ������� Number��������� interval��ִ�С NumberSplit(5, 2)->[2,2,1]
'NumberLCase(NumberStr) As Double ���ִ�дתСд
'NumberUCase(Number) As String ����ת��д
'RMBLCase(NumberStr) As Currency �����Сд
'RMBUCase(curmoney) As String ����Ҵ�д
'NumberRangeInside(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean ����Ƚ� �ڲ�
'NumberRangeExternal(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean ����Ƚ� �ⲿ
'IsEven(Number) As Boolean �ж�ż��
'IsOdd(Number) As Boolean  �ж�����
'Number_Cycle(ByRef Number, ByRef CycleCount) As Long ѭ����� (i,3)->1,2,3,1,2,3,1,2,3
'Number_Repeat(ByRef Number, ByRef RepeatCount) As Long �ظ���� (i,3)->1,1,1,2,2,2,3,3,3
'Number_Separated(ByRef Number, ByRef SeparatedCount) As Long ������ (i,3)->1,4,7,10,13,16,19,22,25
'vbMaxNumber ���� ���ֵ
'vbMinNumber ���� ��Сֵ
'vbPi() As Double Pi��ֵ
'AngleToRadian(Angle) As Double �Ƕ�ת����
'RadianToAngle(Radian, Optional ByVal NumDigitsAfterDecimal = 3) As Double ����ת�Ƕ�




'����-------------------------------------------------------------------------------------------------------------------------------------
'Deconstruc(ParamArray DValue() As Variant, ByRef Value As Variant) �⹹ Deconstruc(����1, ����2, ����3) = Array(1, 2, 3)
'Cover(iValue, jValue) ��ֵ  iValue = jValue
'Exchange(iValue, jValue) ����
'ColToArr(ByRef col) As Variant   Col����ת����
'DictionaryCreate(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� itemΪ�������� �ظ�ֵ����ȡ��ǰ
'DictionaryCreateRev(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� itemΪ�������� ���� �ظ�ֵ����ȡ���
'DictionaryCreateIndex_ItemIsCol(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� �ظ�ֵ��ӵ���������
'DictionaryCreate_DicIndex(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� itemΪ�ֵ���������
'DictionaryCreate_Items(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� ˫���鵽�ֵ�
'DictionaryCreate_ItemsRev(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� ˫���鵽�ֵ� ����
'DictionaryCreate_ItemsIsCol(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object �����ֵ� ˫���鵽�ֵ� �ظ�ֵ��ӵ�����
'DictionaryToArr2D(dic) As Variant �ֵ䵽��ά���� 1����Key 2����Item
'DictionaryGetValues(dic, ByVal arrKey, Optional NoExistsValue = Empty) As Variant �ֵ�ȡ���ֵ  arrKey������һά��ά���鷵�ض�Ӧ��С��Itemֵ���� NoExistsValue��������ֵ
'DictionaryGetValuesParam(dic, ParamArray Keys()) As Variant �ֵ�ȡ���ֵ �����Key
'DictionaryExists(dic, ByVal arrKey) As Variant �ֵ��ж϶��ֵ arrKey������һά��ά���鷵�ض�Ӧ��С��True/False����
'DictionaryAdds(Dic, arrKeys, arrItems) As Object �ֵ�������� �ظ������޸�ԭ��ֵ
'DictionaryAddsRev(Dic, arrKeys, arrItems) As Object �ֵ�������� �ظ��򸲸�ԭ��ֵ
'DictionaryMerge(ParamArray Dics()) As Object �ֵ�ϲ�
'DictionaryMergeRev(ParamArray Dics()) As Object �ֵ�ϲ� ���� ���ظ������滻ǰ��
'Application_Attribute(bol As Boolean) Application_Attribute(False)�ر�һϵ��Ӱ��Ч������  **ע������������� Application_Attribute(True)**
'Sleep(PauseTime)  ������Ĳ�ռCPU�ӳ�,��λ����
'GetTimer() ���ؿ���ʱ�� ��λ����
'PrintEx(ByRef arg, Optional RowCount = 0, Optional DividerLine As Boolean = True) ��ӡ���� arg��ӡ���� RowCount��ӡ��������������  DividerLine�Ƿ��зָ���*��ͨ����Ĭ�ϲ���ӡΪFalseʱ�Ŵ�ӡ�ָ��ߣ���������Ĭ�ϴ�ӡΪFalseʱ����ӡ*
'encodeBase64(Bytes) As String ����Base64
'decodeBase64(String1) As Byte() ����Base64
'ImageSize(ImagePath) As Variant ͼƬ���ؿ���С  ����Array(Width, Height)
'LoadPictureEx(filename) As IPictureDisp ����LoadPicture ֧�ֶ���ͼƬ��ʽ
'CLngEx(Expression) As Variant ��չCLng ֧������ת��
'CDateEx(Expression) As Variant ��չCDate ֧������ת��
'CDblEx(Expression) As Variant ��չCDbl ֧������ת��
'CCurEx(Expression) As Variant ��չCCur ֧������ת��
'CStrEx(Expression) As Variant ��չCStr ֧������ת��
'CVarEx(Expression) As Variant ��չCVar ֧������ת��
'CBoolEx(Expression) As Variant ��չCBool ֧������ת��


'Http-------------------------------------------------------------------------------------------------------------------------------------
'HttpGet(Url, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Get����
'HttpDownload(Url, DownloadFileName, Optional RequestHeaderDic = Nothing) Get�����ļ�
'HttpPost(Url, Optional SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post����
'HttpPost_Form(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post���� ���ͱ�����
'HttpPost_Json(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant Post���� ����Json����
'HttpReadJson(Jsonstr As String, Routestr As String) As Variant ��ȡJSON����
Option Explicit
 
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long
    Private Declare PtrSafe Sub Sleep_ Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetTimer Lib "kernel32" Alias "GetTickCount" () As Long
#Else
    Private Declare Function WaitMessage Lib "user32" () As Long
    Private Declare Sub Sleep_ Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Public Declare Function GetTimer Lib "kernel32" Alias "GetTickCount" () As Long
#End If

'�ַ�������ȡֵģʽ
Public Enum SearchDirection
    LeftLeft = 1
    RightRight = 2
    LeftRight = 3
End Enum

'����ģʽ
Public Enum NumberRangeType
    Include_Exclude = 0
    Exclude_Include = 1
    Include_Include = 2
    Exclude_Exclude = 3
End Enum

'�ֵ�ƥ��ģʽ
Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum

'����ۺ�ģʽ
Public Enum GroupAggregateMethod
    Group_None = -1 - 9000000
    Group_First = -2 - 9000000
    Group_Last = -3 - 9000000
    Group_Sum = -4 - 9000000
    Group_Count = -5 - 9000000
    Group_CountNoEmpty = -6 - 9000000
    Group_Max = -7 - 9000000
    Group_Min = -8 - 9000000
    Group_Average = -9 - 9000000
    Group_Join = -10 - 9000000
    Group_CountClass = -11 - 9000000
End Enum

Public Const vbMaxNumber As Double = 1.79769313486231E+308 '���ֵ
Public Const vbMinNumber As Double = -1.79769313486231E+308  '��Сֵ

Private RangeUndoCollection_ As New Collection '�洢��Ԫ�񱸷�����
'����-------------------------------------------------------------------------------------------------------------------------------------
'��������
Private UFunction_Arr_SetGt_Temp_Cache_ As Variant
Private UFunction_Arr_SetGt_Temp_Cache1_ As Variant
Private UFunction_Arr_SetGt_Temp_Cache2_ As Variant
Private UFunction_Arr_SetGt_Temp_Cache3_ As Variant

'���⻺��
Private UFunction_Dictionary_Title_Temp_Cache_ As Object
Private UFunction_Dictionary_Title_Temp_Cache1_ As Object
Private UFunction_Dictionary_Title_Temp_Cache2_ As Object
Private UFunction_Dictionary_Title_Temp_Cache3_ As Object


'���������Currency @ ���� ����Ϊ�ǵ�n�� ת��Ϊ���� ����Ϊ���� 0Ϊ u + 1 ������빦��
Private Function IndexIsCurrencyToCount_(Index, l, u)
    Dim i As Long
    If IsArray(Index) Then
        For i = LBound(Index) To UBound(Index)
            If VarType(Index(i)) = vbCurrency Then
                If Index(i) > 0 Then
                    Index(i) = VBA.CLng(l + Index(i) - 1)
                Else
                    Index(i) = VBA.CLng(u + Index(i) + 1)
                End If
            End If
        Next
    Else
        If VarType(Index) = vbCurrency Then
            If Index > 0 Then
                Index = VBA.CLng(l + Index - 1)
            Else
                Index = VBA.CLng(u + Index + 1)
            End If
        End If
    End If
End Function

'�����������Currency @ ���� ����Ϊ������ת��Ϊ��n��
Private Function IndexIsLongToCount_(Index, l, u)
    Dim i As Long
    If IsArray(Index) Then
        For i = LBound(Index) To UBound(Index)
            If VarType(Index(i)) <> vbCurrency Then
                Index(i) = VBA.CLng(1 - l + Index(i))
            Else
                If Index(i) > 0 Then
                    Index(i) = VBA.CLng(Index(i))
                Else
                    Index(i) = VBA.CLng(1 - l + u + Index(i) + 1)
                End If
            End If
        Next
    Else
        If VarType(Index) <> vbCurrency Then
            Index = VBA.CLng(1 - l + Index)
        Else
            If Index > 0 Then
                Index = VBA.CLng(Index)
            Else
                Index = VBA.CLng(1 - l + u + Index + 1)
            End If
        End If
    End If
End Function

'�ڲ��ݹ����
Private Sub TitlesGetFlatten_(ByRef dic, ByRef TitleNames)
    Dim v
    For Each v In TitleNames
        If IsArray(v) Then
            TitlesGetFlatten_ dic, v
        ElseIf dic.Exists(v) Then
            ArrayDynamic_ dic(v)
        Else
            ArrayDynamic_ Empty
        End If
    Next
End Sub

'�������ȡֵ ��������
Public Property Get Titles(ParamArray TitleNames()) As Variant
    Dim v
    If Not UFunction_Dictionary_Title_Temp_Cache_ Is Nothing Then
        ArrayDynamic_
        For Each v In TitleNames
            If IsArray(v) Then
                TitlesGetFlatten_ UFunction_Dictionary_Title_Temp_Cache_, v
            ElseIf UFunction_Dictionary_Title_Temp_Cache_.Exists(v) Then
                ArrayDynamic_ UFunction_Dictionary_Title_Temp_Cache_(v)
            Else
                ArrayDynamic_ Empty
            End If
        Next
    End If
    Titles = ArrayDynamic_
End Property

'�������ȡֵһ��ֵ  Title!����
Public Property Get Title() As Object
    If UFunction_Dictionary_Title_Temp_Cache_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache_ = CreateObject("scripting.Dictionary")
        UFunction_Dictionary_Title_Temp_Cache_.CompareMode = TextCompare
    End If
    Set Title = UFunction_Dictionary_Title_Temp_Cache_
End Property

'������⸳ֵ
Public Property Let Titles(ParamArray TitleNames(), ByRef TitleIndexs As Variant)
    If UFunction_Dictionary_Title_Temp_Cache_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache_ = DictionaryCreate_ItemsRev(ArrFlatten(TitleNames), ArrFlatten(TitleIndexs), TextCompare)
    Else
        DictionaryAddsRev UFunction_Dictionary_Title_Temp_Cache_, ArrFlatten(TitleNames), ArrFlatten(TitleIndexs)
    End If
End Property

'�������ȡֵ1
Public Property Get Titles1(ParamArray TitleNames()) As Variant
    Dim v
    If Not UFunction_Dictionary_Title_Temp_Cache1_ Is Nothing Then
        ArrayDynamic_
        For Each v In TitleNames
            If IsArray(v) Then
                TitlesGetFlatten_ UFunction_Dictionary_Title_Temp_Cache1_, v
            ElseIf UFunction_Dictionary_Title_Temp_Cache1_.Exists(v) Then
                ArrayDynamic_ UFunction_Dictionary_Title_Temp_Cache1_(v)
            Else
                ArrayDynamic_ Empty
            End If
        Next
    End If
    Titles1 = ArrayDynamic_
End Property

Public Property Get Title1() As Object
    If UFunction_Dictionary_Title_Temp_Cache1_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache1_ = CreateObject("scripting.Dictionary")
        UFunction_Dictionary_Title_Temp_Cache1_.CompareMode = TextCompare
    End If
    Set Title1 = UFunction_Dictionary_Title_Temp_Cache1_
End Property

'������⸳ֵ1
Public Property Let Titles1(ParamArray TitleNames(), ByRef TitleIndexs As Variant)
    If UFunction_Dictionary_Title_Temp_Cache1_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache1_ = DictionaryCreate_ItemsRev(ArrFlatten(TitleNames), ArrFlatten(TitleIndexs), TextCompare)
    Else
        DictionaryAddsRev UFunction_Dictionary_Title_Temp_Cache1_, ArrFlatten(TitleNames), ArrFlatten(TitleIndexs)
    End If
End Property

'�������ȡֵ2
Public Property Get Titles2(ParamArray TitleNames()) As Variant
    Dim v
    If Not UFunction_Dictionary_Title_Temp_Cache2_ Is Nothing Then
        ArrayDynamic_
        For Each v In TitleNames
            If IsArray(v) Then
                TitlesGetFlatten_ UFunction_Dictionary_Title_Temp_Cache2_, v
            ElseIf UFunction_Dictionary_Title_Temp_Cache2_.Exists(v) Then
                ArrayDynamic_ UFunction_Dictionary_Title_Temp_Cache2_(v)
            Else
                ArrayDynamic_ Empty
            End If
        Next
    End If
    Titles2 = ArrayDynamic_
End Property

Public Property Get Title2() As Object
    If UFunction_Dictionary_Title_Temp_Cache2_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache2_ = CreateObject("scripting.Dictionary")
        UFunction_Dictionary_Title_Temp_Cache2_.CompareMode = TextCompare
    End If
    Set Title2 = UFunction_Dictionary_Title_Temp_Cache2_
End Property

'������⸳ֵ2
Public Property Let Titles2(ParamArray TitleNames(), ByRef TitleIndexs As Variant)
    If UFunction_Dictionary_Title_Temp_Cache2_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache2_ = DictionaryCreate_ItemsRev(ArrFlatten(TitleNames), ArrFlatten(TitleIndexs), TextCompare)
    Else
        DictionaryAddsRev UFunction_Dictionary_Title_Temp_Cache2_, ArrFlatten(TitleNames), ArrFlatten(TitleIndexs)
    End If
End Property

'�������ȡֵ3
Public Property Get Titles3(ParamArray TitleNames()) As Variant
    Dim v
    If Not UFunction_Dictionary_Title_Temp_Cache3_ Is Nothing Then
        ArrayDynamic_
        For Each v In TitleNames
            If IsArray(v) Then
                TitlesGetFlatten_ UFunction_Dictionary_Title_Temp_Cache3_, v
            ElseIf UFunction_Dictionary_Title_Temp_Cache3_.Exists(v) Then
                ArrayDynamic_ UFunction_Dictionary_Title_Temp_Cache3_(v)
            Else
                ArrayDynamic_ Empty
            End If
        Next
    End If
    Titles3 = ArrayDynamic_
End Property

Public Property Get Title3() As Object
    If UFunction_Dictionary_Title_Temp_Cache3_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache3_ = CreateObject("scripting.Dictionary")
        UFunction_Dictionary_Title_Temp_Cache3_.CompareMode = TextCompare
    End If
    Set Title3 = UFunction_Dictionary_Title_Temp_Cache3_
End Property

'������⸳ֵ3
Public Property Let Titles3(ParamArray TitleNames(), ByRef TitleIndexs As Variant)
    If UFunction_Dictionary_Title_Temp_Cache3_ Is Nothing Then
        Set UFunction_Dictionary_Title_Temp_Cache3_ = DictionaryCreate_ItemsRev(ArrFlatten(TitleNames), ArrFlatten(TitleIndexs), TextCompare)
    Else
        DictionaryAddsRev UFunction_Dictionary_Title_Temp_Cache3_, ArrFlatten(TitleNames), ArrFlatten(TitleIndexs)
    End If
End Property


'��������ȡֵ
Public Property Get ArrCache(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) As Variant
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        '��������ȡ����
        ArrCache = UFunction_Arr_SetGt_Temp_Cache_
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        '��RowIndex
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache_)
            Case 1
                If IsArray(RowIndex) Then
                    '������һά ȡ���������Ӧֵ����
                    ArrCache = ArrFromIndex(UFunction_Arr_SetGt_Temp_Cache_, RowIndex)
                Else
                    '������һά ȡֵ
                    Cover ArrCache, UFunction_Arr_SetGt_Temp_Cache_(RowIndex)
                End If
            Case 2
                If IsArray(RowIndex) Then
                    '�����Ƕ�ά ȡ����
                    ArrCache = ArrGetRows(UFunction_Arr_SetGt_Temp_Cache_, RowIndex)
                Else
                    '�����Ƕ�ά ȡһ��
                    ArrCache = ArrGetRow(UFunction_Arr_SetGt_Temp_Cache_, RowIndex, 1, Expansion)
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        '��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(ColumnIndex) Then
            'ȡ����
            ArrCache = ArrGetColumns(UFunction_Arr_SetGt_Temp_Cache_, ColumnIndex)
        Else
            'ȡһ��
            ArrCache = ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache_, ColumnIndex, 1, Expansion)
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        '��RowIndex��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(RowIndex) And IsArray(ColumnIndex) Then
            '������������ȡ�����������򷵻ض�ά����
            ArrCache = ArrGetColumns(ArrGetRows(UFunction_Arr_SetGt_Temp_Cache_, RowIndex), ColumnIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) Then
            'ColumnIndex������  ȡRowIndex�е�ColumnIndex������ֵ  ����һά����
            ArrCache = ArrFromIndex(ArrGetRow(UFunction_Arr_SetGt_Temp_Cache_, RowIndex, 1, Expansion), ColumnIndex)
        ElseIf IsArray(RowIndex) And IsArray(ColumnIndex) = False Then
            'RowIndex������  ȡColumnIndex�е�RowIndex������ֵ  ����һά����
            ArrCache = ArrFromIndex(ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache_, ColumnIndex, 1, Expansion), RowIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) = False Then
            '����������ȡ����ֵ
            Cover ArrCache, UFunction_Arr_SetGt_Temp_Cache_(RowIndex, ColumnIndex)
        End If
    End If
End Property
 
'�������鸳ֵ
Public Property Let ArrCache(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False, ByRef arr As Variant)
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        '��������ֱ��д�뻺������
        UFunction_Arr_SetGt_Temp_Cache_ = arr
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        '��RowIndex
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache_)
            Case 1
                If IsArray(RowIndex) Then
                    '�����������򣬰�����λ������д��
                    ArrSetValues UFunction_Arr_SetGt_Temp_Cache_, RowIndex, arr
                ElseIf IsArray(arr) Then
                    '������һά ��RowIndex������ʼ����д��arr
                    ArrSetArr UFunction_Arr_SetGt_Temp_Cache_, arr, RowIndex, Expansion
                Else
                    '������һά RowIndex����λ�õ�ֵ�޸�Ϊarr
                    Cover UFunction_Arr_SetGt_Temp_Cache_(RowIndex), arr
                End If
            Case 2
                If IsArray(RowIndex) Then
                    '�����������򣬰�����λ������д��
                    ArrSetEntireRowValues UFunction_Arr_SetGt_Temp_Cache_, RowIndex, arr
                Else
                    '�����Ƕ�ά ��RowIndex�е�1������д��arr
                    ArrSetRow UFunction_Arr_SetGt_Temp_Cache_, arr, RowIndex, Expansion
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        '��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(ColumnIndex) Then
            '�����������򣬰�����λ������д��
            ArrSetEntireColumnValues UFunction_Arr_SetGt_Temp_Cache_, ColumnIndex, arr
        Else
            '��ColumnIndex�е�1������д��arr
            ArrSetColumn UFunction_Arr_SetGt_Temp_Cache_, arr, ColumnIndex, Expansion
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        '��RowIndex��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(RowIndex) Or IsArray(ColumnIndex) Then
            '�����������򣬰�����λ�� ���ϵ���һ��һ������д��
            Arr2DSetValues UFunction_Arr_SetGt_Temp_Cache_, RowIndex, ColumnIndex, arr
        ElseIf IsArray(arr) Then
            'arr������ ��RowIndex��ColumnIndex�п�ʼ����д��arr
            Select Case ArrDimension(arr)
                Case 1
                    'arr��һά����������д��
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache_, ArrTranspose(arr), RowIndex, ColumnIndex, Expansion
                Case 2
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache_, arr, RowIndex, ColumnIndex, Expansion
            End Select
        Else
            'arr��������ֱ���޸� RowIndex��ColumnIndex��λ�õ�ֵΪarr
            Cover UFunction_Arr_SetGt_Temp_Cache_(RowIndex, ColumnIndex), arr
        End If
    End If
End Property

Public Property Get ArrCache1(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) As Variant
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        ArrCache1 = UFunction_Arr_SetGt_Temp_Cache1_
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache1_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrCache1 = ArrFromIndex(UFunction_Arr_SetGt_Temp_Cache1_, RowIndex)
                Else
                    Cover ArrCache1, UFunction_Arr_SetGt_Temp_Cache1_(RowIndex)
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrCache1 = ArrGetRows(UFunction_Arr_SetGt_Temp_Cache1_, RowIndex)
                Else
                    ArrCache1 = ArrGetRow(UFunction_Arr_SetGt_Temp_Cache1_, RowIndex, 1, Expansion)
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrCache1 = ArrGetColumns(UFunction_Arr_SetGt_Temp_Cache1_, ColumnIndex)
        Else
            ArrCache1 = ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache1_, ColumnIndex, 1, Expansion)
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) And IsArray(ColumnIndex) Then
            ArrCache1 = ArrGetColumns(ArrGetRows(UFunction_Arr_SetGt_Temp_Cache1_, RowIndex), ColumnIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) Then
            ArrCache1 = ArrFromIndex(ArrGetRow(UFunction_Arr_SetGt_Temp_Cache1_, RowIndex, 1, Expansion), ColumnIndex)
        ElseIf IsArray(RowIndex) And IsArray(ColumnIndex) = False Then
            ArrCache1 = ArrFromIndex(ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache1_, ColumnIndex, 1, Expansion), RowIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) = False Then
            Cover ArrCache1, UFunction_Arr_SetGt_Temp_Cache1_(RowIndex, ColumnIndex)
        End If
    End If
End Property

Public Property Let ArrCache1(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False, ByRef arr As Variant)
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        UFunction_Arr_SetGt_Temp_Cache1_ = arr
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache1_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrSetValues UFunction_Arr_SetGt_Temp_Cache1_, RowIndex, arr
                ElseIf IsArray(arr) Then
                    ArrSetArr UFunction_Arr_SetGt_Temp_Cache1_, arr, RowIndex, Expansion
                Else
                    Cover UFunction_Arr_SetGt_Temp_Cache1_(RowIndex), arr
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrSetEntireRowValues UFunction_Arr_SetGt_Temp_Cache1_, RowIndex, arr
                Else
                    ArrSetRow UFunction_Arr_SetGt_Temp_Cache1_, arr, RowIndex, Expansion
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrSetEntireColumnValues UFunction_Arr_SetGt_Temp_Cache1_, ColumnIndex, arr
        Else
            ArrSetColumn UFunction_Arr_SetGt_Temp_Cache1_, arr, ColumnIndex, Expansion
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) Or IsArray(ColumnIndex) Then
            Arr2DSetValues UFunction_Arr_SetGt_Temp_Cache1_, RowIndex, ColumnIndex, arr
        ElseIf IsArray(arr) Then
            Select Case ArrDimension(arr)
                Case 1
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache1_, ArrTranspose(arr), RowIndex, ColumnIndex, Expansion
                Case 2
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache1_, arr, RowIndex, ColumnIndex, Expansion
            End Select
        Else
            Cover UFunction_Arr_SetGt_Temp_Cache1_(RowIndex, ColumnIndex), arr
        End If
    End If
End Property

Public Property Get ArrCache2(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) As Variant
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        ArrCache2 = UFunction_Arr_SetGt_Temp_Cache2_
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache2_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrCache2 = ArrFromIndex(UFunction_Arr_SetGt_Temp_Cache2_, RowIndex)
                Else
                    Cover ArrCache2, UFunction_Arr_SetGt_Temp_Cache2_(RowIndex)
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrCache2 = ArrGetRows(UFunction_Arr_SetGt_Temp_Cache2_, RowIndex)
                Else
                    ArrCache2 = ArrGetRow(UFunction_Arr_SetGt_Temp_Cache2_, RowIndex, 1, Expansion)
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrCache2 = ArrGetColumns(UFunction_Arr_SetGt_Temp_Cache2_, ColumnIndex)
        Else
            ArrCache2 = ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache2_, ColumnIndex, 1, Expansion)
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) And IsArray(ColumnIndex) Then
            ArrCache2 = ArrGetColumns(ArrGetRows(UFunction_Arr_SetGt_Temp_Cache2_, RowIndex), ColumnIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) Then
            ArrCache2 = ArrFromIndex(ArrGetRow(UFunction_Arr_SetGt_Temp_Cache2_, RowIndex, 1, Expansion), ColumnIndex)
        ElseIf IsArray(RowIndex) And IsArray(ColumnIndex) = False Then
            ArrCache2 = ArrFromIndex(ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache2_, ColumnIndex, 1, Expansion), RowIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) = False Then
            Cover ArrCache2, UFunction_Arr_SetGt_Temp_Cache2_(RowIndex, ColumnIndex)
        End If
    End If
End Property

Public Property Let ArrCache2(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False, ByRef arr As Variant)
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        UFunction_Arr_SetGt_Temp_Cache2_ = arr
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache2_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrSetValues UFunction_Arr_SetGt_Temp_Cache2_, RowIndex, arr
                ElseIf IsArray(arr) Then
                    ArrSetArr UFunction_Arr_SetGt_Temp_Cache2_, arr, RowIndex, Expansion
                Else
                    Cover UFunction_Arr_SetGt_Temp_Cache2_(RowIndex), arr
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrSetEntireRowValues UFunction_Arr_SetGt_Temp_Cache2_, RowIndex, arr
                Else
                    ArrSetRow UFunction_Arr_SetGt_Temp_Cache2_, arr, RowIndex, Expansion
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrSetEntireColumnValues UFunction_Arr_SetGt_Temp_Cache2_, ColumnIndex, arr
        Else
            ArrSetColumn UFunction_Arr_SetGt_Temp_Cache2_, arr, ColumnIndex, Expansion
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) Or IsArray(ColumnIndex) Then
            Arr2DSetValues UFunction_Arr_SetGt_Temp_Cache2_, RowIndex, ColumnIndex, arr
        ElseIf IsArray(arr) Then
            Select Case ArrDimension(arr)
                Case 1
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache2_, ArrTranspose(arr), RowIndex, ColumnIndex, Expansion
                Case 2
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache2_, arr, RowIndex, ColumnIndex, Expansion
            End Select
        Else
            Cover UFunction_Arr_SetGt_Temp_Cache2_(RowIndex, ColumnIndex), arr
        End If
    End If
End Property

Public Property Get ArrCache3(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) As Variant
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        ArrCache3 = UFunction_Arr_SetGt_Temp_Cache3_
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache3_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrCache3 = ArrFromIndex(UFunction_Arr_SetGt_Temp_Cache3_, RowIndex)
                Else
                    Cover ArrCache3, UFunction_Arr_SetGt_Temp_Cache3_(RowIndex)
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrCache3 = ArrGetRows(UFunction_Arr_SetGt_Temp_Cache3_, RowIndex)
                Else
                    ArrCache3 = ArrGetRow(UFunction_Arr_SetGt_Temp_Cache3_, RowIndex, 1, Expansion)
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrCache3 = ArrGetColumns(UFunction_Arr_SetGt_Temp_Cache3_, ColumnIndex)
        Else
            ArrCache3 = ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache3_, ColumnIndex, 1, Expansion)
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) And IsArray(ColumnIndex) Then
            ArrCache3 = ArrGetColumns(ArrGetRows(UFunction_Arr_SetGt_Temp_Cache3_, RowIndex), ColumnIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) Then
            ArrCache3 = ArrFromIndex(ArrGetRow(UFunction_Arr_SetGt_Temp_Cache3_, RowIndex, 1, Expansion), ColumnIndex)
        ElseIf IsArray(RowIndex) And IsArray(ColumnIndex) = False Then
            ArrCache3 = ArrFromIndex(ArrGetColumn(UFunction_Arr_SetGt_Temp_Cache3_, ColumnIndex, 1, Expansion), RowIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) = False Then
            Cover ArrCache3, UFunction_Arr_SetGt_Temp_Cache3_(RowIndex, ColumnIndex)
        End If
    End If
End Property

Public Property Let ArrCache3(Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False, ByRef arr As Variant)
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        UFunction_Arr_SetGt_Temp_Cache3_ = arr
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        Select Case ArrDimension(UFunction_Arr_SetGt_Temp_Cache3_)
            Case 1
                If IsArray(RowIndex) Then
                    ArrSetValues UFunction_Arr_SetGt_Temp_Cache3_, RowIndex, arr
                ElseIf IsArray(arr) Then
                    ArrSetArr UFunction_Arr_SetGt_Temp_Cache3_, arr, RowIndex, Expansion
                Else
                    Cover UFunction_Arr_SetGt_Temp_Cache3_(RowIndex), arr
                End If
            Case 2
                If IsArray(RowIndex) Then
                    ArrSetEntireRowValues UFunction_Arr_SetGt_Temp_Cache3_, RowIndex, arr
                Else
                    ArrSetRow UFunction_Arr_SetGt_Temp_Cache3_, arr, RowIndex, Expansion
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        If IsArray(ColumnIndex) Then
            ArrSetEntireColumnValues UFunction_Arr_SetGt_Temp_Cache3_, ColumnIndex, arr
        Else
            ArrSetColumn UFunction_Arr_SetGt_Temp_Cache3_, arr, ColumnIndex, Expansion
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        If IsArray(RowIndex) Or IsArray(ColumnIndex) Then
            Arr2DSetValues UFunction_Arr_SetGt_Temp_Cache3_, RowIndex, ColumnIndex, arr
        ElseIf IsArray(arr) Then
            Select Case ArrDimension(arr)
                Case 1
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache3_, ArrTranspose(arr), RowIndex, ColumnIndex, Expansion
                Case 2
                    Arr2DSetArr2D UFunction_Arr_SetGt_Temp_Cache3_, arr, RowIndex, ColumnIndex, Expansion
            End Select
        Else
            Cover UFunction_Arr_SetGt_Temp_Cache3_(RowIndex, ColumnIndex), arr
        End If
    End If
End Property

'�������򸴺ϲ��� ����ȡֵ ����ArrCache
Public Property Get ArrBlend(ByRef arrC, Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False) As Variant
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        '��������ȡ����
        ArrBlend = arrC
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        '��RowIndex
        Select Case ArrDimension(arrC)
            Case 1
                If IsArray(RowIndex) Then
                    '������һά ȡ���������Ӧֵ����
                    ArrBlend = ArrFromIndex(arrC, RowIndex)
                Else
                    '������һά ȡֵ
                    Cover ArrBlend, arrC(RowIndex)
                End If
            Case 2
                If IsArray(RowIndex) Then
                    '�����Ƕ�ά ȡ����
                    ArrBlend = ArrGetRows(arrC, RowIndex)
                Else
                    '�����Ƕ�ά ȡһ��
                    ArrBlend = ArrGetRow(arrC, RowIndex, 1, Expansion)
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        '��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(ColumnIndex) Then
            'ȡ����
            ArrBlend = ArrGetColumns(arrC, ColumnIndex)
        Else
            'ȡһ��
            ArrBlend = ArrGetColumn(arrC, ColumnIndex, 1, Expansion)
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        '��RowIndex��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(RowIndex) And IsArray(ColumnIndex) Then
            '������������ȡ�����������򷵻ض�ά����
            ArrBlend = ArrGetColumns(ArrGetRows(arrC, RowIndex), ColumnIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) Then
            'ColumnIndex������  ȡRowIndex�е�ColumnIndex������ֵ  ����һά����
            ArrBlend = ArrFromIndex(ArrGetRow(arrC, RowIndex, 1, Expansion), ColumnIndex)
        ElseIf IsArray(RowIndex) And IsArray(ColumnIndex) = False Then
            'RowIndex������  ȡColumnIndex�е�RowIndex������ֵ  ����һά����
            ArrBlend = ArrFromIndex(ArrGetColumn(arrC, ColumnIndex, 1, Expansion), RowIndex)
        ElseIf IsArray(RowIndex) = False And IsArray(ColumnIndex) = False Then
            '����������ȡ����ֵ
            Cover ArrBlend, arrC(RowIndex, ColumnIndex)
        End If
    End If
End Property
 
 '�������򸴺ϲ��� ���鸳ֵ
Public Property Let ArrBlend(ByRef arrC, Optional ByRef RowIndex, Optional ByRef ColumnIndex, Optional Expansion As Boolean = False, ByRef arr As Variant)
    If IsMissing(RowIndex) And IsMissing(ColumnIndex) Then
        '��������ֱ��д�뻺������
        arrC = arr
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) Then
        '��RowIndex
        Select Case ArrDimension(arrC)
            Case 1
                If IsArray(RowIndex) Then
                    '�����������򣬰�����λ������д��
                    ArrSetValues arrC, RowIndex, arr
                ElseIf IsArray(arr) Then
                    '������һά ��RowIndex������ʼ����д��arr
                    ArrSetArr arrC, arr, RowIndex, Expansion
                Else
                    '������һά RowIndex����λ�õ�ֵ�޸�Ϊarr
                    Cover arrC(RowIndex), arr
                End If
            Case 2
                If IsArray(RowIndex) Then
                    '�����������򣬰�����λ������д��
                    ArrSetEntireRowValues arrC, RowIndex, arr
                Else
                    '�����Ƕ�ά ��RowIndex�е�1������д��arr
                    ArrSetRow arrC, arr, RowIndex, Expansion
                End If
        End Select
    ElseIf IsMissing(RowIndex) And IsMissing(ColumnIndex) = False Then
        '��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(ColumnIndex) Then
            '�����������򣬰�����λ������д��
            ArrSetEntireColumnValues arrC, ColumnIndex, arr
        Else
            '��ColumnIndex�е�1������д��arr
            ArrSetColumn arrC, arr, ColumnIndex, Expansion
        End If
    ElseIf IsMissing(RowIndex) = False And IsMissing(ColumnIndex) = False Then
        '��RowIndex��ColumnIndex  ��Ϊ����һ���Ƕ�ά����
        If IsArray(RowIndex) Or IsArray(ColumnIndex) Then
            '�����������򣬰�����λ�� ���ϵ���һ��һ������д��
            Arr2DSetValues arrC, RowIndex, ColumnIndex, arr
        ElseIf IsArray(arr) Then
            'arr������ ��RowIndex��ColumnIndex�п�ʼ����д��arr
            Select Case ArrDimension(arr)
                Case 1
                    'arr��һά����������д��
                    Arr2DSetArr2D arrC, ArrTranspose(arr), RowIndex, ColumnIndex, Expansion
                Case 2
                    Arr2DSetArr2D arrC, arr, RowIndex, ColumnIndex, Expansion
            End Select
        Else
            'arr��������ֱ���޸� RowIndex��ColumnIndex��λ�õ�ֵΪarr
            Cover arrC(RowIndex, ColumnIndex), arr
        End If
    End If
End Property

'����ȡֵ��������Ԫ�ص�RowCount,ColumnCount��ȡ,�������޷���EmptyContent
'��������ʱ��Զ����arr,����Ԫ������Ϊ1ʱ��Զ�������Ԫ�أ�����Ϊһ������ʱֻ����ColumnCount RowCount��=1������Ϊһ�л�һά����ʱֻ����RowCount ColumnCount��=1
Public Function ArrGetValue(arr, ByVal RowCount, Optional ByVal ColumnCount, Optional EmptyContent = "") As Variant
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(arr)
        Case 1
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            If u1 - l1 = 0 Then
                Cover ArrGetValue, arr(l1)
            ElseIf RowCount + l1 - 1 <= u1 Then
                Cover ArrGetValue, arr(RowCount + l1 - 1)
            Else
                Cover ArrGetValue, EmptyContent
            End If
        Case 2
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            If u1 - l1 = 0 Then RowCount = 1 'һ����Զȡһ��
            If u2 - l2 = 0 Then ColumnCount = 1 'һ����Զȡһ��
            If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                Cover ArrGetValue, arr(RowCount + l1 - 1, ColumnCount + l2 - 1)
            Else
                Cover ArrGetValue, EmptyContent
            End If
        Case 0 '����������Զȡ���ֵ
            Cover ArrGetValue, arr
    End Select
End Function

Private Function ArrGetValue_(arr, ByVal RowCount, Optional ByVal ColumnCount, Optional EmptyContent = "") As Variant
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(arr)
        Case 1
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            If u1 - l1 = 0 Then
                Cover ArrGetValue_, arr(l1)
            ElseIf RowCount + l1 - 1 <= u1 Then
                Cover ArrGetValue_, arr(RowCount + l1 - 1)
            Else
                Cover ArrGetValue_, EmptyContent
            End If
        Case 2
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            If u1 - l1 = 0 Then RowCount = 1 'һ����Զȡһ��
            If u2 - l2 = 0 Then ColumnCount = 1 'һ����Զȡһ��
            If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                Cover ArrGetValue_, arr(RowCount + l1 - 1, ColumnCount + l2 - 1)
            Else
                Cover ArrGetValue_, EmptyContent
            End If
        Case 0 '����������Զȡ���ֵ
            Cover ArrGetValue_, arr
    End Select
End Function

'����ȡֵ���� ͬArrGetValue ��ͬ����arr,EmptyContentд�뺯�������� ���ټ���ӿ��ȡ�ٶ�
'WriteArr=Trueʱд��arr���� WriteArr=Falseʱ����RowCount,ColumnCount��ȡ��������
'���û�������ʾ����ArrGetValueCache WriteArr:=True, arr:=arr, EmptyContent:=""
'��ȡ��������ʾ����v = ArrGetValueCache(i, j)
Public Function ArrGetValueCache(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache, arrRE(l1, l2)
        End Select
    End If
End Function

Public Function ArrGetValueCache1(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache1, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache1, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache1, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache1, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache1, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache1, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache1, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache1, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache1, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache1, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache1, arrRE(l1, l2)
        End Select
    End If
End Function

Public Function ArrGetValueCache2(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache2, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache2, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache2, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache2, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache2, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache2, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache2, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache2, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache2, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache2, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache2, arrRE(l1, l2)
        End Select
    End If
End Function

Public Function ArrGetValueCache3(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache3, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache3, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache3, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache3, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache3, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache3, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache3, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache3, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache3, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache3, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache3, arrRE(l1, l2)
        End Select
    End If
End Function

Public Function ArrGetValueCache4(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache4, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache4, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache4, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache4, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache4, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache4, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache4, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache4, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache4, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache4, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache4, arrRE(l1, l2)
        End Select
    End If
End Function

Public Function ArrGetValueCache5(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache5, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache5, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache5, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache5, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache5, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache5, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache5, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache5, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache5, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache5, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache5, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache_, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache1_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache1_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache1_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache1_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache1_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache1_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache1_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache1_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache1_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache1_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache1_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache1_, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache2_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache2_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache2_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache2_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache2_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache2_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache2_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache2_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache2_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache2_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache2_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache2_, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache3_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache3_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache3_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache3_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache3_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache3_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache3_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache3_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache3_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache3_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache3_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache3_, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache4_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache4_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache4_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache4_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache4_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache4_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache4_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache4_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache4_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache4_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache4_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache4_, arrRE(l1, l2)
        End Select
    End If
End Function

Private Function ArrGetValueCache5_(Optional ByVal RowCount, Optional ByVal ColumnCount, Optional WriteArr As Boolean = False, Optional arr, Optional EmptyContent = "") As Variant
    Static l1 As Long, u1 As Long
    Static l2 As Long, u2 As Long
    Static ArrDimension1 As Long, arrRE, EmptyContentRE
    If WriteArr Then
        Cover EmptyContentRE, EmptyContent
        ArrDimension1 = ArrDimension(arr)
        Select Case ArrDimension1
            Case 1
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                If u1 - l1 = 0 Then ArrDimension1 = -1
                arrRE = arr
            Case 2
                l1 = LBound(arr, 1): u1 = UBound(arr, 1)
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                If u1 - l1 = 0 Then ArrDimension1 = -3 'RowIndex = l1
                If u2 - l2 = 0 Then ArrDimension1 = -4 'ColumnIndex = l2
                If u1 - l1 = 0 And u2 - l2 = 0 Then ArrDimension1 = -5
                arrRE = arr
            Case Else
                Cover arrRE, arr
        End Select
    Else
        Select Case ArrDimension1
            Case 1 'һά��������ȡֵ
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache5_, arrRE(RowCount + l1 - 1)
                Else
                    Cover ArrGetValueCache5_, EmptyContentRE
                End If
            Case 2 '��ά��������ȡֵ
                If RowCount + l1 - 1 <= u1 And ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache5_, arrRE(RowCount + l1 - 1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache5_, EmptyContentRE
                End If
            Case 0 '����������Զȡ���ֵ
                Cover ArrGetValueCache5_, arrRE
            Case -3 'һ����Զȡһ��
                If ColumnCount + l2 - 1 <= u2 Then
                    Cover ArrGetValueCache5_, arrRE(l1, ColumnCount + l2 - 1)
                Else
                    Cover ArrGetValueCache5_, EmptyContentRE
                End If
            Case -4 'һ����Զȡһ��
                If RowCount + l1 - 1 <= u1 Then
                    Cover ArrGetValueCache5_, arrRE(RowCount + l1 - 1, l2)
                Else
                    Cover ArrGetValueCache5_, EmptyContentRE
                End If
            Case -1 'һά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache5_, arrRE(l1)
            Case -5 '��ά����ֻ��һ��Ԫ����Զȡ���Ԫ��
                Cover ArrGetValueCache5_, arrRE(l1, l2)
        End Select
    End If
End Function

'��������ӣ���������ȡֵ���ʼ��
Public Function ArrayDynamic(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic = arr
        Else
            ArrayDynamic = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic = i
    i = i + 1
End Function

Public Function ArrayDynamic1(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic1 = arr
        Else
            ArrayDynamic1 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic1 = i
    i = i + 1
End Function

Public Function ArrayDynamic2(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic2 = arr
        Else
            ArrayDynamic2 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic2 = i
    i = i + 1
End Function

Public Function ArrayDynamic3(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic3 = arr
        Else
            ArrayDynamic3 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic3 = i
    i = i + 1
End Function
 
'�ڲ�ArrayDynamic
Private Function ArrayDynamic_(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic_ = arr
        Else
            ArrayDynamic_ = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic_ = i
    i = i + 1
End Function
 
'�ڲ�ArrayDynamic
Private Function ArrayDynamic2_(Optional ByRef v) As Variant
    Static arr(), i As Long
    Const init = 20
    If IsMissing(v) And IsError(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To i - 1)
            ArrayDynamic2_ = arr
        Else
            ArrayDynamic2_ = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To init)
        i = 1
    ElseIf i > UBound(arr) Then
        ReDim Preserve arr(1 To UBound(arr) * 2)
    End If
    If VBA.IsObject(v) Then
        Set arr(i) = v
    Else
        arr(i) = v
    End If
    ArrayDynamic2_ = i
    i = i + 1
End Function
 
'��ά���� ��������ӣ���������ȡֵ���ʼ��
Public Function ArrayDynamic2D(ParamArray v()) As Variant
    Static arr(), i As Long
    Dim arrRE(), i1 As Long, j As Long
    Const init = 50
    If LBound(v) > UBound(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To i - 1)
            ArrayDynamic2D = ArrTranspose(arr)
        Else
            ArrayDynamic2D = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To UBound(v) + 1, 1 To init)
        i = 1
    ElseIf UBound(v) + 1 > UBound(arr, 1) Then
        arrRE = arr
        ReDim arr(1 To UBound(v) + 1, 1 To UBound(arr, 2) + IIf(i > UBound(arr, 2), init, 0))
        For i1 = 1 To UBound(arrRE, 1)
            For j = 1 To i - 1
               Cover arr(i1, j), arrRE(i1, j)
            Next
        Next
        Erase arrRE
    ElseIf i > UBound(arr, 2) Then
        ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + init)
    End If
    For j = 0 To UBound(v)
       Cover arr(j + 1, i), v(j)
    Next
    ArrayDynamic2D = i
    i = i + 1
End Function
 
'��ά���� ��������ӣ���������ȡֵ���ʼ��
Public Function ArrayDynamic2D1(ParamArray v()) As Variant
    Static arr(), i As Long
    Dim arrRE(), i1 As Long, j As Long
    Const init = 50
    If LBound(v) > UBound(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To i - 1)
            ArrayDynamic2D1 = ArrTranspose(arr)
        Else
            ArrayDynamic2D1 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To UBound(v) + 1, 1 To init)
        i = 1
    ElseIf UBound(v) + 1 > UBound(arr, 1) Then
        arrRE = arr
        ReDim arr(1 To UBound(v) + 1, 1 To UBound(arr, 2) + IIf(i > UBound(arr, 2), init, 0))
        For i1 = 1 To UBound(arrRE, 1)
            For j = 1 To i - 1
               Cover arr(i1, j), arrRE(i1, j)
            Next
        Next
        Erase arrRE
    ElseIf i > UBound(arr, 2) Then
        ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + init)
    End If
    For j = 0 To UBound(v)
       Cover arr(j + 1, i), v(j)
    Next
    ArrayDynamic2D1 = i
    i = i + 1
End Function
 
'��ά���� ��������ӣ���������ȡֵ���ʼ��
Public Function ArrayDynamic2D2(ParamArray v()) As Variant
    Static arr(), i As Long
    Dim arrRE(), i1 As Long, j As Long
    Const init = 50
    If LBound(v) > UBound(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To i - 1)
            ArrayDynamic2D2 = ArrTranspose(arr)
        Else
            ArrayDynamic2D2 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To UBound(v) + 1, 1 To init)
        i = 1
    ElseIf UBound(v) + 1 > UBound(arr, 1) Then
        arrRE = arr
        ReDim arr(1 To UBound(v) + 1, 1 To UBound(arr, 2) + IIf(i > UBound(arr, 2), init, 0))
        For i1 = 1 To UBound(arrRE, 1)
            For j = 1 To i - 1
               Cover arr(i1, j), arrRE(i1, j)
            Next
        Next
        Erase arrRE
    ElseIf i > UBound(arr, 2) Then
        ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + init)
    End If
    For j = 0 To UBound(v)
       Cover arr(j + 1, i), v(j)
    Next
    ArrayDynamic2D2 = i
    i = i + 1
End Function
 
'��ά���� ��������ӣ���������ȡֵ���ʼ��
Public Function ArrayDynamic2D3(ParamArray v()) As Variant
    Static arr(), i As Long
    Dim arrRE(), i1 As Long, j As Long
    Const init = 50
    If LBound(v) > UBound(v) Then
        If i > 1 Then
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To i - 1)
            ArrayDynamic2D3 = ArrTranspose(arr)
        Else
            ArrayDynamic2D3 = Array()
        End If
        i = 0
        Erase arr
        Exit Function
    End If
    If i = 0 Then
        ReDim arr(1 To UBound(v) + 1, 1 To init)
        i = 1
    ElseIf UBound(v) + 1 > UBound(arr, 1) Then
        arrRE = arr
        ReDim arr(1 To UBound(v) + 1, 1 To UBound(arr, 2) + IIf(i > UBound(arr, 2), init, 0))
        For i1 = 1 To UBound(arrRE, 1)
            For j = 1 To i - 1
               Cover arr(i1, j), arrRE(i1, j)
            Next
        Next
        Erase arrRE
    ElseIf i > UBound(arr, 2) Then
        ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + init)
    End If
    For j = 0 To UBound(v)
       Cover arr(j + 1, i), v(j)
    Next
    ArrayDynamic2D3 = i
    i = i + 1
End Function
 
'����ת��
Public Function ArrTranspose(ByRef arr) As Variant
    Dim arrRE(), i As Long, j As Long
    Select Case ArrDimension(arr)
        Case 2
            Dim l As Long, r As Long
            l = LBound(arr, 2): r = UBound(arr, 2)
            ReDim arrRE(1 To r - l + 1, 1 To UBound(arr, 1) - LBound(arr, 1) + 1)
            Dim k As Long
            Dim n As Long: n = 1
            For i = LBound(arr, 1) To UBound(arr, 1)
                k = 1
                For j = l To r
                    Cover arrRE(k, n), arr(i, j)
                    k = k + 1
                Next
                n = n + 1
            Next
        Case 1
            ReDim arrRE(1 To UBound(arr) - LBound(arr) + 1, 1 To 1)
            j = 1
            For i = LBound(arr) To UBound(arr)
                Cover arrRE(j, 1), arr(i)
                j = j + 1
            Next
    End Select
    ArrTranspose = arrRE
End Function
 
'���鷭ת
Public Function ArrFlip(arr) As Variant
    Dim i As Long, j As Long, k As Long, arr2()
    Dim l As Long, u As Long
    Select Case ArrDimension(arr)
        Case 1
            l = LBound(arr, 1): u = UBound(arr, 1)
            ReDim arr2(l To u)
            j = u
            For i = l To u
                Cover arr2(j), arr(i)
                j = j - 1
            Next
        Case 2
            l = LBound(arr, 2): u = UBound(arr, 2)
            ReDim arr2(LBound(arr, 1) To UBound(arr, 1), l To u)
            j = UBound(arr, 1)
            For i = LBound(arr, 1) To UBound(arr, 1)
                For k = l To u
                    Cover arr2(j, k), arr(i, k)
                Next
                j = j - 1
            Next
    End Select
    ArrFlip = arr2
End Function

'һά����ת��ά����
Public Function ArrTo2D(ByRef arr1D, ByVal DCount As Long) As Variant
    Dim arrRE(), i As Long, j As Long
    Dim l As Long, r As Long
    l = LBound(arr1D): r = UBound(arr1D)
    Dim n As Long: n = l
    ReDim arrRE(1 To IntUp((r - l + 1) / DCount), 1 To DCount)
    Dim k As Long
    For i = 1 To UBound(arrRE, 1)
        For j = 1 To DCount
            If n > r Then GoTo ArrTo2DEnd
            Cover arrRE(i, j), arr1D(n)
            n = n + 1
        Next
    Next
ArrTo2DEnd:
    ArrTo2D = arrRE
End Function

'��ά����תһά����
Public Function Arr2DTo1D(ByRef arr2D, Optional RowFirst As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    ArrayDynamic_
    If RowFirst Then
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            For j = l To u
                ArrayDynamic_ arr2D(i, j)
            Next
        Next
    Else
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        For j = LBound(arr2D, 2) To UBound(arr2D, 2)
            For i = l To u
                ArrayDynamic_ arr2D(i, j)
            Next
        Next
    End If
    Arr2DTo1D = ArrayDynamic_
End Function

'�������������  ColumnCount =0ȡ����� >0ʹ��ColumnCount��Ϊ������������ȥ <0����һ��Ԫ�ص�����Ϊ����
Public Function ArrF_T(ByRef arr, Optional ColumnCount = 0) As Variant
    Dim arrRE(), i As Long, j As Long
    Dim l As Long: l = LBound(arr)
    Dim maxColumnCount As Long, maxColumnCountTMP As Long
    If ColumnCount > 0 Then
        maxColumnCount = ColumnCount
    ElseIf ColumnCount = 0 Then
        maxColumnCount = 0
        For i = l To UBound(arr)
            If IsArray(arr(i)) Then
                maxColumnCountTMP = UBound(arr(i)) - LBound(arr(i)) + 1
                If maxColumnCount < maxColumnCountTMP Then maxColumnCount = maxColumnCountTMP
            Else
                If maxColumnCount < 1 Then maxColumnCount = 1
            End If
        Next
    Else
        If IsArray(arr(l)) Then
            maxColumnCount = UBound(arr(l)) - LBound(arr(l)) + 1
        Else
            maxColumnCount = 1
        End If
    End If
    If maxColumnCount > 0 Then
        ReDim arrRE(1 To UBound(arr) - l + 1, 1 To maxColumnCount)
        Dim k As Long
        Dim n As Long: n = 1
        For i = l To UBound(arr)
            k = 1
            If IsArray(arr(i)) Then
                For j = LBound(arr(i)) To MinParams2(UBound(arr(i)), maxColumnCount + LBound(arr(i)) - 1)
                    Cover arrRE(n, k), arr(i)(j)
                    k = k + 1
                Next
            Else
                Cover arrRE(n, k), arr(i)
            End If
            n = n + 1
        Next
        ArrF_T = arrRE
    Else
        ArrF_T = Array()
    End If
End Function

'������������� �����������±� *�����ϱ����һ��*
Public Function ArrF_T_LIndexToUIndex(ByRef arr) As Variant
    Dim arrRE(), i As Long, j As Long
    Dim larr As Long, l As Long, u As Long
    larr = LBound(arr)
    l = LBound(arr(larr)): u = UBound(arr(larr))
    ReDim arrRE(LBound(arr) To UBound(arr), l To u)
    For i = larr To UBound(arr)
        For j = l To u
            Cover arrRE(i, j), arr(i)(j)
        Next
    Next
    ArrF_T_LIndexToUIndex = arrRE
End Function

'չƽ����(һά��) ����
Public Function ArrFlatten_Single(ParamArray arr()) As Variant
    ArrayDynamic_
    Dim v, vv
    For Each v In arr
        If IsArray(v) Then
            For Each vv In v
                ArrayDynamic_ vv
            Next
        Else
            ArrayDynamic_ v
        End If
    Next
    ArrFlatten_Single = ArrayDynamic_
End Function
 
'չƽ����(һά��) �ݹ�
Public Function ArrFlatten(ParamArray arr()) As Variant
    ArrayDynamic_
    Dim v
    For Each v In arr
        If IsArray(v) Then
            ArrFlatten_ v
        Else
            ArrayDynamic_ v
        End If
    Next
    ArrFlatten = ArrayDynamic_
End Function
 
'�ڲ��ݹ�չƽ
Private Sub ArrFlatten_(ByRef arr)
    Dim v
    For Each v In arr
        If IsArray(v) Then
             ArrFlatten_ v
        Else
            ArrayDynamic_ v
        End If
    Next
End Sub
 
'��ά�����ں�����������,����Ӧ���и��ƶ���չ��
'  |  1 | 2     |  3 |  4
'1 | A1 | [1,2] | B1 | C1
'2 | A2 | [1,2] | B2 | C2
'Arr2DFlatten(arr, 2)
' 1 | A1 | 1 | B1 | C1
' 2 | A1 | 2 | B1 | C1
' 3 | A2 | 1 | B2 | C2
' 4 | A2 | 2 | B2 | C2
Public Function Arr2DFlatten(ByRef arr2D, ByVal ColumnIndex) As Variant
    ArrayDynamic_
    Dim i As Long, j As Long, arrRE(), v
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    IndexIsCurrencyToCount_ ColumnIndex, l2, u2
    ReDim arrRE(l2 To u2)
    For i = l1 To u1
        If IsArray(arr2D(i, ColumnIndex)) Or TypeName(arr2D(i, ColumnIndex)) = "Collection" Then
            For Each v In arr2D(i, ColumnIndex)
                For j = l2 To u2
                    If j = ColumnIndex Then
                        Cover arrRE(j), v
                    Else
                        Cover arrRE(j), arr2D(i, j)
                    End If
                Next
                ArrayDynamic_ arrRE
            Next
        Else
            For j = l2 To u2
                Cover arrRE(j), arr2D(i, j)
            Next
            ArrayDynamic_ arrRE
        End If
    Next
    Arr2DFlatten = ArrF_T_LIndexToUIndex(ArrayDynamic_)
End Function

'�ϲ����飬���ºϲ�
Public Function ArrMergeRow(ByRef arr) As Variant
    Dim v, arrtmp, hangshu As Long, lieshu As Long
    Dim i As Long, j As Long
    Dim ii As Long, jj As Long
    hangshu = 0: lieshu = 0
    For Each v In arr
        If IsArray(v) Then
            If ArrDimension(v) = 1 Then
                hangshu = hangshu + 1
                lieshu = MaxParams2(lieshu, UBound(v) - LBound(v) + 1)
            Else
                hangshu = hangshu + UBound(v, 1) - LBound(v, 1) + 1
                lieshu = MaxParams2(lieshu, UBound(v, 2) - LBound(v, 2) + 1)
            End If
        End If
    Next
    If hangshu > 0 And lieshu > 0 Then
        ReDim arrtmp(1 To hangshu, 1 To lieshu)
        ii = 1
        For Each v In arr
            If IsArray(v) Then
                If ArrDimension(v) = 1 Then
                    jj = 1
                    For j = LBound(v) To UBound(v)
                        Cover arrtmp(ii, jj), v(j)
                        jj = jj + 1
                    Next
                    ii = ii + 1
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        jj = 1
                        For j = LBound(v, 2) To UBound(v, 2)
                            Cover arrtmp(ii, jj), v(i, j)
                            jj = jj + 1
                        Next
                        ii = ii + 1
                    Next
                End If
            End If
        Next
        ArrMergeRow = arrtmp
    Else
        ArrMergeRow = Array()
    End If
End Function
 
'�ϲ����飬���ºϲ�
Public Function ArrMergeRowParam(ParamArray arr()) As Variant
    Dim v, arrtmp, hangshu As Long, lieshu As Long
    Dim i As Long, j As Long
    Dim ii As Long, jj As Long
    hangshu = 0: lieshu = 0
    For Each v In arr
        If IsArray(v) Then
            If ArrDimension(v) = 1 Then
                hangshu = hangshu + 1
                lieshu = MaxParams2(lieshu, UBound(v) - LBound(v) + 1)
            Else
                hangshu = hangshu + UBound(v, 1) - LBound(v, 1) + 1
                lieshu = MaxParams2(lieshu, UBound(v, 2) - LBound(v, 2) + 1)
            End If
        End If
    Next
    If hangshu > 0 And lieshu > 0 Then
        ReDim arrtmp(1 To hangshu, 1 To lieshu)
        ii = 1
        For Each v In arr
            If IsArray(v) Then
                If ArrDimension(v) = 1 Then
                    jj = 1
                    For j = LBound(v) To UBound(v)
                        Cover arrtmp(ii, jj), v(j)
                        jj = jj + 1
                    Next
                    ii = ii + 1
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        jj = 1
                        For j = LBound(v, 2) To UBound(v, 2)
                            Cover arrtmp(ii, jj), v(i, j)
                            jj = jj + 1
                        Next
                        ii = ii + 1
                    Next
                End If
            End If
        Next
        ArrMergeRowParam = arrtmp
    Else
        ArrMergeRowParam = Array()
    End If
End Function
 
'�ϲ����飬���Һϲ�
Public Function ArrMergeColumn(ByRef arr) As Variant
    Dim v, arrtmp, hangshu As Long, lieshu As Long
    Dim i As Long, j As Long
    Dim ii As Long, jj As Long
    hangshu = 0: lieshu = 0
    For Each v In arr
        If IsArray(v) Then
            If ArrDimension(v) = 1 Then
                lieshu = lieshu + 1
                hangshu = MaxParams2(hangshu, UBound(v) - LBound(v) + 1)
            Else
                lieshu = lieshu + UBound(v, 2) - LBound(v, 2) + 1
                hangshu = MaxParams2(hangshu, UBound(v, 1) - LBound(v, 1) + 1)
            End If
        End If
    Next
    If hangshu > 0 And lieshu > 0 Then
        ReDim arrtmp(1 To hangshu, 1 To lieshu)
        jj = 1
        For Each v In arr
            If IsArray(v) Then
                If ArrDimension(v) = 1 Then
                    ii = 1
                    For i = LBound(v) To UBound(v)
                        Cover arrtmp(ii, jj), v(i)
                        ii = ii + 1
                    Next
                    jj = jj + 1
                Else
                    For j = LBound(v, 2) To UBound(v, 2)
                        ii = 1
                        For i = LBound(v, 1) To UBound(v, 1)
                            Cover arrtmp(ii, jj), v(i, j)
                            ii = ii + 1
                        Next
                        jj = jj + 1
                    Next
                End If
            End If
        Next
        ArrMergeColumn = arrtmp
    Else
        ArrMergeColumn = Array()
    End If
End Function
 
'�ϲ����飬���Һϲ�
Public Function ArrMergeColumnParam(ParamArray arr()) As Variant
    Dim v, arrtmp, hangshu As Long, lieshu As Long
    Dim i As Long, j As Long
    Dim ii As Long, jj As Long
    hangshu = 0: lieshu = 0
    For Each v In arr
        If IsArray(v) Then
            If ArrDimension(v) = 1 Then
                lieshu = lieshu + 1
                hangshu = MaxParams2(hangshu, UBound(v) - LBound(v) + 1)
            Else
                lieshu = lieshu + UBound(v, 2) - LBound(v, 2) + 1
                hangshu = MaxParams2(hangshu, UBound(v, 1) - LBound(v, 1) + 1)
            End If
        End If
    Next
    If hangshu > 0 And lieshu > 0 Then
        ReDim arrtmp(1 To hangshu, 1 To lieshu)
        jj = 1
        For Each v In arr
            If IsArray(v) Then
                If ArrDimension(v) = 1 Then
                    ii = 1
                    For i = LBound(v) To UBound(v)
                        Cover arrtmp(ii, jj), v(i)
                        ii = ii + 1
                    Next
                    jj = jj + 1
                Else
                    For j = LBound(v, 2) To UBound(v, 2)
                        ii = 1
                        For i = LBound(v, 1) To UBound(v, 1)
                            Cover arrtmp(ii, jj), v(i, j)
                            ii = ii + 1
                        Next
                        jj = jj + 1
                    Next
                End If
            End If
        Next
        ArrMergeColumnParam = arrtmp
    Else
        ArrMergeColumnParam = Array()
    End If
End Function

'һά���� ����Ԫ�� ArrEleCountΪ��Ӧarr��С���������� ArrCopyElement([1,2,3],[2,3])->[1,1,2,2,2,3]
Public Function ArrCopyElement(ByRef arr, ParamArray ArrEleCount()) As Variant
    Dim i As Long, j As Long, k As Long, u As Long
    ArrEleCount = ArrFlatten(ArrEleCount)
    ArrayDynamic_
    k = LBound(ArrEleCount): u = UBound(ArrEleCount)
    For i = LBound(arr) To UBound(arr)
        If k > u Then
            ArrayDynamic_ arr(i)
        Else
            For j = 1 To ArrEleCount(k)
                ArrayDynamic_ arr(i)
            Next
        End If
        k = k + 1
    Next
    ArrCopyElement = ArrayDynamic_
End Function

'�������� ArrEleCountΪ��Ӧarr2D����������������
Public Function ArrCopyColumn(ByRef arr2D, ParamArray ArrEleCount()) As Variant
    Dim arrindex
    arrindex = ArrGetIndex(arr2D, False)
    arrindex = ArrCopyElement(arrindex, ArrEleCount)
    ArrCopyColumn = ArrGetColumns(arr2D, arrindex)
End Function

'�������� ArrEleCountΪ��Ӧarr2D����������������
Public Function ArrCopyRow(ByRef arr2D, ParamArray ArrEleCount()) As Variant
    Dim arrindex
    arrindex = ArrGetIndex(arr2D)
    arrindex = ArrCopyElement(arrindex, ArrEleCount)
    ArrCopyRow = ArrGetRows(arr2D, arrindex)
End Function

'һά���� ����Ԫ�� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount�� ArrCopyElement2([1,2,3],[2,3],[2,3])->[1,2,2,3,3,3]
Public Function ArrCopyElement2(ByRef arr, ArrCopyIndex, ArrCopyCount) As Variant
    Dim u As Long
    Dim ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyIndexRE = ArrFlatten_Single(ArrCopyIndex)
    u = UBound(ArrCopyIndexRE)
    If IsArray(ArrCopyCount) Then
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, 1)
    Else
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, ArrCopyCount)
    End If
    
    Dim arrRE(), i As Long
    ReDim arrRE(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1)
        arrRE(i) = 1
    Next
    ArrSetValues arrRE, ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyElement2 = ArrCopyElement(arr, arrRE)
End Function

'�������� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount��
Public Function ArrCopyColumn2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant
    Dim u As Long
    Dim ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyIndexRE = ArrFlatten_Single(ArrCopyIndex)
    u = UBound(ArrCopyIndexRE)
    If IsArray(ArrCopyCount) Then
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, 1)
    Else
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, ArrCopyCount)
    End If
    
    Dim arrRE(), i As Long
    ReDim arrRE(LBound(arr2D, 2) To UBound(arr2D, 2))
    For i = LBound(arr2D, 2) To UBound(arr2D, 2)
        arrRE(i) = 1
    Next
    ArrSetValues arrRE, ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyColumn2 = ArrCopyColumn(arr2D, arrRE)
End Function

'�������� ArrCopyIndexλ�ö�Ӧ�ĸ���ArrCopyCount��
Public Function ArrCopyRow2(ByRef arr2D, ArrCopyIndex, ArrCopyCount) As Variant
    Dim u As Long
    Dim ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyIndexRE = ArrFlatten_Single(ArrCopyIndex)
    u = UBound(ArrCopyIndexRE)
    If IsArray(ArrCopyCount) Then
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, 1)
    Else
        ArrCopyCountRE = ArrSizeExpansion2(ArrCopyCount, u, ArrCopyCount)
    End If
    
    Dim arrRE(), i As Long
    ReDim arrRE(LBound(arr2D, 1) To UBound(arr2D, 1))
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        arrRE(i) = 1
    Next
    ArrSetValues arrRE, ArrCopyIndexRE, ArrCopyCountRE
    ArrCopyRow2 = ArrCopyRow(arr2D, arrRE)
End Function

'һά���� ����һ����ֵ������ֵ
Public Function ArrInsert(ByRef arr, Optional ByVal Index, Optional ByVal EleCount = 1, Optional EleCopy As Boolean = False) As Variant
    Dim arrRE, i As Long, u As Long
    If IsMissing(Index) Then
        u = UBound(arr)
        ReDim Preserve arr(LBound(arr) To u + EleCount)
        If EleCopy And EleCount > 0 Then
            Cover arrRE, arr(u)
            For i = u + 1 To u + EleCount
                Cover arr(i), arrRE
            Next
        End If
    Else
        IndexIsCurrencyToCount_ Index, LBound(arr), UBound(arr)
        arrRE = arr
        ReDim arr(LBound(arr) To UBound(arr) + EleCount)
        For i = LBound(arrRE) To Index - 1
            Cover arr(i), arrRE(i)
        Next
        If EleCopy Then
            If Index = UBound(arrRE) + 1 Then
                For i = Index To Index + EleCount - 1
                    Cover arr(i), arrRE(Index - 1)
                Next
            Else
                For i = Index To Index + EleCount - 1
                    Cover arr(i), arrRE(Index)
                Next
            End If
        End If
        For i = Index To UBound(arrRE)
            Cover arr(i + EleCount), arrRE(i)
        Next
    End If
    ArrInsert = arr
End Function

'���� ����һ�л����
Public Function ArrInsertColumn(ByRef arr2D, Optional ByVal ColumnIndex, Optional ByVal ColumnCount = 1, Optional EleCopy As Boolean = False) As Variant
    Dim arrRE(), i As Long, j As Long
    Dim l2 As Long, u2 As Long
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    If IsMissing(ColumnIndex) Then
        ReDim Preserve arr2D(LBound(arr2D, 1) To UBound(arr2D, 1), l2 To u2 + ColumnCount)
        If EleCopy And ColumnCount > 0 Then
            arrRE = ArrGetColumn(arr2D, u2)
            For i = LBound(arr2D, 1) To UBound(arr2D, 1)
                For j = u2 + 1 To u2 + ColumnCount
                    Cover arr2D(i, j), arrRE(i)
                Next
            Next
        End If
    Else
        IndexIsCurrencyToCount_ ColumnIndex, l2, u2
        arrRE = arr2D
        ReDim arr2D(LBound(arr2D, 1) To UBound(arr2D, 1), l2 To u2 + ColumnCount)
        For i = LBound(arrRE, 1) To UBound(arrRE, 1)
            For j = l2 To ColumnIndex - 1
                Cover arr2D(i, j), arrRE(i, j)
            Next
        Next
        If EleCopy Then
            If ColumnIndex = UBound(arrRE, 2) + 1 Then
                For i = LBound(arrRE, 1) To UBound(arrRE, 1)
                    For j = ColumnIndex To ColumnIndex + ColumnCount - 1
                        Cover arr2D(i, j), arrRE(i, ColumnIndex - 1)
                    Next
                Next
            Else
                For i = LBound(arrRE, 1) To UBound(arrRE, 1)
                    For j = ColumnIndex To ColumnIndex + ColumnCount - 1
                        Cover arr2D(i, j), arrRE(i, ColumnIndex)
                    Next
                Next
            End If
        End If
        For i = LBound(arrRE, 1) To UBound(arrRE, 1)
            For j = ColumnIndex To u2
                Cover arr2D(i, j + ColumnCount), arrRE(i, j)
            Next
        Next
    End If
    ArrInsertColumn = arr2D
End Function

'���� ����һ�л����
Public Function ArrInsertRow(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal RowCount = 1, Optional EleCopy As Boolean = False) As Variant
    Dim arrRE(), i As Long, j As Long
    Dim l2 As Long, u2 As Long
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    If IsMissing(RowIndex) Then RowIndex = UBound(arr2D, 1) + 1
    IndexIsCurrencyToCount_ RowIndex, LBound(arr2D, 1), UBound(arr2D, 1)
    arrRE = arr2D
    ReDim arr2D(LBound(arr2D, 1) To UBound(arr2D, 1) + RowCount, l2 To u2)
    For i = LBound(arrRE, 1) To RowIndex - 1
        For j = l2 To u2
            Cover arr2D(i, j), arrRE(i, j)
        Next
    Next
    For i = RowIndex To UBound(arrRE, 1)
        For j = l2 To u2
            Cover arr2D(i + RowCount, j), arrRE(i, j)
        Next
    Next
    If EleCopy Then
        If RowIndex = UBound(arrRE, 1) + 1 Then
            For i = RowIndex To RowIndex + RowCount - 1
                For j = l2 To u2
                    Cover arr2D(i, j), arrRE(RowIndex - 1, j)
                Next
            Next
        Else
            For i = RowIndex To RowIndex + RowCount - 1
                For j = l2 To u2
                    Cover arr2D(i, j), arrRE(RowIndex, j)
                Next
            Next
        End If
    End If
    ArrInsertRow = arr2D
End Function

'���� ȡ����
Public Function ArrGetIndex(ByRef arr, Optional GetRowIndex As Boolean = True) As Variant()
    If GetRowIndex Then
        ArrGetIndex = ArrBetween(LBound(arr, 1), UBound(arr, 1))
    Else
        ArrGetIndex = ArrBetween(LBound(arr, 2), UBound(arr, 2))
    End If
End Function

'һά���� ɾ��һ��Ԫ�ػ���Ԫ��
Public Function ArrRemoveRegion(ByRef arr, ByVal Index, Optional ByVal Count = 1) As Variant
    Dim arri
    IndexIsCurrencyToCount_ Index, LBound(arr, 1), UBound(arr, 1)
    arri = ArrFilterRemove(ArrGetIndex(arr), ArrBetween(Index, Index + Count - 1))
    If LBound(arri) <= UBound(arri) Then
        ArrRemoveRegion = ArrFromIndex(arr, arri)
    Else
        ArrRemoveRegion = Array()
    End If
End Function
 
'���� ɾ��һ�л����
Public Function ArrRemoveColumn(ByRef arr2D, ByVal Index, Optional ByVal ColumnCount = 1) As Variant
    Dim arri
    IndexIsCurrencyToCount_ Index, LBound(arr2D, 2), UBound(arr2D, 2)
    arri = ArrFilterRemove(ArrGetIndex(arr2D, False), ArrBetween(Index, Index + ColumnCount - 1))
    If LBound(arri) <= UBound(arri) Then
        ArrRemoveColumn = ArrGetColumns(arr2D, arri)
    Else
        ArrRemoveColumn = Array()
    End If
End Function
 
'���� ɾ��һ�л���� �����
Public Function ArrRemoveColumns(ByRef arr2D, ParamArray arrindex()) As Variant
    Dim arrIndex1
    arrIndex1 = ArrFlatten(arrindex)
    Dim arri
    IndexIsCurrencyToCount_ arrIndex1, LBound(arr2D, 2), UBound(arr2D, 2)
    arri = ArrFilterRemove(ArrGetIndex(arr2D, False), arrIndex1)
    If LBound(arri) <= UBound(arri) Then
        ArrRemoveColumns = ArrGetColumns(arr2D, arri)
    Else
        ArrRemoveColumns = Array()
    End If
End Function
 
'���� ɾ��һ�л����
Public Function ArrRemoveRow(ByRef arr2D, ByVal Index, Optional ByVal RowCount = 1) As Variant
    Dim arri
    IndexIsCurrencyToCount_ Index, LBound(arr2D, 1), UBound(arr2D, 1)
    arri = ArrFilterRemove(ArrGetIndex(arr2D), ArrBetween(Index, Index + RowCount - 1))
    If LBound(arri) <= UBound(arri) Then
        ArrRemoveRow = ArrGetRows(arr2D, arri)
    Else
        ArrRemoveRow = Array()
    End If
End Function
 
'���� ɾ��һ�л���� �����
Public Function ArrRemoveRows(ByRef arr2D, ParamArray arrindex()) As Variant
    Dim arrIndex1
    arrIndex1 = ArrFlatten(arrindex)
    Dim arri
    IndexIsCurrencyToCount_ arrIndex1, LBound(arr2D, 1), UBound(arr2D, 1)
    arri = ArrFilterRemove(ArrGetIndex(arr2D), arrIndex1)
    If LBound(arri) <= UBound(arri) Then
        ArrRemoveRows = ArrGetRows(arr2D, arri)
    Else
        ArrRemoveRows = Array()
    End If
End Function

'����ȡ���� һ��Ϊһά����
Public Function ArrGetRow(ByRef arr2D, ByVal Index, Optional ByVal RowCount = 1, Optional Expansion As Boolean = False) As Variant
    Dim l1 As Long, r1 As Long
    Dim l As Long, r As Long, i As Long, j As Long
    Dim arrtmp()
    l1 = LBound(arr2D, 1): r1 = UBound(arr2D, 1)
    l = LBound(arr2D, 2): r = UBound(arr2D, 2)
    IndexIsCurrencyToCount_ Index, l1, r1
    If RowCount <= 0 Then RowCount = MaxParams2(r1 - Index + 1, 1)
    If RowCount = 1 Then
        If Index < l1 Or Index > r1 Then
            If Expansion Then
                ReDim arrtmp(l To r)
            Else
                arrtmp = Array()
            End If
        Else
            ReDim arrtmp(l To r)
            For i = l To r
                Cover arrtmp(i), arr2D(Index, i)
            Next
        End If
    Else
        If Index < l1 Then
            If Expansion Then
                ReDim arrtmp(1 To RowCount, l To r)
                If Index + RowCount > r1 Then RowCount = r1 - Index + 1
                For i = l To r
                    For j = l1 - Index + 1 To RowCount
                        Cover arrtmp(j, i), arr2D(Index + j - 1, i)
                    Next
                Next
            Else
                arrtmp = Array()
            End If
        ElseIf Index > r1 Then
            If Expansion Then
                ReDim arrtmp(1 To RowCount, l To r)
            Else
                arrtmp = Array()
            End If
        Else
            If Expansion Then
                ReDim arrtmp(1 To RowCount, l To r)
                If Index + RowCount > r1 Then RowCount = r1 - Index + 1
            Else
                If Index + RowCount > r1 Then RowCount = r1 - Index + 1
                ReDim arrtmp(1 To RowCount, l To r)
            End If
            For i = l To r
                For j = 1 To RowCount
                    Cover arrtmp(j, i), arr2D(Index + j - 1, i)
                Next
            Next
        End If
    End If
    ArrGetRow = arrtmp
End Function
 
'����ȡ���е���ά����
Public Function ArrGetRows(ByRef arr2D, ParamArray arrindex()) As Variant
    Dim arrIndex1
    arrIndex1 = ArrFlatten(arrindex)
    Dim l As Long, r As Long, i As Long, j As Long
    l = LBound(arr2D, 2): r = UBound(arr2D, 2)
    Dim lI As Long, ri As Long
    lI = LBound(arrIndex1): ri = UBound(arrIndex1)
    Dim l2 As Long, r2 As Long
    l2 = LBound(arr2D, 1): r2 = UBound(arr2D, 1)
    
    IndexIsCurrencyToCount_ arrIndex1, l2, r2
    
    Dim Index As Long
    Dim arrtmp(): ReDim arrtmp(lI To ri, l To r)
    For j = lI To ri
        Index = arrIndex1(j)
        If Index >= l2 And Index <= r2 Then
            For i = l To r
                Cover arrtmp(j, i), arr2D(Index, i)
            Next
        End If
    Next
    ArrGetRows = arrtmp
End Function

'����ȡ���� һ��Ϊһά����
Public Function ArrGetColumn(ByRef arr2D, ByVal Index, Optional ByVal ColumnCount = 1, Optional Expansion As Boolean = False) As Variant
    Dim l As Long, r As Long, i As Long, j As Long
    l = LBound(arr2D, 1): r = UBound(arr2D, 1)
    Dim l2 As Long, r2 As Long
    l2 = LBound(arr2D, 2): r2 = UBound(arr2D, 2)
    IndexIsCurrencyToCount_ Index, l2, r2
    Dim arrtmp()
    If ColumnCount <= 0 Then ColumnCount = MaxParams2(r2 - Index + 1, 1)
    If ColumnCount = 1 Then
        If Index < l2 Or Index > r2 Then
            If Expansion Then
                ReDim arrtmp(l To r)
            Else
                arrtmp = Array()
            End If
        Else
            ReDim arrtmp(l To r)
            For i = l To r
                Cover arrtmp(i), arr2D(i, Index)
            Next
        End If
    Else
        If Index < l2 Then
            If Expansion Then
                ReDim arrtmp(l To r, 1 To ColumnCount)
                If Index + ColumnCount > r2 Then ColumnCount = r2 - Index + 1
                For i = l To r
                    For j = l2 - Index + 1 To ColumnCount
                        Cover arrtmp(i, j), arr2D(i, Index + j - 1)
                    Next
                Next
            Else
                arrtmp = Array()
            End If
        ElseIf Index > r2 Then
            If Expansion Then
                ReDim arrtmp(l To r, 1 To ColumnCount)
            Else
                arrtmp = Array()
            End If
        Else
            If Expansion Then
                ReDim arrtmp(l To r, 1 To ColumnCount)
                If Index + ColumnCount > r2 Then ColumnCount = r2 - Index + 1
            Else
                If Index + ColumnCount > r2 Then ColumnCount = r2 - Index + 1
                ReDim arrtmp(l To r, 1 To ColumnCount)
            End If
            For i = l To r
                For j = 1 To ColumnCount
                    Cover arrtmp(i, j), arr2D(i, Index + j - 1)
                Next
            Next
        End If
    End If
    ArrGetColumn = arrtmp
End Function
 
'����ȡ���е���ά����
Public Function ArrGetColumns(ByRef arr2D, ParamArray arrindex()) As Variant
    Dim arrIndex1
    arrIndex1 = ArrFlatten(arrindex)
    Dim l As Long, r As Long, i As Long, j As Long
    l = LBound(arr2D, 1): r = UBound(arr2D, 1)
    Dim lI As Long, ri As Long
    lI = LBound(arrIndex1): ri = UBound(arrIndex1)
    Dim l2 As Long, r2 As Long
    l2 = LBound(arr2D, 2): r2 = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ arrIndex1, l2, r2
    
    Dim Index As Long
    Dim arrtmp(): ReDim arrtmp(l To r, lI To ri)
    For j = lI To ri
        Index = arrIndex1(j)
        If Index >= l2 And Index <= r2 Then
            For i = l To r
                Cover arrtmp(i, j), arr2D(i, Index)
            Next
        End If
    Next
    ArrGetColumns = arrtmp
End Function
 
'����ȡ���� �����Ӵ�С ��ά����
Public Function ArrGetRegion2D(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
Optional ByVal Height = 0, Optional ByVal Width = 0, Optional Expansion As Boolean = False) As Variant
    Dim l1 As Long, r1 As Long
    Dim l2 As Long, r2 As Long
    Dim i As Long, j As Long
    l1 = LBound(arr2D, 1): r1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): r2 = UBound(arr2D, 2)
    If IsMissing(ColumnIndex) Then ColumnIndex = l2
    If IsMissing(RowIndex) Then RowIndex = l1
    IndexIsCurrencyToCount_ RowIndex, l1, r1
    IndexIsCurrencyToCount_ ColumnIndex, l2, r2
    If Width = 0 Then Width = (r2 - l2) - (ColumnIndex - l2) + 1
    If Height = 0 Then Height = (r1 - l1) - (RowIndex - l1) + 1
    'ѭ��ĩβ����
    Dim ws As Long, hs As Long
    ws = MinParams2(ColumnIndex + Width - 1, r2)
    hs = MinParams2(RowIndex + Height - 1, r1)
    Dim arrtmp()
    If Expansion Then
        ReDim arrtmp(1 To Height, 1 To Width)
    Else
        ReDim arrtmp(1 To hs - RowIndex + 1, 1 To ws - ColumnIndex + 1)
    End If
    Dim i2 As Long, j2 As Long
    i2 = 1
    For i = RowIndex To hs
        j2 = 1
        For j = ColumnIndex To ws
           Cover arrtmp(i2, j2), arr2D(i, j)
           j2 = j2 + 1
        Next
        i2 = i2 + 1
    Next
    ArrGetRegion2D = arrtmp
End Function

'����ȡ���� ���������� ��ά����
Public Function ArrGetRegion2D_To(ByRef arr2D, Optional ByVal RowIndex, Optional ByVal ColumnIndex, _
        Optional ByVal RowIndexTo, Optional ByVal ColumnIndexTo, Optional Expansion As Boolean = False) As Variant
    If IsMissing(ColumnIndex) Then ColumnIndex = LBound(arr2D, 2)
    If IsMissing(RowIndex) Then RowIndex = LBound(arr2D, 1)
    If IsMissing(ColumnIndexTo) Then ColumnIndexTo = UBound(arr2D, 2)
    If IsMissing(RowIndexTo) Then RowIndexTo = UBound(arr2D, 1)
    
    IndexIsCurrencyToCount_ RowIndex, LBound(arr2D, 1), UBound(arr2D, 1)
    IndexIsCurrencyToCount_ ColumnIndex, LBound(arr2D, 2), UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ RowIndexTo, LBound(arr2D, 1), UBound(arr2D, 1)
    IndexIsCurrencyToCount_ ColumnIndexTo, LBound(arr2D, 2), UBound(arr2D, 2)
    
    If RowIndexTo - RowIndex + 1 > 0 And ColumnIndexTo - ColumnIndex + 1 > 0 Then
        ArrGetRegion2D_To = ArrGetRegion2D(arr2D, RowIndex, ColumnIndex, RowIndexTo - RowIndex + 1, ColumnIndexTo - ColumnIndex + 1, Expansion)
    Else
        ArrGetRegion2D_To = Array()
    End If
End Function

'����ȡ���� �����Ӵ�С һά����
Public Function ArrGetRegion(ByRef arr, Optional ByVal Index, Optional ByVal Count = 0, Optional Expansion As Boolean = False) As Variant
    Dim l1 As Long, r1 As Long
    Dim i As Long
    l1 = LBound(arr, 1): r1 = UBound(arr, 1)
    If IsMissing(Index) Then Index = l1
    IndexIsCurrencyToCount_ Index, l1, r1
    If Count = 0 Then Count = (r1 - l1) - (Index - l1) + 1
    'ѭ��ĩβ����
    Dim rs As Long
    rs = MinParams2(Index + Count - 1, r1)
    Dim arrtmp()
    If Expansion Then
        ReDim arrtmp(1 To Count)
    Else
        ReDim arrtmp(1 To rs - Index + 1)
    End If
    Dim i2 As Long
    i2 = 1
    For i = Index To rs
        Cover arrtmp(i2), arr(i)
        i2 = i2 + 1
    Next
    ArrGetRegion = arrtmp
End Function

'����ȡ���� ���������� һά����
Public Function ArrGetRegion_To(ByRef arr, Optional ByVal Index, Optional ByVal IndexTo, Optional Expansion As Boolean = False) As Variant
    If IsMissing(Index) Then Index = LBound(arr, 1)
    If IsMissing(IndexTo) Then IndexTo = UBound(arr, 1)
    IndexIsCurrencyToCount_ Index, LBound(arr, 1), UBound(arr, 1)
    IndexIsCurrencyToCount_ IndexTo, LBound(arr, 1), UBound(arr, 1)
    If IndexTo - Index + 1 > 0 Then
        ArrGetRegion_To = ArrGetRegion(arr, Index, IndexTo - Index + 1, Expansion)
    Else
        ArrGetRegion_To = Array()
    End If
End Function

'���������С �������鶼���һά  **�����±��1**  �������������
Public Function ArrSizeExpansion2(ByRef arr, ByRef ArrSizeCount, Optional FillValue = Empty)
    Dim arrRE()
    ReDim arrRE(1 To ArrSizeCount)
    Dim i As Long, v
    If IsArray(arr) Then
        i = 1
        For Each v In arr
            If i > ArrSizeCount Then Exit For
            Cover arrRE(i), v
            i = i + 1
        Next
        If Not IsEmpty(FillValue) Then
            For i = i To ArrSizeCount
                Cover arrRE(i), FillValue
            Next
        End If
    Else
        Cover arrRE(1), arr
        If Not IsEmpty(FillValue) Then
            For i = 2 To ArrSizeCount
                Cover arrRE(i), FillValue
            Next
        End If
    End If
    ArrSizeExpansion2 = arrRE
End Function

'���������С  **�����±��1**
Public Function ArrSizeExpansion(ByRef arr, ByRef RowCount, Optional ByRef ColumnCount, Optional FillValue = Empty)
    Dim arrRE()
    Dim i As Long, j As Long, ia As Long, ja As Long, c As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(arr)
        Case 1
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            If IsMissing(ColumnCount) Then
                If l1 = 1 And u1 = RowCount Then ArrSizeExpansion = arr: Exit Function   '����
                ReDim arrRE(1 To RowCount)
                ia = l1
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    Cover arrRE(i), arr(ia)
                    ia = ia + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For i = u1 - l1 + 2 To RowCount
                        Cover arrRE(i), FillValue
                    Next
                End If
            Else
                ReDim arrRE(1 To RowCount, 1 To ColumnCount)
                ia = l1
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    Cover arrRE(i, 1), arr(ia)
                    ia = ia + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For i = u1 - l1 + 2 To RowCount
                        Cover arrRE(i, 1), FillValue
                    Next
                    For i = 1 To RowCount
                        For j = 2 To ColumnCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                End If
            End If
        Case 2
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            If l1 = 1 And u1 = RowCount And l2 = 1 And u2 = ColumnCount Then ArrSizeExpansion = arr: Exit Function '����
            ReDim arrRE(1 To RowCount, 1 To ColumnCount)
            ia = l1
            c = MinParams2(u2 - l2 + 1, ColumnCount)
            For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                ja = l2
                For j = 1 To c
                    Cover arrRE(i, j), arr(ia, ja)
                    ja = ja + 1
                Next
                ia = ia + 1
            Next
            If Not IsEmpty(FillValue) Then
                For i = u1 - l1 + 2 To RowCount
                    For j = 1 To ColumnCount
                        Cover arrRE(i, j), FillValue
                    Next
                Next
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    For j = c + 1 To ColumnCount
                        Cover arrRE(i, j), FillValue
                    Next
                Next
            End If
    End Select
    ArrSizeExpansion = arrRE
End Function

'���������С ���������������  **�����±��1**
'��������ʱ�������Ԫ��,����Ԫ������Ϊ1ʱ�������Ԫ�أ�����Ϊһ������ʱ��������У�����Ϊһ�л�һά����ʱ���������
Public Function ArrSizeExpansionEx(ByRef arr, ByRef RowCount, ByRef ColumnCount, Optional FillValue = Empty)
    Dim arrRE()
    Dim i As Long, j As Long, ia As Long, ja As Long, c As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(arr)
        Case 1
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            ReDim arrRE(1 To RowCount, 1 To ColumnCount)
            If u1 - l1 = 0 Then
                For i = 1 To RowCount
                    For j = 1 To ColumnCount
                        Cover arrRE(i, j), arr(l1)
                    Next
                Next
            Else
                ia = l1
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    For j = 1 To ColumnCount
                        Cover arrRE(i, j), arr(ia)
                    Next
                    ia = ia + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For i = u1 - l1 + 2 To RowCount
                        For j = 1 To ColumnCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                End If
            End If
        Case 2
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            If l1 = 1 And u1 = RowCount And l2 = 1 And u2 = ColumnCount Then ArrSizeExpansionEx = arr: Exit Function '����
            ReDim arrRE(1 To RowCount, 1 To ColumnCount)
            If u1 - l1 = 0 And u2 - l2 = 0 Then
                For i = 1 To RowCount
                    For j = 1 To ColumnCount
                        Cover arrRE(i, j), arr(l1, l2)
                    Next
                Next
            ElseIf u1 - l1 = 0 Then
                ja = l2
                For j = 1 To MinParams2(u2 - l2 + 1, ColumnCount)
                    For i = 1 To RowCount
                        Cover arrRE(i, j), arr(l1, ja)
                    Next
                    ja = ja + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For j = u2 - l2 + 2 To ColumnCount
                        For i = 1 To RowCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                End If
            ElseIf u2 - l2 = 0 Then
                ia = l1
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    For j = 1 To ColumnCount
                        Cover arrRE(i, j), arr(ia, l2)
                    Next
                    ia = ia + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For i = u1 - l1 + 2 To RowCount
                        For j = 1 To ColumnCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                End If
            Else
                ia = l1
                c = MinParams2(u2 - l2 + 1, ColumnCount)
                For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                    ja = l2
                    For j = 1 To c
                        Cover arrRE(i, j), arr(ia, ja)
                        ja = ja + 1
                    Next
                    ia = ia + 1
                Next
                If Not IsEmpty(FillValue) Then
                    For i = u1 - l1 + 2 To RowCount
                        For j = 1 To ColumnCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                    For i = 1 To MinParams2(u1 - l1 + 1, RowCount)
                        For j = c + 1 To ColumnCount
                            Cover arrRE(i, j), FillValue
                        Next
                    Next
                End If
            End If
        Case 0
            ReDim arrRE(1 To RowCount, 1 To ColumnCount)
            For i = 1 To RowCount
                For j = 1 To ColumnCount
                    Cover arrRE(i, j), arr
                Next
            Next
    End Select
    ArrSizeExpansionEx = arrRE
End Function

'����������������������������ʱ�ᱻ����
Public Function ArrIndexExpansion(ByRef arr, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional FillValue = Empty)
    Dim l1 As Long, u1 As Long, l2 As Long, u2 As Long
    Dim i As Long, j As Long
    Dim arrRE(): arrRE = arr
    l1 = LBound(arr, 1): u1 = UBound(arr, 1)
    Select Case ArrDimension(arr)
        Case 1
            If IsMissing(RowIndex) Then Exit Function
            IndexIsCurrencyToCount_ RowIndex, l1, u1
            If RowIndex < l1 Then
                ReDim arr(RowIndex To u1)
                For i = l1 To u1
                    Cover arr(i), arrRE(i)
                Next
            ElseIf RowIndex > u1 Then
                ReDim Preserve arr(l1 To RowIndex)
            End If
            If Not IsEmpty(FillValue) Then
                For i = RowIndex To l1 - 1
                    Cover arr(i), FillValue
                Next
                For i = u1 + 1 To RowIndex
                    Cover arr(i), FillValue
                Next
            End If
        Case 2
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            If IsMissing(RowIndex) Then RowIndex = l1
            If IsMissing(ColumnIndex) Then ColumnIndex = l2
            IndexIsCurrencyToCount_ RowIndex, l1, u1
            IndexIsCurrencyToCount_ ColumnIndex, l2, u2
            Dim p As Boolean: p = False
            If RowIndex < l1 Then
                If ColumnIndex < l2 Then
                    ReDim arr(RowIndex To u1, ColumnIndex To u2): p = True
                ElseIf ColumnIndex > u2 Then
                    ReDim arr(RowIndex To u1, l2 To ColumnIndex): p = True
                Else
                    ReDim arr(RowIndex To u1, l2 To u2): p = True
                End If
            ElseIf RowIndex > u1 Then
                If ColumnIndex < l2 Then
                    ReDim arr(l1 To RowIndex, ColumnIndex To u2): p = True
                ElseIf ColumnIndex > u2 Then
                    ReDim arr(l1 To RowIndex, l2 To ColumnIndex): p = True
                Else
                    ReDim arr(l1 To RowIndex, l2 To u2): p = True
                End If
            ElseIf ColumnIndex < l2 Then
                ReDim arr(l1 To u1, ColumnIndex To u2): p = True
            ElseIf ColumnIndex > u2 Then
                ReDim Preserve arr(l1 To u1, l2 To ColumnIndex)
            End If
            If p Then
                For i = l1 To u1
                    For j = l2 To u2
                        Cover arr(i, j), arrRE(i, j)
                    Next
                Next
            End If
            If Not IsEmpty(FillValue) Then
                For i = RowIndex To l1 - 1
                    For j = ColumnIndex To l2 - 1
                        Cover arr(i, j), FillValue
                    Next
                    For j = l2 To u2
                        Cover arr(i, j), FillValue
                    Next
                    For j = u2 + 1 To ColumnIndex
                        Cover arr(i, j), FillValue
                    Next
                Next
                For i = u1 + 1 To RowIndex
                    For j = ColumnIndex To l2 - 1
                        Cover arr(i, j), FillValue
                    Next
                    For j = l2 To u2
                        Cover arr(i, j), FillValue
                    Next
                    For j = u2 + 1 To ColumnIndex
                        Cover arr(i, j), FillValue
                    Next
                Next
                For i = l1 To u1
                    For j = ColumnIndex To l2 - 1
                        Cover arr(i, j), FillValue
                    Next
                    For j = u2 + 1 To ColumnIndex
                        Cover arr(i, j), FillValue
                    Next
                Next
            End If
    End Select
    ArrIndexExpansion = arr
End Function

Public Function Generator_(Optional v)
    Static arr, i As Long
    If IsMissing(v) And IsError(v) Then
        If IsArray(arr) Then
            Cover Generator_, i
            i = i + 1
        Else
            Cover Generator_, arr
        End If
    Else
        Cover arr, v
        i = LBound(v)
    End If
End Function


'���ֵ��RowIndexArr��ColumnIndexArr����λ�����θ�ֵ������  ���ϵ���һ��һ��д�� ��ά
Public Function Arr2DSetValues(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values())
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    Values = ArrFlatten(Values)
    Dim lv1 As Long, uv1 As Long, p As Boolean
    lv1 = LBound(Values, 1): uv1 = UBound(Values, 1)
    p = lv1 <> uv1
    RowIndexArr = ArrFlatten_Single(RowIndexArr)
    ColumnIndexArr = ArrFlatten_Single(ColumnIndexArr)
    
    IndexIsCurrencyToCount_ RowIndexArr, l1, u1
    IndexIsCurrencyToCount_ ColumnIndexArr, l2, u2
    
    Dim RIndex, CIndex
    For Each RIndex In RowIndexArr
        If RIndex >= l1 And RIndex <= u1 Then
            For Each CIndex In ColumnIndexArr
                If CIndex >= l2 And CIndex <= u2 Then
                    If lv1 > uv1 Then Exit For
                    Cover arr2D(RIndex, CIndex), Values(lv1)
                    If p Then lv1 = lv1 + 1
                End If
            Next
        End If
    Next
    Arr2DSetValues = arr2D
End Function

'���ֵ��RowIndexArr��ColumnIndexArr����λ�����θ�ֵ������  ������һ��һ��д�� ��ά
Public Function Arr2DSetValues_LtoR(ByRef arr2D, ByVal RowIndexArr, ByVal ColumnIndexArr, ParamArray Values())
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    Values = ArrFlatten(Values)
    Dim lv1 As Long, uv1 As Long, p As Boolean
    lv1 = LBound(Values, 1): uv1 = UBound(Values, 1)
    p = lv1 <> uv1
    RowIndexArr = ArrFlatten_Single(RowIndexArr)
    ColumnIndexArr = ArrFlatten_Single(ColumnIndexArr)
    
    IndexIsCurrencyToCount_ RowIndexArr, l1, u1
    IndexIsCurrencyToCount_ ColumnIndexArr, l2, u2
    
    Dim RIndex, CIndex
    For Each CIndex In ColumnIndexArr
        If CIndex >= l2 And CIndex <= u2 Then
            For Each RIndex In RowIndexArr
                If RIndex >= l1 And RIndex <= u1 Then
                    
                    If lv1 > uv1 Then Exit For
                    Cover arr2D(RIndex, CIndex), Values(lv1)
                    If p Then lv1 = lv1 + 1
                End If
            Next
        End If
    Next
    Arr2DSetValues_LtoR = arr2D
End Function

'���ֵ��IndexArrλ�����θ�ֵ������ һά
Public Function ArrSetValues(ByRef arr1D, ByVal IndexArr, ParamArray Values())
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    l1 = LBound(arr1D, 1): u1 = UBound(arr1D, 1)
    
    IndexIsCurrencyToCount_ IndexArr, l1, u1
    
    Values = ArrFlatten(Values)
    Dim lv1 As Long, uv1 As Long, p As Boolean
    lv1 = LBound(Values, 1): uv1 = UBound(Values, 1)
    p = lv1 <> uv1
    If IsArray(IndexArr) Then
        Dim Index
        For Each Index In IndexArr
            If Index >= l1 And Index <= u1 Then
                If lv1 > uv1 Then Exit For
                Cover arr1D(Index), Values(lv1)
                If p Then lv1 = lv1 + 1
            End If
        Next
    Else
        If IndexArr >= l1 And IndexArr <= u1 Then
            For i = l1 To u1
                Cover arr1D(IndexArr), Values(lv1)
            Next
        End If
    End If
    ArrSetValues = arr1D
End Function

'��ֵ������һ���� ��ֵ��Ӧ���� ��ά
Public Function ArrSetEntireColumnValues(ByRef arr2D, ByVal ColumnIndexArr, ParamArray Values())
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndexArr, l2, u2
    
    Values = ArrFlatten(Values)
    Dim lv1 As Long, uv1 As Long, p As Boolean
    lv1 = LBound(Values, 1): uv1 = UBound(Values, 1)
    p = lv1 <> uv1
    If IsArray(ColumnIndexArr) Then
        Dim Index
        For Each Index In ColumnIndexArr
            If Index >= l2 And Index <= u2 Then
                If lv1 > uv1 Then Exit For
                For i = l1 To u1
                    Cover arr2D(i, Index), Values(lv1)
                Next
                If p Then lv1 = lv1 + 1
            End If
        Next
    Else
        If ColumnIndexArr >= l2 And ColumnIndexArr <= u2 Then
            For i = l1 To u1
                Cover arr2D(i, ColumnIndexArr), Values(lv1)
            Next
        End If
    End If
    ArrSetEntireColumnValues = arr2D
End Function

'��ֵ������һ���� ��ֵ��Ӧ���� ��ά
Public Function ArrSetEntireRowValues(ByRef arr2D, ByVal RowIndexArr, ParamArray Values())
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ RowIndexArr, l1, u1
    
    Values = ArrFlatten(Values)
    Dim lv1 As Long, uv1 As Long, p As Boolean
    lv1 = LBound(Values, 1): uv1 = UBound(Values, 1)
    p = lv1 <> uv1
    If IsArray(RowIndexArr) Then
        Dim Index
        For Each Index In RowIndexArr
            If Index >= l1 And Index <= u1 Then
                If lv1 > uv1 Then Exit For
                For i = l2 To u2
                    Cover arr2D(Index, i), Values(lv1)
                Next
                If p Then lv1 = lv1 + 1
            End If
        Next
    Else
        If RowIndexArr >= l1 And RowIndexArr <= u1 Then
            For i = l2 To u2
                Cover arr2D(RowIndexArr, i), Values(lv1)
            Next
        End If
    End If
    ArrSetEntireRowValues = arr2D
End Function

'���鸳ֵ������ ��ά
Public Function Arr2DSetArr2D(ByRef arrL, ByRef arrR, Optional ByVal RowIndex, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False)
    Dim l1 As Long, r1 As Long
    Dim l2 As Long, r2 As Long
    Dim i As Long, j As Long
    l1 = LBound(arrL, 1): r1 = UBound(arrL, 1)
    l2 = LBound(arrL, 2): r2 = UBound(arrL, 2)
    Dim lR1 As Long, rR1 As Long
    Dim lR2 As Long, rR2 As Long
    lR1 = LBound(arrR, 1): rR1 = UBound(arrR, 1)
    lR2 = LBound(arrR, 2): rR2 = UBound(arrR, 2)
    If IsMissing(ColumnIndex) Then ColumnIndex = l2
    If IsMissing(RowIndex) Then RowIndex = l1
    
    IndexIsCurrencyToCount_ RowIndex, l1, r1
    IndexIsCurrencyToCount_ ColumnIndex, l2, r2
    
    'ѭ��ĩβ����
    Dim ws As Long, hs As Long
    If Expansion Then
        If ArrValid_Index(arrL, RowIndex, ColumnIndex) = False Then
            ArrIndexExpansion arrL, RowIndex, ColumnIndex
        End If
        If ArrValid_Index(arrL, RowIndex + rR1 - lR1, ColumnIndex + rR2 - lR2) = False Then
            ArrIndexExpansion arrL, RowIndex + rR1 - lR1, ColumnIndex + rR2 - lR2
        End If
        ws = rR2 - lR2 + ColumnIndex
        hs = rR1 - lR1 + RowIndex
    ElseIf ArrValid_Index(arrL, RowIndex, ColumnIndex) = False Then
        Arr2DSetArr2D = arrL
        Exit Function
    Else
        ws = MinParams2(rR2 - lR2 + ColumnIndex, r2)
        hs = MinParams2(rR1 - lR1 + RowIndex, r1)
    End If
    Dim i2 As Long, j2 As Long
    i2 = lR1
    For i = RowIndex To hs
        j2 = lR2
        For j = ColumnIndex To ws
            Cover arrL(i, j), arrR(i2, j2)
            j2 = j2 + 1
        Next
        i2 = i2 + 1
    Next
    Arr2DSetArr2D = arrL
End Function
 
'���鸳ֵ������ һά
Public Function ArrSetArr(ByRef arrL, ByRef arrR, Optional ByVal Index, Optional Expansion As Boolean = False)
    Dim l1 As Long, r1 As Long
    Dim i As Long, v
    l1 = LBound(arrL, 1): r1 = UBound(arrL, 1)
    If IsMissing(Index) Then Index = l1
    IndexIsCurrencyToCount_ Index, l1, r1
    If Expansion Then
        If ArrValid_Index(arrL, Index) = False Then
            ArrIndexExpansion arrL, Index
        End If
        Dim n As Long
        n = ArrCount(arrR)
        If ArrValid_Index(arrL, Index + n - 1) = False Then
            ArrIndexExpansion arrL, Index + n - 1
        End If
        l1 = LBound(arrL, 1): r1 = UBound(arrL, 1)
    ElseIf ArrValid_Index(arrL, Index) = False Then
        ArrSetArr = arrL
        Exit Function
    End If
    i = Index
    For Each v In arrR
        If i > r1 Then Exit For
        Cover arrL(i), v
        i = i + 1
    Next
    ArrSetArr = arrL
End Function

'���鸳ֵ������һ��
Public Function ArrSetColumn(ByRef arrL2D, ByRef arrR, Optional ByVal ColumnIndex, Optional Expansion As Boolean = False)
    Dim l1 As Long, r1 As Long
    Dim i As Long, v
    l1 = LBound(arrL2D, 1): r1 = UBound(arrL2D, 1)
    If IsMissing(ColumnIndex) Then ColumnIndex = LBound(arrL2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LBound(arrL2D, 2), UBound(arrL2D, 2)
    
    Select Case ArrDimension(arrR)
        Case 1, 0
            If Expansion Then
                Dim n As Long
                n = ArrCount(arrR)
                If ArrValid_Index(arrL2D, l1 + n - 1, ColumnIndex) = False Then
                    ArrIndexExpansion arrL2D, l1 + n - 1, ColumnIndex
                    r1 = UBound(arrL2D, 1)
                End If
            ElseIf ArrValid_Index(arrL2D, l1, ColumnIndex) = False Then
                ArrSetColumn = arrL2D
                Exit Function
            End If
            i = l1
            For Each v In arrR
                If i > r1 Then Exit For
                Cover arrL2D(i, ColumnIndex), v
                i = i + 1
            Next
        Case 2
            Arr2DSetArr2D arrL2D, arrR, l1, ColumnIndex, Expansion
    End Select
    ArrSetColumn = arrL2D
End Function

'���鸳ֵ������һ��
Public Function ArrSetRow(ByRef arrL2D, ByRef arrR, Optional ByVal RowIndex, Optional Expansion As Boolean = False)
    Dim l1 As Long, r1 As Long
    Dim i As Long, v
    l1 = LBound(arrL2D, 2): r1 = UBound(arrL2D, 2)
    If IsMissing(RowIndex) Then RowIndex = LBound(arrL2D, 1)
    
    IndexIsCurrencyToCount_ RowIndex, LBound(arrL2D, 1), UBound(arrL2D, 1)
    
    Select Case ArrDimension(arrR)
        Case 1, 0
            If Expansion Then
                Dim n As Long
                n = ArrCount(arrR)
                If ArrValid_Index(arrL2D, RowIndex, l1 + n - 1) = False Then
                    ArrIndexExpansion arrL2D, RowIndex, l1 + n - 1
                    r1 = UBound(arrL2D, 2)
                End If
            ElseIf ArrValid_Index(arrL2D, RowIndex, l1) = False Then
                ArrSetRow = arrL2D
                Exit Function
            End If
            i = l1
            For Each v In arrR
                If i > r1 Then Exit For
                Cover arrL2D(RowIndex, i), v
                i = i + 1
            Next
        Case 2
            Arr2DSetArr2D arrL2D, arrR, RowIndex, l1, Expansion
    End Select
    ArrSetRow = arrL2D
End Function
 
'����������˳��ȡ������ֵ��������ԭ������
Public Function ArrFromIndex(arr, ByVal arrindex) As Variant
    Dim br, i&, i2&, j&, l&, l2&, u&, u2&, l1&, u1&
    l = LBound(arrindex): u = UBound(arrindex)
    If u < l Then ArrFromIndex = Array(): Exit Function
    Select Case ArrDimension(arr)
        Case 2
            l1 = LBound(arr, 1): u1 = UBound(arr, 1)
            l2 = LBound(arr, 2): u2 = UBound(arr, 2)
            
            IndexIsCurrencyToCount_ arrindex, l1, u1
            
            ReDim br(l To u, l2 To u2)
            For i = l To u
                i2 = arrindex(i)
                If i2 >= l1 And i2 <= u1 Then
                    For j = l2 To u2
                        Cover br(i, j), arr(i2, j)
                    Next
                End If
            Next
        Case 1
            ReDim br(l To u)
            l1 = LBound(arr): u1 = UBound(arr)
            
            IndexIsCurrencyToCount_ arrindex, l1, u1
            
            For i = l To u
                i2 = arrindex(i)
                If i2 >= l1 And i2 <= u1 Then
                    Cover br(i), arr(i2)
                End If
            Next
        Case Else
            br = Array()
    End Select
    ArrFromIndex = br
End Function
 
'��������������=Trueȡ������ֵ������ɸѡ����
Public Function ArrFromBoolea(arr, arrBoolea) As Variant
    Dim br, i&, i2&, j&, l2&, u2&, l1&, u1&
    Dim v, n As Long
    l1 = LBound(arr, 1): u1 = UBound(arr, 1)
    n = 0
    i = l1
    For Each v In arrBoolea
        If i > u1 Then Exit For
        If v Then n = n + 1
        i = i + 1
    Next
    If n > 0 Then
        Select Case ArrDimension(arr)
            Case 2
                l2 = LBound(arr, 2): u2 = UBound(arr, 2)
                ReDim br(1 To n, l2 To u2)
                i = l1
                i2 = 1
                For Each v In arrBoolea
                    If i > u1 Then Exit For
                    If v Then
                        For j = l2 To u2
                            Cover br(i2, j), arr(i, j)
                        Next
                        i2 = i2 + 1
                    End If
                    i = i + 1
                Next
            Case 1
                ReDim br(1 To n)
                i = l1
                i2 = 1
                For Each v In arrBoolea
                    If i > u1 Then Exit For
                    If v Then
                        Cover br(i2), arr(i)
                        i2 = i2 + 1
                    End If
                    i = i + 1
                Next
            Case Else
                br = Array()
        End Select
        ArrFromBoolea = br
    Else
        ArrFromBoolea = Array()
    End If
End Function
 
'�����������
Public Function ArrRandSort(ByVal arr) As Variant
    Randomize
    Dim l As Long, u As Long
    l = LBound(arr): u = UBound(arr)
    Dim j As Long, i As Long
    For i = l To u - 1
        j = RandBetween(i + 1, u)
        Exchange arr(i), arr(j)
    Next
    ArrRandSort = arr
End Function
 
'��ά�����ȶ�����
Public Function ArrSort2D(arr, Index, Optional Order As Boolean = True) As Variant
    ArrSort2D = ArrFromIndex(arr, ArrSort(ArrGetColumn(arr, Index), Order))
End Function

'��ά��������ȶ�����
Public Function ArrSort2Ds(arr, Indexs, Optional Orders = True) As Variant
    Dim i As Long, u As Long, arrindex
    Dim IndexsRE, OrdersRE
    IndexsRE = ArrFlatten_Single(Indexs)
    u = UBound(IndexsRE)
    If IsArray(Orders) Then
        OrdersRE = ArrSizeExpansion2(Orders, u, True)
    Else
        OrdersRE = ArrSizeExpansion2(Orders, u, Orders)
    End If
    arrindex = ArrSort(ArrGetColumn(arr, IndexsRE(u)), CBool(OrdersRE(u)))
    For i = UBound(IndexsRE) To 1 Step -1
        arrindex = ArrSortNext(ArrGetColumn(arr, IndexsRE(i)), arrindex, CBool(OrdersRE(i)))
    Next
    ArrSort2Ds = ArrFromIndex(arr, arrindex)
End Function
 
'һά�����ȶ�����
Public Function ArrSort1D(arr, Optional Order As Boolean = True) As Variant
    ArrSort1D = ArrFromIndex(arr, ArrSort(arr, Order))
End Function
 
'һά�����ȶ�����  ����������Order=True ��������
'���ӣ�����arr��ά����
'ArrColumns = ArrGetColumn(arr, 1)  'ȡ��arr������
'arrIndex = ArrSort(ArrColumns)  '�����������򷵻���������
'arrOrder = ArrFromIndex(arr, arrIndex) '������������ȡ����������
Public Function ArrSort(arr, Optional Order As Boolean = True) As Variant
    Dim i As Long, i2 As Long, l As Long, u As Long, s As Long, T
    l = LBound(arr): u = UBound(arr)
    ReDim x&(l To u), Z(l To u + 1) As Boolean
    For i = l To u
        x(i) = i
    Next
    Z(u + 1) = True '���������λ��
    '����
    If Order Then Call QuickSort1(arr, x, l, u) Else Call QuickSort2(arr, x, l, u)
    If Order Then Call AZE(arr, x, l, u) '��ֵ����
    '��֤�ȶ����򣬶���ֵͬ��������
    i = l: T = arr(x(i)): i2 = i
    Do
        Do
            i2 = i2 + 1: If Z(i2) Then Exit Do Else If arr(x(i2)) <> T Then Exit Do
        Loop
        If i2 - i > 1 Then Call QuickSort(x, i, i2 - 1)
        If i2 > u Then Exit Do Else i = i2: T = arr(x(i))
    Loop
    ArrSort = x
End Function
 
'������������
'���ӣ���1,2������
'arrIndex = ArrSort(ArrGetColumn(arr, 1)) '��һ������
'arrIndex = ArrSortNext(ArrGetColumn(arr, 2), arrIndex) '��2������
'arrorder = ArrFromIndex(arr, arrIndex) '���ؽ��
Public Function ArrSortNext(arr, Indexs, Optional Order As Boolean = True) As Variant
    ArrSortNext = ArrFromIndex(Indexs, ArrSort(ArrFromIndex(arr, Indexs), Order))
End Function

Private Function QuickSort(x, l&, u&) 'A-Z QuickSort '����ȶ�����ʱ����ͬkey��Indexֵ��������
    Dim i&, j&, n&, r&
    i = l: j = u: r = x((l + u) \ 2)
    While i < j
        While x(i) < r: i = i + 1: Wend 'A-Z
        While x(j) > r: j = j - 1: Wend 'A-Z
        If i <= j Then: n = x(i): x(i) = x(j): x(j) = n: i = i + 1: j = j - 1
    Wend
    If l < j Then Call QuickSort(x, l, j)
    If i < u Then Call QuickSort(x, i, u)
End Function

Private Function QuickSort1(ar, x, l&, u&)   'A-Z QuickSort ��ԭ����j2�ж�Ӧ���ݽ�����������
    Dim i&, j&, n&, r
    i = l: j = u: r = ar(x((l + u) \ 2))
    While i < j
        While ar(x(i)) < r And i < u: i = i + 1: Wend    'A-Z
        While ar(x(j)) > r And j > l: j = j - 1: Wend    'A-Z
        If i <= j Then n = x(i): x(i) = x(j): x(j) = n: i = i + 1: j = j - 1
    Wend
    If l < j Then Call QuickSort1(ar, x, l, j)
    If i < u Then Call QuickSort1(ar, x, i, u)
End Function

Private Function QuickSort2(ar, x, l&, u&)   'Z-A QuickSort ��ԭ����j2�ж�Ӧ���ݽ��н�������
    Dim i&, j&, n&, r
    i = l: j = u: r = ar(x((l + u) \ 2))
    While i < j
        While ar(x(i)) > r And i < u: i = i + 1: Wend  'Z-A
        While ar(x(j)) < r And j > l: j = j - 1: Wend  'Z-A
        If i <= j Then n = x(i): x(i) = x(j): x(j) = n: i = i + 1: j = j - 1
    Wend
    If l < j Then Call QuickSort2(ar, x, l, j)
    If i < u Then Call QuickSort2(ar, x, i, u)
End Function

Private Function AZE(ar, x, l&, u&)   '��������ɺ�Ŀ�ֵ�ƶ������
    Dim i&, i2&, y
    For i = l To u
        If ar(x(i)) <> "" Then
            y = x
            For i2 = l To i - 1
                x(u - i + i2 + 1) = y(i2)
            Next
            For i2 = i To u
                x(i2 - i + l) = y(i2)
            Next
            Exit For
        End If
    Next
End Function
 
'��ά�����Զ�������
Public Function ArrCustomSort2D(arrValue, arrKey, Index, Optional IsLike As Boolean = False) As Variant
    ArrCustomSort2D = ArrFromIndex(arrValue, ArrCustomSort(ArrGetColumn(arrValue, Index), arrKey, IsLike))
End Function
 
'�Զ�������  CustomSort(��������, �Զ�������, Likeƥ��) ������������
Public Function ArrCustomSort(arrValue, ByVal arrKey, Optional IsLike As Boolean = False)
    Dim i As Long, j As Long, k As Long, tmp As Long
    Dim l As Long, u As Long, x() As Long
    arrKey = ArrFlatten(arrKey)
    l = LBound(arrValue): u = UBound(arrValue)
    ReDim x(l To u)
    For i = l To u
        x(i) = i
    Next
    k = l
    If IsLike Then
        For i = LBound(arrKey) To UBound(arrKey)
            For j = k To u
                If arrValue(x(j)) Like arrKey(i) Then
                    If j <> k Then
                        Cover tmp, x(j)
                        Cover x(j), x(k)
                        Cover x(k), tmp
                    End If
                    k = k + 1
                End If
            Next
        Next
    Else
        For i = LBound(arrKey) To UBound(arrKey)
            For j = k To u
                If arrValue(x(j)) = arrKey(i) Then
                    If j <> k Then
                        Cover tmp, x(j)
                        Cover x(j), x(k)
                        Cover x(k), tmp
                    End If
                    k = k + 1
                End If
            Next
        Next
    End If
 
    '��֤��ƥ����ȶ����򣬶Բ�ƥ��ֵ��������
    If u - k > 0 Then Call QuickSort(x, k, u)
    ArrCustomSort = x
End Function

'����Number��arrInterval�������λ�� λ��������LBound(arrInterval)��UBound(arr)+1 arrInterval��������˳��
Public Function ArrInInterval(ByVal arrInterval, Number) As Long
    Dim i As Long
    arrInterval = ArrSort1D(arrInterval)
    For i = LBound(arrInterval) To UBound(arrInterval)
        If Number < arrInterval(i) Then
            ArrInInterval = i
            Exit Function
        End If
    Next
    ArrInInterval = i
End Function

'����Number��arrInterval�������λ�� ������ λ��������LBound(arrInterval)��UBound(arr)+1 arrInterval��������˳��
Public Function ArrInIntervalEqual(ByVal arrInterval, Number) As Long
    Dim i As Long
    arrInterval = ArrSort1D(arrInterval)
    For i = LBound(arrInterval) To UBound(arrInterval)
        If Number <= arrInterval(i) Then
            ArrInIntervalEqual = i
            Exit Function
        End If
    Next
    ArrInIntervalEqual = i
End Function

'����С��v������
Public Function ArrFindLessIndex(arr_Small, V_Large, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLessIndex = LBound(arr_Small) - 1
    If IsMissing(Start) Then
        Start = LBound(arr_Small)
    ElseIf Start < LBound(arr_Small) Then
        Start = LBound(arr_Small)
    ElseIf Start > UBound(arr_Small) Then
        Exit Function
    End If
    For i = Start To UBound(arr_Small)
        If arr_Small(i) < V_Large Then
            ArrFindLessIndex = i
            Exit For
        End If
    Next
End Function
 
'����С��v������ ����
Public Function ArrFindLessIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLessIndexRev = LBound(arr_Small) - 1
    If IsMissing(Start) Then
        Start = UBound(arr_Small)
    ElseIf Start > UBound(arr_Small) Then
        Start = UBound(arr_Small)
    ElseIf Start < LBound(arr_Small) Then
        Exit Function
    End If
    For i = Start To LBound(arr_Small) Step -1
        If arr_Small(i) < V_Large Then
            ArrFindLessIndexRev = i
            Exit For
        End If
    Next
End Function
 
'����С�ڵ���v������
Public Function ArrFindLessEqualIndex(arr_Small, V_Large, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLessEqualIndex = LBound(arr_Small) - 1
    If IsMissing(Start) Then
        Start = LBound(arr_Small)
    ElseIf Start < LBound(arr_Small) Then
        Start = LBound(arr_Small)
    ElseIf Start > UBound(arr_Small) Then
        Exit Function
    End If
    For i = Start To UBound(arr_Small)
        If arr_Small(i) <= V_Large Then
            ArrFindLessEqualIndex = i
            Exit For
        End If
    Next
End Function
 
'����С�ڵ���v������ ����
Public Function ArrFindLessEqualIndexRev(arr_Small, V_Large, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLessEqualIndexRev = LBound(arr_Small) - 1
    If IsMissing(Start) Then
        Start = UBound(arr_Small)
    ElseIf Start > UBound(arr_Small) Then
        Start = UBound(arr_Small)
    ElseIf Start < LBound(arr_Small) Then
        Exit Function
    End If
    For i = Start To LBound(arr_Small) Step -1
        If arr_Small(i) <= V_Large Then
            ArrFindLessEqualIndexRev = i
            Exit For
        End If
    Next
End Function
 
'���Ҵ���v������
Public Function ArrFindGreaterIndex(arr_Large, V_Small, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindGreaterIndex = LBound(arr_Large) - 1
    If IsMissing(Start) Then
        Start = LBound(arr_Large)
    ElseIf Start < LBound(arr_Large) Then
        Start = LBound(arr_Large)
    ElseIf Start > UBound(arr_Large) Then
        Exit Function
    End If
    For i = Start To UBound(arr_Large)
        If arr_Large(i) > V_Small Then
            ArrFindGreaterIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҵ���v������ ����
Public Function ArrFindGreaterIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindGreaterIndexRev = LBound(arr_Large) - 1
    If IsMissing(Start) Then
        Start = UBound(arr_Large)
    ElseIf Start > UBound(arr_Large) Then
        Start = UBound(arr_Large)
    ElseIf Start < LBound(arr_Large) Then
        Exit Function
    End If
    For i = Start To LBound(arr_Large) Step -1
        If arr_Large(i) > V_Small Then
            ArrFindGreaterIndexRev = i
            Exit For
        End If
    Next
End Function
 
'���Ҵ��ڵ���v������
Public Function ArrFindGreaterEqualIndex(arr_Large, V_Small, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindGreaterEqualIndex = LBound(arr_Large) - 1
    If IsMissing(Start) Then
        Start = LBound(arr_Large)
    ElseIf Start < LBound(arr_Large) Then
        Start = LBound(arr_Large)
    ElseIf Start > UBound(arr_Large) Then
        Exit Function
    End If
    For i = Start To UBound(arr_Large)
        If arr_Large(i) >= V_Small Then
            ArrFindGreaterEqualIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҵ��ڵ���v������ ����
Public Function ArrFindGreaterEqualIndexRev(arr_Large, V_Small, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindGreaterEqualIndexRev = LBound(arr_Large) - 1
    If IsMissing(Start) Then
        Start = UBound(arr_Large)
    ElseIf Start > UBound(arr_Large) Then
        Start = UBound(arr_Large)
    ElseIf Start < LBound(arr_Large) Then
        Exit Function
    End If
    For i = Start To LBound(arr_Large) Step -1
        If arr_Large(i) >= V_Small Then
            ArrFindGreaterEqualIndexRev = i
            Exit For
        End If
    Next
End Function

'���Ҷ�Ӧֵ���� Like
Public Function ArrFindLikeIndex(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLikeIndex = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = LBound(arr)
    ElseIf Start < LBound(arr) Then
        Start = LBound(arr)
    ElseIf Start > UBound(arr) Then
        Exit Function
    End If
    For i = Start To UBound(arr)
        If arr(i) Like v Then
            ArrFindLikeIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҷ�Ӧֵ�������� Like
Public Function ArrFindLikeIndexRev(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindLikeIndexRev = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = UBound(arr)
    ElseIf Start > UBound(arr) Then
        Start = UBound(arr)
    ElseIf Start < LBound(arr) Then
        Exit Function
    End If
    For i = Start To LBound(arr) Step -1
        If arr(i) Like v Then
            ArrFindLikeIndexRev = i
            Exit For
        End If
    Next
End Function

'���Ҷ�Ӧֵ���� Not Like
Public Function ArrFindNotLikeIndex(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindNotLikeIndex = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = LBound(arr)
    ElseIf Start < LBound(arr) Then
        Start = LBound(arr)
    ElseIf Start > UBound(arr) Then
        Exit Function
    End If
    For i = Start To UBound(arr)
        If Not arr(i) Like v Then
            ArrFindNotLikeIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҷ�Ӧֵ�������� Not Like
Public Function ArrFindNotLikeIndexRev(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindNotLikeIndexRev = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = UBound(arr)
    ElseIf Start > UBound(arr) Then
        Start = UBound(arr)
    ElseIf Start < LBound(arr) Then
        Exit Function
    End If
    For i = Start To LBound(arr) Step -1
        If Not arr(i) Like v Then
            ArrFindNotLikeIndexRev = i
            Exit For
        End If
    Next
End Function
 
'���Ҷ�Ӧֵ����
Public Function ArrFindIndex(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindIndex = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = LBound(arr)
    ElseIf Start < LBound(arr) Then
        Start = LBound(arr)
    ElseIf Start > UBound(arr) Then
        Exit Function
    End If
    For i = Start To UBound(arr)
        If arr(i) = v Then
            ArrFindIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҷ�Ӧֵ��������
Public Function ArrFindIndexRev(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindIndexRev = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = UBound(arr)
    ElseIf Start > UBound(arr) Then
        Start = UBound(arr)
    ElseIf Start < LBound(arr) Then
        Exit Function
    End If
    For i = Start To LBound(arr) Step -1
        If arr(i) = v Then
            ArrFindIndexRev = i
            Exit For
        End If
    Next
End Function

'���Ҷ�Ӧֵ���� ������
Public Function ArrFindNotIndex(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindNotIndex = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = LBound(arr)
    ElseIf Start < LBound(arr) Then
        Start = LBound(arr)
    ElseIf Start > UBound(arr) Then
        Exit Function
    End If
    For i = Start To UBound(arr)
        If arr(i) <> v Then
            ArrFindNotIndex = i
            Exit For
        End If
    Next
End Function
 
'���Ҳ����ڵ�ֵ��������
Public Function ArrFindNotIndexRev(arr, v, Optional ByVal Start) As Long
    Dim i As Long
    ArrFindNotIndexRev = LBound(arr) - 1
    If IsMissing(Start) Then
        Start = UBound(arr)
    ElseIf Start > UBound(arr) Then
        Start = UBound(arr)
    ElseIf Start < LBound(arr) Then
        Exit Function
    End If
    For i = Start To LBound(arr) Step -1
        If arr(i) <> v Then
            ArrFindNotIndexRev = i
            Exit For
        End If
    Next
End Function

'���Ҷ�Ӧֵ���� ����
Public Function ArrFindRegIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim i As Long
        ArrFindRegIndex = LBound(arr) - 1
        If IsMissing(Start) Then
            Start = LBound(arr)
        ElseIf Start < LBound(arr) Then
            Start = LBound(arr)
        ElseIf Start > UBound(arr) Then
            Exit Function
        End If
        For i = Start To UBound(arr)
            If .test(arr(i)) Then
                ArrFindRegIndex = i
                Exit For
            End If
        Next
    End With
End Function
 
'���Ҷ�Ӧֵ���� ���� ����
Public Function ArrFindRegIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim i As Long
        ArrFindRegIndexRev = LBound(arr) - 1
        If IsMissing(Start) Then
            Start = UBound(arr)
        ElseIf Start > UBound(arr) Then
            Start = UBound(arr)
        ElseIf Start < LBound(arr) Then
            Exit Function
        End If
        For i = Start To LBound(arr) Step -1
            If .test(arr(i)) Then
                ArrFindRegIndexRev = i
                Exit For
            End If
        Next
    End With
End Function

'���Ҷ�Ӧֵ���� ����������
Public Function ArrFindRegNotIndex(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim i As Long
        ArrFindRegNotIndex = LBound(arr) - 1
        If IsMissing(Start) Then
            Start = LBound(arr)
        ElseIf Start < LBound(arr) Then
            Start = LBound(arr)
        ElseIf Start > UBound(arr) Then
            Exit Function
        End If
        For i = Start To UBound(arr)
            If Not .test(arr(i)) Then
                ArrFindRegNotIndex = i
                Exit For
            End If
        Next
    End With
End Function

'���Ҷ�Ӧֵ���� ���������� ����
Public Function ArrFindRegNotIndexRev(arr, Pattern, Optional ByVal Start, Optional ByVal ignoreCase As Boolean = False) As Long
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim i As Long
        ArrFindRegNotIndexRev = LBound(arr) - 1
        If IsMissing(Start) Then
            Start = UBound(arr)
        ElseIf Start > UBound(arr) Then
            Start = UBound(arr)
        ElseIf Start < LBound(arr) Then
            Exit Function
        End If
        For i = Start To LBound(arr) Step -1
            If Not .test(arr(i)) Then
                ArrFindRegNotIndexRev = i
                Exit For
            End If
        Next
    End With
End Function

'��ά����������� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    If RowFirst Then
        If StartColumn > u2 Then
            StartColumn = l2
            StartRow = StartRow + 1
        End If
        If StartRow > u1 Then
            Exit Function
        End If
        For j = StartColumn To u2 '��һ�δ�����λ�ò���
            If arr2D(StartRow, j) = v Then     '*********����*********
                ArrFindIndex2D = Array(StartRow, j)
                Exit Function
            End If
        Next
        For i = StartRow + 1 To u1
            For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) = v Then     '*********����*********
                    ArrFindIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    Else
        If StartRow > u1 Then
            StartRow = l1
            StartColumn = StartColumn + 1
        End If
        If StartColumn > u2 Then
            Exit Function
        End If
        For i = StartRow To u1 '��һ�δ�����λ�ò���
            If arr2D(i, StartColumn) = v Then     '*********����*********
                ArrFindIndex2D = Array(i, StartColumn)
                Exit Function
            End If
        Next
        For j = StartColumn + 1 To u2
            For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) = v Then     '*********����*********
                    ArrFindIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    End If
End Function

'��ά����������� ������ �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindNotIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindNotIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    If RowFirst Then
        If StartColumn > u2 Then
            StartColumn = l2
            StartRow = StartRow + 1
        End If
        If StartRow > u1 Then
            Exit Function
        End If
        For j = StartColumn To u2 '��һ�δ�����λ�ò���
            If arr2D(StartRow, j) <> v Then     '*********����*********
                ArrFindNotIndex2D = Array(StartRow, j)
                Exit Function
            End If
        Next
        For i = StartRow + 1 To u1
            For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) <> v Then     '*********����*********
                    ArrFindNotIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    Else
        If StartRow > u1 Then
            StartRow = l1
            StartColumn = StartColumn + 1
        End If
        If StartColumn > u2 Then
            Exit Function
        End If
        For i = StartRow To u1 '��һ�δ�����λ�ò���
            If arr2D(i, StartColumn) <> v Then     '*********����*********
                ArrFindNotIndex2D = Array(i, StartColumn)
                Exit Function
            End If
        Next
        For j = StartColumn + 1 To u2
            For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) <> v Then     '*********����*********
                    ArrFindNotIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    End If
End Function

'��ά����������� Like���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindLikeIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    If RowFirst Then
        If StartColumn > u2 Then
            StartColumn = l2
            StartRow = StartRow + 1
        End If
        If StartRow > u1 Then
            Exit Function
        End If
        For j = StartColumn To u2 '��һ�δ�����λ�ò���
            If arr2D(StartRow, j) Like v Then     '*********����*********
                ArrFindLikeIndex2D = Array(StartRow, j)
                Exit Function
            End If
        Next
        For i = StartRow + 1 To u1
            For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) Like v Then     '*********����*********
                    ArrFindLikeIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    Else
        If StartRow > u1 Then
            StartRow = l1
            StartColumn = StartColumn + 1
        End If
        If StartColumn > u2 Then
            Exit Function
        End If
        For i = StartRow To u1 '��һ�δ�����λ�ò���
            If arr2D(i, StartColumn) Like v Then     '*********����*********
                ArrFindLikeIndex2D = Array(i, StartColumn)
                Exit Function
            End If
        Next
        For j = StartColumn + 1 To u2
            For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If arr2D(i, j) Like v Then     '*********����*********
                    ArrFindLikeIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    End If
End Function

'��ά����������� Not Like���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindNotLikeIndex2D(ByRef arr2D, v, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindNotLikeIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    If RowFirst Then
        If StartColumn > u2 Then
            StartColumn = l2
            StartRow = StartRow + 1
        End If
        If StartRow > u1 Then
            Exit Function
        End If
        For j = StartColumn To u2 '��һ�δ�����λ�ò���
            If Not arr2D(StartRow, j) Like v Then     '*********����*********
                ArrFindNotLikeIndex2D = Array(StartRow, j)
                Exit Function
            End If
        Next
        For i = StartRow + 1 To u1
            For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If Not arr2D(i, j) Like v Then     '*********����*********
                    ArrFindNotLikeIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    Else
        If StartRow > u1 Then
            StartRow = l1
            StartColumn = StartColumn + 1
        End If
        If StartColumn > u2 Then
            Exit Function
        End If
        For i = StartRow To u1 '��һ�δ�����λ�ò���
            If Not arr2D(i, StartColumn) Like v Then     '*********����*********
                ArrFindNotLikeIndex2D = Array(i, StartColumn)
                Exit Function
            End If
        Next
        For j = StartColumn + 1 To u2
            For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                If Not arr2D(i, j) Like v Then     '*********����*********
                    ArrFindNotLikeIndex2D = Array(i, j)
                    Exit Function
                End If
            Next
        Next
    End If
End Function

'��ά����������� ���� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindRegIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindRegIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        If RowFirst Then
            If StartColumn > u2 Then
                StartColumn = l2
                StartRow = StartRow + 1
            End If
            If StartRow > u1 Then
                Exit Function
            End If
            For j = StartColumn To u2 '��һ�δ�����λ�ò���
                If .test(arr2D(StartRow, j)) Then      '*********����*********
                    ArrFindRegIndex2D = Array(StartRow, j)
                    Exit Function
                End If
            Next
            For i = StartRow + 1 To u1
                For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                    If .test(arr2D(i, j)) Then       '*********����*********
                        ArrFindRegIndex2D = Array(i, j)
                        Exit Function
                    End If
                Next
            Next
        Else
            If StartRow > u1 Then
                StartRow = l1
                StartColumn = StartColumn + 1
            End If
            If StartColumn > u2 Then
                Exit Function
            End If
            For i = StartRow To u1 '��һ�δ�����λ�ò���
                If .test(arr2D(i, StartColumn)) Then      '*********����*********
                    ArrFindRegIndex2D = Array(i, StartColumn)
                    Exit Function
                End If
            Next
            For j = StartColumn + 1 To u2
                For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                    If .test(arr2D(i, j)) Then       '*********����*********
                        ArrFindRegIndex2D = Array(i, j)
                        Exit Function
                    End If
                Next
            Next
        End If
    End With
End Function

'��ά����������� ���������� �ҵ�����Array(RowIndex, ColumnIndex) �Ҳ������ؿ�����
Public Function ArrFindRegNotIndex2D(ByRef arr2D, Pattern, Optional ByVal StartRow, Optional ByVal StartColumn, Optional RowFirst As Boolean = True, Optional ByVal ignoreCase As Boolean = False) As Variant
    Dim i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    l1 = LBound(arr2D, 1): u1 = UBound(arr2D, 1)
    l2 = LBound(arr2D, 2): u2 = UBound(arr2D, 2)
    ArrFindRegNotIndex2D = Array()
    If IsMissing(StartRow) Then StartRow = l1
    IndexIsCurrencyToCount_ StartRow, l1, u1
    If StartRow < l1 Then
        StartRow = l1
    End If
    If IsMissing(StartColumn) Then StartColumn = l2
    IndexIsCurrencyToCount_ StartColumn, l2, u2
    If StartColumn < l2 Then
        StartColumn = l2
    End If
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        If RowFirst Then
            If StartColumn > u2 Then
                StartColumn = l2
                StartRow = StartRow + 1
            End If
            If StartRow > u1 Then
                Exit Function
            End If
            For j = StartColumn To u2 '��һ�δ�����λ�ò���
                If Not .test(arr2D(StartRow, j)) Then      '*********����*********
                    ArrFindRegNotIndex2D = Array(StartRow, j)
                    Exit Function
                End If
            Next
            For i = StartRow + 1 To u1
                For j = l2 To u2  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                    If Not .test(arr2D(i, j)) Then        '*********����*********
                        ArrFindRegNotIndex2D = Array(i, j)
                        Exit Function
                    End If
                Next
            Next
        Else
            If StartRow > u1 Then
                StartRow = l1
                StartColumn = StartColumn + 1
            End If
            If StartColumn > u2 Then
                Exit Function
            End If
            For i = StartRow To u1 '��һ�δ�����λ�ò���
                If Not .test(arr2D(i, StartColumn)) Then       '*********����*********
                    ArrFindRegNotIndex2D = Array(i, StartColumn)
                    Exit Function
                End If
            Next
            For j = StartColumn + 1 To u2
                For i = l1 To u1  '�ڶ��ο�ʼ�ָ�������ͷ���� �����ܱ�֤���ҳ�����������
                    If Not .test(arr2D(i, j)) Then        '*********����*********
                        ArrFindRegNotIndex2D = Array(i, j)
                        Exit Function
                    End If
                Next
            Next
        End If
    End With
End Function

'��������Ч�� �д��󷵻�True
Public Function ArrValid_Index(ByRef arr, Optional ByVal RowIndex, Optional ByVal ColumnIndex) As Boolean
    Dim p As Boolean: p = True
    If Not IsMissing(RowIndex) Then
        p = RowIndex >= LBound(arr, 1) And RowIndex <= UBound(arr, 1)
    End If
    If Not IsMissing(ColumnIndex) And p Then
        p = ColumnIndex >= LBound(arr, 2) And ColumnIndex <= UBound(arr, 2)
    End If
    ArrValid_Index = p
End Function
 
'��������Ч�� �д��󷵻�True
Public Function ArrValid_InError(arr) As Boolean
    Dim v
    ArrValid_InError = False
    For Each v In arr
        If IsError(v) Then
            ArrValid_InError = True
            Exit For
        End If
    Next
End Function
 
'��������Ч�� ȫ�������ַ���True
Public Function ArrValid_NumericAll(arr, Optional InEmpty As Boolean = True, Optional IsStr As Boolean = True) As Boolean
    Dim v
    ArrValid_NumericAll = True
    For Each v In arr
        If IsNumeric(v) Then
            If InEmpty = False Then
                If IsEmpty(v) Then
                    ArrValid_NumericAll = False
                    Exit For
                End If
            ElseIf IsStr = False Then
                If TypeName(v) = "String" Then
                    ArrValid_NumericAll = False
                    Exit For
                End If
            End If
        Else
            ArrValid_NumericAll = False
            Exit For
        End If
    Next
End Function
 
'��������Ч�� ȫ�������ڷ���True
Public Function ArrValid_DateAll(arr, Optional IsStr As Boolean = True) As Boolean
    Dim v
    ArrValid_DateAll = True
    For Each v In arr
        If IsDate(v) Then
            If IsStr = False Then
                If TypeName(v) = "String" Then
                    ArrValid_DateAll = False
                    Exit For
                End If
            End If
        Else
            ArrValid_DateAll = False
            Exit For
        End If
    Next
End Function
 
'��������Ч������һ�� ���� ƥ�䷵��True
Public Function ArrValid_Reg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim v
        ArrValid_Reg = False
        For Each v In arr
            If .test(v) Then
                ArrValid_Reg = True
                Exit For
            End If
        Next
    End With
End Function
 
'��������Ч������ȫ�� ���� ȫ��ƥ�䷵��True
Public Function ArrValid_RegAll(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        Dim v
        ArrValid_RegAll = True
        For Each v In arr
            If .test(v) = False Then
                 ArrValid_RegAll = False
            End If
        Next
    End With
End Function
 
'��������Ч���Ƿ����ظ� �ظ�����True
Public Function ArrValid_Repeat(arr, Optional CompareMode As CompareMethod = BinaryCompare) As Boolean
    Dim dic As Object, v As Variant
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    ArrValid_Repeat = False
    For Each v In arr
        If dic.Exists(v) Then
            ArrValid_Repeat = True
            Exit For
        Else
            dic.Add v, 0
        End If
    Next
End Function
 
'��������Ч���Ƿ��������
Public Function ArrValid_Incremental(ParamArray arr()) As Boolean
    Dim arrRE, i As Long
    arrRE = ArrFlatten(arr)
    ArrValid_Incremental = True
    For i = LBound(arrRE) To UBound(arrRE) - 1
        If arrRE(i + 1) <= arrRE(i) Then '���űȽ�
            ArrValid_Incremental = False
            Exit For
        End If
    Next
End Function
 
'��������Ч���Ƿ�������а������
Public Function ArrValid_IncrementalEqual(ParamArray arr()) As Boolean
    Dim arrRE, i As Long
    arrRE = ArrFlatten(arr)
    ArrValid_IncrementalEqual = True
    For i = LBound(arrRE) To UBound(arrRE) - 1
        If arrRE(i + 1) < arrRE(i) Then '���űȽ�
            ArrValid_IncrementalEqual = False
            Exit For
        End If
    Next
End Function
 
'��������Ч���Ƿ�ݼ�����
Public Function ArrValid_Decrement(ParamArray arr()) As Boolean
    Dim arrRE, i As Long
    arrRE = ArrFlatten(arr)
    ArrValid_Decrement = True
    For i = LBound(arrRE) To UBound(arrRE) - 1
        If arrRE(i + 1) >= arrRE(i) Then '���űȽ�
            ArrValid_Decrement = False
            Exit For
        End If
    Next
End Function
 
'��������Ч���Ƿ�ݼ����а������
Public Function ArrValid_DecrementEqual(ParamArray arr()) As Boolean
    Dim arrRE, i As Long
    arrRE = ArrFlatten(arr)
    ArrValid_DecrementEqual = True
    For i = LBound(arrRE) To UBound(arrRE) - 1
        If arrRE(i + 1) > arrRE(i) Then '���űȽ�
            ArrValid_DecrementEqual = False
            Exit For
        End If
    Next
End Function

'ɸѡ �ظ�����  ,*����ɸѡ����*
Public Function ArrFilterRepeatCount(arr, Optional CountSmall = 0, Optional CountLarge = 1.79769313486231E+308, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim i As Long, dic As Object, c As Long
    Set dic = DictionaryCreate_Count(arr, CompareMode)
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        c = dic(arr(i))
        If c >= CountSmall And c <= CountLarge Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterRepeatCount = ArrayDynamic_
End Function
 
'ɸѡ ���� �ڲ� ,*����ɸѡ����*
Public Function ArrFilterRangeInside(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        If NumberRangeInside(arr(i), NumberL, NumberR, NumberRangeRule) Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterRangeInside = ArrayDynamic_
End Function
 
'ɸѡ ���� �ⲿ ,*����ɸѡ����*
Public Function ArrFilterRangeExternal(arr, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        If NumberRangeExternal(arr(i), NumberL, NumberR, NumberRangeRule) Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterRangeExternal = ArrayDynamic_
End Function
 
'ɸѡ ����V_Small��ֵ ,*����ɸѡ����*
Public Function ArrFilterGreater(arr_Large, V_Small) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr_Large) To UBound(arr_Large)
        If arr_Large(i) > V_Small Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterGreater = ArrayDynamic_
End Function
 
'ɸѡ ���ڵ���V_Small��ֵ ,*����ɸѡ����*
Public Function ArrFilterGreaterEqual(arr_Large, V_Small) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr_Large) To UBound(arr_Large)
        If arr_Large(i) >= V_Small Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterGreaterEqual = ArrayDynamic_
End Function
 
'ɸѡ С��V_Large��ֵ ,*����ɸѡ����*
Public Function ArrFilterLess(arr_Small, V_Large) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr_Small) To UBound(arr_Small)
        If arr_Small(i) < V_Large Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterLess = ArrayDynamic_
End Function
 
'ɸѡ С��V_Large��ֵ ,*����ɸѡ����*
Public Function ArrFilterLessEqual(arr_Small, V_Large) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    For i = LBound(arr_Small) To UBound(arr_Small)
        If arr_Small(i) <= V_Large Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterLessEqual = ArrayDynamic_
End Function
 
'ɸѡ ,*����ɸѡ����*
Public Function ArrFilter(arr, ByVal arrV) As Variant
    Dim i As Long, l As Long, j As Long
    arrV = ArrFlatten(arrV)
    l = UBound(arrV)
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        For j = 1 To l
            If arr(i) = arrV(j) Then
                ArrayDynamic_ i
                Exit For
            End If
        Next
    Next
    ArrFilter = ArrayDynamic_
End Function
 
'ɸѡlikeƥ�� ,*����ɸѡ����*
Public Function ArrFilterLike(arr, ByVal arrV) As Variant
    Dim i As Long, l As Long, j As Long
    arrV = ArrFlatten(arrV)
    l = UBound(arrV)
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        For j = 1 To l
            If arr(i) Like arrV(j) Then
                ArrayDynamic_ i
                Exit For
            End If
        Next
    Next
    ArrFilterLike = ArrayDynamic_
End Function
 
'ɸѡ����ƥ�� ,*����ɸѡ����*
Public Function ArrFilterReg(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        For i = LBound(arr) To UBound(arr)
            If .test(arr(i)) Then
                ArrayDynamic_ i
            End If
        Next
    End With
    ArrFilterReg = ArrayDynamic_
End Function
 
'ɸѡ�ų� ,*����ɸѡ����*
Public Function ArrFilterRemove(arr, ByVal arrV) As Variant
    Dim i As Long, l As Long, j As Long
    arrV = ArrFlatten(arrV)
    l = UBound(arrV)
    ArrayDynamic_
    Dim p As Boolean
    For i = LBound(arr) To UBound(arr)
        p = True
        For j = 1 To l
            If arr(i) = arrV(j) Then
                p = False
                Exit For
            End If
        Next
        If p Then ArrayDynamic_ i
    Next
    ArrFilterRemove = ArrayDynamic_
End Function
 
'ɸѡlike�ų� ,*����ɸѡ����*
Public Function ArrFilterLikeRemove(arr, ByVal arrV) As Variant
    Dim i As Long, l As Long, j As Long
    arrV = ArrFlatten(arrV)
    l = UBound(arrV)
    ArrayDynamic_
    Dim p As Boolean
    For i = LBound(arr) To UBound(arr)
        p = True
        For j = 1 To l
            If arr(i) Like arrV(j) Then
                p = False
                Exit For
            End If
        Next
        If p Then ArrayDynamic_ i
    Next
    ArrFilterLikeRemove = ArrayDynamic_
End Function
 
'ɸѡ�����ų� ,*����ɸѡ����*
Public Function ArrFilterRegRemove(arr, Pattern, Optional ByVal ignoreCase As Boolean = False) As Variant
    Dim i As Long, l As Long, j As Long
    ArrayDynamic_
    With CreateObject("VBScript.RegExp")
        .Global = False
        .ignoreCase = ignoreCase
        .multiline = False
        .Pattern = Pattern
        For i = LBound(arr) To UBound(arr)
            If Not .test(arr(i)) Then
                ArrayDynamic_ i
            End If
        Next
    End With
    ArrFilterRegRemove = ArrayDynamic_
End Function
 
'ȥ�� ������ͷһ��ֵ
Public Function ArrDistinct(arr, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim d As Object, i As Long
    Set d = CreateObject("scripting.Dictionary")
    d.CompareMode = CompareMode
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        If Not d.Exists(arr(i)) Then
            d.Add arr(i), i
            ArrayDynamic_ arr(i)
        End If
    Next
    ArrDistinct = ArrayDynamic_
End Function
 
'ȥ�أ������������� ������ͷ����
Public Function ArrDistinctIndex(arr, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim d As Object, i As Long
    Set d = CreateObject("scripting.Dictionary")
    d.CompareMode = CompareMode
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        If Not d.Exists(arr(i)) Then
            d.Add arr(i), i
            ArrayDynamic_ i
        End If
    Next
    ArrDistinctIndex = ArrayDynamic_
End Function
 
'ȥ�أ������������� �����������
Public Function ArrDistinctIndexRev(arr, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim d As Object, i As Long
    Set d = CreateObject("scripting.Dictionary")
    d.CompareMode = CompareMode
    For i = LBound(arr) To UBound(arr)
        d(arr(i)) = i
    Next
    ArrDistinctIndexRev = ArrLBoundToN_1D(d.Items)
End Function

'�����±��StartLBound һά����
Public Function ArrLBoundToN_1D(arr, Optional StartLBound = 1) As Variant
    Dim arrRE(), i As Long, l1 As Long, u1 As Long, l1RE As Long
    If LBound(arr) <> StartLBound Then
        l1 = LBound(arr): u1 = UBound(arr)
        ReDim arrRE(StartLBound To StartLBound + u1 - l1)
        l1RE = StartLBound
        For i = l1 To u1
            Cover arrRE(l1RE), arr(i)
            l1RE = l1RE + 1
        Next
        ArrLBoundToN_1D = arrRE
    Else
        ArrLBoundToN_1D = arr
    End If
End Function
 
'�����±��StartLBound1,StartLBound2 ��ά����
Public Function ArrLBoundToN_2D(arr, Optional StartLBound1 = 1, Optional StartLBound2 = 1) As Variant
    Dim l1 As Long, u1 As Long, l2 As Long, u2 As Long
    Dim l1RE As Long, l2RE As Long
    Dim arrRE(), i As Long, j As Long
    If LBound(arr, 1) <> StartLBound1 Or LBound(arr, 2) <> StartLBound2 Then
        l1 = LBound(arr, 1): u1 = UBound(arr, 1)
        l2 = LBound(arr, 2): u2 = UBound(arr, 2)
        ReDim arrRE(StartLBound1 To StartLBound1 + u1 - l1, StartLBound2 To StartLBound2 + u2 - l2)
        l1RE = StartLBound1
        For i = l1 To u1
            l2RE = StartLBound2
            For j = l2 To u2
                Cover arrRE(l1RE, l2RE), arr(i, j)
                l2RE = l2RE + 1
            Next
            l1RE = l1RE + 1
        Next
        ArrLBoundToN_2D = arrRE
    Else
        ArrLBoundToN_2D = arr
    End If
End Function

'Evaluate�޸�����
Public Function ArrMap(ByRef arr, EvaluateStr) As Variant
    Dim i As Long, v
    For i = LBound(arr) To UBound(arr)
        Cover arr(i), Application.evaluate(VBA.Replace(EvaluateStr, "$", arr(i)))
    Next
    ArrMap = arr
End Function
 
'�����滻������������Ԫ�� FindValueArr֧�ֵ�ֵ������
Public Function ArrReplace(ByRef arr, FindValueArr, ReplaceValue) As Variant
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    Dim v
    If IsArray(FindValueArr) Then
        If ArrDimension(arr) = 1 Then
            For i = LBound(arr) To UBound(arr)
                For Each v In FindValueArr
                    If arr(i) Like v Then
                        Cover arr(i), ReplaceValue
                        Exit For
                    End If
                Next
            Next
        Else
            l = LBound(arr, 2): u = UBound(arr, 2)
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = l To u
                    For Each v In FindValueArr
                        If arr(i, j) Like v Then
                            Cover arr(i, j), ReplaceValue
                            Exit For
                        End If
                    Next
                Next
            Next
        End If
    Else
        If ArrDimension(arr) = 1 Then
            For i = LBound(arr) To UBound(arr)
                If arr(i) Like FindValueArr Then
                    Cover arr(i), ReplaceValue
                End If
            Next
        Else
            l = LBound(arr, 2): u = UBound(arr, 2)
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = l To u
                    If arr(i, j) Like FindValueArr Then
                        Cover arr(i, j), ReplaceValue
                    End If
                Next
            Next
        End If
    End If
    ArrReplace = arr
End Function
 
'����������ֵ
Public Function ArrErrorClear(ByRef arr, Optional EmptyValue = Empty) As Variant
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            If IsError(arr(i)) Then
                arr(i) = EmptyValue
            End If
        Next
    Else
        l = LBound(arr, 2): u = UBound(arr, 2)
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = l To u
                If IsError(arr(i, j)) Then
                    arr(i, j) = EmptyValue
                End If
            Next
        Next
    End If
    ArrErrorClear = arr
End Function
 
'Evaluateɸѡ ,*����ɸѡ����*
Public Function ArrFilterEval(ByRef arr, EvaluateStr) As Variant
    Dim i As Long, v
    ArrayDynamic_
    For i = LBound(arr) To UBound(arr)
        If Application.evaluate(VBA.Replace(EvaluateStr, "$", arr(i))) Then
            ArrayDynamic_ i
        End If
    Next
    ArrFilterEval = ArrayDynamic_
End Function
 
'�����Ƿ���Ч
Public Function ArrIsValid(ByRef arr) As Boolean
    On Error Resume Next
    Dim u As Long
    Err.Clear
    u = UBound(arr)
    If Err Then
        Err.Clear
        ArrIsValid = False
    Else
        ArrIsValid = u >= LBound(arr)
    End If
End Function
 
'����ά��
Public Function ArrDimension(ByRef arr) As Long
    On Error Resume Next
    Dim s As Long, i As Long
    Err.Clear
    For i = 1 To 9
        s = UBound(arr, i)
        If Err Then ArrDimension = i - 1: Err.Clear: Exit For
    Next
End Function

'����Ԫ�ظ���
Public Function ArrCount(ByRef arr) As Long
    Select Case ArrDimension(arr)
        Case 1
            ArrCount = UBound(arr) - LBound(arr) + 1
        Case 2
            ArrCount = (UBound(arr, 1) - LBound(arr, 1) + 1) * (UBound(arr, 2) - LBound(arr, 2) + 1)
        Case 0
            Select Case TypeName(arr)
                Case "Collection", "Dictionary"
                    ArrCount = arr.Count
                Case Else
                    Dim n As Long, v
                    n = 0
                    For Each v In arr
                        n = n + 1
                    Next
                    ArrCount = n
            End Select
    End Select
End Function

'��������
Public Function ArrCountRow(ByRef arr) As Long
    If ArrDimension(arr) = 1 Then
        ArrCountRow = UBound(arr) - LBound(arr) + 1
    Else
        ArrCountRow = UBound(arr, 1) - LBound(arr, 1) + 1
    End If
End Function
 
'��������
Public Function ArrCountColumn(ByRef arr) As Long
    If ArrDimension(arr) = 1 Then
        ArrCountColumn = UBound(arr) - LBound(arr) + 1
    Else
        ArrCountColumn = UBound(arr, 2) - LBound(arr, 2) + 1
    End If
End Function

'ͬʱ�������������ñ���RowCount,ColumnCount���շ���ֵ��һά����ColumnCount=1����������RowCount=ColumnCount=1
Public Sub ArrCountRowAndColumn(arr, ByRef RowCount, ByRef ColumnCount)
    Select Case ArrDimension(arr)
        Case 1
            RowCount = UBound(arr, 1) - LBound(arr, 1) + 1
            ColumnCount = 1
        Case 2
            RowCount = UBound(arr, 1) - LBound(arr, 1) + 1
            ColumnCount = UBound(arr, 2) - LBound(arr, 2) + 1
        Case 0
            RowCount = 1
            ColumnCount = 1
    End Select
End Sub

'������Ԫ�ظ�����������������
Public Function ArrCountElement(ByRef arr, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim i As Long, dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    For i = LBound(arr) To UBound(arr)
        If dic.Exists(arr(i)) Then
            dic(arr(i)) = dic(arr(i)) + 1
        Else
            dic.Add arr(i), 1
        End If
    Next
    For i = LBound(arr) To UBound(arr)
        arr(i) = dic(arr(i))
    Next
    ArrCountElement = arr
End Function

'�����Ǻϲ���Ԫ����ʽԪ�ظ��������ظ�������
Public Function ArrCountMergeElement(ByRef arr, Optional EmptyContent = "") As Variant
    Dim i As Long, j As Long, c As Long, stari As Long
    c = 0: stari = LBound(arr)
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> EmptyContent Then
            For j = stari To i - 1
                arr(j) = c
            Next
            stari = i
            c = 0
        End If
        c = c + 1
    Next
    For j = stari To i - 1
        arr(j) = c
    Next
    ArrCountMergeElement = arr
End Function

'������Χ��������
Public Function ArrBetween(l, u) As Variant()
    Dim arr()
    ReDim arr(l To u)
    Dim i As Long
    For i = l To u
        arr(i) = i
    Next
    ArrBetween = arr
End Function
 
'��������
Public Function ArrCreate(RowCount, Optional ColumnCount = 0, Optional FillValue = Empty) As Variant()
    Dim arr(), i As Long, j As Long
    If ColumnCount = 0 Then
        ReDim arr(1 To RowCount)
        If Not IsEmpty(FillValue) Then
            For i = 1 To RowCount
                Cover arr(i), FillValue
            Next
        End If
    Else
        ReDim arr(1 To RowCount, 1 To ColumnCount)
        If Not IsEmpty(FillValue) Then
            For i = 1 To RowCount
                For j = 1 To ColumnCount
                    Cover arr(i, j), FillValue
                Next
            Next
        End If
    End If
    ArrCreate = arr
End Function

'�������������
Public Function ArrCreateRand(l, r, RowCount, Optional ColumnCount = 0) As Variant()
    Dim arr(), i As Long, j As Long
    Randomize
    If ColumnCount = 0 Then
        ReDim arr(1 To RowCount)
        For i = 1 To RowCount
            arr(i) = IntDown((r - l + 1) * Rnd()) + l
        Next
    Else
        ReDim arr(1 To RowCount, 1 To ColumnCount)
        For i = 1 To RowCount
            For j = 1 To ColumnCount
                arr(i, j) = IntDown((r - l + 1) * Rnd()) + l
            Next
        Next
    End If
    ArrCreateRand = arr
End Function

'������������� ���ظ������
Public Function ArrCreateRandDic(l, r, RowCount, Optional ColumnCount = 0) As Variant()
    Dim arr(), i As Long, j As Long, n As Long
    Randomize
    Dim col As New Collection
    For i = l To r
        col.Add i
    Next
    If ColumnCount = 0 Then
        ReDim arr(1 To RowCount)
        For i = 1 To RowCount
            If col.Count > 0 Then
                n = RandBetween(1, col.Count)
                arr(i) = col(n)
                col.Remove n
            End If
        Next
        arr = ArrGetRegion(arr, l, RowCount)
    Else
        ReDim arr(1 To RowCount, 1 To ColumnCount)
        For i = 1 To RowCount
            For j = 1 To ColumnCount
                If col.Count > 0 Then
                    n = RandBetween(1, col.Count)
                    arr(i, j) = col(n)
                    col.Remove n
                End If
            Next
        Next
    End If
    ArrCreateRandDic = arr
End Function



'��ֵ�������  arrһά���ά���� index��ά����������  EmptyContent������ֵ������
Public Function ArrFillDown(ByRef arr, Optional ByVal Index = 1, Optional EmptyContent = "") As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) + 1 To UBound(arr)
            If arr(i) = EmptyContent Then Cover arr(i), arr(i - 1)
        Next
    Else
        IndexIsCurrencyToCount_ Index, LBound(arr, 2), UBound(arr, 2)
        For i = LBound(arr, 1) + 1 To UBound(arr, 1)
            If arr(i, Index) = EmptyContent Then Cover arr(i, Index), arr(i - 1, Index)
        Next
    End If
    ArrFillDown = arr
End Function
 
'��ֵ�������  arrһά���ά���� index��ά����������  EmptyContent������ֵ������
Public Function ArrFillUp(ByRef arr, Optional ByVal Index = 1, Optional EmptyContent = "") As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = UBound(arr) - 1 To LBound(arr) Step -1
            If arr(i) = EmptyContent Then Cover arr(i), arr(i + 1)
        Next
    Else
        IndexIsCurrencyToCount_ Index, LBound(arr, 2), UBound(arr, 2)
        For i = UBound(arr, 1) - 1 To LBound(arr, 1) Step -1
            If arr(i, Index) = EmptyContent Then Cover arr(i, Index), arr(i + 1, Index)
        Next
    End If
    ArrFillUp = arr
End Function
 
'��͸�� arrH������(�����Ƕ���)  arrV�����(ֻ��һ��) arrRegion2D��������(�д�С������arrH���� �д�С������arrV����)
Public Function ArrPerspectiveRev(ByRef arrH, ByRef arrV, Optional ByRef arrRegion2D = "") As Variant
    Dim arrRE(), i As Long, j As Long, k As Long, n As Long
    Dim arrHRE, arrVRE
    Dim LH As Long, LV As Long
    If ArrDimension(arrH) = 1 Then
        arrHRE = ArrTranspose(arrH)
    Else
        arrHRE = arrH
    End If
    arrVRE = ArrFlatten_Single(arrV)
    LH = UBound(arrHRE, 2)
    LV = UBound(arrVRE)
    n = 1
    If IsArray(arrRegion2D) Then
        ReDim arrRE(1 To UBound(arrHRE, 1) * UBound(arrVRE), 1 To UBound(arrHRE, 2) + 2)
        For i = 1 To UBound(arrHRE, 1)
            For j = 1 To LV
                For k = 1 To LH
                    arrRE(n, k) = arrHRE(i, k)
                Next
                arrRE(n, LH + 1) = arrVRE(j)
                arrRE(n, LH + 2) = arrRegion2D(i, j)
                n = n + 1
            Next
        Next
    Else
        ReDim arrRE(1 To UBound(arrHRE, 1) * UBound(arrVRE), 1 To UBound(arrHRE, 2) + 1)
        n = 1
        For i = 1 To UBound(arrHRE, 1)
            For j = 1 To LV
                For k = 1 To LH
                    arrRE(n, k) = arrHRE(i, k)
                Next
                arrRE(n, LH + 1) = arrVRE(j)
                n = n + 1
            Next
        Next
    End If
    ArrPerspectiveRev = arrRE
End Function
 
'͸�� ���н������ظ�����ʱȡ���һֵ arr2D��ά��  VIndex���������  DataIndex�������������
Public Function ArrPerspective(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant
    Dim arrRE(), i As Long, j As Long, l As Long, s As String, arrS, n As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ VIndex, LV, UV
    IndexIsCurrencyToCount_ DataIndex, LV, UV
    
    Dim arrindex, dicV As Object, dicH As Object
    Set dicV = CreateObject("scripting.Dictionary")
    Set dicH = CreateObject("scripting.Dictionary")
    ArrayDynamic_
    For i = LV To UV
        If i <> VIndex And i <> DataIndex Then
            ArrayDynamic_ i
        End If
    Next
    arrindex = ArrayDynamic_ 'ȥ���������VIndex��������DataIndex��������
    l = UBound(arrindex)
    ReDim arrS(LH To UH, 1 To 2) '1 arrRE������,2 arrRE������
    For i = LH To UH
        '���arrRE������
        s = ""
        For j = 1 To l 'arrindex
            s = s & "@" & arr2D(i, arrindex(j))
        Next
        If Not dicH.Exists(s) Then
            dicH.Add s, dicH.Count + 2
        End If
        arrS(i, 1) = dicH(s)
        '���arrRE������
        If Not dicV.Exists(arr2D(i, VIndex)) Then
            dicV.Add arr2D(i, VIndex), dicV.Count + 1 + l
        End If
        arrS(i, 2) = dicV(arr2D(i, VIndex))
    Next
    'д����
    ReDim arrRE(1 To dicH.Count + 1, 1 To l + dicV.Count)
    For i = LH To UH
        n = arrS(i, 1)
        For j = 1 To l 'arrindex
            arrRE(n, j) = arr2D(i, arrindex(j))
        Next
        arrRE(n, arrS(i, 2)) = arr2D(i, DataIndex)
    Next
    '�ӱ���
    arrS = dicV.Keys
    For i = 0 To UBound(arrS)
        arrRE(1, dicV(arrS(i))) = arrS(i)
    Next
    ArrPerspective = arrRE
End Function
 
'͸�� ���н������ظ�����ʱд���� arr2D��ά��  VIndex���������  DataIndex�������������
Public Function ArrPerspective_Repeating(ByRef arr2D, ByVal VIndex, ByVal DataIndex) As Variant
    Dim arrRE(), i As Long, j As Long, l As Long, s, arrS, n As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
        
    IndexIsCurrencyToCount_ VIndex, LV, UV
    IndexIsCurrencyToCount_ DataIndex, LV, UV
    
    Dim arrindex, dicV As Object, dicH As Object, dicHindex As Object, arrjoins, dicCount As Object, dicCountSub As Object
    ReDim arrjoins(LH To UH)
    ArrayDynamic_
    For i = LV To UV
        If i <> VIndex And i <> DataIndex Then
            ArrayDynamic_ i
        End If
    Next
    arrindex = ArrayDynamic_ 'ȥ���������VIndex��������DataIndex��������
    l = UBound(arrindex)
    '�ظ�����
    Set dicCount = CreateObject("scripting.Dictionary")
    For i = LH To UH
        s = ""
        For j = 1 To l 'arrindex
            s = s & "@" & arr2D(i, arrindex(j))
        Next
        arrjoins(i) = s
        s = s & "@" & arr2D(i, VIndex)
        If dicCount.Exists(arrjoins(i)) Then
            Set dicCountSub = dicCount(arrjoins(i))
            dicCountSub(s) = dicCountSub(s) + 1
        Else
            Set dicCountSub = CreateObject("scripting.Dictionary")
            dicCountSub(s) = 1
            dicCount.Add arrjoins(i), dicCountSub
        End If
    Next
    For Each s In dicCount
        dicCount(s) = ArrMax(dicCount(s).Items)
    Next
    Set dicV = CreateObject("scripting.Dictionary")
    Set dicHindex = CreateObject("scripting.Dictionary")
    Set dicH = CreateObject("scripting.Dictionary")
    ReDim arrS(LH To UH, 1 To 2) As Long '1 arrRE������,2 arrRE������
    Dim rsum As Long
    rsum = 2
    For i = LH To UH
        '���arrRE������
        If Not dicV.Exists(arr2D(i, VIndex)) Then
            dicV.Add arr2D(i, VIndex), dicV.Count + 1 + l
        End If
        arrS(i, 2) = dicV(arr2D(i, VIndex))
        '���arrRE������
        s = arrjoins(i)
        If Not dicH.Exists(s) Then 'ÿ�����������
            dicH.Add s, rsum
            rsum = rsum + dicCount(s)
        End If
        arrS(i, 1) = dicH(s)
        s = s & "@" & arr2D(i, VIndex)
        If Not dicHindex.Exists(s) Then '�ظ�����ĵ�ǰ��
            dicHindex.Add s, 1
        Else
            dicHindex(s) = dicHindex(s) + 1
        End If
        arrS(i, 1) = arrS(i, 1) + dicHindex(s) - 1
    Next
    'д����
    ReDim arrRE(1 To rsum, 1 To l + dicV.Count)
    For i = LH To UH
        n = arrS(i, 1)
        For j = 1 To l 'arrindex
            arrRE(n, j) = arr2D(i, arrindex(j))
        Next
        arrRE(n, arrS(i, 2)) = arr2D(i, DataIndex)
    Next
    '�ӱ���
    arrS = dicV.Keys
    For i = 0 To UBound(arrS)
        arrRE(1, dicV(arrS(i))) = arrS(i)
    Next
    ArrPerspective_Repeating = arrRE
End Function
 
'������� arr2D��ά�� ArrGroupIndex����������֧������ ArrSumIndex���������֧������
Public Function ArrGroupSum(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrSumIndex, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim arrRE(), i As Long, j As Long, s As String, arrS, l As Long, ls As Long, p As Boolean, parr() As Boolean, n As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    ArrGroupIndex = ArrFlatten(ArrGroupIndex)
    ArrSumIndex = ArrFlatten(ArrSumIndex)
    
    IndexIsCurrencyToCount_ ArrGroupIndex, LV, UV
    IndexIsCurrencyToCount_ ArrSumIndex, LV, UV
    
    l = UBound(ArrGroupIndex)
    ReDim arrS(LH To UH)
    For i = LH To UH
        '���arrRE������
        s = ""
        For j = 1 To l 'GroupIndex
            s = s & "@" & arr2D(i, ArrGroupIndex(j))
        Next
        If Not dic.Exists(s) Then
            dic.Add s, dic.Count + 1
        End If
        arrS(i) = dic(s)
    Next
    ls = UBound(ArrSumIndex)
    ArrayDynamic_
    For i = LV To UV
        p = True
        For j = 1 To ls
            If i = ArrSumIndex(j) Then
                p = False
                Exit For
            End If
        Next
        If p Then ArrayDynamic_ i
    Next
    Dim arrindex
    arrindex = ArrayDynamic_ 'ȥSumIndex��������
    l = UBound(arrindex)
    'д����
    ReDim arrRE(1 To dic.Count, LV To UV)
    ReDim parr(1 To UBound(arrRE, 1)) As Boolean
    For i = LH To UH
        n = arrS(i)
        If parr(n) = False Then
            For j = 1 To l 'arrindex
                arrRE(n, arrindex(j)) = arr2D(i, arrindex(j))
            Next
            parr(n) = True
        End If
        For j = 1 To ls 'ArrSumIndex
            arrRE(n, ArrSumIndex(j)) = arrRE(n, ArrSumIndex(j)) + arr2D(i, ArrSumIndex(j))
        Next
    Next
    ArrGroupSum = arrRE
End Function
 
'������� arr2D��ά�� ArrGroupIndex����������֧������ ArrCountIndex����������֧������ NoEmpty = True����ǿ�ֵ����
Public Function ArrGroupCount(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrCountIndex, Optional NoEmpty As Boolean = True, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim arrRE(), i As Long, j As Long, s As String, arrS, l As Long, ls As Long, p As Boolean, parr() As Boolean, n As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    ArrGroupIndex = ArrFlatten(ArrGroupIndex)
    ArrCountIndex = ArrFlatten(ArrCountIndex)
    
    IndexIsCurrencyToCount_ ArrGroupIndex, LV, UV
    IndexIsCurrencyToCount_ ArrCountIndex, LV, UV
    
    l = UBound(ArrGroupIndex)
    ReDim arrS(LH To UH)
    For i = LH To UH
        '���arrRE������
        s = ""
        For j = 1 To l 'GroupIndex
            s = s & "@" & arr2D(i, ArrGroupIndex(j))
        Next
        If Not dic.Exists(s) Then
            dic.Add s, dic.Count + 1
        End If
        arrS(i) = dic(s)
    Next
    ls = UBound(ArrCountIndex)
    ArrayDynamic_
    For i = LV To UV
        p = True
        For j = 1 To ls
            If i = ArrCountIndex(j) Then
                p = False
                Exit For
            End If
        Next
        If p Then ArrayDynamic_ i
    Next
    Dim arrindex
    arrindex = ArrayDynamic_ 'ȥSumIndex��������
    l = UBound(arrindex)
    'д����
    ReDim arrRE(1 To dic.Count, LV To UV)
    ReDim parr(1 To UBound(arrRE, 1)) As Boolean
    For i = LH To UH
        n = arrS(i)
        If parr(n) = False Then
            For j = 1 To l 'arrindex
                arrRE(n, arrindex(j)) = arr2D(i, arrindex(j))
            Next
            For j = 1 To ls
                arrRE(n, ArrCountIndex(j)) = 0
            Next
            parr(n) = True
        End If
        For j = 1 To ls 'ArrCountIndex
            If NoEmpty = False Or arr2D(i, ArrCountIndex(j)) <> "" Then arrRE(n, ArrCountIndex(j)) = arrRE(n, ArrCountIndex(j)) + 1
        Next
    Next
    ArrGroupCount = arrRE
End Function
 
'����ƴ���ַ��� arr2D��ά�� ArrGroupIndex����������֧������ ArrJoinIndex���������֧������ Delimiter�ָ��� OmittedEmpty���Կ��ַ���
Public Function ArrGroupJoin(ByRef arr2D, ByVal ArrGroupIndex, ByVal ArrJoinIndex, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim arrRE(), i As Long, j As Long, s As String, arrS, l As Long, ls As Long, p As Boolean, parr() As Boolean, n As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    ArrGroupIndex = ArrFlatten(ArrGroupIndex)
    ArrJoinIndex = ArrFlatten(ArrJoinIndex)
    
    IndexIsCurrencyToCount_ ArrGroupIndex, LV, UV
    IndexIsCurrencyToCount_ ArrJoinIndex, LV, UV
    
    l = UBound(ArrGroupIndex)
    ReDim arrS(LH To UH)
    For i = LH To UH
        '���arrRE������
        s = ""
        For j = 1 To l 'GroupIndex
            s = s & "@" & arr2D(i, ArrGroupIndex(j))
        Next
        If Not dic.Exists(s) Then
            dic.Add s, dic.Count + 1
        End If
        arrS(i) = dic(s)
    Next
    ls = UBound(ArrJoinIndex)
    ArrayDynamic_
    For i = LV To UV
        p = True
        For j = 1 To ls
            If i = ArrJoinIndex(j) Then
                p = False
                Exit For
            End If
        Next
        If p Then ArrayDynamic_ i
    Next
    Dim arrindex
    arrindex = ArrayDynamic_ 'ȥSumIndex��������
    l = UBound(arrindex)
    'д����
    ReDim arrRE(1 To dic.Count, LV To UV)
    ReDim parr(1 To UBound(arrRE, 1)) As Boolean
    For i = LH To UH
        n = arrS(i)
        If parr(n) = False Then '�����n��ĵ�һ��д��
            For j = 1 To l 'arrindex  д�����
                arrRE(n, arrindex(j)) = arr2D(i, arrindex(j))
            Next
            parr(n) = True
            For j = 1 To ls 'ArrJoinIndex д���ַ���
                arrRE(n, ArrJoinIndex(j)) = arr2D(i, ArrJoinIndex(j))
            Next
        Else
            For j = 1 To ls 'ArrJoinIndex
                If OmittedEmpty = False Then
                    arrRE(n, ArrJoinIndex(j)) = arrRE(n, ArrJoinIndex(j)) & Delimiter & arr2D(i, ArrJoinIndex(j))
                Else
                    If arr2D(i, ArrJoinIndex(j)) <> "" Then
                        If arrRE(n, ArrJoinIndex(j)) = "" Then
                            arrRE(n, ArrJoinIndex(j)) = arr2D(i, ArrJoinIndex(j))
                        Else
                            arrRE(n, ArrJoinIndex(j)) = arrRE(n, ArrJoinIndex(j)) & Delimiter & arr2D(i, ArrJoinIndex(j))
                        End If
                    End If
                End If
            Next
        End If
    Next
    ArrGroupJoin = arrRE
End Function

'����ۺϺ���  ArrGroup���麯�����ص�����������  OmittedNoneArgû��дCn���������Ƿ�ʡ�� Delimiterƴ���ַ��ָ��� OmittedEmptyƴ���ַ����Ƿ���Կ�ֵ
Public Function ArrGroupAgg(ByRef ArrGroup, Optional OmittedNoneArg As Boolean = True, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, _
        Optional ByRef C1 As GroupAggregateMethod = Group_None, Optional ByRef C2 As GroupAggregateMethod = Group_None, _
        Optional ByRef C3 As GroupAggregateMethod = Group_None, Optional ByRef C4 As GroupAggregateMethod = Group_None, _
        Optional ByRef C5 As GroupAggregateMethod = Group_None, Optional ByRef C6 As GroupAggregateMethod = Group_None, _
        Optional ByRef C7 As GroupAggregateMethod = Group_None, Optional ByRef C8 As GroupAggregateMethod = Group_None, _
        Optional ByRef C9 As GroupAggregateMethod = Group_None, Optional ByRef C10 As GroupAggregateMethod = Group_None, _
        Optional ByRef C11 As GroupAggregateMethod = Group_None, Optional ByRef C12 As GroupAggregateMethod = Group_None, _
        Optional ByRef C13 As GroupAggregateMethod = Group_None, Optional ByRef C14 As GroupAggregateMethod = Group_None, _
        Optional ByRef C15 As GroupAggregateMethod = Group_None, Optional ByRef C16 As GroupAggregateMethod = Group_None, _
        Optional ByRef C17 As GroupAggregateMethod = Group_None, Optional ByRef C18 As GroupAggregateMethod = Group_None, _
        Optional ByRef C19 As GroupAggregateMethod = Group_None, Optional ByRef C20 As GroupAggregateMethod = Group_None, _
        Optional ByRef C21 As GroupAggregateMethod = Group_None, Optional ByRef C22 As GroupAggregateMethod = Group_None, _
        Optional ByRef C23 As GroupAggregateMethod = Group_None, Optional ByRef C24 As GroupAggregateMethod = Group_None, _
        Optional ByRef C25 As GroupAggregateMethod = Group_None, Optional ByRef C26 As GroupAggregateMethod = Group_None, _
        Optional ByRef C27 As GroupAggregateMethod = Group_None, Optional ByRef C28 As GroupAggregateMethod = Group_None, _
        Optional ByRef C29 As GroupAggregateMethod = Group_None, Optional ByRef C30 As GroupAggregateMethod = Group_None, _
        Optional ByRef C31 As GroupAggregateMethod = Group_None, Optional ByRef C32 As GroupAggregateMethod = Group_None, _
        Optional ByRef C33 As GroupAggregateMethod = Group_None, Optional ByRef C34 As GroupAggregateMethod = Group_None, _
        Optional ByRef C35 As GroupAggregateMethod = Group_None, Optional ByRef C36 As GroupAggregateMethod = Group_None, _
        Optional ByRef C37 As GroupAggregateMethod = Group_None, Optional ByRef C38 As GroupAggregateMethod = Group_None, _
        Optional ByRef C39 As GroupAggregateMethod = Group_None, Optional ByRef C40 As GroupAggregateMethod = Group_None, _
        Optional ByRef C41 As GroupAggregateMethod = Group_None, Optional ByRef C42 As GroupAggregateMethod = Group_None, _
        Optional ByRef C43 As GroupAggregateMethod = Group_None, Optional ByRef C44 As GroupAggregateMethod = Group_None, _
        Optional ByRef C45 As GroupAggregateMethod = Group_None, Optional ByRef C46 As GroupAggregateMethod = Group_None _
        ) As Variant
    Dim arrRE, i As Long, j As Long
    arrRE = Array(C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, C14, C15, C16, C17, C18, C19, C20, _
    C21, C22, C23, C24, C25, C26, C27, C28, C29, C30, C31, C32, C33, C34, C35, C36, C37, C38, C39, C40, C41, C42, C43, C44, C45, C46)
    ArrayDynamic_
    For i = LBound(ArrGroup) To UBound(ArrGroup)
        If UBound(ArrGroup(i), 1) >= LBound(ArrGroup(i), 1) Then
            ArrayDynamic2_
            For j = 0 To MinParams2(UBound(arrRE), UBound(ArrGroup(i), 2) - LBound(ArrGroup(i), 2))
                ArrGroupAgg_Argument_ ArrGroup(i), OmittedNoneArg, Delimiter, OmittedEmpty, arrRE(j), j + 1
            Next
            ArrayDynamic_ ArrayDynamic2_
        End If
    Next
    ArrGroupAgg = ArrF_T(ArrayDynamic_, -1)
End Function

'���оۺ� CI�ۺ�ģʽ ColumnIndex�ۺ���(*���������ǵ�n��*)   �ڲ�ʹ��
Private Function ArrGroupAgg_Argument_(ByRef ArrGroup, OmittedNoneArg As Boolean, Delimiter, OmittedEmpty As Boolean, CI, ColumnIndex)
    If ColumnIndex >= 1 And ColumnIndex <= (UBound(ArrGroup, 2) - LBound(ArrGroup, 2) + 1) Then
        Select Case CI
            Case Group_None
                If OmittedNoneArg = False Then ArrayDynamic2_ ArrGroup(LBound(ArrGroup, 1), LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_First
                ArrayDynamic2_ ArrGroup(LBound(ArrGroup, 1), LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Last
                ArrayDynamic2_ ArrGroup(UBound(ArrGroup, 1), LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Sum
                ArrayDynamic2_ ArrSumColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Count
                ArrayDynamic2_ UBound(ArrGroup, 1) - LBound(ArrGroup, 1) + 1
            Case Group_CountNoEmpty
                ArrayDynamic2_ ArrCountNoEmptyColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_CountClass
                ArrayDynamic2_ ArrCountClassColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Max
                ArrayDynamic2_ ArrMaxColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Min
                ArrayDynamic2_ ArrMinColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1)
            Case Group_Average
                ArrayDynamic2_ ArrAverageColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1, 5)
            Case Group_Join
                ArrayDynamic2_ ArrJoinColumn(ArrGroup, LBound(ArrGroup, 2) + ColumnIndex - 1, Delimiter, OmittedEmpty)
            Case Else
                If CI < 0 Then
                    ArrayDynamic2_ ArrGroup(UBound(ArrGroup, 1) + CI + 1, LBound(ArrGroup, 2) + ColumnIndex - 1)
                Else
                    ArrayDynamic2_ ArrGroup(LBound(ArrGroup, 1) + CI - 1, LBound(ArrGroup, 2) + ColumnIndex - 1)
                End If
        End Select
    Else
        ArrayDynamic2_ Empty
    End If
End Function

'����ۺϺ��� ֧��һ�ж��־ۺ� ArrGroup���麯�����ص����������� ArrGroupIndex�ۺ��� ArrAggregateMethod��Ӧ�ľۺ�ģʽ Delimiterƴ���ַ��ָ��� OmittedEmptyƴ���ַ����Ƿ���Կ�ֵ
Public Function ArrGroupAgg2(ByRef ArrGroup, ArrGroupIndex, ArrAggregateMethod, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
    Dim arrRE, i As Long, j As Long, l As Long, u As Long
    Dim ArrGroupIndexRE, ArrAggregateMethodRE, ArrGroupIndexRE2
    ArrGroupIndexRE = ArrFlatten_Single(ArrGroupIndex)
    u = UBound(ArrGroupIndexRE)
    If IsArray(ArrAggregateMethod) Then
        ArrAggregateMethodRE = ArrSizeExpansion2(ArrAggregateMethod, u, Group_None)
    Else
        ArrAggregateMethodRE = ArrSizeExpansion2(ArrAggregateMethod, u, ArrAggregateMethod)
    End If

    ArrayDynamic_
    For i = LBound(ArrGroup) To UBound(ArrGroup)
        If UBound(ArrGroup(i), 1) >= LBound(ArrGroup(i), 1) Then
            ArrGroupIndexRE2 = ArrGroupIndexRE
            '����ת����n��
            IndexIsLongToCount_ ArrGroupIndexRE2, LBound(ArrGroup(i), 2), UBound(ArrGroup(i), 2)
            ArrayDynamic2_
            For j = LBound(ArrGroupIndexRE) To UBound(ArrGroupIndexRE)
                ArrGroupAgg_Argument_ ArrGroup(i), False, Delimiter, OmittedEmpty, ArrAggregateMethodRE(j), ArrGroupIndexRE2(j)
            Next
            ArrayDynamic_ ArrayDynamic2_
        End If
    Next
    ArrGroupAgg2 = ArrF_T(ArrayDynamic_, -1)
End Function

'������� ����� ArrClassIndex��������֧������  ��������������ķ���
Public Function ArrGroup_Class(ByRef arr2D, ByVal ArrClassIndex, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim arrRE(), i As Long, j As Long, s As String, arrS, l As Long, arrREindex() As Long, n As Long, k As Long, arrtmp()
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim dic As Object, Dics As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    Set Dics = CreateObject("scripting.Dictionary")
    Dics.CompareMode = CompareMode
    ArrClassIndex = ArrFlatten(ArrClassIndex)
    
    IndexIsCurrencyToCount_ ArrClassIndex, LV, UV
    
    l = UBound(ArrClassIndex)
    ReDim arrS(LH To UH)
    For i = LH To UH
        s = ""
        For j = 1 To l 'GroupIndex
            s = s & "@" & arr2D(i, ArrClassIndex(j))
        Next
        'ÿ���������
        If Dics.Exists(s) Then
            Dics(s) = Dics(s) + 1
        Else
            Dics(s) = 1
        End If
 
        'ÿ�ж�Ӧ��������
        If Not dic.Exists(s) Then
            dic.Add s, dic.Count + 1
        End If
        arrS(i) = dic(s)
    Next
    Dim dicSitem
    dicSitem = Dics.Items 'ÿ������
    ReDim arrRE(1 To dic.Count)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        ReDim arrtmp(1 To dicSitem(i - 1), LV To UV)
        arrRE(i) = arrtmp
        arrREindex(i) = 1
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_Class = arrRE
End Function
 
'������� ����������Ϊ������� ���޷��ڷ���*����*  FindIndex������ FindValue��������  ��������������ķ���
Public Function ArrGroup_Find_First(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant
    Dim arrRE(), i As Long, j As Long, arrtmp(), n As Long, k As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ FindIndex, LV, UV
    
    Dim col As New Collection '�ֽ�����
    col.Add LH
    For i = LH + 1 To UH
        If arr2D(i, FindIndex) Like FindValue Then
            col.Add i
        End If
    Next
    col.Add UH + 1
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        k = 1
        For i = col(n) To col(n + 1) - 1
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Find_First = arrRE
End Function
 
'������� ����������Ϊ������� ���޷��ڷ���*ĩβ*  FindIndex������ FindValue��������  ��������������ķ���
Public Function ArrGroup_Find_Last(ByRef arr2D, ByVal FindIndex, ByVal FindValue) As Variant
    Dim arrRE(), i As Long, j As Long, arrtmp(), n As Long, k As Long
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ FindIndex, LV, UV
    
    Dim col As New Collection '�ֽ�����
    col.Add LH - 1
    For i = LH To UH - 1
        If arr2D(i, FindIndex) Like FindValue Then
            col.Add i
        End If
    Next
    col.Add UH
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        k = 1
        For i = col(n) + 1 To col(n + 1)
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Find_Last = arrRE
End Function
 
'������� �����������ݲ���Ϊ�������  ArrDifferIndex��ͬ��������֧������  ��������������ķ���
Public Function ArrGroup_Differ(ByRef arr2D, ByVal ArrDifferIndex) As Variant
    Dim arrRE(), i As Long, j As Long, l As Long, n As Long, k As Long, arrtmp()
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    ArrDifferIndex = ArrFlatten(ArrDifferIndex)
    
    IndexIsCurrencyToCount_ ArrDifferIndex, LV, UV
    
    l = UBound(ArrDifferIndex)
    Dim col As New Collection '�ֽ�����
    col.Add LH
    For i = LH + 1 To UH
        For j = 1 To l
            If arr2D(i, ArrDifferIndex(j)) <> arr2D(i - 1, ArrDifferIndex(j)) Then
                col.Add i
                Exit For
            End If
        Next
    Next
    col.Add UH + 1
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        k = 1
        For i = col(n) To col(n + 1) - 1
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Differ = arrRE
End Function

'������� ��������  number����  ��������������ķ���
Public Function ArrGroup_Number_Column(ByRef arr2D, ByVal Number, Optional FixedSize As Boolean = False) As Variant
    Dim arrRE(), i As Long, j As Long, n As Long, k As Long, arrtmp()
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim col As New Collection '�ֽ�����
    col.Add LV
    k = 1
    For i = LV + 1 To UV
        If (k Mod Number) = 0 Then
            col.Add i
        End If
        k = k + 1
    Next
    col.Add UV + 1
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        If FixedSize Then
            ReDim arrtmp(LH To UH, 1 To Number)
        Else
            ReDim arrtmp(LH To UH, 1 To col(n + 1) - col(n))
        End If
        For i = LH To UH
            k = 1
            For j = col(n) To col(n + 1) - 1
                arrtmp(i, k) = arr2D(i, j)
                k = k + 1
            Next
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Number_Column = arrRE
End Function
 
'������� ������  number����  ��������������ķ���
Public Function ArrGroup_Number(ByRef arr2D, ByVal Number, Optional FixedSize As Boolean = False) As Variant
    Dim arrRE(), i As Long, j As Long, n As Long, k As Long, arrtmp()
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    Dim col As New Collection '�ֽ�����
    col.Add LH
    k = 1
    For i = LH + 1 To UH
        If (k Mod Number) = 0 Then
            col.Add i
        End If
        k = k + 1
    Next
    col.Add UH + 1
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        If FixedSize Then
            ReDim arrtmp(1 To Number, LV To UV)
        Else
            ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        End If
        k = 1
        For i = col(n) To col(n + 1) - 1
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Number = arrRE
End Function
 
'������� ��������Ϊ���޷���  ���޷��ڷ���*����* ArrRowIndex������֧������  ��������������ķ���
Public Function ArrGroup_Row_First(ByRef arr2D, ParamArray ArrRowIndexs()) As Variant
    Dim arrRE(), i As Long, j As Long, n As Long, k As Long, arrtmp(), v
    Dim ArrRowIndex
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    ArrRowIndex = ArrFlatten(ArrRowIndexs)
    
    IndexIsCurrencyToCount_ ArrRowIndex, LH, UH
    
    Dim col As New Collection '�ֽ�����
    col.Add LH
    For Each v In ArrRowIndex
        If v <> LH Then col.Add v
    Next
    col.Add UH + 1
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        k = 1
        For i = col(n) To col(n + 1) - 1
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Row_First = arrRE
End Function
 
'������� ��������Ϊ���޷���  ���޷��ڷ���*ĩβ* ArrRowIndex������֧������  ��������������ķ���
Public Function ArrGroup_Row_Last(ByRef arr2D, ParamArray ArrRowIndexs()) As Variant
    Dim arrRE(), i As Long, j As Long, n As Long, k As Long, arrtmp(), v
    Dim ArrRowIndex
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    ArrRowIndex = ArrFlatten(ArrRowIndexs)
    
    IndexIsCurrencyToCount_ ArrRowIndex, LH, UH
    
    Dim col As New Collection '�ֽ�����
    col.Add LH - 1
    For Each v In ArrRowIndex
        If v <> UH Then col.Add v
    Next
    col.Add UH
    ReDim arrRE(1 To col.Count - 1)
    '����
    For n = 1 To col.Count - 1
        ReDim arrtmp(1 To col(n + 1) - col(n), LV To UV)
        k = 1
        For i = col(n) + 1 To col(n + 1)
            For j = LV To UV
                arrtmp(k, j) = arr2D(i, j)
            Next
            k = k + 1
        Next
        arrRE(n) = arrtmp
    Next
    ArrGroup_Row_Last = arrRE
End Function

'������� ����ֵ����������  С�ڲ����ڱ���һ�� ArrInterval��������  ��������������ķ���
Public Function ArrGroup_Interval(ByVal arr2D, ByVal ColumnIndex, ParamArray arrInterval()) As Variant
    arrInterval = ArrSort1D(ArrDistinct(ArrFlatten(arrInterval)), True)
    Dim arrRE(), arrRECount() As Long, arrREindex() As Long, i As Long, j As Long, n As Long, k As Long, arrtmp(), p As Boolean
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LV, UV
    
    Dim lI As Long, UI As Long
    lI = LBound(arrInterval, 1): UI = UBound(arrInterval, 1)
    ReDim arrS(LH To UH) '���������
    ReDim arrRECount(1 To UI - lI + 2) '������
    For i = LH To UH
        'ѭ���Աȼ���������
        For j = lI To UI
            If arr2D(i, ColumnIndex) < arrInterval(j) Then
                arrS(i) = j
                arrRECount(j) = arrRECount(j) + 1
                GoTo AlreadyWritten_
            End If
        Next
        'ʣ�µĶ��Ǵ�ķŵ����һ��
        arrS(i) = UI + 1
        arrRECount(UI + 1) = arrRECount(UI + 1) + 1
AlreadyWritten_:
    Next
    
    '����
    ReDim arrRE(1 To UI - lI + 2)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        If arrRECount(i) > 0 Then
            ReDim arrtmp(1 To arrRECount(i), LV To UV)
            arrRE(i) = arrtmp
            arrREindex(i) = 1
        Else
            arrRE(i) = Array()
        End If
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_Interval = arrRE
End Function
 
'������� ����ֵ����������  С�ڵ��ڱ���һ�� ArrInterval��������  ��������������ķ���
Public Function ArrGroup_Interval_Equal(ByVal arr2D, ByVal ColumnIndex, ParamArray arrInterval()) As Variant
    arrInterval = ArrSort1D(ArrDistinct(ArrFlatten(arrInterval)), True)
    Dim arrRE(), arrRECount() As Long, arrREindex() As Long, i As Long, j As Long, n As Long, k As Long, arrtmp(), p As Boolean
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LV, UV
    
    Dim lI As Long, UI As Long
    lI = LBound(arrInterval, 1): UI = UBound(arrInterval, 1)
    ReDim arrS(LH To UH) '���������
    ReDim arrRECount(1 To UI - lI + 2) '������
    For i = LH To UH
        'ѭ���Աȼ���������
        For j = lI To UI
            If arr2D(i, ColumnIndex) <= arrInterval(j) Then
                arrS(i) = j
                arrRECount(j) = arrRECount(j) + 1
                GoTo AlreadyWritten_
            End If
        Next
        'ʣ�µĶ��Ǵ�ķŵ����һ��
        arrS(i) = UI + 1
        arrRECount(UI + 1) = arrRECount(UI + 1) + 1
AlreadyWritten_:
    Next
    
    '����
    ReDim arrRE(1 To UI - lI + 2)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        If arrRECount(i) > 0 Then
            ReDim arrtmp(1 To arrRECount(i), LV To UV)
            arrRE(i) = arrtmp
            arrREindex(i) = 1
        Else
            arrRE(i) = Array()
        End If
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_Interval_Equal = arrRE
End Function
 
'������� ���Զ������ ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
Public Function ArrGroup_CustomClass(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant
    arrCustomValue = ArrFlatten(arrCustomValue)
    Dim arrRE(), arrRECount() As Long, arrREindex() As Long, i As Long, j As Long, n As Long, k As Long, arrtmp(), p As Boolean
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LV, UV
    
    Dim lI As Long, UI As Long
    lI = LBound(arrCustomValue, 1): UI = UBound(arrCustomValue, 1)
    ReDim arrS(LH To UH) '���������
    ReDim arrRECount(1 To UI - lI + 2) '������
    For i = LH To UH
        'ѭ���Աȼ���������
        For j = lI To UI
            If arr2D(i, ColumnIndex) = arrCustomValue(j) Then
                arrS(i) = j
                arrRECount(j) = arrRECount(j) + 1
                GoTo AlreadyWritten_
            End If
        Next
        'ʣ�µĶ��ŵ����һ��
        arrS(i) = UI + 1
        arrRECount(UI + 1) = arrRECount(UI + 1) + 1
AlreadyWritten_:
    Next
    
    '����
    ReDim arrRE(1 To UI - lI + 2)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        If arrRECount(i) > 0 Then
            ReDim arrtmp(1 To arrRECount(i), LV To UV)
            arrRE(i) = arrtmp
            arrREindex(i) = 1
        Else
            arrRE(i) = Array()
        End If
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_CustomClass = arrRE
End Function

'������� ���Զ������Likeƥ��  ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
Public Function ArrGroup_CustomClass_Like(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomValue()) As Variant
    arrCustomValue = ArrFlatten(arrCustomValue)
    Dim arrRE(), arrRECount() As Long, arrREindex() As Long, i As Long, j As Long, n As Long, k As Long, arrtmp(), p As Boolean
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LV, UV
    
    Dim lI As Long, UI As Long
    lI = LBound(arrCustomValue, 1): UI = UBound(arrCustomValue, 1)
    ReDim arrS(LH To UH) '���������
    ReDim arrRECount(1 To UI - lI + 2) '������
    For i = LH To UH
        'ѭ���Աȼ���������
        For j = lI To UI
            If arr2D(i, ColumnIndex) Like arrCustomValue(j) Then
                arrS(i) = j
                arrRECount(j) = arrRECount(j) + 1
                GoTo AlreadyWritten_
            End If
        Next
        'ʣ�µĶ��ŵ����һ��
        arrS(i) = UI + 1
        arrRECount(UI + 1) = arrRECount(UI + 1) + 1
AlreadyWritten_:
    Next
    
    '����
    ReDim arrRE(1 To UI - lI + 2)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        If arrRECount(i) > 0 Then
            ReDim arrtmp(1 To arrRECount(i), LV To UV)
            arrRE(i) = arrtmp
            arrREindex(i) = 1
        Else
            arrRE(i) = Array()
        End If
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_CustomClass_Like = arrRE
End Function

'������� ���Զ������ ����ƥ��  ��ƥ��ķ����һ�� arrCustomValueƥ������  ��������������ķ���
Public Function ArrGroup_CustomClass_Reg(ByVal arr2D, ByVal ColumnIndex, ParamArray arrCustomPattern()) As Variant
    arrCustomPattern = ArrFlatten(arrCustomPattern)
    Dim arrRE(), arrRECount() As Long, arrREindex() As Long, i As Long, j As Long, n As Long, k As Long, arrtmp(), p As Boolean
    Dim LH As Long, UH As Long
    Dim LV As Long, UV As Long
    LH = LBound(arr2D, 1): UH = UBound(arr2D, 1)
    LV = LBound(arr2D, 2): UV = UBound(arr2D, 2)
    
    IndexIsCurrencyToCount_ ColumnIndex, LV, UV
    
    Dim lI As Long, UI As Long
    lI = LBound(arrCustomPattern, 1): UI = UBound(arrCustomPattern, 1)
    ReDim arrS(LH To UH) '���������
    ReDim arrRECount(1 To UI - lI + 2) '������
    Dim Regex As Object
    For j = lI To UI
        Set Regex = CreateObject("VBScript.RegExp")
        With Regex
            .Global = False
            .ignoreCase = False
            .multiline = False
            .Pattern = arrCustomPattern(j)
        End With
        Set arrCustomPattern(j) = Regex
    Next
    For i = LH To UH
        'ѭ���Աȼ���������
        For j = lI To UI
            If arrCustomPattern(j).test(arr2D(i, ColumnIndex)) Then
                arrS(i) = j
                arrRECount(j) = arrRECount(j) + 1
                GoTo AlreadyWritten_
            End If
        Next
        'ʣ�µĶ��ŵ����һ��
        arrS(i) = UI + 1
        arrRECount(UI + 1) = arrRECount(UI + 1) + 1
AlreadyWritten_:
    Next
    
    '����
    ReDim arrRE(1 To UI - lI + 2)
    ReDim arrREindex(1 To UBound(arrRE)) As Long 'ÿ��ĵ�ǰ��
    '��ʼ����С
    For i = 1 To UBound(arrRE)
        If arrRECount(i) > 0 Then
            ReDim arrtmp(1 To arrRECount(i), LV To UV)
            arrRE(i) = arrtmp
            arrREindex(i) = 1
        Else
            arrRE(i) = Array()
        End If
    Next
    '����
    For i = LH To UH
        n = arrS(i)
        k = arrREindex(n)
        For j = LV To UV
            arrRE(n)(k, j) = arr2D(i, j)
        Next
        arrREindex(n) = k + 1
    Next
    ArrGroup_CustomClass_Reg = arrRE
End Function

'ArrGroup_IntervalRight �������ֵ���Ҳ�
'�������  ȡ�������Ԫ��
Public Function ArrUnions(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrUnion(arrRE, arr(i))
    Next
    ArrUnions = arrRE
End Function
 
'�������  ȥ��
Public Function ArrUnions_Distinct(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrUnion_Distinct(arrRE, arr(i))
    Next
    ArrUnions_Distinct = arrRE
End Function
 
'�������  ����
Public Function ArrUnions_Sort(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrUnion_Sort(arrRE, arr(i))
    Next
    ArrUnions_Sort = arrRE
End Function
 
'�������  ȥ������
Public Function ArrUnions_DistinctSort(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrUnion_DistinctSort(arrRE, arr(i))
    Next
    ArrUnions_DistinctSort = arrRE
End Function
 
'���� ȡ��������Ԫ��
Public Function ArrUnion(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object, dic2 As Object
    Set dic1 = DictionaryCreate(arr1)
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        ArrayDynamic_ arr1(i)
    Next
    For i = LBound(arr2) To UBound(arr2)
        ArrayDynamic_ arr2(i)
    Next
    ArrUnion = ArrayDynamic_
End Function
 
'���� ȥ��
Public Function ArrUnion_Distinct(ByRef arr1, ByRef arr2) As Variant
    ArrUnion_Distinct = ArrDistinct(ArrUnion(arr1, arr2))
End Function
 
'���� ����
Public Function ArrUnion_Sort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant
    ArrUnion_Sort = ArrSort1D(ArrUnion(arr1, arr2), Order)
End Function
 
'���� ȥ������
Public Function ArrUnion_DistinctSort(ByRef arr1, ByRef arr2, Optional Order As Boolean = True) As Variant
    ArrUnion_DistinctSort = ArrSort1D(ArrDistinct(ArrUnion(arr1, arr2)), Order)
End Function
 
'�������  ȡ�������Ԫ��
Public Function ArrIntersects(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrIntersect(arrRE, arr(i))
    Next
    ArrIntersects = arrRE
End Function
 
'�������  ȥ��
Public Function ArrIntersects_Distinct(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrIntersect_Distinct(arrRE, arr(i))
    Next
    ArrIntersects_Distinct = arrRE
End Function
 
'������� ȡ��һ������Ԫ��
Public Function ArrIntersects_arr1(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrIntersect_arr1(arrRE, arr(i))
    Next
    ArrIntersects_arr1 = arrRE
End Function
 
'������� ȡ��һ������Ԫ������
Public Function ArrIntersects_arr1_Index(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    If UBound(arr) = 0 Then
        ArrayDynamic_
        For i = LBound(arr(0)) To UBound(arr(0))
            ArrayDynamic_ i
        Next
        ArrIntersects_arr1_Index = ArrayDynamic_
    Else
        arrRE = ArrIntersect_arr1_Index(arr(0), arr(1))
        For i = 2 To UBound(arr)
            arrRE = ArrIntersect_arr1_Index_(arrRE, arr(0), arr(i))
        Next
        ArrIntersects_arr1_Index = arrRE
    End If
End Function
 
'���� ȡarr1�����ڲ�ʹ��  arr1_Index��������
Private Function ArrIntersect_arr1_Index_(ByRef arr1_Index, ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1_Index) To UBound(arr1_Index)
        If dic2.Exists(arr1(arr1_Index(i))) Then
            ArrayDynamic_ arr1_Index(i)
        End If
    Next
    ArrIntersect_arr1_Index_ = ArrayDynamic_
End Function
 
'���� ȡ��������Ԫ��
Public Function ArrIntersect(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object, dic2 As Object
    Set dic1 = DictionaryCreate(arr1)
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If dic2.Exists(arr1(i)) Then
            ArrayDynamic_ arr1(i)
        End If
    Next
    For i = LBound(arr2) To UBound(arr2)
        If dic1.Exists(arr2(i)) Then
            ArrayDynamic_ arr2(i)
        End If
    Next
    ArrIntersect = ArrayDynamic_
End Function
 
'���� ȥ��
Public Function ArrIntersect_Distinct(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If dic2.Exists(arr1(i)) Then
            ArrayDynamic_ arr1(i)
        End If
    Next
    ArrIntersect_Distinct = ArrDistinct(ArrayDynamic_)
End Function
 
'���� ȡarr1Ԫ��
Public Function ArrIntersect_arr1(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If dic2.Exists(arr1(i)) Then
            ArrayDynamic_ arr1(i)
        End If
    Next
    ArrIntersect_arr1 = ArrayDynamic_
End Function
 
'���� ȡarr2Ԫ��
Public Function ArrIntersect_arr2(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object
    Set dic1 = DictionaryCreate(arr1)
    ArrayDynamic_
    For i = LBound(arr2) To UBound(arr2)
        If dic1.Exists(arr2(i)) Then
            ArrayDynamic_ arr2(i)
        End If
    Next
    ArrIntersect_arr2 = ArrayDynamic_
End Function
 
'���� ȡarr1����
Public Function ArrIntersect_arr1_Index(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If dic2.Exists(arr1(i)) Then
            ArrayDynamic_ i
        End If
    Next
    ArrIntersect_arr1_Index = ArrayDynamic_
End Function
 
'���� ȡarr2����
Public Function ArrIntersect_arr2_Index(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object
    Set dic1 = DictionaryCreate(arr1)
    ArrayDynamic_
    For i = LBound(arr2) To UBound(arr2)
        If dic1.Exists(arr2(i)) Then
            ArrayDynamic_ i
        End If
    Next
    ArrIntersect_arr2_Index = ArrayDynamic_
End Function
 
'����  ȡ�������Ԫ��(������������������û�е�Ԫ��)[1,2,3,4,5,5][1,2,3][2,3,4,6]->[5,5,6]
Public Function ArrExcepts_Single(ParamArray arr()) As Variant
    Dim arrRE, i As Long, j As Long
    ArrayDynamic2_
    For i = 0 To UBound(arr)
        arrRE = arr(i)
        For j = 0 To UBound(arr)
            If i <> j Then
                arrRE = ArrExcept_arr1(arrRE, arr(j))
            End If
        Next
        ArrayDynamic2_ arrRE
    Next
    ArrExcepts_Single = ArrFlatten(ArrayDynamic2_)
End Function
 
'����  ȡ�������Ԫ��(ȥ���������鶼������Ԫ��)[1,2,3,4,5,5][1,2,3][2,3,4,6]->ȥ������Ԫ��2,3�õ�[1,4,5,5,1,4,6]
Public Function ArrExcepts_RemoveAllIntersect(ParamArray arr()) As Variant
    Dim arrRE, arrRE1, i As Long
    If UBound(arr) = 0 Then
        ArrExcepts_RemoveAllIntersect = ArrFlatten_Single(arr(0))
    Else
        arrRE = arr(0)
        For i = 1 To UBound(arr)
            arrRE = ArrIntersect_arr1(arrRE, arr(i))
        Next
        arrRE1 = ArrFlatten(arr)
        ArrExcepts_RemoveAllIntersect = ArrExcept_arr1(arrRE1, arrRE)
    End If
End Function
 
'����  ȡ��һ��Ԫ��
Public Function ArrExcepts_arr1(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    arrRE = arr(0)
    For i = 1 To UBound(arr)
        arrRE = ArrExcept_arr1(arrRE, arr(i))
    Next
    ArrExcepts_arr1 = arrRE
End Function
 
'���� ȡ��һ������Ԫ������
Public Function ArrExcepts_arr1_Index(ParamArray arr()) As Variant
    Dim arrRE, i As Long
    If UBound(arr) = 0 Then
        ArrayDynamic_
        For i = LBound(arr(0)) To UBound(arr(0))
            ArrayDynamic_ i
        Next
        ArrExcepts_arr1_Index = ArrayDynamic_
    Else
        arrRE = ArrExcept_arr1_Index(arr(0), arr(1))
        For i = 2 To UBound(arr)
            arrRE = ArrExcept_arr1_Index_(arrRE, arr(0), arr(i))
        Next
        ArrExcepts_arr1_Index = arrRE
    End If
End Function
 
'�  ȡarr1�����ڲ�ʹ��  arr1_Index��������
Private Function ArrExcept_arr1_Index_(ByRef arr1_Index, ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1_Index) To UBound(arr1_Index)
        If Not dic2.Exists(arr1(arr1_Index(i))) Then
            ArrayDynamic_ arr1_Index(i)
        End If
    Next
    ArrExcept_arr1_Index_ = ArrayDynamic_
End Function
 
'� ȡ��������Ԫ��
Public Function ArrExcept(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object, dic2 As Object
    Set dic1 = DictionaryCreate(arr1)
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If Not dic2.Exists(arr1(i)) Then
            ArrayDynamic_ arr1(i)
        End If
    Next
    For i = LBound(arr2) To UBound(arr2)
        If Not dic1.Exists(arr2(i)) Then
            ArrayDynamic_ arr2(i)
        End If
    Next
    ArrExcept = ArrayDynamic_
End Function
 
'� ȡarr1Ԫ��
Public Function ArrExcept_arr1(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If Not dic2.Exists(arr1(i)) Then
            ArrayDynamic_ arr1(i)
        End If
    Next
    ArrExcept_arr1 = ArrayDynamic_
End Function
 
'� ȡarr2Ԫ��
Public Function ArrExcept_arr2(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object
    Set dic1 = DictionaryCreate(arr1)
    ArrayDynamic_
    For i = LBound(arr2) To UBound(arr2)
        If Not dic1.Exists(arr2(i)) Then
            ArrayDynamic_ arr2(i)
        End If
    Next
    ArrExcept_arr2 = ArrayDynamic_
End Function
 
'� ȡarr1����
Public Function ArrExcept_arr1_Index(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic2 As Object
    Set dic2 = DictionaryCreate(arr2)
    ArrayDynamic_
    For i = LBound(arr1) To UBound(arr1)
        If Not dic2.Exists(arr1(i)) Then
            ArrayDynamic_ i
        End If
    Next
    ArrExcept_arr1_Index = ArrayDynamic_
End Function
 
'� ȡarr2����
Public Function ArrExcept_arr2_Index(ByRef arr1, ByRef arr2) As Variant
    Dim i As Long
    Dim dic1 As Object
    Set dic1 = DictionaryCreate(arr1)
    ArrayDynamic_
    For i = LBound(arr2) To UBound(arr2)
        If Not dic1.Exists(arr2(i)) Then
            ArrayDynamic_ i
        End If
    Next
    ArrExcept_arr2_Index = ArrayDynamic_
End Function
 
'arrTitle(һά)��arrOrder(һά)���ض�Ӧ��˳��ı�����������,���ص�����ΪarrTitle������ƥ��λ�÷���(LBound-1),���ص������С��arrOrder��ͬ
Public Function ArrTitleToIndex(ByRef arrTitle, ByRef arrOrder) As Variant
    Dim i As Long, j As Long, p As Boolean
    Dim l As Long, u As Long
    Dim arrRE(): ReDim arrRE(LBound(arrOrder) To UBound(arrOrder))
    l = LBound(arrTitle): u = UBound(arrTitle)
    For i = LBound(arrOrder) To UBound(arrOrder)
        p = True
        For j = l To u
            If arrOrder(i) = arrTitle(j) Then
                arrRE(i) = j
                p = False
                Exit For
            End If
        Next
        If p Then
            arrRE(i) = LBound(arrTitle) - 1
        End If
    Next
    ArrTitleToIndex = arrRE
End Function
 
'���鲼���Ҽ���
Public Function ArrBoolea_And(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v And arr2(j) '����
                        j = j + 1
                    Next
                    For i2 = j To UBound(arr2)
                        arr2(i2) = False And arr2(i2) '����λ����
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) And v '����
                        j = j + 1
                    Next
                    For i2 = j To UBound(arr)
                        arr(i2) = False And arr(i2) '����λ����
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) And v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr And arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr And Calculates(i) '����
            End If
        End If
    Next
    ArrBoolea_And = arr
End Function
 
'���鲼�������
Public Function ArrBoolea_Or(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v Or arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) Or v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) Or v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr Or arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr Or Calculates(i) '����
            End If
        End If
    Next
    ArrBoolea_Or = arr
End Function
 
'���鲼���Ǽ���
Public Function ArrBoolea_Not(ByVal arr) As Variant
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = Not arr(i) '����
    Next
    ArrBoolea_Not = arr
End Function
 
'����IFs�жϼ��� ArrIFs(����,ֵ,����,ֵ,����ֵ)
Public Function ArrIFs(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v, arrRE(), arr, l As Long
    Dim maxL As Long: maxL = 0
    For i = LBound(Calculates) To UBound(Calculates)
        If IsArray(Calculates(i)) Then
            j = ArrCount(Calculates(i))
            If maxL < j Then maxL = j
        Else
            If maxL < 1 Then maxL = 1
        End If
    Next
    ReDim arrRE(1 To maxL)
    ArrayDynamic2_
    l = UBound(Calculates) - LBound(Calculates) + 1
    If IsOdd(l) Then j = UBound(Calculates) - 1 Else j = UBound(Calculates)
    For i = LBound(Calculates) To j
        If IsArray(Calculates(i)) Then
            If IsOdd(i) Then
                ArrayDynamic2_ ArrSizeExpansion2(Calculates(i), maxL, False)
            Else
                ArrayDynamic2_ ArrSizeExpansion2(Calculates(i), maxL)
            End If
        Else
            ArrayDynamic2_ ArrSizeExpansion2(Calculates(i), maxL, Calculates(i))
        End If
    Next
    If IsOdd(l) Then
        If IsArray(Calculates(UBound(Calculates))) Then
            ArrayDynamic2_ ArrSizeExpansion2(Calculates(UBound(Calculates)), maxL)
        Else
            ArrayDynamic2_ ArrSizeExpansion2(Calculates(UBound(Calculates)), maxL, Calculates(UBound(Calculates)))
        End If
    End If
    arr = ArrayDynamic2_
    For i = 1 To maxL
        For j = 1 To l - 1 Step 2
            If arr(j)(i) Then
                Cover arrRE(i), arr(j + 1)(i)
                Exit For
            ElseIf j = l - 2 Then
                Cover arrRE(i), arr(l)(i)
                Exit For
            End If
        Next
    Next
    ArrIFs = arrRE
End Function

'��������Ƚϼ��� �ڲ�
Public Function ArrComp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Select Case NumberRangeRule
        Case Include_Exclude: ArrComp_RangeInside = ArrBoolea_And(ArrComp_SizeEqual(arr, arrL), ArrComp_Size(arrR, arr))
        Case Exclude_Include: ArrComp_RangeInside = ArrBoolea_And(ArrComp_Size(arr, arrL), ArrComp_SizeEqual(arrR, arr))
        Case Include_Include: ArrComp_RangeInside = ArrBoolea_And(ArrComp_SizeEqual(arr, arrL), ArrComp_SizeEqual(arrR, arr))
        Case Exclude_Exclude: ArrComp_RangeInside = ArrBoolea_And(ArrComp_Size(arr, arrL), ArrComp_Size(arrR, arr))
    End Select
End Function
 
'��������Ƚϼ��� �ⲿ
Public Function ArrComp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Select Case NumberRangeRule
        Case Include_Exclude: ArrComp_RangeExternal = ArrBoolea_Or(ArrComp_SizeEqual(arrL, arr), ArrComp_Size(arr, arrR))
        Case Exclude_Include: ArrComp_RangeExternal = ArrBoolea_Or(ArrComp_Size(arrL, arr), ArrComp_SizeEqual(arr, arrR))
        Case Include_Include: ArrComp_RangeExternal = ArrBoolea_Or(ArrComp_SizeEqual(arrL, arr), ArrComp_SizeEqual(arr, arrR))
        Case Exclude_Exclude: ArrComp_RangeExternal = ArrBoolea_Or(ArrComp_Size(arrL, arr), ArrComp_Size(arr, arrR))
    End Select
End Function
 
'����Ƚ�Like����
Public Function ArrComp_Like(ByVal arr, ByVal arr2) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr) Then
        If IsArray(arr2) Then
            If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                j = LBound(arr2)
                For Each v In arr
                    arr2(j) = v Like arr2(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr2)
                    arr2(i2) = False '����λ����
                Next
                arr = arr2
            Else
                j = LBound(arr)
                For Each v In arr2
                    arr(j) = arr(j) Like v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr)
                    arr(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr) To UBound(arr)
                arr(j) = arr(j) Like arr2 '����
            Next
        End If
    Else
        If IsArray(arr2) Then
            For j = LBound(arr2) To UBound(arr2)
                arr2(j) = arr Like arr2(j) '����
            Next
            arr = arr2
        Else
            arr = arr Like arr2 '����
        End If
    End If
    ArrComp_Like = arr
End Function
 
'����Ƚ�Not Like����
Public Function ArrComp_NotLike(ByVal arr, ByVal arr2) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr) Then
        If IsArray(arr2) Then
            If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                j = LBound(arr2)
                For Each v In arr
                    arr2(j) = Not v Like arr2(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr2)
                    arr2(i2) = False '����λ����
                Next
                arr = arr2
            Else
                j = LBound(arr)
                For Each v In arr2
                    arr(j) = Not arr(j) Like v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr)
                    arr(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr) To UBound(arr)
                arr(j) = Not arr(j) Like arr2 '����
            Next
        End If
    Else
        If IsArray(arr2) Then
            For j = LBound(arr2) To UBound(arr2)
                arr2(j) = Not arr Like arr2(j) '����
            Next
            arr = arr2
        Else
            arr = Not arr Like arr2 '����
        End If
    End If
    ArrComp_NotLike = arr
End Function
 
'����Ƚϵ��ڼ���
Public Function ArrComp_Equal(ByVal arr, ByVal arr2) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr) Then
        If IsArray(arr2) Then
            If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                j = LBound(arr2)
                For Each v In arr
                    arr2(j) = v = arr2(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr2)
                    arr2(i2) = False '����λ����
                Next
                arr = arr2
            Else
                j = LBound(arr)
                For Each v In arr2
                    arr(j) = arr(j) = v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr)
                    arr(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr) To UBound(arr)
                arr(j) = arr(j) = arr2 '����
            Next
        End If
    Else
        If IsArray(arr2) Then
            For j = LBound(arr2) To UBound(arr2)
                arr2(j) = arr = arr2(j) '����
            Next
            arr = arr2
        Else
            arr = arr = arr2 '����
        End If
    End If
    ArrComp_Equal = arr
End Function
 
'����Ƚϲ����ڼ���
Public Function ArrComp_NotEqual(ByVal arr, ByVal arr2) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr) Then
        If IsArray(arr2) Then
            If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                j = LBound(arr2)
                For Each v In arr
                    arr2(j) = v <> arr2(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr2)
                    arr2(i2) = False '����λ����
                Next
                arr = arr2
            Else
                j = LBound(arr)
                For Each v In arr2
                    arr(j) = arr(j) <> v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr)
                    arr(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr) To UBound(arr)
                arr(j) = arr(j) <> arr2 '����
            Next
        End If
    Else
        If IsArray(arr2) Then
            For j = LBound(arr2) To UBound(arr2)
                arr2(j) = arr <> arr2(j) '����
            Next
            arr = arr2
        Else
            arr = arr <> arr2 '����
        End If
    End If
    ArrComp_NotEqual = arr
End Function
 
'����Ƚϴ�С����
Public Function ArrComp_Size(ByVal arr_Large, ByVal arr_Small) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr_Large) Then
        If IsArray(arr_Small) Then
            If UBound(arr_Small) - LBound(arr_Small) > UBound(arr_Large) - LBound(arr_Large) Then
                j = LBound(arr_Small)
                For Each v In arr_Large
                    arr_Small(j) = v > arr_Small(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr_Small)
                    arr_Small(i2) = False '����λ����
                Next
                arr_Large = arr_Small
            Else
                j = LBound(arr_Large)
                For Each v In arr_Small
                    arr_Large(j) = arr_Large(j) > v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr_Large)
                    arr_Large(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr_Large) To UBound(arr_Large)
                arr_Large(j) = arr_Large(j) > arr_Small '����
            Next
        End If
    Else
        If IsArray(arr_Small) Then
            For j = LBound(arr_Small) To UBound(arr_Small)
                arr_Small(j) = arr_Large > arr_Small(j) '����
            Next
            arr_Large = arr_Small
        Else
            arr_Large = arr_Large > arr_Small '����
        End If
    End If
    ArrComp_Size = arr_Large
End Function
 
'����Ƚϴ�С�������ڼ���
Public Function ArrComp_SizeEqual(ByVal arr_Large, ByVal arr_Small) As Variant
    Dim i As Long, j As Long, v, i2 As Long
    If IsArray(arr_Large) Then
        If IsArray(arr_Small) Then
            If UBound(arr_Small) - LBound(arr_Small) > UBound(arr_Large) - LBound(arr_Large) Then
                j = LBound(arr_Small)
                For Each v In arr_Large
                    arr_Small(j) = v >= arr_Small(j) '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr_Small)
                    arr_Small(i2) = False '����λ����
                Next
                arr_Large = arr_Small
            Else
                j = LBound(arr_Large)
                For Each v In arr_Small
                    arr_Large(j) = arr_Large(j) >= v '����
                    j = j + 1
                Next
                For i2 = j To UBound(arr_Large)
                    arr_Large(i2) = False '����λ����
                Next
            End If
        Else
            For j = LBound(arr_Large) To UBound(arr_Large)
                arr_Large(j) = arr_Large(j) >= arr_Small '����
            Next
        End If
    Else
        If IsArray(arr_Small) Then
            For j = LBound(arr_Small) To UBound(arr_Small)
                arr_Small(j) = arr_Large >= arr_Small(j) '����
            Next
            arr_Large = arr_Small
        Else
            arr_Large = arr_Large >= arr_Small '����
        End If
    End If
    ArrComp_SizeEqual = arr_Large
End Function
 
'����ӷ�����
Public Function ArrMath_Add(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v + arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) + v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) + v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr + arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr + Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Add = arr
End Function
 
'�����������
Public Function ArrMath_Sub(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v - arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) - v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) - v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr - arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr - Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Sub = arr
End Function
 
'����˷�����
Public Function ArrMath_Multipli(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v * arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) * v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) * v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr * arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr * Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Multipli = arr
End Function
 
'�����������
Public Function ArrMath_Division(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v / arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) / v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) / v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr / arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr / Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Division = arr
End Function
 
'����˷�����
Public Function ArrMath_Power(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v ^ arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) ^ v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) ^ v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr ^ arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr ^ Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Power = arr
End Function
 
'�������Ӽ���
Public Function ArrMath_Join(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, v
    Dim arr, arr2
    arr = Calculates(LBound(Calculates))
    For i = LBound(Calculates) + 1 To UBound(Calculates)
        If IsArray(arr) Then
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                If UBound(arr2) - LBound(arr2) > UBound(arr) - LBound(arr) Then
                    j = LBound(arr2)
                    For Each v In arr
                        arr2(j) = v & arr2(j) '����
                        j = j + 1
                    Next
                    arr = arr2
                Else
                    j = LBound(arr)
                    For Each v In arr2
                        arr(j) = arr(j) & v '����
                        j = j + 1
                    Next
                End If
            Else
                v = Calculates(i)
                For j = LBound(arr) To UBound(arr)
                    arr(j) = arr(j) & v '����
                Next
            End If
        Else
            If IsArray(Calculates(i)) Then
                arr2 = Calculates(i)
                For j = LBound(arr2) To UBound(arr2)
                    arr2(j) = arr & arr2(j) '����
                Next
                arr = arr2
            Else
                arr = arr & Calculates(i) '����
            End If
        End If
    Next
    ArrMath_Join = arr
End Function
 
'������������
Public Function ArrMath_Round(ByVal arr, Optional Number = 0, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = RoundEX(arr(i), Number)     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = RoundEX(arr(i, ColumnIndex), Number)    '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = RoundEX(arr(i, ColumnIndexArr), Number)    '����
            Next
        End If
    End If
    ArrMath_Round = arr
End Function
 
'���� ������� [1,2,3,4,5]->[1,3,6,10,15]
Public Function ArrMath_SumIncrease(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) + 1 To UBound(arr)
            arr(i) = arr(i) + arr(i - 1) '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = arr(i, ColumnIndex) + arr(i - 1, ColumnIndex) '����
                Next
            Next
        Else
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                arr(i, ColumnIndexArr) = arr(i, ColumnIndexArr) + arr(i - 1, ColumnIndexArr) '����
            Next
        End If
    End If
    ArrMath_SumIncrease = arr
End Function
 
'����ת����
Public Function ArrMath_Val(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Val(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Val(arr(i, ColumnIndex))     '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Val(arr(i, ColumnIndexArr))     '����
            Next
        End If
    End If
    ArrMath_Val = arr
End Function

'�������ֵAbs
Public Function ArrMath_Abs(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Abs(arr(i))      '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Abs(arr(i, ColumnIndex))      '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Abs(arr(i, ColumnIndexArr))      '����
            Next
        End If
    End If
    ArrMath_Abs = arr
End Function

'����Format
Public Function ArrMath_Format(ByVal arr, Pormat, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Format(arr(i), Pormat, vbMonday, vbFirstFullWeek)  '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Format(arr(i, ColumnIndex), Pormat, vbMonday, vbFirstFullWeek) '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Format(arr(i, ColumnIndexArr), Pormat, vbMonday, vbFirstFullWeek) '����
            Next
        End If
    End If
    ArrMath_Format = arr
End Function

'����Trim
Public Function ArrStr_Trim(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Trim(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Trim(arr(i, ColumnIndex))     '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Trim(arr(i, ColumnIndexArr))      '����
            Next
        End If
    End If
    ArrStr_Trim = arr
End Function

'����RTrim
Public Function ArrStr_RTrim(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.RTrim(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.RTrim(arr(i, ColumnIndex))     '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.RTrim(arr(i, ColumnIndexArr))      '����
            Next
        End If
    End If
    ArrStr_RTrim = arr
End Function

'����LTrim
Public Function ArrStr_LTrim(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.LTrim(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.LTrim(arr(i, ColumnIndex))      '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.LTrim(arr(i, ColumnIndexArr))      '����
            Next
        End If
    End If
    ArrStr_LTrim = arr
End Function

'����ת��д
Public Function ArrStr_Ucase(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Ucase(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Ucase(arr(i, ColumnIndex))     '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Ucase(arr(i, ColumnIndexArr))     '����
            Next
        End If
    End If
    ArrStr_Ucase = arr
End Function

'����תСд
Public Function ArrStr_Lcase(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Lcase(arr(i))      '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Lcase(arr(i, ColumnIndex))     '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Lcase(arr(i, ColumnIndexArr))     '����
            Next
        End If
    End If
    ArrStr_Lcase = arr
End Function

'����ѭ������ַ��� ��������������
Public Function ArrStr_Split(ByVal arr, Delimiter, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = Str_Split(arr(i), Delimiter)      '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = Str_Split(arr(i, ColumnIndex), Delimiter)   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = Str_Split(arr(i, ColumnIndexArr), Delimiter)   '����
            Next
        End If
    End If
    ArrStr_Split = arr
End Function
 
'�����滻
Public Function ArrStr_Replace(ByVal arr, FindStr, ReplaceStr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Replace(arr(i), FindStr, ReplaceStr)    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Replace(arr(i, ColumnIndex), FindStr, ReplaceStr)    '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Replace(arr(i, ColumnIndexArr), FindStr, ReplaceStr)    '����
            Next
        End If
    End If
    ArrStr_Replace = arr
End Function
 
'�����滻������������
Public Function ArrStr_ReplaceAll(ByVal arr, FindStr, ReplaceStr) As Variant
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Replace(arr(i), FindStr, ReplaceStr)    '����
        Next
    Else
        l = LBound(arr, 2): u = UBound(arr, 2)
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = l To u
                arr(i, j) = VBA.Replace(arr(i, j), FindStr, ReplaceStr)    '����
            Next
        Next
    End If
    ArrStr_ReplaceAll = arr
End Function
 
'��������ȡֵ
Public Function ArrStr_RegexSearch(ByVal arr, Pattern, Optional RegIndex = 0, Optional ByVal ColumnIndexArr = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant
 
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = StrRegexSearch(arr(i), Pattern, RegIndex, True, ignoreCase, multiline)   '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = StrRegexSearch(arr(i, ColumnIndex), Pattern, RegIndex, True, ignoreCase, multiline)      '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = StrRegexSearch(arr(i, ColumnIndexArr), Pattern, RegIndex, True, ignoreCase, multiline)       '����
            Next
        End If
    End If
    ArrStr_RegexSearch = arr
End Function
 
'��������ȡ����ֵ��������������
Public Function ArrStr_RegexSearchs(ByVal arr, Pattern, Optional ByVal ColumnIndexArr = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant
 
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = StrRegexSearchs(arr(i), Pattern, True, ignoreCase, multiline)      '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = StrRegexSearchs(arr(i, ColumnIndex), Pattern, True, ignoreCase, multiline)        '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = StrRegexSearchs(arr(i, ColumnIndexArr), Pattern, True, ignoreCase, multiline)        '����
            Next
        End If
    End If
    ArrStr_RegexSearchs = arr
End Function

'�������򷵻�ƥ������
Public Function ArrStr_RegexCount(ByVal arr, Pattern, Optional ByVal ColumnIndexArr = 1, Optional ByVal NumberAdd = 0, _
         Optional ByRef ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant
 
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = StrRegexCount(arr(i), Pattern, ignoreCase, multiline) + NumberAdd    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = StrRegexCount(arr(i, ColumnIndex), Pattern, ignoreCase, multiline) + NumberAdd       '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = StrRegexCount(arr(i, ColumnIndexArr), Pattern, ignoreCase, multiline) + NumberAdd       '����
            Next
        End If
    End If
    ArrStr_RegexCount = arr
End Function

'���������滻
Public Function ArrStr_RegexReplace(ByVal arr, Pattern, ReplaceStr, Optional ByVal ColumnIndexArr = 1, _
        Optional ByVal ignoreCase As Boolean = False, Optional ByVal multiline As Boolean = False) As Variant
 
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = StrRegexReplace(arr(i), Pattern, ReplaceStr, True, ignoreCase, multiline)     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = StrRegexReplace(arr(i, ColumnIndex), Pattern, ReplaceStr, True, ignoreCase, multiline)       '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = StrRegexReplace(arr(i, ColumnIndexArr), Pattern, ReplaceStr, True, ignoreCase, multiline)      '����
            Next
        End If
    End If
    ArrStr_RegexReplace = arr
End Function
 
'����MID
Public Function ArrStr_Mid(ByVal arr, Start, Optional Length, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        If IsMissing(Length) Then
            For i = LBound(arr) To UBound(arr)
                arr(i) = VBA.Mid(arr(i), Start)    '����
            Next
        Else
            For i = LBound(arr) To UBound(arr)
                arr(i) = VBA.Mid(arr(i), Start, Length)   '����
            Next
        End If
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsMissing(Length) Then
            If IsArray(ColumnIndexArr) Then
                Dim ColumnIndex
                For i = LBound(arr, 1) To UBound(arr, 1)
                    For Each ColumnIndex In ColumnIndexArr
                        arr(i, ColumnIndex) = VBA.Mid(arr(i, ColumnIndex), Start)     '����
                    Next
                Next
            Else
                For i = LBound(arr, 1) To UBound(arr, 1)
                    arr(i, ColumnIndexArr) = VBA.Mid(arr(i, ColumnIndexArr), Start)     '����
                Next
            End If
        Else
            If IsArray(ColumnIndexArr) Then
                For i = LBound(arr, 1) To UBound(arr, 1)
                    For Each ColumnIndex In ColumnIndexArr
                        arr(i, ColumnIndex) = VBA.Mid(arr(i, ColumnIndex), Start, Length)      '����
                    Next
                Next
            Else
                For i = LBound(arr, 1) To UBound(arr, 1)
                    arr(i, ColumnIndexArr) = VBA.Mid(arr(i, ColumnIndexArr), Start, Length)     '����
                Next
            End If
        End If
    End If
    ArrStr_Mid = arr
End Function

'����������ڲ�ֵ ����DateDiff
Public Function ArrDate_DateSub(Interval, Date1, Date2) As Variant
    Dim i As Long, j As Long, v, arrRE(), arr, l As Long
    Dim maxL As Long: maxL = 1
    If IsArray(Interval) Then j = ArrCount(Interval): If maxL < j Then maxL = j
    If IsArray(Date1) Then j = ArrCount(Date1): If maxL < j Then maxL = j
    If IsArray(Date2) Then j = ArrCount(Date2): If maxL < j Then maxL = j
    
    ReDim arrRE(1 To maxL)
    Dim IntervalRE, Date1RE, Date2RE
    If IsArray(Interval) Then
        IntervalRE = ArrSizeExpansion2(Interval, maxL, "")
    Else
        IntervalRE = ArrSizeExpansion2(Interval, maxL, Interval)
    End If
    If IsArray(Date1) Then
        Date1RE = ArrSizeExpansion2(Date1, maxL)
    Else
        Date1RE = ArrSizeExpansion2(Date1, maxL, Date1)
    End If
    If IsArray(Date2) Then
        Date2RE = ArrSizeExpansion2(Date2, maxL)
    Else
        Date2RE = ArrSizeExpansion2(Date2, maxL, Date2)
    End If
    
    For i = 1 To maxL
        arrRE(i) = VBA.DateDiff(IntervalRE(i), Date1RE(i), Date2RE(i), vbMonday)
    Next
    ArrDate_DateSub = arrRE
End Function

'����ȡ��
Public Function ArrDate_Year(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Year(arr(i))    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Year(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Year(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrDate_Year = arr
End Function
 
'����ȡ��
Public Function ArrDate_Month(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Month(arr(i))    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Month(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Month(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrDate_Month = arr
End Function
 
'����ȡ��
Public Function ArrDate_Day(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Day(arr(i))    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Day(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Day(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrDate_Day = arr
End Function
 
'����ȡ����
Public Function ArrDate_Weekday(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Weekday(arr(i), vbMonday)    '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Weekday(arr(i, ColumnIndex), vbMonday)   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Weekday(arr(i, ColumnIndexArr), vbMonday)   '����
            Next
        End If
    End If
    ArrDate_Weekday = arr
End Function
 
'����ȡСʱ
Public Function ArrTime_Hour(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Hour(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Hour(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Hour(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrTime_Hour = arr
End Function
 
'����ȡ����
Public Function ArrTime_Minute(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Minute(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Minute(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Minute(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrTime_Minute = arr
End Function
 
'����ȡ��
Public Function ArrTime_Second(ByVal arr, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = VBA.Second(arr(i))     '����
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr, 2), UBound(arr, 2)
    
        If IsArray(ColumnIndexArr) Then
            Dim ColumnIndex
            For i = LBound(arr, 1) To UBound(arr, 1)
                For Each ColumnIndex In ColumnIndexArr
                    arr(i, ColumnIndex) = VBA.Second(arr(i, ColumnIndex))   '����
                Next
            Next
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                arr(i, ColumnIndexArr) = VBA.Second(arr(i, ColumnIndexArr))   '����
            Next
        End If
    End If
    ArrTime_Second = arr
End Function
 
'����� �������鷵��1++���
Public Function ArrSerialNumber(ByVal arr, Optional ByVal ColumnIndex = 1, Optional StartNumber = 1) As Variant
    Dim i As Long, k As Long
    k = StartNumber
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            arr(i) = k
            k = k + 1
        Next
    Else
        IndexIsCurrencyToCount_ ColumnIndex, LBound(arr, 2), UBound(arr, 2)
    
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr(i, ColumnIndex) = k
            k = k + 1
        Next
    End If
    ArrSerialNumber = arr
End Function
 
'����� �����鲻ͬ���� ��ͬ����1++ ����1++���
Public Function ArrSerialNumberCalssSelf(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim dic As Object, i As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            If dic.Exists(arr(i)) Then
                dic(arr(i)) = dic(arr(i)) + 1
                arr(i) = dic(arr(i))
            Else
                dic.Add arr(i), StartNumber
                arr(i) = StartNumber
            End If
        Next
    Else
        IndexIsCurrencyToCount_ InputIndex, LBound(arr, 2), UBound(arr, 2)
        IndexIsCurrencyToCount_ CalssIndex, LBound(arr, 2), UBound(arr, 2)
    
        For i = LBound(arr, 1) To UBound(arr, 1)
            If dic.Exists(arr(i, CalssIndex)) Then
                dic(arr(i, CalssIndex)) = dic(arr(i, CalssIndex)) + 1
                arr(i, InputIndex) = dic(arr(i, CalssIndex))
            Else
                dic.Add arr(i, CalssIndex), StartNumber
                arr(i, InputIndex) = StartNumber
            End If
        Next
    End If
    ArrSerialNumberCalssSelf = arr
End Function

'����� �����鲻ͬ����1++ ����1++���
Public Function ArrSerialNumberCalss(ByVal arr, Optional ByVal InputIndex = 1, Optional ByVal CalssIndex = 1, Optional StartNumber = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim dic As Object, i As Long, j As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    j = StartNumber
    If ArrDimension(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            If dic.Exists(arr(i)) Then
                arr(i) = dic(arr(i))
            Else
                dic.Add arr(i), j
                arr(i) = j
                j = j + 1
            End If
        Next
    Else
        IndexIsCurrencyToCount_ InputIndex, LBound(arr, 2), UBound(arr, 2)
        IndexIsCurrencyToCount_ CalssIndex, LBound(arr, 2), UBound(arr, 2)
        For i = LBound(arr, 1) To UBound(arr, 1)
            If dic.Exists(arr(i, CalssIndex)) Then
                arr(i, InputIndex) = dic(arr(i, CalssIndex))
            Else
                dic.Add arr(i, CalssIndex), j
                arr(i, InputIndex) = j
                j = j + 1
            End If
        Next
    End If
    ArrSerialNumberCalss = arr
End Function

'����ȡ���ֵ���� ColumnIndex ��ά����������  Front = True ��ǰ������
Public Function ArrMaxIndex(ByRef arr, Optional ByVal ColumnIndex = 1, Optional Front As Boolean = True) As Long
    Dim i As Long, MaxIndex As Long
    If ArrDimension(arr) = 1 Then
        MaxIndex = LBound(arr)
        If Front Then
            For i = LBound(arr) + 1 To UBound(arr)
                If arr(MaxIndex) * 1 < arr(i) * 1 Then MaxIndex = i
            Next
        Else
            For i = LBound(arr) + 1 To UBound(arr)
                If arr(MaxIndex) * 1 <= arr(i) * 1 Then MaxIndex = i
            Next
        End If
    Else
        IndexIsCurrencyToCount_ ColumnIndex, LBound(arr, 2), UBound(arr, 2)
        MaxIndex = LBound(arr, 1)
        If Front Then
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                If arr(MaxIndex, ColumnIndex) * 1 < arr(i, ColumnIndex) * 1 Then MaxIndex = i
            Next
        Else
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                If arr(MaxIndex, ColumnIndex) * 1 <= arr(i, ColumnIndex) * 1 Then MaxIndex = i
            Next
        End If
    End If
    ArrMaxIndex = MaxIndex
End Function
 
'����ȡ��Сֵ���� ColumnIndex ��ά����������  Front = True ��ǰ������
Public Function ArrMinIndex(ByRef arr, Optional ByVal ColumnIndex = 1, Optional Front As Boolean = True) As Long
    Dim i As Long, MinIndex As Long
    If ArrDimension(arr) = 1 Then
        MinIndex = LBound(arr)
        If Front Then
            For i = LBound(arr) + 1 To UBound(arr)
                If arr(MinIndex) * 1 > arr(i) * 1 Then MinIndex = i
            Next
        Else
            For i = LBound(arr) + 1 To UBound(arr)
                If arr(MinIndex) * 1 >= arr(i) * 1 Then MinIndex = i
            Next
        End If
    Else
        IndexIsCurrencyToCount_ ColumnIndex, LBound(arr, 2), UBound(arr, 2)
        MinIndex = LBound(arr, 1)
        If Front Then
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                If arr(MinIndex, ColumnIndex) * 1 > arr(i, ColumnIndex) * 1 Then MinIndex = i
            Next
        Else
            For i = LBound(arr, 1) + 1 To UBound(arr, 1)
                If arr(MinIndex, ColumnIndex) * 1 >= arr(i, ColumnIndex) * 1 Then MinIndex = i
            Next
        End If
    End If
    ArrMinIndex = MinIndex
End Function

'�������
Public Function ArrSum(ByRef arr) As Double
    Dim v
    For Each v In arr
        ArrSum = ArrSum + VBA.Val(v)
    Next
End Function
 
'���������ֵ
Public Function ArrMax(ByRef arr) As Double
    Dim v
    ArrMax = -1.79769313486231E+308
    For Each v In arr
        If IsNumeric(v) Then
            If ArrMax < VBA.Val(v) Then ArrMax = VBA.Val(v)
        End If
    Next
End Function
 
'��������Сֵ
Public Function ArrMin(ByRef arr) As Double
    Dim v
    ArrMin = 1.79769313486231E+308
    For Each v In arr
        If IsNumeric(v) Then
            If ArrMin > VBA.Val(v) Then ArrMin = VBA.Val(v)
        End If
    Next
End Function

'�������ǿ�ֵ����
Public Function ArrCountNoEmpty(ByRef arr) As Double
    Dim v
    ArrCountNoEmpty = 0
    For Each v In arr
        If v <> "" Then
            ArrCountNoEmpty = ArrCountNoEmpty + 1
        End If
    Next
End Function

'���������ƽ��ֵ
Public Function ArrAverage(ByRef arr, Optional NumDigitsAfterDecimal As Long = 2) As Double
    Dim v, REcount, REsum
    REcount = 0
    ArrAverage = 0
    For Each v In arr
        If v <> "" Then
            REcount = REcount + 1
            REsum = REsum + VBA.Val(v)
        End If
    Next
    If REcount <> 0 Then ArrAverage = RoundEX(REsum / REcount, NumDigitsAfterDecimal)
End Function

'���鰴�����
Public Function ArrSumColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            For i = l To u
                arrRE(j) = arrRE(j) + VBA.Val(arr2D(i, ColumnIndex))   '����
            Next
        Next
        ArrSumColumn = arrRE
    Else
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            ArrSumColumn = ArrSumColumn + VBA.Val(arr2D(i, ColumnIndexArr))    '����
        Next
    End If
End Function

'���鰴�����
Public Function ArrSumRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            For i = l To u
                arrRE(j) = arrRE(j) + VBA.Val(arr2D(RowIndex, i))    '����
            Next
        Next
        ArrSumRow = arrRE
    Else
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            ArrSumRow = ArrSumRow + VBA.Val(arr2D(RowIndexArr, i))     '����
        Next
    End If
End Function

'���鰴�������ֵ
Public Function ArrMaxColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long, j As Long, v
    Dim arrRE()
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            arrRE(j) = VBA.Val(arr2D(l, ColumnIndex))
            For i = l + 1 To u
                v = VBA.Val(arr2D(i, ColumnIndex))
                If arrRE(j) < v Then arrRE(j) = v   '����
            Next
        Next
        ArrMaxColumn = arrRE
    Else
        ArrMaxColumn = VBA.Val(arr2D(LBound(arr2D, 1), ColumnIndexArr))
        For i = LBound(arr2D, 1) + 1 To UBound(arr2D, 1)
            v = VBA.Val(arr2D(i, ColumnIndexArr))
            If ArrMaxColumn < v Then ArrMaxColumn = v   '����
        Next
    End If
End Function

'���鰴�������ֵ
Public Function ArrMaxRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant
    Dim i As Long, j As Long, v
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            arrRE(j) = VBA.Val(arr2D(RowIndex, l))
            For i = l + 1 To u
                v = VBA.Val(arr2D(RowIndex, i))
                If arrRE(j) < v Then arrRE(j) = v      '����
            Next
        Next
        ArrMaxRow = arrRE
    Else
        ArrMaxRow = VBA.Val(arr2D(RowIndexArr, LBound(arr2D, 2)))
        For i = LBound(arr2D, 2) + 1 To UBound(arr2D, 2)
            v = VBA.Val(arr2D(RowIndexArr, i))
            If ArrMaxRow < v Then ArrMaxRow = v      '����
        Next
    End If
End Function


'���鰴������Сֵ
Public Function ArrMinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1) As Variant
    Dim i As Long, j As Long, v
    Dim arrRE()
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            arrRE(j) = VBA.Val(arr2D(l, ColumnIndex))
            For i = l + 1 To u
                v = VBA.Val(arr2D(i, ColumnIndex))
                If arrRE(j) > v Then arrRE(j) = v   '����
            Next
        Next
        ArrMinColumn = arrRE
    Else
        ArrMinColumn = VBA.Val(arr2D(LBound(arr2D, 1), ColumnIndexArr))
        For i = LBound(arr2D, 1) + 1 To UBound(arr2D, 1)
            v = VBA.Val(arr2D(i, ColumnIndexArr))
            If ArrMinColumn > v Then ArrMinColumn = v   '����
        Next
    End If
End Function

'���鰴������Сֵ
Public Function ArrMinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1) As Variant
    Dim i As Long, j As Long, v
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            arrRE(j) = VBA.Val(arr2D(RowIndex, l))
            For i = l + 1 To u
                v = VBA.Val(arr2D(RowIndex, i))
                If arrRE(j) > v Then arrRE(j) = v       '����
            Next
        Next
        ArrMinRow = arrRE
    Else
        ArrMinRow = VBA.Val(arr2D(RowIndexArr, LBound(arr2D, 2)))
        For i = LBound(arr2D, 2) + 1 To UBound(arr2D, 2)
            v = VBA.Val(arr2D(RowIndexArr, i))
            If ArrMinRow > v Then ArrMinRow = v       '����
        Next
    End If
End Function

'���鰴��ƴ���ַ���
Public Function ArrJoinColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
    Dim i As Long, j As Long, s As String
    Dim arrRE()
    s = ""
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            StringBuilder_
            For i = l To u
                If OmittedEmpty = False Then
                    StringBuilder_ Delimiter & arr2D(i, ColumnIndex)  '����
                Else
                    If arr2D(i, ColumnIndex) <> "" Then
                        StringBuilder_ Delimiter & arr2D(i, ColumnIndex)   '����
                    End If
                End If
            Next
            arrRE(j) = Mid(StringBuilder_, Len(Delimiter) + 1)
        Next
        ArrJoinColumn = arrRE
    Else
        StringBuilder_
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            If OmittedEmpty = False Then
                StringBuilder_ Delimiter & arr2D(i, ColumnIndexArr)  '����
            Else
                If arr2D(i, ColumnIndexArr) <> "" Then
                    StringBuilder_ Delimiter & arr2D(i, ColumnIndexArr)   '����
                End If
            End If
        Next
        ArrJoinColumn = Mid(StringBuilder_, Len(Delimiter) + 1)
    End If
End Function

'���鰴��ƴ���ַ���
Public Function ArrJoinRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional ByRef Delimiter = "", Optional OmittedEmpty As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            StringBuilder_
            For i = l To u
                If OmittedEmpty = False Then
                    StringBuilder_ Delimiter & arr2D(RowIndex, i)  '����
                Else
                    If arr2D(RowIndex, i) <> "" Then
                        StringBuilder_ Delimiter & arr2D(RowIndex, i)   '����
                    End If
                End If
            Next
            arrRE(j) = Mid(StringBuilder_, Len(Delimiter) + 1)
        Next
        ArrJoinRow = arrRE
    Else
        StringBuilder_
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            If OmittedEmpty = False Then
                StringBuilder_ Delimiter & arr2D(RowIndexArr, i)  '����
            Else
                If arr2D(RowIndexArr, i) <> "" Then
                    StringBuilder_ Delimiter & arr2D(RowIndexArr, i) '����
                End If
            End If
        Next
        ArrJoinRow = Mid(StringBuilder_, Len(Delimiter) + 1)
    End If
End Function

'���鰴�м���ǿ�ֵ����
Public Function ArrCountNoEmptyColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            arrRE(j) = 0
            For i = l To u
                If arr2D(i, ColumnIndex) <> EmptyContent Then arrRE(j) = arrRE(j) + 1  '����
            Next
        Next
        ArrCountNoEmptyColumn = arrRE
    Else
        ArrCountNoEmptyColumn = 0
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            If arr2D(i, ColumnIndexArr) <> EmptyContent Then
                ArrCountNoEmptyColumn = ArrCountNoEmptyColumn + 1    '����
            End If
        Next
    End If
End Function

'���鰴�м���ǿ�ֵ����
Public Function ArrCountNoEmptyRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            arrRE(j) = 0
            For i = l To u
                If arr2D(RowIndex, i) <> EmptyContent Then arrRE(j) = arrRE(j) + 1  '����
            Next
        Next
        ArrCountNoEmptyRow = arrRE
    Else
        ArrCountNoEmptyRow = 0
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            If arr2D(RowIndexArr, i) <> EmptyContent Then
                ArrCountNoEmptyRow = ArrCountNoEmptyRow + 1    '����
            End If
        Next
    End If
End Function

'���鰴�м�����������
Public Function ArrCountClassColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    Dim dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            dic.RemoveAll
            ColumnIndex = ColumnIndexArr(j)
            arrRE(j) = 0
            For i = l To u
                If arr2D(i, ColumnIndex) <> EmptyContent Then
                    If Not dic.Exists(arr2D(i, ColumnIndex)) Then
                        dic.Add arr2D(i, ColumnIndex), i
                    End If
                End If
            Next
            arrRE(j) = dic.Count
        Next
        ArrCountClassColumn = arrRE
    Else
        dic.RemoveAll
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            If arr2D(i, ColumnIndexArr) <> EmptyContent Then
                If Not dic.Exists(arr2D(i, ColumnIndexArr)) Then
                    dic.Add arr2D(i, ColumnIndexArr), i
                End If
            End If
        Next
        ArrCountClassColumn = dic.Count
    End If
End Function

'���鰴�м�����������
Public Function ArrCountClassRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "", Optional CompareMode As CompareMethod = BinaryCompare) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    Dim dic As Object
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            dic.RemoveAll
            RowIndex = RowIndexArr(j)
            arrRE(j) = 0
            For i = l To u
                If arr2D(RowIndex, i) <> EmptyContent Then
                    If Not dic.Exists(arr2D(RowIndex, i)) Then
                        dic.Add arr2D(RowIndex, i), i
                    End If
                End If
            Next
            arrRE(j) = dic.Count
        Next
        ArrCountClassRow = arrRE
    Else
        dic.RemoveAll
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            If arr2D(RowIndexArr, i) <> EmptyContent Then
                If Not dic.Exists(arr2D(RowIndexArr, i)) Then
                    dic.Add arr2D(RowIndexArr, i), i
                End If
            End If
        Next
        ArrCountClassRow = dic.Count
    End If
End Function

'���鰴�м���ƽ��ֵ
Public Function ArrAverageColumn(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        ReDim arrRE(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        ReDim arrREsum(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        ReDim arrRECount(LBound(ColumnIndexArr) To UBound(ColumnIndexArr))
        For j = LBound(ColumnIndexArr) To UBound(ColumnIndexArr)
            ColumnIndex = ColumnIndexArr(j)
            arrRECount(j) = 0
            arrRE(j) = 0
            For i = l To u
                If arr2D(i, ColumnIndex) <> "" Then
                    arrRECount(j) = arrRECount(j) + 1  '����
                    arrREsum(j) = arrREsum(j) + VBA.Val(arr2D(i, ColumnIndex))
                End If
            Next
            If arrRECount(j) <> 0 Then arrRE(j) = RoundEX(arrREsum(j) / arrRECount(j), NumDigitsAfterDecimal)
        Next
        ArrAverageColumn = arrRE
    Else
        Dim REsum, REcount
        REcount = 0
        ArrAverageColumn = 0
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            If arr2D(i, ColumnIndexArr) <> "" Then
                REcount = REcount + 1    '����
                REsum = REsum + VBA.Val(arr2D(i, ColumnIndexArr))
            End If
            
        Next
        If REcount <> 0 Then ArrAverageColumn = RoundEX(REsum / REcount, NumDigitsAfterDecimal)
    End If
End Function

'���鰴�м���ƽ��ֵ
Public Function ArrAverageRow(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional NumDigitsAfterDecimal As Long = 2) As Variant
    Dim i As Long, j As Long
    Dim arrRE()
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex As Long, l As Long, u As Long
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        ReDim arrRE(LBound(RowIndexArr) To UBound(RowIndexArr))
        ReDim arrREsum(LBound(RowIndexArr) To UBound(RowIndexArr))
        ReDim arrRECount(LBound(RowIndexArr) To UBound(RowIndexArr))
        For j = LBound(RowIndexArr) To UBound(RowIndexArr)
            RowIndex = RowIndexArr(j)
            arrRECount(j) = 0
            arrRE(j) = 0
            For i = l To u
                If arr2D(RowIndex, i) <> "" Then
                    arrRECount(j) = arrRECount(j) + 1  '����
                    arrREsum(j) = arrREsum(j) + VBA.Val(arr2D(RowIndex, i))
                End If
            Next
            If arrRECount(j) <> 0 Then arrRE(j) = RoundEX(arrREsum(j) / arrRECount(j), NumDigitsAfterDecimal)
        Next
        ArrAverageRow = arrRE
    Else
        Dim REsum, REcount
        REcount = 0
        ArrAverageRow = 0
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            If arr2D(RowIndexArr, i) <> "" Then
                REcount = REcount + 1    '����
                REsum = REsum + VBA.Val(arr2D(RowIndexArr, i))
            End If
        Next
        If REcount <> 0 Then ArrAverageRow = RoundEX(REsum / REcount, NumDigitsAfterDecimal)
    End If
End Function

'��ֵ�ƶ� ����
Public Function ArrMoveUp(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex
        For Each ColumnIndex In ColumnIndexArr
            k = LBound(arr2D, 1)
            For i = LBound(arr2D, 1) To UBound(arr2D, 1)
                If arr2D(i, ColumnIndex) <> EmptyContent Then
                    If i <> k Then
                        Cover arr2D(k, ColumnIndex), arr2D(i, ColumnIndex)
                        Cover arr2D(i, ColumnIndex), EmptyContent
                    End If
                    k = k + 1
                End If
            Next
        Next
    Else
        k = LBound(arr2D, 1)
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            If arr2D(i, ColumnIndexArr) <> EmptyContent Then
                If i <> k Then
                    Cover arr2D(k, ColumnIndexArr), arr2D(i, ColumnIndexArr)
                    Cover arr2D(i, ColumnIndexArr), EmptyContent
                End If
                k = k + 1
            End If
        Next
    End If
    ArrMoveUp = arr2D
End Function

'��ֵ�ƶ� ����
Public Function ArrMoveDown(ByRef arr2D, Optional ByVal ColumnIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    IndexIsCurrencyToCount_ ColumnIndexArr, LBound(arr2D, 2), UBound(arr2D, 2)
    If IsArray(ColumnIndexArr) Then
        Dim ColumnIndex
        For Each ColumnIndex In ColumnIndexArr
            k = UBound(arr2D, 1)
            For i = UBound(arr2D, 1) To LBound(arr2D, 1) Step -1
                If arr2D(i, ColumnIndex) <> EmptyContent Then
                    If i <> k Then
                        Cover arr2D(k, ColumnIndex), arr2D(i, ColumnIndex)
                        Cover arr2D(i, ColumnIndex), EmptyContent
                    End If
                    k = k - 1
                End If
            Next
        Next
    Else
        k = UBound(arr2D, 1)
        For i = UBound(arr2D, 1) To LBound(arr2D, 1) Step -1
            If arr2D(i, ColumnIndexArr) <> EmptyContent Then
                If i <> k Then
                    Cover arr2D(k, ColumnIndexArr), arr2D(i, ColumnIndexArr)
                    Cover arr2D(i, ColumnIndexArr), EmptyContent
                End If
                k = k - 1
            End If
        Next
    End If
    ArrMoveDown = arr2D
End Function

'��ֵ�ƶ� ����
Public Function ArrMoveLeft(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex
        For Each RowIndex In RowIndexArr
            k = LBound(arr2D, 2)
            For i = LBound(arr2D, 2) To UBound(arr2D, 2)
                If arr2D(RowIndex, i) <> EmptyContent Then
                    If i <> k Then
                        Cover arr2D(RowIndex, k), arr2D(RowIndex, i)
                        Cover arr2D(RowIndex, i), EmptyContent
                    End If
                    k = k + 1
                End If
            Next
        Next
    Else
        k = LBound(arr2D, 2)
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            If arr2D(RowIndexArr, i) <> EmptyContent Then
                If i <> k Then
                    Cover arr2D(RowIndexArr, k), arr2D(RowIndexArr, i)
                    Cover arr2D(RowIndexArr, i), EmptyContent
                End If
                k = k + 1
            End If
        Next
    End If
    ArrMoveLeft = arr2D
End Function

'��ֵ�ƶ� ����
Public Function ArrMoveRight(ByRef arr2D, Optional ByVal RowIndexArr = 1, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    IndexIsCurrencyToCount_ RowIndexArr, LBound(arr2D, 1), UBound(arr2D, 1)
    If IsArray(RowIndexArr) Then
        Dim RowIndex
        For Each RowIndex In RowIndexArr
            k = UBound(arr2D, 2)
            For i = UBound(arr2D, 2) To LBound(arr2D, 2) Step -1
                If arr2D(RowIndex, i) <> EmptyContent Then
                    If i <> k Then
                        Cover arr2D(RowIndex, k), arr2D(RowIndex, i)
                        Cover arr2D(RowIndex, i), EmptyContent
                    End If
                    k = k - 1
                End If
            Next
        Next
    Else
        k = UBound(arr2D, 2)
        For i = UBound(arr2D, 2) To LBound(arr2D, 2) Step -1
            If arr2D(RowIndexArr, i) <> EmptyContent Then
                If i <> k Then
                    Cover arr2D(RowIndexArr, k), arr2D(RowIndexArr, i)
                    Cover arr2D(RowIndexArr, i), EmptyContent
                End If
                k = k - 1
            End If
        Next
    End If
    ArrMoveRight = arr2D
End Function

'��ֵ�ƶ� һά���� ����
Public Function ArrMove(ByRef arr1D, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    k = LBound(arr1D)
    For i = LBound(arr1D) To UBound(arr1D)
        If arr1D(i) <> EmptyContent Then
            If i <> k Then
                Cover arr1D(k), arr1D(i)
                Cover arr1D(i), EmptyContent
            End If
            k = k + 1
        End If
    Next
    ArrMove = arr1D
End Function

'��ֵ�ƶ� һά���� ����
Public Function ArrMoveRev(ByRef arr1D, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long
    k = UBound(arr1D)
    For i = UBound(arr1D) To LBound(arr1D) Step -1
        If arr1D(i) <> EmptyContent Then
            If i <> k Then
                Cover arr1D(k), arr1D(i)
                Cover arr1D(i), EmptyContent
            End If
            k = k - 1
        End If
    Next
    ArrMoveRev = arr1D
End Function

'��ֵ�ƶ� һά���� ���� ��������
Public Function ArrMove_Index(ByRef arr1D, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long, arrRE(), n As Long
    ReDim arrRE(LBound(arr1D) To UBound(arr1D))
    For i = LBound(arr1D) To UBound(arr1D)
        arrRE(i) = i
    Next
    k = LBound(arrRE)
    For i = LBound(arrRE) To UBound(arrRE)
        If arr1D(arrRE(i)) <> EmptyContent Then
            If i <> k Then
                n = arrRE(k)
                arrRE(k) = arrRE(i)
                arrRE(i) = n
            End If
            k = k + 1
        End If
    Next
    ArrMove_Index = arrRE
End Function

'��ֵ�ƶ� һά���� ���� ��������
Public Function ArrMoveRev_Index(ByRef arr1D, Optional EmptyContent = "") As Variant
    Dim i As Long, k As Long, arrRE(), n As Long
    ReDim arrRE(LBound(arr1D) To UBound(arr1D))
    For i = LBound(arr1D) To UBound(arr1D)
        arrRE(i) = i
    Next
    k = UBound(arrRE)
    For i = UBound(arr1D) To LBound(arr1D) Step -1
        If arr1D(arrRE(i)) <> EmptyContent Then
            If i <> k Then
                n = arrRE(k)
                arrRE(k) = arrRE(i)
                arrRE(i) = n
            End If
            k = k - 1
        End If
    Next
    ArrMoveRev_Index = arrRE
End Function


'������� ���� Index������������ͷ
Public Function ArrScroll(ByRef arr, Index) As Variant
    ArrScroll = ArrFromIndex(arr, ArrScroll_Index(arr, Index))
End Function

'������� ���� Index����������ĩβ
Public Function ArrScrollRev(ByRef arr, Index) As Variant
    ArrScrollRev = ArrFromIndex(arr, ArrScrollRev_Index(arr, Index))
End Function

'������� ���� Index������������ͷ ��������
Public Function ArrScroll_Index(ByRef arr, ByVal Index) As Variant
    Dim i As Long, k As Long, v, arrRE()
    IndexIsCurrencyToCount_ Index, LBound(arr), UBound(arr)
    ReDim arrRE(LBound(arr) To UBound(arr))
    k = LBound(arr)
    For i = Index To UBound(arr)
        arrRE(k) = i
        k = k + 1
    Next
    For i = LBound(arr) To Index - 1
        arrRE(k) = i
        k = k + 1
    Next
    ArrScroll_Index = arrRE
End Function

'������� ���� Index����������ĩβ ��������
Public Function ArrScrollRev_Index(ByRef arr, ByVal Index) As Variant
    Dim i As Long, k As Long, v, arrRE()
    IndexIsCurrencyToCount_ Index, LBound(arr), UBound(arr)
    ReDim arrRE(LBound(arr) To UBound(arr))
    k = LBound(arr)
    For i = Index + 1 To UBound(arr)
        arrRE(k) = i
        k = k + 1
    Next
    For i = LBound(arr) To Index
        arrRE(k) = i
        k = k + 1
    Next
    ArrScrollRev_Index = arrRE
End Function

'��ά�����й��� ���� Index������������ͷ
Public Function ArrScrollColumn(ByRef arr2D, Index) As Variant
    ArrScrollColumn = ArrGetColumns(arr2D, ArrScrollColumn_Index(arr2D, Index))
End Function

'��ά�����й��� ���� Index����������ĩβ
Public Function ArrScrollColumnRev(ByRef arr2D, Index) As Variant
    ArrScrollColumnRev = ArrGetColumns(arr2D, ArrScrollColumnRev_Index(arr2D, Index))
End Function

'��ά�����й���  ���� Index������������ͷ ��������
Public Function ArrScrollColumn_Index(ByRef arr2D, ByVal Index) As Variant
    Dim i As Long, k As Long, v, arrRE()
    IndexIsCurrencyToCount_ Index, LBound(arr2D, 2), UBound(arr2D, 2)
    ReDim arrRE(LBound(arr2D, 2) To UBound(arr2D, 2))
    k = LBound(arr2D, 2)
    For i = Index To UBound(arr2D, 2)
        arrRE(k) = i
        k = k + 1
    Next
    For i = LBound(arr2D, 2) To Index - 1
        arrRE(k) = i
        k = k + 1
    Next
    ArrScrollColumn_Index = arrRE
End Function

'��ά�����й��� ���� Index����������ĩβ ��������
Public Function ArrScrollColumnRev_Index(ByRef arr2D, ByVal Index) As Variant
    Dim i As Long, k As Long, v, arrRE()
    IndexIsCurrencyToCount_ Index, LBound(arr2D, 2), UBound(arr2D, 2)
    ReDim arrRE(LBound(arr2D, 2) To UBound(arr2D, 2))
    k = LBound(arr2D, 2)
    For i = Index + 1 To UBound(arr2D, 2)
        arrRE(k) = i
        k = k + 1
    Next
    For i = LBound(arr2D, 2) To Index
        arrRE(k) = i
        k = k + 1
    Next
    ArrScrollColumnRev_Index = arrRE
End Function

'���  arr һά���� r��ȡ����
Public Function ArrCombinCon(arr, r) As Variant
    Dim arrOri, arrRst, arrOri2, arrRst2, rw&, i&, j&, k&, st&, en&, n&, l&, M&
    ReDim arrOri(1 To 1, 0 To 0)
    ReDim arrOri2(1 To 1)
    l = LBound(arr, 1)
    en = UBound(arr, 1) - l + 1 - r
    arrOri2(1) = l
    For i = 1 To r
        n = n + 1: rw = 1
        ReDim arrRst(1 To Application.WorksheetFunction.Combin(en + 1, n), 1 To i)
        ReDim arrRst2(1 To UBound(arrRst, 1))
        For j = 1 To UBound(arrOri2, 1)
            st = arrOri2(j)
            For k = st To en + l
                For M = 1 To i - 1
                    arrRst(rw, M) = arrOri(j, M)
                Next
                arrRst(rw, i) = arr(k)
                arrRst2(rw) = k + 1
                rw = rw + 1
            Next k
        Next j
        arrOri = arrRst
        arrOri2 = arrRst2
        en = en + 1
    Next i
    ArrCombinCon = arrRst
End Function

'����  arr һά���� r��ȡ����
Public Function ArrPermutCon(arr, r) As Variant
    Dim arrOri, arrRst, arrOri2, arrRst2, rw&, i&, j&, k&, en&, n&, l&, u&, M&, arrb() As Boolean
    ReDim arrOri(1 To 1)
    ReDim arrOri2(1 To 1)
    l = LBound(arr, 1): u = UBound(arr, 1)
    en = UBound(arr, 1) - l + 1
    ReDim arrb(l To u)
    arrOri2(1) = arrb
    For i = 1 To r
        n = n + 1: rw = 1
        ReDim arrRst(1 To Application.WorksheetFunction.Permut(en, n), 1 To i)
        ReDim arrRst2(1 To UBound(arrRst, 1))
        For j = 1 To UBound(arrOri2, 1)
            For k = l To u
                If arrOri2(j)(k) = False Then
                    For M = 1 To i - 1
                        arrRst(rw, M) = arrOri(j, M)
                    Next
                    arrRst(rw, i) = arr(k)
                    arrRst2(rw) = arrOri2(j)
                    arrRst2(rw)(k) = True
                    rw = rw + 1
                End If
            Next k
        Next j
        arrOri = arrRst
        arrOri2 = arrRst2
    Next i
    ArrPermutCon = arrRst
End Function










'����-------------------------------------------------------------------------------------------------------------------------------------

''����ӷ�����  ��������ȡֵд�� ��������������д�����
'Public Function Matrix_Add(ParamArray Calculates()) As Variant
'    Dim i As Long, j As Long, n As Long, v
'    Dim arr, arrre()
'    Dim maxRowCount As Long, maxColumnCount As Long
'    Dim RowCountRE As Long, ColumnCountRE As Long
'    maxRowCount = 1: maxColumnCount = 1
'    For n = LBound(Calculates) To UBound(Calculates)
'        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
'        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
'        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
'    Next
'    ReDim arrre(1 To maxRowCount, 1 To maxColumnCount)
'    For n = LBound(Calculates) To UBound(Calculates)
'        ArrGetValueCache_ WriteArr:=True, arr:=Calculates(n), EmptyContent:=0
'        For i = 1 To maxRowCount
'            For j = 1 To maxColumnCount
'                arrre(i, j) = arrre(i, j) + ArrGetValueCache_(i, j)
'            Next
'        Next
'    Next
'    Matrix_Add = arrre
'End Function

''����IF����  ��������ȡֵд�� ��������������д�����
'Public Function Matrix_IF(Expression, TruePart, FalsePart) As Variant
'    Dim i As Long, j As Long, n As Long, v
'    Dim arr, arrre()
'    Dim maxRowCount As Long, maxColumnCount As Long
'    Dim RowCountRE As Long, ColumnCountRE As Long
'    maxRowCount = 1: maxColumnCount = 1
'
'    ArrCountRowAndColumn Expression, RowCountRE, ColumnCountRE
'    If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
'    If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
'
'    ArrCountRowAndColumn TruePart, RowCountRE, ColumnCountRE
'    If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
'    If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
'
'    ArrCountRowAndColumn FalsePart, RowCountRE, ColumnCountRE
'    If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
'    If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
'
'    ReDim arrre(1 To maxRowCount, 1 To maxColumnCount)
'    ArrGetValueCache_ WriteArr:=True, arr:=Expression, EmptyContent:=False
'    ArrGetValueCache1_ WriteArr:=True, arr:=TruePart
'    ArrGetValueCache2_ WriteArr:=True, arr:=FalsePart
'
'    For i = 1 To maxRowCount
'        For j = 1 To maxColumnCount
'            If ArrGetValueCache_(i, j) Then
'                Cover arrre(i, j), ArrGetValueCache1_(i, j)
'            Else
'                Cover arrre(i, j), ArrGetValueCache2_(i, j)
'            End If
'        Next
'    Next
'    Matrix_IF = arrre
'End Function

'���Calculates����������������� ��maxRowCount maxColumnCount���շ���ֵ
Private Sub ArrMaxCountRowColumn_(ByRef maxRowCount As Long, ByRef maxColumnCount As Long, ParamArray Calculates())
    Dim i As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    For i = LBound(Calculates) To UBound(Calculates)
        ArrCountRowAndColumn Calculates(i), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
End Sub

'����ӷ�����
Public Function Matrix_Add(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, 0)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For n = l To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) + CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Add = arrRE
End Function

'�����������
Public Function Matrix_Sub(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, 0)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) - CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Sub = arrRE
End Function

'����˷�����
Public Function Matrix_Multipli(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, 0)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) * CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Multipli = arrRE
End Function

'�����������
Public Function Matrix_Division(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, 0)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) / CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Division = arrRE
End Function

'����˷�����
Public Function Matrix_Power(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, 0)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) ^ CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Power = arrRE
End Function

'�������Ӽ���
Public Function Matrix_Join(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, "")
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) & CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Join = arrRE
End Function

'����Ƚϵ���
Public Function Matrix_Comp_Equal(ByRef arr, ByRef arr2) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arr2
    
    Dim CalculatesRE, Calculates2RE
    CalculatesRE = ArrSizeExpansionEx(arr, maxRowCount, maxColumnCount)
    Calculates2RE = ArrSizeExpansionEx(arr2, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(i, j) = Calculates2RE(i, j)
        Next
    Next
    Matrix_Comp_Equal = arrRE
End Function

'����Ƚϲ�����
Public Function Matrix_Comp_NotEqual(ByRef arr, ByRef arr2) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arr2
    
    Dim CalculatesRE, Calculates2RE
    CalculatesRE = ArrSizeExpansionEx(arr, maxRowCount, maxColumnCount)
    Calculates2RE = ArrSizeExpansionEx(arr2, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(i, j) <> Calculates2RE(i, j)
        Next
    Next
    Matrix_Comp_NotEqual = arrRE
End Function

'����Ƚϴ�С
Public Function Matrix_Comp_Size(ByRef arr_Large, ByRef arr_Small) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr_Large, arr_Small
    
    Dim arr_LargeRE, arr_SmallRE
    arr_LargeRE = ArrSizeExpansionEx(arr_Large, maxRowCount, maxColumnCount)
    arr_SmallRE = ArrSizeExpansionEx(arr_Small, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = arr_LargeRE(i, j) > arr_SmallRE(i, j)
        Next
    Next
    Matrix_Comp_Size = arrRE
End Function

'����Ƚϴ�С��������
Public Function Matrix_Comp_SizeEqual(ByRef arr_Large, ByRef arr_Small) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr_Large, arr_Small
    
    Dim arr_LargeRE, arr_SmallRE
    arr_LargeRE = ArrSizeExpansionEx(arr_Large, maxRowCount, maxColumnCount)
    arr_SmallRE = ArrSizeExpansionEx(arr_Small, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = arr_LargeRE(i, j) >= arr_SmallRE(i, j)
        Next
    Next
    Matrix_Comp_SizeEqual = arrRE
End Function

'��������Ƚϼ��� �ڲ�
Public Function Matrix_Comp_RangeInside(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arrL, arrR
    
    Dim arr1RE, arrLRE, arrRRE
    arr1RE = ArrSizeExpansionEx(arr1RE, maxRowCount, maxColumnCount)
    arrLRE = ArrSizeExpansionEx(arrLRE, maxRowCount, maxColumnCount)
    arrRRE = ArrSizeExpansionEx(arrRRE, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = NumberRangeInside(arr1RE(i, j), arrLRE(i, j), arrRRE(i, j), NumberRangeRule)
        Next
    Next
    Matrix_Comp_RangeInside = arrRE
End Function

'��������Ƚϼ��� �ⲿ
Public Function Matrix_Comp_RangeExternal(ByRef arr, ByRef arrL, ByRef arrR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arrL, arrR
    
    Dim arr1RE, arrLRE, arrRRE
    arr1RE = ArrSizeExpansionEx(arr1RE, maxRowCount, maxColumnCount)
    arrLRE = ArrSizeExpansionEx(arrLRE, maxRowCount, maxColumnCount)
    arrRRE = ArrSizeExpansionEx(arrRRE, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = NumberRangeExternal(arr1RE(i, j), arrLRE(i, j), arrRRE(i, j), NumberRangeRule)
        Next
    Next
    Matrix_Comp_RangeExternal = arrRE
End Function

'����Ƚ�Like
Public Function Matrix_Comp_Like(ByRef arr, ByRef arr2) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arr2
    
    Dim CalculatesRE, Calculates2RE
    CalculatesRE = ArrSizeExpansionEx(arr, maxRowCount, maxColumnCount)
    Calculates2RE = ArrSizeExpansionEx(arr2, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(i, j) Like Calculates2RE(i, j)
        Next
    Next
    Matrix_Comp_Like = arrRE
End Function

'����Ƚ�Not Like
Public Function Matrix_Comp_NotLike(ByRef arr, ByRef arr2) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr, arr2
    
    Dim CalculatesRE, Calculates2RE
    CalculatesRE = ArrSizeExpansionEx(arr, maxRowCount, maxColumnCount)
    Calculates2RE = ArrSizeExpansionEx(arr2, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = Not CalculatesRE(i, j) Like Calculates2RE(i, j)
        Next
    Next
    Matrix_Comp_NotLike = arrRE
End Function

'���󲼶��Ҽ���
Public Function Matrix_Boolea_And(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, False)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) And CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Boolea_And = arrRE
End Function

'���󲼶������
Public Function Matrix_Boolea_Or(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = l To u
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, False)
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = CalculatesRE(l)(i, j)
        Next
    Next
    For n = l + 1 To u
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = arrRE(i, j) Or CalculatesRE(n)(i, j)
            Next
        Next
    Next
    Matrix_Boolea_Or = arrRE
End Function

'���󲼶��Ǽ���
Public Function Matrix_Boolea_Not(ByRef arr) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, arr
    
    Dim CalculatesRE
    CalculatesRE = ArrSizeExpansionEx(arr, maxRowCount, maxColumnCount, True)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = Not CalculatesRE(i, j)
        Next
    Next
    Matrix_Boolea_Not = arrRE
End Function
 
'����IF
Public Function Matrix_IF(Expression, TruePart, FalsePart) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, Expression, TruePart, FalsePart
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    Dim ExpressionRE, TruePartRE, FalsePartRE
    ExpressionRE = ArrSizeExpansionEx(Expression, maxRowCount, maxColumnCount, False)
    TruePartRE = ArrSizeExpansionEx(TruePart, maxRowCount, maxColumnCount)
    FalsePartRE = ArrSizeExpansionEx(FalsePart, maxRowCount, maxColumnCount)
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            If ExpressionRE(i, j) Then
                Cover arrRE(i, j), TruePartRE(i, j)
            Else
                Cover arrRE(i, j), FalsePartRE(i, j)
            End If
        Next
    Next
    Matrix_IF = arrRE
End Function

'����IFs
Public Function Matrix_IFs(ParamArray Calculates()) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    Dim RowCountRE As Long, ColumnCountRE As Long
    maxRowCount = 1: maxColumnCount = 1
    For n = LBound(Calculates) To UBound(Calculates)
        ArrCountRowAndColumn Calculates(n), RowCountRE, ColumnCountRE
        If maxRowCount < RowCountRE Then maxRowCount = RowCountRE
        If maxColumnCount < ColumnCountRE Then maxColumnCount = ColumnCountRE
    Next
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    Dim l As Long, u As Long
    l = LBound(Calculates): u = UBound(Calculates)
    Dim CalculatesRE
    ReDim CalculatesRE(l To u)
    For n = l To u - 1 Step 2
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount, False)
    Next
    For n = l + 1 To u Step 2
        CalculatesRE(n) = ArrSizeExpansionEx(Calculates(n), maxRowCount, maxColumnCount)
    Next
    If IsOdd(u - l + 1) Then CalculatesRE(u) = ArrSizeExpansionEx(Calculates(u), maxRowCount, maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            For n = l To u - 1 Step 2
                If CalculatesRE(n)(i, j) Then
                    Cover arrRE(i, j), CalculatesRE(n + 1)(i, j)
                    Exit For
                ElseIf n = u - 2 Then
                    Cover arrRE(i, j), CalculatesRE(u)(i, j)
                    Exit For
                End If
            Next
        Next
    Next
    Matrix_IFs = arrRE
End Function

'����Mid ���������String1, Start, Length
Public Function Matrix_Str_Mid(String1, Start, Optional Length) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    If IsMissing(Length) Then
        ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, String1, Start
    Else
        ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, String1, Start, Length
    End If
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    Dim String1RE, StartRE, LengthRE
    String1RE = ArrSizeExpansionEx(String1, maxRowCount, maxColumnCount, "")
    StartRE = ArrSizeExpansionEx(Start, maxRowCount, maxColumnCount, 1)
    
    If IsMissing(Length) Then
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = VBA.Mid(String1RE(i, j), StartRE(i, j))
            Next
        Next
    Else
        LengthRE = ArrSizeExpansionEx(Length, maxRowCount, maxColumnCount)
        For i = 1 To maxRowCount
            For j = 1 To maxColumnCount
                arrRE(i, j) = VBA.Mid(String1RE(i, j), StartRE(i, j), LengthRE(i, j))
            Next
        Next
    End If
    Matrix_Str_Mid = arrRE
End Function

'����Left ���������String1, Length
Public Function Matrix_Str_Left(String1, Length) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, String1, Length
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    
    Dim String1RE, LengthRE
    String1RE = ArrSizeExpansionEx(String1, maxRowCount, maxColumnCount, "")
    LengthRE = ArrSizeExpansionEx(Length, maxRowCount, maxColumnCount)
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.Left(String1RE(i, j), LengthRE(i, j))
        Next
    Next
    Matrix_Str_Left = arrRE
End Function

'����Right ���������String1, Length
Public Function Matrix_Str_Right(String1, Length) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, String1, Length
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    
    Dim String1RE, LengthRE
    String1RE = ArrSizeExpansionEx(String1, maxRowCount, maxColumnCount, "")
    LengthRE = ArrSizeExpansionEx(Length, maxRowCount, maxColumnCount)
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.Right(String1RE(i, j), LengthRE(i, j))
        Next
    Next
    Matrix_Str_Right = arrRE
End Function

'����InStr ���������StringLarge, StringSmall, Start
Public Function Matrix_Str_InStr(StringLarge, StringSmall, Optional Start = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, StringLarge, StringSmall, Start
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)

    Dim StringLargeRE, StringSmallRE, StartRE
    StringLargeRE = ArrSizeExpansionEx(StringLarge, maxRowCount, maxColumnCount, "")
    StringSmallRE = ArrSizeExpansionEx(StringSmall, maxRowCount, maxColumnCount, "")
    StartRE = ArrSizeExpansionEx(Start, maxRowCount, maxColumnCount, 1)
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.InStr(StartRE(i, j), StringLargeRE(i, j), StringSmallRE(i, j), Compare)
        Next
    Next
    Matrix_Str_InStr = arrRE
End Function

'����InStr ���������StringLarge, StringSmall, Start
Public Function Matrix_Str_InStrRev(StringLarge, StringSmall, Optional Start = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, StringLarge, StringSmall, Start
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)

    Dim StringLargeRE, StringSmallRE, StartRE
    StringLargeRE = ArrSizeExpansionEx(StringLarge, maxRowCount, maxColumnCount, "")
    StringSmallRE = ArrSizeExpansionEx(StringSmall, maxRowCount, maxColumnCount, "")
    StartRE = ArrSizeExpansionEx(Start, maxRowCount, maxColumnCount, -1)
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.InStrRev(StringLargeRE(i, j), StringSmallRE(i, j), StartRE(i, j), Compare)
        Next
    Next
    Matrix_Str_InStrRev = arrRE
End Function

'����Len ���������String1
Public Function Matrix_Str_Len(ByRef String1) As Variant
    Dim i As Long, j As Long, n As Long
    Dim arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, String1
    
    Dim String1RE
    String1RE = ArrSizeExpansionEx(String1, maxRowCount, maxColumnCount)
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.Len(String1RE(i, j))
        Next
    Next
    Matrix_Str_Len = arrRE
End Function

'�����滻 ���������Expression, Find, Replace
Public Function Matrix_Str_Replace(Expression, Find, Replace, Optional Start = 1, Optional Count = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, Expression, Find, Replace
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    Dim ExpressionRE, FindRE, ReplaceRE
    ExpressionRE = ArrSizeExpansionEx(Expression, maxRowCount, maxColumnCount, "")
    FindRE = ArrSizeExpansionEx(Find, maxRowCount, maxColumnCount, "")
    ReplaceRE = ArrSizeExpansionEx(Replace, maxRowCount, maxColumnCount, "")
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.Replace(ExpressionRE(i, j), FindRE(i, j), ReplaceRE(i, j), Start, Count, Compare)
        Next
    Next
    Matrix_Str_Replace = arrRE
End Function

'�������ڼ�� ����DateDiff ���������Interval, Date1, Date2
Public Function Matrix_DateSub(Interval, Date1, Date2) As Variant
    Dim i As Long, j As Long, n As Long, v
    Dim arr, arrRE()
    Dim maxRowCount As Long, maxColumnCount As Long
    maxRowCount = 1: maxColumnCount = 1
    ArrMaxCountRowColumn_ maxRowCount, maxColumnCount, Interval, Date1, Date2
    
    ReDim arrRE(1 To maxRowCount, 1 To maxColumnCount)
    Dim IntervalRE, Date1RE, Date2RE
    IntervalRE = ArrSizeExpansionEx(Interval, maxRowCount, maxColumnCount, "")
    Date1RE = ArrSizeExpansionEx(Date1, maxRowCount, maxColumnCount, "")
    Date2RE = ArrSizeExpansionEx(Date2, maxRowCount, maxColumnCount, "")
    
    For i = 1 To maxRowCount
        For j = 1 To maxColumnCount
            arrRE(i, j) = VBA.DateDiff(IntervalRE(i, j), Date1RE(i, j), Date2RE(i, j), vbMonday)
        Next
    Next
    Matrix_DateSub = arrRE
End Function


















'�ַ���-----------------------------------------------------------------------------------------------------------------------------------
'��������ӣ���������ȡֵ���ʼ��
Public Function StringBuilder(Optional ByRef s) As Variant
    Static str As String, i As Long
    Const init = 20
    If IsMissing(s) And IsError(s) Then
        If i > 1 Then StringBuilder = Left(str, i - 1) Else StringBuilder = ""
        i = 0
        str = ""
        Exit Function
    End If
    If i = 0 Then
        str = VBA.Space$(init)
        i = 1
    ElseIf i + Len(s) > Len(str) Then
        Dim ds As String
        ds = str
        str = VBA.Space$(Len(str) * 2 + Len(s))
        LSet str = ds
    End If
    Mid(str, i) = s
    i = i + Len(s)
    StringBuilder = i - 1
End Function

Public Function StringBuilder1(Optional ByRef s) As Variant
    Static str As String, i As Long
    Const init = 20
    If IsMissing(s) And IsError(s) Then
        If i > 1 Then StringBuilder1 = Left(str, i - 1) Else StringBuilder1 = ""
        i = 0
        str = ""
        Exit Function
    End If
    If i = 0 Then
        str = VBA.Space$(init)
        i = 1
    ElseIf i + Len(s) > Len(str) Then
        Dim ds As String
        ds = str
        str = VBA.Space$(Len(str) * 2 + Len(s))
        LSet str = ds
    End If
    Mid(str, i) = s
    i = i + Len(s)
    StringBuilder1 = i - 1
End Function

Public Function StringBuilder2(Optional ByRef s) As Variant
    Static str As String, i As Long
    Const init = 20
    If IsMissing(s) And IsError(s) Then
        If i > 1 Then StringBuilder2 = Left(str, i - 1) Else StringBuilder2 = ""
        i = 0
        str = ""
        Exit Function
    End If
    If i = 0 Then
        str = VBA.Space$(init)
        i = 1
    ElseIf i + Len(s) > Len(str) Then
        Dim ds As String
        ds = str
        str = VBA.Space$(Len(str) * 2 + Len(s))
        LSet str = ds
    End If
    Mid(str, i) = s
    i = i + Len(s)
    StringBuilder2 = i - 1
End Function

Public Function StringBuilder3(Optional ByRef s) As Variant
    Static str As String, i As Long
    Const init = 20
    If IsMissing(s) And IsError(s) Then
        If i > 1 Then StringBuilder3 = Left(str, i - 1) Else StringBuilder3 = ""
        i = 0
        str = ""
        Exit Function
    End If
    If i = 0 Then
        str = VBA.Space$(init)
        i = 1
    ElseIf i + Len(s) > Len(str) Then
        Dim ds As String
        ds = str
        str = VBA.Space$(Len(str) * 2 + Len(s))
        LSet str = ds
    End If
    Mid(str, i) = s
    i = i + Len(s)
    StringBuilder3 = i - 1
End Function
 
'�ڲ�StringBuilder
Private Function StringBuilder_(Optional ByRef s) As Variant
    Static str As String, i As Long
    Const init = 20
    If IsMissing(s) And IsError(s) Then
        If i > 1 Then StringBuilder_ = Left(str, i - 1) Else StringBuilder_ = ""
        i = 0
        str = ""
        Exit Function
    End If
    If i = 0 Then
        str = VBA.Space$(init)
        i = 1
    ElseIf i + Len(s) > Len(str) Then
        Dim ds As String
        ds = str
        str = VBA.Space$(Len(str) * 2 + Len(s))
        LSet str = ds
    End If
    Mid(str, i) = s
    i = i + Len(s)
    StringBuilder_ = i - 1
End Function

'��ά����ƴ��
Public Function StrJoinArr2D(ByRef arr2D, Optional Delimiter = "", Optional OmittedEmpty As Boolean = True, Optional RowFirst As Boolean = True) As String
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    StringBuilder_
    If RowFirst Then
        l = LBound(arr2D, 2): u = UBound(arr2D, 2)
        For i = LBound(arr2D, 1) To UBound(arr2D, 1)
            For j = l To u
                If OmittedEmpty = False Then
                    StringBuilder_ Delimiter & arr2D(i, j)
                Else
                    If arr2D(i, j) <> "" Then
                        StringBuilder_ Delimiter & arr2D(i, j)
                    End If
                End If
            Next
        Next
    Else
        l = LBound(arr2D, 1): u = UBound(arr2D, 1)
        For j = LBound(arr2D, 2) To UBound(arr2D, 2)
            For i = l To u
                If OmittedEmpty = False Then
                    StringBuilder_ Delimiter & arr2D(i, j)
                Else
                    If arr2D(i, j) <> "" Then
                        StringBuilder_ Delimiter & arr2D(i, j)
                    End If
                End If
            Next
        Next
    End If
    StrJoinArr2D = Mid(StringBuilder_, Len(Delimiter) + 1)
End Function
        
'���齻��ƴ��
Public Function StrJoin_ArrDelimiter(ByRef arr, ParamArray ArrDelimiter()) As String
    Dim v, arrRE, i As Long, u As Long
    arrRE = ArrFlatten(ArrDelimiter)
    StringBuilder_
    i = 1: u = UBound(arrRE)
    For Each v In arr
        StringBuilder_ v
        If i <= u Then
            StringBuilder_ arrRE(i)
            i = i + 1
        End If
    Next
    StrJoin_ArrDelimiter = StringBuilder_
End Function

'Likeƥ��
Public Function StrLike(str1, LikeStr) As Boolean
    StrLike = str1 Like LikeStr
End Function

'֧�ָ�Length��Left
Public Function StrLeft(String1, Length) As String
    If Length > 0 Then
        StrLeft = VBA.Left$(String1, Length)
    Else
        StrLeft = VBA.Left$(String1, VBA.Len(String1) + Length)
    End If
End Function

'֧�ָ�Length��Right
Public Function StrRight(String1, Length) As String
    If Length > 0 Then
        StrRight = VBA.Right$(String1, Length)
    Else
        StrRight = VBA.Right$(String1, VBA.Len(String1) + Length)
    End If
End Function

'֧�ָ�Start��Length��Mid
Public Function StrMid(String1, ByVal Start, Optional ByVal Length) As String
    If Start < 0 Then
        Start = VBA.Len(String1) + Start + 1
    End If
    If IsMissing(Length) Then
        StrMid = VBA.Mid$(String1, Start): Exit Function
    ElseIf Length < 0 Then
        Start = Start + Length + 1
        Length = -Length
    End If
    If Start > 0 Then
        StrMid = VBA.Mid$(String1, Start, Length)
    Else
        StrMid = VBA.Mid$(String1, 1, Start + Length - 1)
    End If
End Function

'��ʼ����ȡֵ
Public Function StrMidBetween(String1, ByVal Start, Optional ByVal EndIndex = 0) As String
    Dim Length As Long
    Length = VBA.Len(String1)
    If Start < 0 Then
        Start = Length + Start + 1
    End If
    If Start <= 0 Then Start = 1
    If EndIndex < 0 Then
        EndIndex = Length + EndIndex + 1
        If EndIndex < 0 Then EndIndex = 1
    ElseIf EndIndex > Length Or EndIndex = 0 Then
        EndIndex = Length
    End If
    Length = EndIndex - Start + 1
    If Length > 0 Then
        StrMidBetween = VBA.Mid$(String1, Start, Length)
    Else
        StrMidBetween = ""
    End If
End Function

'ȡstr������ݣ��������
Public Function StrGetLeft(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim i As Long
    i = VBA.InStr(1, str1, str2, Compare)
    If i = 0 Then
        StrGetLeft = ""
    Else
        StrGetLeft = VBA.Left$(str1, i - 1)
    End If
End Function
 
'ȡstr������ݣ����Ҳ���
Public Function StrGetLeftRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim i As Long
    i = VBA.InStrRev(str1, str2, -1, Compare)
    If i = 0 Then
        StrGetLeftRev = ""
    Else
        StrGetLeftRev = VBA.Left$(str1, i - 1)
    End If
End Function
 
'ȡstr�ұ����ݣ��������
Public Function StrGetRight(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim i As Long
    i = VBA.InStr(1, str1, str2, Compare)
    If i = 0 Then
        StrGetRight = ""
    Else
        StrGetRight = VBA.Right$(str1, VBA.Len(str1) - i - VBA.Len(str2) + 1)
    End If
End Function
 
'ȡstr�ұ����ݣ����Ҳ���
Public Function StrGetRightRev(str1, str2, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim i As Long
    i = VBA.InStrRev(str1, str2, -1, Compare)
    If i = 0 Then
        StrGetRightRev = ""
    Else
        StrGetRightRev = VBA.Right$(str1, VBA.Len(str1) - i - VBA.Len(str2) + 1)
    End If
End Function

'ȡ����str�м�����,
'LeftLeft�������str�м�����
'RightRight�����ұ�str�м�����
'LeftRight����������str�м����ݣ����Χ��
Public Function StrGetCentre(String1, str1, str2, Optional SearchType As SearchDirection = LeftLeft) As String
    Dim i As Long, j As Long
    Select Case SearchType
    Case LeftLeft
        i = VBA.InStr(1, String1, str1)
        j = VBA.InStr(i + Len(str1), String1, str2)
    Case RightRight
        j = VBA.InStrRev(String1, str2, -1)
        i = VBA.InStrRev(String1, str1, j - 1)
    Case LeftRight
        i = VBA.InStr(1, String1, str1)
        j = VBA.InStrRev(String1, str2, -1)
    End Select
    If i = 0 Or j = 0 Or i >= j Then
        StrGetCentre = ""
    Else
        StrGetCentre = VBA.Mid$(String1, i + VBA.Len(str1), j - i - VBA.Len(str1))
    End If
End Function
 
'��Chrs����ַ�ȥ�������ַ���
Public Function StrTrimChr(String1, Optional Chrs = " ") As String
    Dim l As Long, r As Long, i As Long
    l = 1: r = Len(String1)
    For i = 1 To Len(String1)
        If VBA.InStr(Chrs, VBA.Mid$(String1, i, 1)) > 0 Then
            l = i + 1
        Else
            Exit For
        End If
    Next
    For i = Len(String1) To 1 Step -1
        If VBA.InStr(Chrs, VBA.Mid$(String1, i, 1)) > 0 Then
            r = i - 1
        Else
            Exit For
        End If
    Next
    If r >= l Then
        StrTrimChr = VBA.Mid(String1, l, r - l + 1)
    Else
        StrTrimChr = ""
    End If
End Function

'��Chrs����ַ�ȥ������ַ���
Public Function StrLTrimChr(String1, Optional Chrs = " ") As String
    Dim s As String, l As Long, i As Long
    l = 1
    For i = 1 To Len(String1)
        If VBA.InStr(Chrs, VBA.Mid$(String1, i, 1)) > 0 Then
            l = i + 1
        Else
            Exit For
        End If
    Next
    StrLTrimChr = VBA.Mid(String1, l)
End Function

'��Chrs����ַ�ȥ���Ҷ��ַ���
Public Function StrRTrimChr(String1, Optional Chrs = " ") As String
    Dim r As Long, i As Long
    r = Len(String1)
    For i = Len(String1) To 1 Step -1
        If VBA.InStr(Chrs, VBA.Mid$(String1, i, 1)) > 0 Then
            r = i - 1
        Else
            Exit For
        End If
    Next
    StrRTrimChr = VBA.Left(String1, r)
End Function

'�ظ��ַ���
Public Function StrRepeat(String1, numberOfRepeats) As String
    Dim i As Long
    Dim combinedString As String
    StringBuilder_
    For i = 1 To numberOfRepeats
        StringBuilder_ String1
    Next
    StrRepeat = StringBuilder_
End Function

'�ַ�����ֵݹ��ƴ�� ���������滻�ַ��� �ڲ�ʹ��
Public Function StrReplaces_Split_Recursion_(Expression, Find, Replace, Count, Compare As VbCompareMethod, Index, MaxIndex) As String
    Dim arrRE, i As Long, sRE As String
    If Expression = "" Then
        StrReplaces_Split_Recursion_ = "": Exit Function
    ElseIf Count(Index) = 0 Then
        If Index = MaxIndex Then
            StrReplaces_Split_Recursion_ = Expression
        Else
            StrReplaces_Split_Recursion_ = StrReplaces_Split_Recursion_(Expression, Find, Replace, Count, Compare, Index + 1, MaxIndex)
        End If
        Exit Function
    ElseIf Count(Index) > 0 Then
        arrRE = VBA.Split(Expression, Find(Index), Count(Index) + 1, Compare)
    Else
        arrRE = VBA.Split(Expression, Find(Index), Count(Index), Compare)
    End If
    
    If Index = MaxIndex Then
        For i = 0 To UBound(arrRE) - 1
            sRE = sRE & arrRE(i) & Replace(Index)
        Next
        StrReplaces_Split_Recursion_ = sRE & arrRE(i)
    Else
        For i = 0 To UBound(arrRE) - 1
            If arrRE(i) = "" Then
                sRE = sRE & Replace(Index)
            Else
                sRE = sRE & StrReplaces_Split_Recursion_(arrRE(i), Find, Replace, Count, Compare, Index + 1, MaxIndex) & Replace(Index)
            End If
        Next
        If arrRE(i) = "" Then
            StrReplaces_Split_Recursion_ = sRE & Replace(Index)
        Else
            StrReplaces_Split_Recursion_ = sRE & StrReplaces_Split_Recursion_(arrRE(i), Find, Replace, Count, Compare, Index + 1, MaxIndex)
        End If
    End If
End Function

'�����滻 Finds,Replaces,Counts֧������
Public Function StrReplaces(Expression, Finds, Replaces, Optional Counts = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim j As Long
    Dim maxL As Long: maxL = 1
    If IsArray(Finds) Then j = ArrCount(Finds): If maxL < j Then maxL = j
    If IsArray(Replaces) Then j = ArrCount(Replaces): If maxL < j Then maxL = j
    If IsArray(Counts) Then j = ArrCount(Counts): If maxL < j Then maxL = j
    
    Dim FindsRE, ReplacesRE, CountsRE
    If IsArray(Finds) Then
        FindsRE = ArrSizeExpansion2(Finds, maxL, "")
    Else
        FindsRE = ArrSizeExpansion2(Finds, maxL, Finds)
    End If
    If IsArray(Replaces) Then
        ReplacesRE = ArrSizeExpansion2(Replaces, maxL, "")
    Else
        ReplacesRE = ArrSizeExpansion2(Replaces, maxL, Replaces)
    End If
    If IsArray(Counts) Then
        CountsRE = ArrSizeExpansion2(Counts, maxL, -1)
    Else
        CountsRE = ArrSizeExpansion2(Counts, maxL, Counts)
    End If
    StrReplaces = StrReplaces_Split_Recursion_(Expression, FindsRE, ReplacesRE, CountsRE, Compare, 1, maxL)
End Function

'�滻ռλ��placeholder    StrReplacePlaceholder("a%b%c", "%", 1, 2) '"a1b2c"
Public Function StrReplacePlaceholder(ByVal String1, placeholder, ParamArray ValueStrs()) As String
    Dim vst
    ValueStrs = ArrFlatten(ValueStrs)
    For Each vst In ValueStrs
        String1 = VBA.Replace(String1, placeholder, vst, 1, 1)
    Next
    StrReplacePlaceholder = String1
End Function

'��StrKey����ַ� �滻��Ӧλ�õ�StrItem  StrReplaceChr("aabbccdd","abc","123") 112233dd
Public Function StrReplaceChr(ByVal String1, StrKey, StrItem) As String
    Dim i As Long, n As Long, s As String
    For i = 1 To VBA.Len(String1)
        s = Mid(String1, i, 1)
        n = VBA.InStr(StrKey, s)
        If n > 0 Then
            Mid(String1, i, 1) = Mid(StrItem, n, 1)
        End If
    Next
    StrReplaceChr = String1
End Function

'������λ���滻
Public Function StrReplaceIndex(String1, ReplaceStr, ByVal Start, ByVal Length) As String
    Dim ri As Long
    If Start < 0 Then
        Start = VBA.Len(String1) + Start + 1
    End If
    If Start > 0 Then
        ri = VBA.Len(String1) - Start - Length + 1
        If ri > 0 Then
            StrReplaceIndex = VBA.Left$(String1, Start - 1) & ReplaceStr & VBA.Right$(String1, ri)
        Else
            StrReplaceIndex = VBA.Left$(String1, Start - 1) & ReplaceStr
        End If
    Else
        ri = Len(String1) - Start - Length + 1
        If ri > 0 Then
            StrReplaceIndex = ReplaceStr & VBA.Right$(String1, ri)
        Else
            StrReplaceIndex = ReplaceStr
        End If
    End If
End Function

'����ַ��� ֧�ֶ���ָ��
Public Function Str_Split(ByVal Expression, Optional Delimitre = "", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String()
    If IsArray(Delimitre) Then
        Dim s, s1, p As Boolean
        p = True
        For Each s In Delimitre
            If p Then
                s1 = s
                p = False
            Else
                Expression = VBA.Replace(Expression, s, s1, 1, -1, Compare)
            End If
        Next
        If s1 = "" Then
            Str_Split = VBA.Split(Expression, , -1, Compare)
        Else
            Str_Split = VBA.Split(Expression, s1, -1, Compare)
        End If
    Else
        If Delimitre = "" Then
            Str_Split = VBA.Split(Expression, , Limit, Compare)
        Else
            Str_Split = VBA.Split(Expression, Delimitre, Limit, Compare)
        End If
    End If
End Function

'���� "���=1,����=abc,����=1" ���͵����ݣ�Str_SplitMatch("���=1,����=abc,����=1", "���=",",����=",",����=")�������飬����(0)��"���="�������
Public Function Str_SplitMatch(String1, ParamArray Delimitre()) As Variant
    Dim UpPointer As Long, LowPointer As Long
    UpPointer = 1
    Dim arr() As String, Ul As Long, Ll As Long, i As Long
    Delimitre = ArrFlatten(Delimitre)
    Ll = LBound(Delimitre)
    Ul = UBound(Delimitre)
    ReDim arr(Ll To Ul + 1) As String
    For i = Ll To Ul
        LowPointer = VBA.InStr(UpPointer, String1, Delimitre(i))
        If LowPointer = 0 Then
            arr(i) = VBA.Mid(String1, UpPointer, Len(String1) - UpPointer + 1)
            Str_SplitMatch = arr
            Exit Function
        End If
        arr(i) = VBA.Mid(String1, UpPointer, LowPointer - UpPointer)
        UpPointer = LowPointer + Len(Delimitre(i))
    Next
    arr(Ul + 1) = VBA.Mid(String1, UpPointer, Len(String1) - UpPointer + 1)
    Str_SplitMatch = arr
End Function

'�ַ�����ֶ�ά����
Public Function Str_Split2D(ByRef String1, DelimitreRow, DelimitreColumn, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim i As Long, j As Long, arrRE1, arrRE2, maxj As Long, arrRE()
    arrRE1 = Split(String1, DelimitreRow, -1, Compare)
    ReDim arrRE2(0 To UBound(arrRE1))
    maxj = 0
    For i = 0 To UBound(arrRE1)
        arrRE2(i) = Split(arrRE1(i), DelimitreColumn, -1, Compare)
        maxj = MaxParams2(maxj, UBound(arrRE2(i)) + 1)
    Next
    ReDim arrRE(1 To UBound(arrRE1) + 1, 1 To maxj)
    For i = 0 To UBound(arrRE2)
        For j = 0 To UBound(arrRE2(i))
            arrRE(i + 1, j + 1) = arrRE2(i)(j)
        Next
    Next
    Str_Split2D = arrRE
End Function

 '������
Public Function StrReg_Split(ByVal Expression, ByVal Pattern As Variant, Optional ByVal ignoreCase As Boolean = True) As Variant
    On Error Resume Next
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = True
            .ignoreCase = ignoreCase
            .multiline = False
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(Expression)
    Dim arr(), i As Long
    Dim UpPointer As Long, LowPointer As Long
    UpPointer = 1
    If searchResults.Count > 0 Then
        ReDim arr(0 To searchResults.Count)
        For i = 0 To searchResults.Count - 1
            arr(i) = VBA.Mid(Expression, UpPointer, searchResults(i).FirstIndex - UpPointer + 1)
            UpPointer = searchResults(i).FirstIndex + searchResults(i).Length + 1
        Next
        arr(searchResults.Count) = VBA.Mid(Expression, UpPointer)
    End If
    StrReg_Split = arr
End Function

'��ƴ������������дƴ������ ע�������ֺ���Ƨ�֣����ܲ�׼
Public Function PinYin(Txt As Variant, Optional Delimiter = " ") As String
    On Error Resume Next
    If Txt = "" Then PinYin = "": Exit Function
    Static PY_DB(1 To 72, 1 To 94) As String
    Static PY_Index(72) As Integer
    If PY_DB(72, 94) <> "a" Then
        Dim i As Long, j As Long
        Dim db As Variant
        Dim PYDB(72) As String
        PYDB(0) = "-2050,-2306,-2562,-2818,-3074,-3330,-3586,-3842,-4098,-4354,-4610,-4866,-5122,-5378,-5634,-5890,-6146,-6402,-6658,-6914,-7170,-7426,-7682,-7938,-8194,-8450,-8706,-8962,-9218,-9474,-9730,-9986,-10247,-10498,-10754,-11010,-11266,-11522,-11778,-12034,-12290,-12546,-12802,-13058,-13314,-13570,-13826,-14082,-14338,-14594,-14850,-15106,-15362,-15618,-15874,-16130,-16386,-16642,-16898,-17154,-17410,-17666,-17922,-18178,-18434,-18690,-18946,-19202,-19458,-19714,-19970,-20226,"
        PYDB(1) = "zha,han,qiu,xi,yan,wu,you,fen,an,can,qing,li,du,qu,yi,xia,you,chu,dai,lin,she,ao,qi,mi,zhu,Jun,ji,mi,hui,me,lie,huan,bin,jiu,quan,xiu,zi,ji,tiao,ran,mao,kun,biao,yong,tao,tie,yan,xiang,chi,wang,xiao,liang,yan,ba,mei,du,bin,kuan,qia,lou,bi,ke,ge,hou,di,gu,ku,tou,jie,bei,gou,rou,ju,jian,man,qiao,da,yang,da,li,zun,shan,gui,yong,min,man,xue,biao,le,yao,guan,ta,qi,ao,"
        PYDB(2) = "sao,bian,huang,fu,qiu,e,die,fen,zi,shi,diao,nian,ni,gu,chang,kun,fei,zou,ling,qing,ji,huan,sha,gun,tiao,shi,jian,lian,li,geng,xun,xiang,jiao,ji,er,wei,jie,gui,tai,hou,fu,su,lu,nian,ping,ba,fang,you,xin,bei,liu,ao,mou,zan,wu,luan,qiong,chou,qu,luo,ju,Jun,sun,zhui,tuo,yuan,mian,wo,chuo,yu,yin,zi,tiao,bao,ju,chen,mai,xian,ai,yin,sha,fei,pei,ji,ting,wen,li,yu,liang,qing,zi,zhi,su,gong,"
        PYDB(3) = "zi,gu,shang,jue,hu,pi,mo,xiu,mo,diao,zhi,xie,zuan,lie,chan,lin,zhu,cu,fan,pu,jue,chu,qi,pan,nie,rou,pian,cuo,duo,ju,zhong,chuai,die,jian,zhi,bo,dian,zhi,chi,huai,chuo,ji,liang,jiao,ji,xian,xian,bi,qiao,kui,tai,bo,jia,tuo,li,shan,fu,zhi,qiang,fu,jian,ta,bo,bie,cu,xue,qiong,dun,cuo,shi,xun,li,ju,xi,jiao,bu,lao,tang,hai,xu,ti,hu,pei,kun,lei,tu,cheng,shai,yan,zhi,ming,xian,tuo,zuo,"
        PYDB(4) = "gu,yi,zhou,gan,ding,chi,jiang,zhe,nan,zan,zi,lie,ju,jiu,qu,fu,dao,yao,qi,qi,zhi,jiao,yi,he,pian,jian,fei,zhu,xi,ling,yi,ji,gen,jiang,qiu,rou,xu,ci,zan,hou,shen,zong,lin,can,zi,xi,tiao,li,ba,mi,xian,xi,tang,jie,suo,qiang,di,bi,sha,qiu,jia,niao,qin,meng,chong,cao,shou,meng,wei,shao,xi,ze,zhu,lu,ge,fang,ban,zhong,bi,yi,shan,chuan,nv,nie,xi,chong,yu,yu,zhou,lai,bo,deng,zan,dian,"
        PYDB(5) = "gui,duan,lu,dou,mie,su,chi,bi,li,fei,gou,hou,huang,kui,zhen,xiao,yuan,kong,dan,bi,tuo,qian,ruo,zhu,qie,ze,qing,xiao,shao,pa,gang,shi,Jun,zheng,quan,yan,xian,bi,kou,chi,bian,jia,tiao,si,li,gou,ze,sheng,da,po,qiong,hu,zi,zhao,jian,ji,du,ji,yu,zhu,shi,xia,qing,ying,fou,qu,du,li,mie,lian,chan,meng,huo,shan,pan,hui,peng,mao,shuai,zhang,zhong,xiang,xi,tang,piao,cao,huang,shi,pang,tang,chi,xi,yuan,ma,"
        PYDB(6) = "mang,man,ao,qin,mou,bian,you,lou,you,yu,sou,fu,ke,kui,fu,nan,rong,chun,meng,lang,wan,quan,tiao,pi,yi,guo,guo,fei,yu,xi,qi,qing,qiang,fu,chu,li,wu,shao,zhe,shen,mou,yang,jiao,qi,kuo,ting,qu,si,zhi,nao,jia,qiong,you,cheng,ling,qiu,zha,ran,you,li,ke,gu,han,chi,yin,dou,gong,jie,hao,xian,rui,pi,fu,meng,ge,hui,chai,ji,qiu,qian,hu,pin,ru,hao,sang,man,nie,zhuan,e,han,ke,ying,he,jie,"
        PYDB(7) = "hang,qi,han,tan,ao,kui,guo,ning,ling,dan,ding,die,mo,nou,jiang,lou,ou,tang,lao,huo,si,chao,zi,lei,jin,cun,Jun,xu,ya,pan,ru,qiang,zhe,chi,lan,bian,lv,bao,bei,da,duo,ju,bi,ti,chu,biao,jian,lian,cheng,lian,ken,ge,jia,dang,pan,mei,jin,ren,na,cha,yi,yu,ju,xun,yu,ke,dou,tiao,yao,bian,zhun,qiong,xi,song,yi,qu,dian,pi,dian,yi,lai,ban,chou,yin,long,zhai,ying,luo,biao,huang,ji,ban,mo,chi,"
        PYDB(8) = "sao,jia,lou,chai,hou,yi,la,dan,yu,yu,wei,gu,fei,zhu,sha,xian,cuo,wu,lao,zhi,yi,ya,jia,xuan,zhu,pao,zha,da,ke,gan,you,li,shan,li,jie,ding,bing,guan,lu,hu,yu,jiu,jiao,liao,liu,zhe,ying,jian,yao,wu,mei,ci,e,hu,chun,bei,an,miao,wu,ti,xian,yu,hu,li,bo,luan,xiu,gua,zhi,er,si,chi,qu,lu,dong,gu,zhen,bao,yuan,jiu,yong,hu,die,po,xi,hao,jiao,gui,rang,fu,nian,se,ji,zhen,"
        PYDB(9) = "ren,ke,lang,fu,ji,lv,shu,mo,zi,bi,zhi,cuo,shen,zhong,biao,cha,yi,zhuo,huo,deng,qiang,cuan,pu,lan,dui,lu,pu,jue,tan,di,xuan,zu,yong,luo,man,tang,biao,bin,jia,yi,liu,na,juan,ge,mo,mei,fei,qiang,lou,ai,huan,sou,cha,e,si,kai,qie,zi,tan,juan,pei,huo,gu,kun,ke,de,ben,nuo,qiang,a,ju,qin,lang,jian,kai,liu,lue,cuo,e,gao,li,zeng,keng,te,lai,lao,ru,an,tang,chong,se,zheng,yao,sha,"
        PYDB(10) = "ha,quan,hua,diu,ding,zhu,kai,yin,diao,cheng,ye,nao,jia,cheng,you,er,lao,kao,duo,pi,ni,bi,ta,xuan,shi,shuo,dian,tan,mu,yue,bo,bu,po,ke,gu,zheng,yu,ba,huo,tou,kang,fang,qian,ban,ju,tai,bu,nv,chai,men,shan,chuan,tu,liao,zhao,po,yi,ga,jin,juan,guan,he,zeng,ji,li,lan,pi,yan,li,gu,gang,fu,tuan,wan,she,zhen,fan,tian,quan,bi,ding,gu,lin,kan,cheng,piao,ming,ke,mao,kui,sou,rui,pi,sui,"
        PYDB(11) = "ni,ya,suo,di,jian,lai,mou,chi,zi,sui,yi,yuan,sheng,dan,miao,dun,kou,mian,xu,fu,fu,zhi,kan,bo,meng,ca,jiang,deng,dun,qu,qing,sang,gun,zhe,bian,xuan,di,jie,zhou,chen,ding,bei,dui,qi,wo,ge,nao,dong,zhai,qiao,xia,mang,xing,fu,tuo,la,di,tong,zha,long,li,ai,fa,feng,bian,zhuo,ya,dun,che,hua,dang,gan,ji,miao,xue,ta,yu,yu,gang,men,mao,dui,qi,te,min,qian,que,zi,yang,nin,nv,hui,jia,dui,"
        PYDB(12) = "tan,te,rang,xi,zhuo,xi,chan,qi,tiao,zhen,ci,zhi,mi,zuo,fu,hu,qu,zhi,xian,si,shi,fei,hu,jiong,hu,li,xi,xu,tao,si,cuan,jue,xian,sui,fan,yu,yi,yun,shang,man,liu,tui,bian,xuan,bao,duan,wei,yu,hu,yan,chao,men,han,wu,yang,ye,tai,xuan,zhu,hu,shi,qiang,dun,wei,yang,yi,liu,ni,jing,zhan,mao,pei,yu,lan,ji,fei,hu,gu,gou,shu,biao,biao,sou,ju,sa,biao,xi,xin,sha,yi,xi,yu,lin,lian,"
        PYDB(13) = "shan,sao,meng,gu,chuai,zhi,teng,bin,lv,ge,ying,cheng,shu,e,wa,mian,nan,cou,jian,ding,zong,yu,fei,yan,jing,niao,wan,pao,cuo,luo,tun,mi,zhen,pian,hai,sa,kuai,yan,dong,guang,jing,zhi,qu,zhen,gua,zuo,zhou,shen,jia,lu,ka,dong,long,qian,yao,na,zhun,gong,tai,ruan,jing,huang,rong,wo,yue,guo,yuan,you,die,du,jiao,chi,fan,yun,ke,yin,ya,dong,chuan,xian,dao,pie,qu,pu,lu,chang,shu,san,jian,cui,mu,mao,bo,ge,"
        PYDB(14) = "bai,suo,qie,kao,pian,jian,ju,ji,gu,wu,gu,mao,pin,jiang,jian,qu,jin,gou,yu,di,xi,ji,chan,fu,dan,ji,qiu,lai,zhen,jin,gai,zi,zhi,yi,kuang,shi,ben,nang,xi,yao,xun,tun,ming,ai,kui,xuan,gui,han,bu,hui,yan,chao,ye,sheng,qi,ni,chang,yu,mao,zan,he,gui,yun,xin,ze,gao,tan,hao,gan,la,ga,po,pi,zeng,beng,bu,ling,ou,zang,jian,gai,deng,kan,ji,ji,jia,qiang,jian,wei,lin,lu,cou,zi,chuo,"
        PYDB(15) = "wang,nian,zhe,lu,quan,zhi,shi,yao,li,hu,zhen,yi,zhi,lu,ke,gu,e,ren,yi,bin,ji,dan,piao,lian,yun,tian,shang,cu,mo,ao,you,cha,bo,lin,yan,lei,yuan,ju,xi,zun,lu,qin,qiao,jue,tuo,qing,yue,gan,hu,zhu,tang,chu,qi,qiang,jin,xie,zhu,rong,bin,shuo,gao,cui,gao,xie,sun,ta,fei,zhen,ying,mei,xuan,ju,cha,lv,chen,chui,duan,qiu,ju,pin,ji,lan,lian,zha,nan,zhen,cou,ju,di,jian,guo,liang,chui,luo,"
        PYDB(16) = "zhao,qian,du,fen,chu,ling,suo,zi,jue,fu,gu,fan,xu,an,juan,luan,jie,gui,heng,hua,jiu,gua,ting,qi,guang,zhen,zhi,rao,ya,lao,kao,cheng,tuo,li,di,gou,ling,zhi,tuo,zhi,you,xiao,xia,lu,ping,jiu,long,zhe,zhi,nai,zhu,pa,fang,xiao,cong,cheng,chu,jian,rui,yao,miao,pi,li,ma,cha,qi,shao,wu,tao,yun,wei,wen,zan,bi,lu,qu,can,pu,zhang,xuan,cong,cui,ying,huang,jin,tang,ai,nao,xia,yuan,yu,mao,ju,chen,"
        PYDB(17) = "wan,cong,yan,kun,hu,qi,ying,qi,lian,hui,xi,luo,yao,heng,ya,xu,gong,er,jia,min,po,dai,dian,long,ke,jue,min,bin,wei,ji,ding,yong,zai,chuan,ji,yao,zuan,huan,qiao,qian,jiang,zeng,liao,xie,sao,miao,lei,man,piao,bin,jian,yi,li,gao,ru,zhen,jin,min,zhui,gou,bian,si,hui,miao,ti,xiang,ke,zi,wan,quan,liu,shou,duo,gun,shang,fei,qi,ling,ti,xiao,geng,jiang,hang,ku,dai,chu,fu,zhou,fu,xie,gan,shu,pi,yun,"
        PYDB(18) = "kuang,wan,ge,zhou,yu,jiao,xiang,ji,chan,cong,biao,shan,liu,ao,wu,zhi,can,zhui,ke,qi,li,pian,hua,xiao,dai,nu,yi,zou,fu,si,zang,bao,jue,jie,zi,nu,fu,ga,ga,shuang,mo,niao,bi,shan,xi,zhang,lei,chang,piao,qiang,yan,li,chi,pin,ai,pi,mo,gou,wu,ting,yuan,ao,nu,chan,bi,chang,jie,biao,jing,e,wei,di,suo,xian,wa,ping,li,cha,pin,jiao,luan,shu,rao,ya,qie,shan,zhou,da,si,yu,niu,gui,zi,jin,"
        PYDB(19) = "bi,yu,wu,yan,fei,shuo,chu,yu,bi,fu,mi,nu,jing,chan,ju,xi,chan,e,ji,zhi,kao,zhi,tuan,hui,xun,la,sui,miao,xie,ju,lin,xian,liu,ta,gou,ao,xia,qiu,huang,chuan,lu,huan,wei,kui,qun,ti,xiao,qiu,li,bu,pang,hou,dai,jing,jia,er,yi,ze,jiong,wu,ya,zou,jian,jian,huan,qian,liao,wu,qian,qian,ning,chen,you,mi,dang,gui,bao,ba,hao,fen,yue,ying,xie,han,zhuo,hao,bi,pu,ru,lian,chan,dan,li,sui,"
        PYDB(20) = "lai,chan,tong,shao,shan,si,shu,gan,xuan,lu,yi,zhu,lian,huan,luo,hu,cao,lan,xiao,ying,huang,ming,pang,tang,fu,xiu,bi,hun,ta,ru,li,pu,ying,mang,she,ke,qin,yan,mei,wo,xuan,jian,pen,xu,huang,sou,qiu,mian,yan,xie,shuan,lu,guan,shen,cong,fei,gan,mian,pi,zhuo,du,song,xi,qi,zhu,huan,mei,xi,bang,cen,juan,zhuo,wei,lai,wu,su,ru,xun,hu,liu,jiang,xun,tao,hui,xu,ji,zhu,hui,yin,zhen,jia,lie,wei,huan,"
        PYDB(21) = "jing,min,hong,tuo,pan,xuan,luo,mao,ling,duo,si,yang,lu,long,shu,gan,le,wei,hang,wen,bian,gu,mi,dun,mian,mu,yuan,feng,cha,si,qi,san,qiang,pan,zhuang,kan,que,tian,he,que,qu,e,hun,wen,xi,chang,yu,lang,jiu,kun,lv,ta,kang,min,hong,wei,yan,shuan,hui,tian,meng,lin,chu,chong,qiao,jing,yong,qie,su,bi,qiao,zhui,leng,e,kui,yun,cui,hu,chou,wang,chang,fei,xing,qie,quan,ti,yi,kun,kui,qian,song,bei,yun,ke,"
        PYDB(22) = "xun,kai,ce,yan,tong,yi,yi,chao,fu,ni,zuo,yang,da,peng,chu,hu,niu,bian,song,chuang,chang,kai,wu,chong,ou,zhi,wu,chan,cun,dao,shu,ying,lin,xie,chan,jin,ao,geng,bi,yu,an,tuo,xiang,xiu,pao,gui,wu,pi,nang,zhuan,san,jin,xiu,mo,sou,zha,hun,yu,bo,xiang,yi,chi,yu,ren,xi,tun,tang,shi,dong,yin,sun,huo,chuan,huan,xun,xie,liao,jue,jing,zhang,nao,mei,wei,wei,cha,hu,mi,cu,she,ni,luo,guo,yi,suan,"
        PYDB(23) = "yin,xian,yu,li,juan,bi,sun,shou,kuai,rong,fei,pao,xia,yun,niu,ma,guang,an,qiu,fan,san,qu,jiao,zheng,yao,huang,chang,xi,lai,hou,yang,xun,cu,pang,chi,dian,yi,bin,deng,lin,zhang,ji,song,sheng,mei,zi,lou,cuo,yu,wei,zai,wai,yao,rong,jue,kong,guo,xiao,gu,yan,song,lai,lao,zheng,xun,qiao,dong,yi,min,mao,gou,dai,xiu,jia,dong,ke,hu,ba,lan,cen,ao,xian,ya,qu,qi,qian,qi,ji,fan,fu,zhang,man,wo,wei,"
        PYDB(24) = "guo,ze,dao,tang,pei,zhi,wei,huan,yu,qing,yu,you,ling,hu,lun,nan,jian,wei,nang,huo,cha,ru,pi,sai,yi,jue,jin,hao,deng,ceng,lu,qin,o,jiao,chuai,pu,liao,jue,peng,mi,di,sou,beng,ying,qi,piao,cao,lei,pei,chi,tong,hei,suo,ai,ai,dia,hao,en,a,ge,suo,chen,he,nie,ke,du,su,ao,qin,hui,wo,ku,lou,jie,chi,yin,sou,jiu,kui,yong,jie,kui,li,nan,da,die,chuo,shua,li,lang,ding,bo,dan,yo,"
        PYDB(25) = "sha,cui,hu,tao,zhou,zhuan,lin,miao,nuo,ze,feng,ji,zuo,xi,zao,suo,zha,en,geng,lao,chi,mai,mou,gen,nong,zha,mi,mie,ji,duo,kuai,pai,yi,xiu,yue,guang,ci,bi,xiao,yi,lie,da,hui,shen,ji,kuang,si,you,nao,duo,ning,dong,ling,gua,ga,ka,za,yin,qin,guo,bei,bi,e,li,tai,yi,mu,fu,yao,a,zha,le,dao,kou,ji,chi,bu,shi,dai,tui,yi,nang,zuan,huo,zhuo,xing,pi,huan,gan,cuan,zun,lu,xie,zhe,"
        PYDB(26) = "han,zhi,ying,luo,sang,nuo,zhan,shuo,jian,chuai,en,shu,yuan,kui,bing,xuan,an,yu,qin,ya,zha,die,guan,qian,lie,pou,ju,bai,guo,ji,na,ai,ye,tian,Jun,lv,yi,za,jiao,jie,niu,pin,fu,chen,tuan,men,ti,gan,ga,liao,you,pao,zang,xi,yi,da,lian,kuang,yi,nong,mi,nie,fan,heng,qu,huo,li,gao,xian,xun,ru,tai,hao,bi,sou,weng,yi,wei,hong,xie,hong,qi,fan,meng,ji,zui,rui,jue,xun,hui,liao,xu,kou,qu,"
        PYDB(27) = "lin,cu,lian,xi,dou,meng,su,yu,ying,lang,shuo,jian,bang,li,ji,hao,weng,bei,en,mo,ru,shi,zhen,jia,xuan,pai,lou,ting,pa,bao,e,xi,kui,qi,kai,chan,wei,xiang,shen,feng,qia,han,gu,ying,wan,jian,dang,zu,yan,cui,dan,tu,fu,bi,huan,yu,tie,chang,shu,ba,qi,nai,jin,song,xi,qi,jing,chun,ying,lang,guan,shen,di,you,sui,piao,xian,tu,li,you,mei,e,you,wo,shi,bi,kan,zhou,hong,sun,mai,jin,gen,qian,"
        PYDB(28) = "ying,luo,jiang,chong,jiao,qi,ming,xun,hui,quan,xing,ren,fu,qiao,ting,zhu,hui,tong,ju,zi,bi,rao,yi,qian,tiao,min,qiong,ying,mao,yin,niao,ling,fu,chi,qing,ran,ju,mu,ba,long,pie,gan,mo,yi,kou,zhu,bian,shan,qi,qian,wu,qin,cong,chang,xian,rui,zhi,pi,ju,e,li,ji,fei,yun,yuan,fu,xiang,qi,xiong,wan,ji,qian,du,nai,jiao,cao,yi,pi,xin,chi,liang,yong,man,yuan,ge,hou,leng,yin,die,ku,dai,sao,tu,peng,"
        PYDB(29) = "nian,pi,yi,an,zhi,yuan,lie,xun,guo,shi,cheng,yin,gai,nao,shang,shan,kai,dong,die,ya,ao,mu,ni,tuo,di,che,lu,dian,long,gan,ban,qi,li,yi,pi,kuang,zhen,ge,wu,xu,he,yong,ji,shu,fa,e,ben,qiu,ben,bian,si,chang,dang,shan,jian,jue,xie,sou,xie,xu,meng,ge,he,shao,qu,mai,huan,chu,feng,ling,zou,shan,po,zhang,yin,yan,juan,yan,tan,pi,fu,xi,gao,ying,li,yun,xun,qie,kuai,zhu,zhi,jia,tai,di,"
        PYDB(30) = "ye,bei,pi,bing,fang,wu,mang,kuang,qiong,han,xi,wei,huang,wei,pi,chui,zou,nie,zhi,gai,xing,bei,zuo,yan,ban,jing,qian,wu,zuo,jin,dan,chen,zhan,yan,jue,qiao,zen,jian,zhe,mi,shi,su,dang,mo,pian,zi,di,an,xuan,yu,e,ye,xue,jian,chen,sui,chan,shen,yu,wei,zhuo,zou,ei,kuang,gao,qiao,xu,hun,zheng,quan,gou,shen,hui,jie,gua,lei,kuang,yi,qu,zhao,di,he,gu,ne,ju,ou,shan,hong,jie,yan,ming,zhong,ping,song,"
        PYDB(31) = "xian,lie,hu,liang,lei,luo,ying,bing,pou,luan,xie,mao,gun,bo,yan,wen,si,su,fu,fu,hong,pu,bao,kui,chan,guo,hong,xun,xi,di,cuan,yue,zu,qian,she,tun,tong,dan,xuan,tong,jiu,jiao,jian,jing,xi,chi,nuo,bin,tang,lv,zong,wei,ji,xie,yan,fen,ju,kong,guan,ti,bi,wo,luo,shu,zhuo,pai,ruo,qian,feng,si,yong,ping,yu,li,qiu,li,yan,chou,mou,nong,jiao,chai,tiao,yi,zhu,kan,kua,you,er,ji,ga,ni,tuo,tong,"
        PYDB(32) = "gou,yi,you,ka,ning,zhu,kang,cang,chang,wu,wa,ya,pi,yu,ren,mu,yi,sa,le,zhang,ding,dan,wang,tong,yi,huo,qiao,jue,piao,kuai,wan,yan,ji,la,kai,gui,ku,jing,wen,yi,ce,you,gua,ze,bian,kui,gui,po,qu,yan,ye,si,jue,yan,cuo,she,ze,gu,se,bo,mi,qi,ji,nie,nai,ji,dian,tao,gao,yu,kui,yin,xin,di,zhi,yao,yao,tuo,bi,pie,yu,shu,e,nao,ge,cheng,gen,pi,sa,nian,gai,wu,ji,chu,"
        PYDB(33) = "zuo,zuo,zuo,zuo,zha,zuo,zuo,zuo,zun,zun,zui,zui,zui,zui,zuan,zuan,zu,zu,zu,zu,zu,zu,zu,zu,zou,zou,zou,zou,zong,zong,zong,zong,zong,zong,zong,zi,zi,zi,zi,zi,zi,zai,zi,zi,zi,zi,zi,zi,zi,zi,zhuo,zhuo,zhe,zhuo,zhuo,zhuo,zhuo,zhuo,zhuo,zhuo,zhuo,zhun,zhun,zhui,zhui,zhui,zhui,zhui,zhui,zhuang,zhuang,zhuang,zhuang,zhuang,zhuang,zhuang,zhuan,zuan,zhuan,zhuan,zhuan,zhuan,zhuai,zhao,zhua,zhu,zhu,zhu,zhu,,,,,,"
        PYDB(34) = "zhu,zhu,zhu,zhu,zhu,zhu,zhe,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhu,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhou,zhong,zhong,zhong,zhong,zhong,zhong,zhong,zhong,zhong,zhong,zhong,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zhi,zi,zhi,zhi,zhi,zheng,zheng,zheng,zhen,"
        PYDB(35) = "zheng,zheng,zheng,zheng,zheng,zheng,zheng,zheng,zheng,zheng,zheng,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhen,zhe,zhe,zhe,zhe,zhe,zhe,zhe,zhe,zhe,zhe,zhao,zhao,zhao,zhao,zhao,zhao,zhao,zhao,zhao,zhao,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhang,zhan,zhan,zhan,zhan,zhan,zhan,zhan,zhan,zhan,nian,zhan,zhan,zhan,zhan,zhan,zhan,zhan,zhai,zhai,zhai,zhai,zhai,zhai,zha,zha,zha,za,zha,shan,zha,zha,zha,"
        PYDB(36) = "zha,zha,zha,zha,zha,zeng,ceng,zeng,zeng,zen,zei,ze,ze,ze,ze,zao,zao,zao,zao,zao,zao,zao,zao,zao,zao,zao,zao,zao,zao,zang,zang,zang,zan,zan,zan,zan,zai,zai,zai,zai,zai,zai,zai,za,za,za,yun,yun,yun,yun,yun,yun,yun,yun,yun,yun,yun,yun,yue,yue,yue,yue,yue,yao,yue,yue,yue,yue,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yuan,yu,yu,yu,yu,yu,yu,"
        PYDB(37) = "yu,yu,yu,yu,yu,yu,yu,yu,yu,xu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,yu,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,you,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yong,yo,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,ying,yin,"
        PYDB(38) = "yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yin,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,yi,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,ye,yao,yao,yao,yao,yao,yao,yao,yao,yao,yao,yao,"
        PYDB(39) = "yao,yao,yao,yao,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yang,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,yan,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,ya,xun,xun,xun,xun,xun,xun,xun,xun,xun,xun,xun,xun,xun,xun,xue,xue,xue,xue,xue,xue,xuan,xuan,xuan,xuan,"
        PYDB(40) = "xuan,xuan,xuan,xuan,xuan,xuan,xu,xu,xu,xu,xu,chu,xu,xu,xu,xu,xu,xu,xu,xu,xu,xu,xu,xu,xu,xiu,xiu,xiu,xiu,xiu,xiu,xiu,xiu,xiu,xiong,xiong,xiong,xiong,xiong,xiong,xiong,xing,xing,xing,xing,xing,hang,xing,xing,xing,xing,xing,xing,xing,xing,xing,xin,xin,xin,xin,xin,xin,xin,xin,xin,xin,xie,xie,xie,xie,xie,xie,xie,xie,xie,xie,xie,xie,xie,xie,jia,xie,xie,xie,xie,xie,xie,xiao,xiao,xiao,xiao,xiao,xiao,xiao,"
        PYDB(41) = "xiao,xiao,xiao,xiao,xiao,xiao,xiao,xue,xiao,xiao,xiao,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xiang,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xian,xia,xia,sha,xia,xia,xia,xia,xia,xia,xia,xia,xia,xia,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,"
        PYDB(42) = "xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,xi,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wu,wo,wo,wo,wo,wo,wo,wo,wo,wo,weng,weng,weng,wen,wen,wen,wen,wen,wen,wen,wen,wen,wen,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,wei,"
        PYDB(43) = "wei,wang,wang,wang,wang,wang,wang,wang,wang,wang,wang,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wan,wai,wai,wa,wa,wa,wa,wa,wa,wa,tuo,tuo,tuo,tuo,tuo,tuo,tuo,tuo,tuo,tuo,tuo,tun,tun,tun,tui,tui,tui,tui,tui,tui,tuan,tuan,tu,tu,tu,tu,tu,tu,tu,tu,tu,tu,tu,tou,tou,tou,tou,tong,tong,tong,tong,tong,tong,tong,tong,tong,tong,tong,tong,tong,ting,ting,ting,ting,ting,ting,ting,"
        PYDB(44) = "ting,ting,ting,tie,tie,tie,tiao,tiao,tiao,tiao,tiao,tian,tian,tian,tian,tian,tian,tian,tian,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,ti,teng,teng,teng,teng,te,tao,tao,tao,tao,tao,tao,tao,tao,tao,tao,tao,tang,tang,tang,tang,tang,tang,tang,tang,tang,tang,tang,tang,tang,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tan,tai,tai,tai,tai,tai,tai,tai,tai,tai,ta,ta,ta,ta,"
        PYDB(45) = "ta,ta,ta,ta,ta,suo,suo,suo,suo,suo,suo,suo,suo,sun,sun,sun,sui,sui,sui,sui,sui,sui,sui,sui,sui,sui,sui,suan,suan,suan,su,su,xiu,su,su,su,su,su,su,su,su,su,sou,sou,sou,sou,song,song,song,song,song,song,song,song,si,si,si,si,si,si,si,si,si,si,si,si,si,si,si,si,shuo,shuo,shuo,shuo,shun,shun,shun,shun,shui,shui,shui,shui,shuang,shuang,shuang,shuan,shuan,shuai,shuai,shuai,shuai,shua,shua,shu,"
        PYDB(46) = "shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shu,shou,shou,shou,shou,shou,shou,shou,shou,shou,shou,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,shi,sheng,sheng,sheng,sheng,sheng,"
        PYDB(47) = "sheng,sheng,sheng,sheng,sheng,sheng,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,shen,she,she,she,she,she,she,she,she,she,she,she,she,shao,shao,shao,shao,shao,shao,shao,shao,shao,shao,shao,shang,shang,shang,shang,shang,shang,shang,shang,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shan,shai,shai,sha,sha,sha,sha,sha,cha,sha,sha,sha,seng,sen,se,se,se,sao,sao,sao,sao,sang,sang,sang,san,san,"
        PYDB(48) = "san,san,sai,sai,sai,sai,sa,sa,sa,ruo,ruo,run,run,rui,rui,rui,ruan,ruan,ru,ru,ru,ru,ru,ru,ru,ru,ru,ru,rou,rou,rou,rong,rong,rong,rong,rong,rong,rong,rong,rong,rong,ri,reng,reng,ren,ren,ren,ren,ren,ren,ren,ren,ren,ren,re,re,rao,rao,rao,rang,rang,rang,rang,rang,ran,ran,ran,ran,qun,qun,que,que,que,que,que,que,gui,que,quan,quan,quan,quan,quan,quan,quan,quan,quan,quan,quan,qu,qu,qu,qu,qu,"
        PYDB(49) = "qu,qu,qu,qu,qu,qu,qu,qu,qiu,qiu,qiu,qiu,qiu,qiu,qiu,qiu,qiong,qiong,qing,qing,qing,qing,qing,qing,qing,qing,qing,qing,qing,qing,qing,qin,qin,qin,qin,qin,qin,qin,qin,qin,qin,qin,qie,qie,qie,qie,qie,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiao,qiang,qiang,qiang,qiang,qiang,qiang,qiang,qiang,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qian,qia,qia,"
        PYDB(50) = "qia,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,qi,bao,pu,pu,pu,pu,pu,pu,pu,pu,pu,pu,pu,pu,pu,pu,pou,po,po,po,po,po,po,po,po,ping,ping,ping,ping,ping,ping,ping,ping,ping,pin,pin,pin,pin,pin,pie,pie,piao,piao,piao,piao,pian,pian,pian,pian,pi,pi,pi,pi,pi,pi,pi,pi,pi,"
        PYDB(51) = "pi,pi,pi,pi,pi,pi,pi,pi,peng,peng,peng,peng,peng,peng,peng,peng,peng,peng,peng,peng,peng,peng,pen,pen,pei,pei,pei,pei,pei,pei,pei,pei,pei,pao,pao,pao,pao,pao,pao,pao,pang,pang,pang,pang,pang,pan,pan,pan,pan,pan,pan,pan,pan,pai,pai,pai,pai,pai,pai,pa,pa,pa,pa,pa,pa,ou,ou,ou,ou,ou,ou,ou,o,nuo,nuo,nuo,nuo,nue,nue,nuan,nv,nu,nu,nu,nong,nong,nong,nong,niu,niu,niu,niu,ning,ning,"
        PYDB(52) = "ning,ning,ning,ning,nin,nie,nie,nie,nie,nie,nie,nie,niao,niao,niang,niang,nian,nian,nian,nian,nian,nian,nian,ni,ni,ni,ni,ni,ni,ni,ni,ni,ni,ni,neng,nen,nei,nei,ne,nao,nao,nao,nao,nao,nang,nan,nan,nan,nai,nai,nai,nai,nai,na,na,na,na,na,na,na,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mu,mou,mou,mou,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,mo,"
        PYDB(53) = "mo,miu,ming,ming,ming,ming,ming,ming,min,min,min,min,min,min,mie,mie,miao,miao,miao,miao,miao,miao,miao,miao,mian,mian,mian,mian,mian,mian,mian,mian,mian,mi,mi,mi,mi,mi,mi,mi,mi,mi,mi,mi,mi,mi,mi,meng,meng,meng,meng,meng,meng,meng,meng,men,men,men,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,mei,me,mao,mao,mao,mao,mao,mao,mao,mao,mao,mao,mao,mao,mang,mang,mang,mang,mang,mang,man,"
        PYDB(54) = "man,man,man,man,man,man,man,man,mai,mai,mai,mai,mai,mai,ma,ma,ma,ma,ma,ma,ma,ma,ma,luo,luo,luo,luo,luo,luo,luo,luo,luo,luo,luo,luo,lun,lun,lun,lun,lun,lun,lun,lue,lue,luan,luan,luan,luan,luan,luan,lv,lv,lv,lv,lv,lv,lv,lv,lv,lv,lv,lv,lv,lv,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lu,lou,lou,lou,lou,lou,lou,long,long,long,long,"
        PYDB(55) = "long,long,long,long,long,liu,liu,liu,liu,liu,liu,liu,liu,liu,liu,liu,ling,ling,ling,ling,ling,ling,ling,ling,ling,ling,ling,ling,ling,ling,lin,lin,lin,lin,lin,lin,lin,lin,lin,lin,lin,lin,lie,lie,lie,lie,lie,liao,liao,liao,liao,le,liao,liao,liao,liao,liao,liao,liao,liao,liang,liang,liang,liang,liang,liang,liang,liang,liang,liang,liang,lian,lian,lian,lian,lian,lian,lian,lian,lian,lian,lian,lian,lian,lian,liang,li,li,li,li,li,li,li,li,"
        PYDB(56) = "li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,li,leng,leng,leng,lei,lei,lei,lei,lei,lei,lei,lei,lei,lei,lei,le,le,lao,lao,lao,lao,lao,lao,lao,lao,lao,lang,lang,lang,lang,lang,lang,lang,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lan,lai,lai,lai,la,la,la,la,la,la,la,kuo,kuo,kuo,kuo,kun,kun,kun,kun,kui,kui,kui,"
        PYDB(57) = "gui,kui,kui,kui,kui,kui,kui,kui,kuang,kuang,kuang,kuang,kuang,kuang,kuang,kuang,kuan,kuan,kuai,kuai,kuai,kuai,kua,kua,kua,kua,kua,ku,ku,ku,ku,ku,ku,ku,kou,kou,kou,kou,kong,kong,kong,kong,keng,keng,ken,ken,ken,ken,ke,ke,ke,ke,ke,ke,hai,ke,ke,ke,ke,ke,ke,ke,ke,kao,kao,kao,kao,kang,kang,kang,kang,kang,kang,kang,kan,kan,kan,kan,kan,kan,kai,kai,kai,kai,kai,ge,ka,ka,ka,Jun,Jun,xun,Jun,Jun,"
        PYDB(58) = "Jun,Jun,Jun,Jun,Jun,Jun,jue,jue,jue,jiao,jue,jue,jue,jue,jue,jue,juan,juan,juan,juan,juan,juan,juan,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,ju,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiu,jiong,jiong,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jing,jin,jin,"
        PYDB(59) = "jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jin,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,jie,ju,jie,jie,jie,jie,jie,jie,jie,jie,jie,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,yao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiao,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jiang,jian,jian,jian,jian,jian,jian,jian,jian,"
        PYDB(60) = "jian,jian,jian,jian,jian,jian,jian,kan,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jian,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,jia,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,ji,"
        PYDB(61) = "ji,ji,ji,ji,ji,ji,ji,ji,huo,huo,huo,huo,huo,huo,huo,huo,huo,huo,hun,hun,hun,hun,hun,hun,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,hui,huang,huang,huang,huang,huang,huang,huang,huang,huang,huang,huang,huang,huang,huang,huan,huan,huan,huan,huan,huan,huan,huan,huan,huan,hai,huan,huan,huan,huai,huai,huai,huai,huai,hua,hua,hua,hua,hua,hua,hua,hua,hua,hu,hu,hu,hu,hu,hu,hu,"
        PYDB(62) = "hu,hu,hu,hu,hu,hu,hu,hu,hu,hu,hu,hou,hou,hou,hou,hou,hou,hou,hong,hong,hong,hong,hong,hong,hong,hong,hong,heng,heng,heng,heng,heng,hen,hen,hen,hen,hei,hei,he,he,he,he,he,he,he,mo,he,he,he,he,he,he,he,he,he,he,hao,hao,hao,hao,hao,hao,hao,hao,hao,hang,hang,ben,han,han,han,han,han,han,han,han,han,han,han,han,han,han,han,han,han,han,han,hai,hai,hai,hai,hai,hai,hai,"
        PYDB(63) = "ha,guo,guo,guo,guo,guo,guo,gun,gun,gun,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,gui,guang,guang,guang,guan,guan,guan,guan,guan,guan,guan,guan,guan,guan,guan,guai,guai,guai,gua,gua,gua,gua,gua,gua,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gu,gou,gou,gou,gou,gou,gou,gou,gou,gou,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,gong,geng,geng,geng,"
        PYDB(64) = "geng,geng,geng,geng,gen,gen,gei,ge,ge,ge,ge,ge,ha,ge,ge,ge,ge,ge,ge,ge,ge,ge,ge,ge,gao,gao,gao,gao,gao,gao,gao,gao,gao,gao,gang,gang,gang,gang,gang,gang,gang,gang,gang,gan,gan,gan,gan,gan,gan,gan,gan,gan,gan,gan,gai,gai,gai,gai,gai,gai,ga,ga,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,pu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,"
        PYDB(65) = "fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fu,fou,fo,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,feng,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fen,fei,fei,fei,fei,fei,fei,fei,fei,fei,fei,fei,fei,fang,fang,fang,fang,fang,fang,fang,fang,fang,fang,fang,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fan,fa,fa,fa,fa,fa,fa,fa,fa,er,"
        PYDB(66) = "er,er,er,er,er,er,er,en,e,e,e,e,e,e,e,e,e,e,e,e,e,duo,duo,duo,duo,duo,duo,duo,duo,duo,duo,duo,duo,dun,dun,dun,dun,dun,dun,dun,dun,dun,dui,dui,dui,dui,duan,duan,duan,duan,duan,duan,du,du,du,du,du,du,du,du,du,du,du,du,du,du,dou,dou,dou,dou,dou,dou,dou,dou,dong,dong,dong,dong,dong,dong,dong,dong,dong,dong,diu,ding,ding,ding,ding,ding,ding,ding,ding,ding,"
        PYDB(67) = "die,die,die,die,die,die,die,diao,diao,diao,diao,diao,diao,diao,diao,diao,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,dian,di,di,di,di,di,di,di,di,di,di,zhai,di,di,di,di,di,di,di,di,deng,deng,deng,deng,deng,deng,deng,de,de,de,dao,dao,dao,dao,dao,dao,dao,dao,dao,dao,dao,dao,dang,dang,dang,dang,dang,dan,tan,dan,dan,dan,dan,dan,dan,dan,dan,dan,dan,dan,dan,dan,dai,"
        PYDB(68) = "dai,dai,dai,dai,dai,dai,dai,dai,dai,dai,dai,da,da,da,da,da,da,cuo,cuo,cuo,cuo,cuo,cuo,cun,cun,cun,cui,cui,cui,cui,cui,cui,cui,cui,cuan,cuan,cuan,cu,cu,cu,cu,cou,cong,cong,cong,cong,cong,cong,ci,ci,ci,ci,ci,ci,ci,ci,ci,ci,ci,ci,chao,chuo,chun,chun,chun,chun,chun,chun,chun,chui,chui,chui,chui,chui,chuang,chuang,chuang,zhuang,chuang,chuang,chuan,chuan,chuan,chuan,chuan,chuan,chuan,chuai,chu,chu,chu,chu,chu,chu,"
        PYDB(69) = "chu,chu,chu,chu,chu,chu,chu,chu,chu,chu,chou,chou,chou,chou,chou,chou,chou,chou,chou,chou,chou,chou,chong,chong,chong,chong,chong,chi,chi,chi,chi,chi,chi,chi,chi,chi,chi,chi,chi,shi,chi,chi,chi,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,cheng,chen,chen,chen,chen,chen,chen,chen,chen,chen,chen,che,che,che,che,che,che,chao,chao,chao,chao,chao,chao,chao,chao,chao,chang,chang,chang,chang,chang,chang,chang,chang,chang,chang,chang,"
        PYDB(70) = "chang,chang,chan,chan,chan,chan,chan,chan,chan,chan,chan,chan,chai,chai,chai,cha,cha,cha,cha,cha,cha,cha,cha,cha,cha,cha,ceng,ceng,ce,ce,ce,ce,ce,cao,cao,cao,cao,cao,cang,cang,cang,cang,cang,can,can,can,can,can,can,can,cai,cai,cai,cai,cai,cai,cai,cai,cai,cai,cai,ca,bu,bu,bu,bu,bu,bu,bu,bu,bu,bu,bu,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bo,bing,bing,"
        PYDB(71) = "bing,bing,bing,bing,bing,bing,bing,bin,bin,bin,bin,bin,bin,bie,bie,bie,bie,biao,biao,biao,biao,bian,bian,bian,bian,bian,bian,bian,bian,bian,bian,bian,bian,bi,bi,bi,bi,pi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,bi,beng,beng,beng,beng,beng,beng,ben,ben,ben,ben,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bei,bao,bao,bao,bao,bao,bao,bao,bao,bao,bao,bao,bao,"
        PYDB(72) = "bao,bao,bao,bao,bao,bang,bang,bang,bang,pang,bang,bang,bang,bang,bang,bang,bang,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,ban,bai,bai,bai,bai,bai,bai,bai,bai,ba,ba,ba,ba,pa,ba,ba,ba,ba,ba,ba,ba,ba,ba,ba,ba,ba,ba,ao,ao,ao,ao,ao,ao,ao,ao,ao,ang,ang,ang,an,an,an,an,an,an,an,an,an,ai,ai,ai,ai,ai,ai,ai,ai,ai,ai,ai,ai,ai,a,a,"
        db = VBA.Split(PYDB(0), ",")
        For i = 1 To UBound(db)
            PY_Index(i) = db(i - 1)
        Next i
        For i = 1 To 72
            db = VBA.Split(PYDB(i), ",")
            For j = 1 To UBound(db)
                PY_DB(i, j) = db(j - 1)
            Next j
        Next i
    End If
    Dim n As Long, ASCID As Long, y As Byte
    Dim M_Txt As String, M_PY As String
    For i = 1 To Len(Trim(Txt))
        M_Txt = Mid(Trim(Txt), i, 1)
        If M_Txt = "" Then
            M_PY = ""
        Else
            ASCID = Asc(M_Txt)
            For n = 1 To UBound(PY_Index)
                If PY_Index(n) < ASCID Then
                    Exit For
                End If
            Next n
            Dim PYDB_Index
            PYDB_Index = PY_Index(n - 1) - ASCID
            If PYDB_Index < 0 Or PYDB_Index > 93 Then
                M_PY = M_Txt
                y = 1
            Else
                M_PY = PY_DB(n - 1, PYDB_Index + 1)
            End If
        End If
        PinYin = PinYin & IIf(M_PY = M_Txt, M_PY, IIf(y = 1, Delimiter & M_PY & Delimiter, IIf(i = Len(Trim(Txt)), M_PY, M_PY & Delimiter)))
        y = IIf(y = 1, IIf(M_PY = M_Txt, 1, 0), 0)
    Next i
End Function
 
'ƴ����ͷ
Public Function PinYinInitial(Txt As Variant) As String
    Dim i As Long, getpychar As String, tmp As Long
    For i = 1 To Len(Txt)
        tmp = 65536 + Asc(Mid(Txt, i, 1))
        If (tmp >= 45217 And tmp <= 45252) Then
            getpychar = "a"
        ElseIf (tmp >= 45253 And tmp <= 45760) Then
            getpychar = "b"
        ElseIf (tmp >= 45761 And tmp <= 46317) Then
            getpychar = "c"
        ElseIf (tmp >= 46318 And tmp <= 46825) Then
            getpychar = "d"
        ElseIf (tmp >= 46826 And tmp <= 47009) Then
            getpychar = "e"
        ElseIf (tmp >= 47010 And tmp <= 47296) Then
            getpychar = "f"
        ElseIf (tmp >= 47297 And tmp <= 47613) Then
            getpychar = "g"
        ElseIf (tmp >= 47614 And tmp <= 48118) Then
            getpychar = "h"
        ElseIf (tmp >= 48119 And tmp <= 49061) Then
            getpychar = "j"
        ElseIf (tmp >= 49062 And tmp <= 49323) Then
            getpychar = "k"
        ElseIf (tmp >= 49324 And tmp <= 49895) Then
            getpychar = "l"
        ElseIf (tmp >= 49896 And tmp <= 50370) Then
            getpychar = "m"
        ElseIf (tmp >= 50371 And tmp <= 50613) Then
            getpychar = "n"
        ElseIf (tmp >= 50614 And tmp <= 50621) Then
            getpychar = "o"
        ElseIf (tmp >= 50622 And tmp <= 50905) Then
            getpychar = "p"
        ElseIf (tmp >= 50906 And tmp <= 51386) Then
            getpychar = "q"
        ElseIf (tmp >= 51387 And tmp <= 51445) Then
            getpychar = "r"
        ElseIf (tmp >= 51446 And tmp <= 52217) Then
            getpychar = "s"
        ElseIf (tmp >= 52218 And tmp <= 52697) Then
            getpychar = "t"
        ElseIf (tmp >= 52698 And tmp <= 52979) Then
            getpychar = "w"
        ElseIf (tmp >= 52980 And tmp <= 53640) Then
            getpychar = "x"
        ElseIf (tmp >= 53679 And tmp <= 54480) Then
            getpychar = "y"
        ElseIf (tmp >= 54481 And tmp <= 62289) Then
            getpychar = "z"
        Else
            getpychar = Mid(Txt, i, 1)
        End If
        PinYinInitial = PinYinInitial & getpychar
    Next i
End Function
 
'�༭�������ƶ��㷨 �����ַ���˳�� ����FindStr��arrλ�� SimilarityΪ��С���ƶ�
'����ôģ����Ч�ʾͱ�Ҫ����
Public Function StrFindSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long
    Dim maxRow As Long, maxSIMILAR As Double, linshiSIMILAR As Double, i As Long, v
    i = 1
    maxSIMILAR = 0
    linshiSIMILAR = 0
    For Each v In arr
        linshiSIMILAR = StrSimilar(FindStr, v)
        If maxSIMILAR < linshiSIMILAR Then
            maxSIMILAR = linshiSIMILAR
            maxRow = i
        End If
        i = i + 1
    Next
    If maxSIMILAR >= Similarity Then StrFindSimilar = maxRow Else StrFindSimilar = 0
End Function
 
'�������ƶ��㷨 �����ַ���˳�� ����FindStr��arrλ�� SimilarityΪ��С���ƶ�
'����ôģ����Ч�ʾͱ�Ҫ����
Public Function StrFindCosineSimilar(FindStr, arr, Optional Similarity As Double = 60) As Long
    Dim maxRow As Long, maxSIMILAR As Double, linshiSIMILAR As Double, i As Long, v
    i = 1
    maxSIMILAR = 0
    linshiSIMILAR = 0
    For Each v In arr
        linshiSIMILAR = StrCosineSimilar(FindStr, v)
        If maxSIMILAR < linshiSIMILAR Then
            maxSIMILAR = linshiSIMILAR
            maxRow = i
        End If
        i = i + 1
    Next
    If maxSIMILAR >= Similarity Then StrFindCosineSimilar = maxRow Else StrFindCosineSimilar = 0
End Function
 
'�༭�������ƶ��㷨 �ж��ַ���S1��S2�����ƶ�,�����ַ���˳��,���ƶ�����Ϊ0-100,100Ϊ��ȫһ��
Public Function StrSimilar(s1, s2) As Double
    Dim Str_l() As String
    Dim Str_s() As String
    Dim str_chg() As Integer
    Dim Str_new() As String
    Dim Matrix1() As Integer
    Dim Matrix2() As Integer
    Dim Matrix3() As Integer
    Dim n  As Integer
    Dim n1 As Integer
    Dim n2 As Integer
    Dim Longer As Integer
    Dim Shorter As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim Max As Integer
    Dim Max1 As Integer
    Dim Max2 As Integer
    Dim Line_Best() As Integer
    If s1 = "" Or s2 = "" Then
        StrSimilar = -1
        Exit Function
    End If
    n1 = Len(s1)
    n2 = Len(s2)
    n = Abs(n1 - n2) + 1
    If n1 >= n2 Then
        Longer = n1
        Shorter = n2
    Else
        Longer = n2
        Shorter = n1
    End If
    ReDim Str_l(Longer)
    ReDim Str_s(Shorter)
    ReDim str_chg(Shorter)
    ReDim Str_new(Longer)
    ReDim Matrix1(Shorter, Longer)
    ReDim Matrix2(n, Shorter)
    ReDim Matrix3(n, Shorter)
    ReDim Line_Best(n)
    If n1 >= n2 Then
        For i = 1 To Longer
            Str_l(i) = VBA.Mid(s1, i, 1)
        Next
        For i = 1 To Shorter
            Str_s(i) = VBA.Mid(s2, i, 1)
        Next
    Else
        For i = 1 To Longer
            Str_l(i) = VBA.Mid(s2, i, 1)
        Next
        For i = 1 To Shorter
            Str_s(i) = VBA.Mid(s1, i, 1)
        Next
    End If
    For i = 1 To Longer
        For j = 1 To Shorter
            If Str_l(i) = Str_s(j) Then
                Matrix1(j, i) = 1
            Else
                Matrix1(j, i) = 0
            End If
        Next
    Next
    For i = 1 To n
        k = 1
        l = i
        For j = 1 To Shorter
            Matrix2(i, j) = Matrix1(k, l)
            k = k + 1
            l = l + 1
        Next
    Next
    For i = 1 To n
        For j = 1 To Shorter
            If Matrix2(i, j) = 1 Then
               Matrix2(i, j) = i
            End If
        Next
    Next
    i = 0
    j = 0
    k = 0
    l = 0
    For i = n To 2 Step -1
        Max1 = 0
        For j = 1 To Shorter
            Max = 0
            For k = 1 To j
                If Matrix2(i - 1, k) <> 0 Then
                    Max = Max + 1
                End If
            Next
            For l = j + 1 To Shorter
                If Matrix2(i, l) <> 0 Then
                    Max = Max + 1
                End If
            Next
            If Max1 < Max Then
                Max1 = Max
                Max2 = j
            End If
        Next
        Line_Best(i - 1) = Max2
    Next
    i = 0
    j = 0
    k = 0
    l = 0
    For i = n To 1 Step -1
        If i = n Then
            For j = 1 To Shorter
                Matrix3(i, j) = Matrix2(i, j)
            Next
        Else
            For j = 1 To Line_Best(i)
                Matrix3(i, j) = Matrix2(i, j)
            Next
            For j = j To Shorter
                Matrix3(i, j) = Matrix3(i + 1, j)
            Next
        End If
    Next
    Matrix3(1, 1) = 1
    For j = 2 To Shorter
        If Matrix3(1, j) = 0 Then
            Matrix3(1, j) = Matrix3(1, j - 1)
        End If
    Next
    For i = 1 To Shorter
        If i = 1 Then
            str_chg(i) = 0
        Else
            str_chg(i) = Matrix3(1, i) - Matrix3(1, i - 1)
        End If
    Next
    k = 1
    j = 1
    l = 1
    For i = 1 To Longer
        If k <= Shorter Then
            If str_chg(k) = 0 Then
                Str_new(i) = Str_s(l)
                i = i + 1
            Else
                For j = 1 To str_chg(k)
                    Str_new(i) = ""
                    i = i + 1
                Next
                Str_new(i) = Str_s(l)
                i = i + 1
            End If
            l = l + 1
            k = k + 1
            i = i - 1
        End If
    Next
    i = 1
    l = 1
    For i = 1 To Longer
        If Str_l(i) <> Str_new(i) Then
            l = l + 1
        End If
    Next
    l = i - l
    StrSimilar = (l / Longer) ^ 2 * 100
End Function
 
'�������ƶ��㷨 �ж��ַ���S1��S2�����ƶ�,�����ַ���˳��,���ƶ�����Ϊ0-100,100Ϊ��ȫһ��
Public Function StrCosineSimilar(strA, strB) As Double
    Dim objDic_All As Object, objDic_1 As Object, objDic_2 As Object
    Dim lngID As Long, StrKey As String
    Dim arrKey As Variant, arrResult As Variant
    Dim dblSum As Double, dblVal_A As Double, dblVal_B As Double
    If strA = "" Or strB = "" Then
        StrCosineSimilar = 0
        Exit Function
    End If
    Set objDic_All = CreateObject("Scripting.Dictionary")
    Set objDic_1 = CreateObject("Scripting.Dictionary")
    Set objDic_2 = CreateObject("Scripting.Dictionary")
    For lngID = 1 To Len(strA)
        StrKey = Trim(Mid(strA, lngID, 1))
        If StrKey <> "" Then
            objDic_All(StrKey) = ""
            objDic_1(StrKey) = Val(objDic_1(StrKey)) + 1
        End If
    Next
    For lngID = 1 To Len(strB)
        StrKey = Trim(Mid(strB, lngID, 1))
        If StrKey <> "" Then
            objDic_All(StrKey) = ""
            objDic_2(StrKey) = Val(objDic_2(StrKey)) + 1
        End If
    Next
    arrKey = objDic_All.Keys
    ReDim arrResult(LBound(arrKey) To UBound(arrKey), 1 To 3)
    For lngID = LBound(arrKey) To UBound(arrKey)
        arrResult(lngID, 1) = arrKey(lngID)
        arrResult(lngID, 2) = objDic_1(arrKey(lngID)) + 0
        arrResult(lngID, 3) = objDic_2(arrKey(lngID)) + 0
    Next
    Set objDic_All = Nothing
    Set objDic_1 = Nothing
    Set objDic_2 = Nothing
    For lngID = LBound(arrResult) To UBound(arrResult)
        dblSum = dblSum + arrResult(lngID, 2) * arrResult(lngID, 3)
        dblVal_A = dblVal_A + arrResult(lngID, 2) ^ 2
        dblVal_B = dblVal_B + arrResult(lngID, 3) ^ 2
    Next
    StrCosineSimilar = dblSum / (Sqr(dblVal_A) * Sqr(dblVal_B)) * 100
End Function

'�ڲ�ʹ�� ���������ת��
Private Function RegExp_Pattern_Modify_(ByVal Pattern) As String
    Pattern = Replace(Pattern, "\", "\\")
    Pattern = Replace(Pattern, ".", "\.")
    Pattern = Replace(Pattern, "?", "\?")
    Pattern = Replace(Pattern, "*", "\*")
    Pattern = Replace(Pattern, "+", "\+")
    Pattern = Replace(Pattern, "$", "\$")
    Pattern = Replace(Pattern, "^", "\^")
    Pattern = Replace(Pattern, "(", "\(")
    Pattern = Replace(Pattern, ")", "\)")
    Pattern = Replace(Pattern, "[", "\[")
    Pattern = Replace(Pattern, "]", "\]")
    Pattern = Replace(Pattern, "{", "\{")
    Pattern = Replace(Pattern, "}", "\}")
    RegExp_Pattern_Modify_ = Pattern
End Function
 
'����ȡ����ֵ
Public Function StrRegexSearch( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByVal Index = 0, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = All
            .ignoreCase = ignoreCase
            .multiline = multiline
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    If Index > 0 Then
        Index = Index - 1
    ElseIf Index < 0 Then
        Index = Index + searchResults.Count
    End If
    If searchResults.Count > 0 Then
        StrRegexSearch = searchResults(Index).Value
    End If
End Function
 
'����ȡ����ƥ�䣬��������
Public Function StrRegexSearchs( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As Variant()
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = All
            .ignoreCase = ignoreCase
            .multiline = multiline
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    Dim i As Long
    ArrayDynamic_
    For i = 0 To searchResults.Count - 1
        ArrayDynamic_ searchResults(i).Value
    Next
    StrRegexSearchs = ArrayDynamic_
End Function
 
'����ȡ��һ��ֵ
Public Function StrRegexSearchOne( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As String
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = False
            .ignoreCase = ignoreCase
            .multiline = False
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    If searchResults.Count > 0 Then StrRegexSearchOne = searchResults(0).Value
End Function
 
'�������λ��
Public Function StrRegexInStr( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = False
            .ignoreCase = ignoreCase
            .multiline = False
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    If searchResults.Count > 0 Then StrRegexInStr = searchResults(0).FirstIndex + 1 Else StrRegexInStr = 0
End Function
 
'�������λ�� ����
Public Function StrRegexInStrRev( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Long
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = True
            .ignoreCase = ignoreCase
            .multiline = False
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    If searchResults.Count > 0 Then StrRegexInStrRev = searchResults(searchResults.Count - 1).FirstIndex + 1 Else StrRegexInStrRev = 0
End Function
 
'����ȡ������ƥ�䣬�����������()�ٶ�ά����
Public Function StrRegexSearchSub( _
        ByRef String1, _
        ByRef Pattern, _
        Optional ByRef All As Boolean = True, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Variant()
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = All
            .ignoreCase = ignoreCase
            .multiline = multiline
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    Dim i As Long, j As Long, arrRE()
    If searchResults.Count > 0 Then
        If searchResults(0).SubMatches.Count > 0 Then
            ReDim arrRE(1 To searchResults.Count, 1 To searchResults(0).SubMatches.Count)
            For i = 0 To searchResults.Count - 1
                For j = 0 To searchResults(i).SubMatches.Count - 1
                    arrRE(i + 1, j + 1) = searchResults(i).SubMatches(j)
                Next
            Next
        Else
            ReDim arrRE(1 To searchResults.Count, 1 To 1)
            For i = 0 To searchResults.Count - 1
                arrRE(i + 1, 1) = searchResults(i).Value
            Next
        End If
        StrRegexSearchSub = arrRE
    Else
        StrRegexSearchSub = Array()
    End If
End Function
 
'�������
Public Function StrRegexCount( _
        ByRef String1, _
        ByRef Pattern, _
        Optional ByRef ignoreCase As Boolean = False, _
        Optional ByRef multiline As Boolean = False) As Long
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = True  'ȫ��ƥ��
            .ignoreCase = ignoreCase '��Сд
            .multiline = multiline '����
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    Set searchResults = Regex.Execute(String1)
    StrRegexCount = searchResults.Count
End Function
 
'������֤
Public Function StrRegexTest( _
    ByRef String1, _
    ByRef Pattern, _
    Optional ByRef ignoreCase As Boolean = False) As Boolean
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        With Regex
            .Global = False
            .ignoreCase = ignoreCase
            .multiline = False
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    StrRegexTest = Regex.test(String1)
End Function
 
'�����滻
Public Function StrRegexReplace( _
    ByRef String1, _
    ByRef Pattern, _
    ByRef replacementString, _
    Optional ByRef All As Boolean = True, _
    Optional ByRef ignoreCase As Boolean = False, _
    Optional ByRef multiline As Boolean = False) As String
 
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> Pattern Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        With Regex
            .Global = All
            .ignoreCase = ignoreCase
            .multiline = multiline
            .Pattern = Pattern
        End With
        stringPattern = Pattern
    End If
    StrRegexReplace = Regex.Replace(String1, replacementString)
End Function
 
'ģ���ַ���
'Formatter("������{1},���䣺{2}","UFO",18)  ����"������UFO,���䣺18"
Public Function StrFormatter(ByVal formatString, ParamArray textArray() As Variant) As String
    Dim i As Byte
    Dim individualTextItem As Variant
    Dim individualValue As Variant
    i = 0
    For Each individualTextItem In textArray
        If IsArray(individualTextItem) Then
            For Each individualValue In individualTextItem
                i = i + 1
                formatString = VBA.Replace(formatString, "{" & i & "}", individualValue)
            Next
        Else
            i = i + 1
            formatString = VBA.Replace(formatString, "{" & i & "}", individualTextItem)
        End If
    Next
    StrFormatter = formatString
End Function
 
'������ת��ָ��������ı�
'"Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
Public Function ByteToStr(arrByte, Optional strCharset = "UTF-8") As String
    With CreateObject("Adodb.Stream")
        .Type = 1 'adTypeBinary
        .Open
        .Write arrByte
        .Position = 0
        .Type = 2 'adTypeText
        .Charset = strCharset
        ByteToStr = .Readtext
        .Close
    End With
End Function
 
'�ı���ָ������תΪ������
'"Unicode", "GB2312", "UTF-8", "ASCII", "GBK"
Public Function StrToByte(strText, Optional strCharset = "UTF-8")
    With CreateObject("adodb.stream")
        .Mode = 3 'adModeReadWrite
        .Type = 2 'adTypeText
        .Charset = strCharset
        .Open
        .Writetext strText
        .Position = 0
        .Type = 1 'adTypeBinary
        '.Position = 2 '����BOMͷ������д��룬ȥ�������ֽڵ�BOMͷ������3��ȥ�������ֽڵľ�����2
        StrToByte = .Read
        .Close
    End With
End Function
 
'URLת��
Public Function StrencodeURI(strText) As String
    Dim oDom, oWindow
    Set oDom = CreateObject("HTMLFILE")
    Set oWindow = oDom.parentWindow
    oDom.Write "<Script></Script>"
    strText = Replace(strText, vbCr, "")
    strText = Replace(strText, vbLf, "")
    StrencodeURI = oWindow.encodeURIComponent(strText)
End Function
 
'URL����
Public Function StrdecodeURI(strText) As String
    Dim oDom, oWindow
    Set oDom = CreateObject("HTMLFILE")
    Set oWindow = oDom.parentWindow
    oDom.Write "<Script></Script>"
    strText = Replace(strText, vbCr, "")
    strText = Replace(strText, vbLf, "")
    StrdecodeURI = oWindow.decodeURIComponent(strText)
End Function
 
'unicode�ַ�ת��������
Public Function StrConvert(ByVal strText) As String
    Dim oDom, oWindow
    Set oDom = CreateObject("HTMLFILE")
    Set oWindow = oDom.parentWindow
    oDom.Write "<Script></Script>"
    strText = Replace(strText, vbCr, "")
    strText = Replace(strText, vbLf, "")
    StrConvert = oWindow.eval("('" & strText & "').replace(/&#\d+;/g,function(b){return String.fromCharCode(b.slice(2,b.length-1))});")
End Function

'����Base64
Public Function StrencodeBase64(String1, Optional Charset = "") As String
    Dim b() As Byte
    With CreateObject("msxml2.domdocument").createelement("b64")
        .DataType = "bin.base64"
        If Charset = "" Then
            b = String1
            .nodetypedvalue = b
        Else
            .nodetypedvalue = StrToByte(String1, Charset)
        End If
        StrencodeBase64 = .Text
    End With
End Function

'����Base64
Public Function StrdecodeBase64(String1, Optional Charset = "") As String
    Dim Dom As Object
    Set Dom = CreateObject("msxml2.domdocument").createelement("b64")
    Dom.DataType = "bin.base64"
    Dom.Text = String1
    If Charset = "" Then
        StrdecodeBase64 = Dom.nodetypedvalue
    Else
        With CreateObject("Adodb.Stream")
            .Type = 1 'adTypeBinary
            .Open
            .Write Dom.nodetypedvalue
            .Position = 0
            .Type = 2 'adTypeText
            .Charset = "ASCII"
            StrdecodeBase64 = .Readtext
            .Close
        End With
    End If
End Function














 
'ϵͳ-------------------------------------------------------------------------------------------------------------------------------------
'�������ȡ
Public Function Clipboard_GetData() As String
    Dim oHTML As Object, clp As Object
    Set oHTML = CreateObject("htmlfile")
    Set clp = oHTML.parentWindow.clipboardData
    Dim s As Variant
    s = clp.GetData("text")
    If IsNull(s) Then
        Clipboard_GetData = ""
    Else
        Clipboard_GetData = s
    End If
End Function
 
'������д��
Public Function Clipboard_SetData(strData) As Boolean
    Dim oHTML As Object, clp As Object
    Set oHTML = CreateObject("htmlfile")
    Set clp = oHTML.parentWindow.clipboardData
    Clipboard_SetData = clp.setData("text", CStr(strData))
End Function
 
'���������
Public Function Clipboard_ClearData() As Boolean
    Dim oHTML As Object, clp As Object
    Set oHTML = CreateObject("htmlfile")
    Set clp = oHTML.parentWindow.clipboardData
    Clipboard_ClearData = clp.clearData("text")
End Function
 
'�û���
Public Function UserName() As String
    UserName = VBA.Environ("USERNAME")
End Function
 
'�û�������
Public Function UserDomain() As String
    UserDomain = VBA.Environ("USERDOMAIN")
End Function
 
'�������
Public Function ComputerName() As String
    ComputerName = VBA.Environ("COMPUTERNAME")
End Function





















'�ļ�-------------------------------------------------------------------------------------------------------------------------------------
'��ȡtxt�ļ�(ANSI����)
Public Function TextRead(TextPath) As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        With .OpenTextFile(TextPath, 1, False)
            TextRead = .ReadAll
            .Close
        End With
    End With
    Err.Clear
End Function
 
'д��txt�ļ�(ANSI����)
Public Function TextWrite(TextPath, str) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        With .OpenTextFile(TextPath, 2, True)
            .Write str
            .Close
        End With
    End With
    TextWrite = True
    If Err Then Err.Clear: TextWrite = False
End Function
 
'׷��txt�ļ�(ANSI����)
Public Function TextAppend(TextPath, str) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        With .OpenTextFile(TextPath, 8, True)
            .Write str
            .Close
        End With
    End With
    TextAppend = True
    If Err Then Err.Clear: TextAppend = False
End Function
 
'��ȡtxt�ļ�(�Զ������)
'"Unicode", "GB2312", "UTF-7", "UTF-8", "ASCII", "GBK", "Big5", "unicodeFEFF", "unicodeFFFE"
Public Function TextRead2(TextPath, Optional strCharset = "UTF-8") As String
    With CreateObject("Adodb.Stream")
        .Open
        .Type = 2
         .Charset = strCharset '"UTF-8"
        .LoadFromFile TextPath
        TextRead2 = .Readtext
        .Close
    End With
End Function
 
'д��txt�ļ�(�Զ������)
Public Function TextWrite2(TextPath, str, Optional strCharset = "UTF-8") As Boolean
    On Error Resume Next
    With CreateObject("Adodb.Stream")
        .Type = 2
        .Charset = strCharset
        .Open
        .Writetext str
        .SaveToFile TextPath, 2
        .Close
    End With
    TextWrite2 = True
    If Err Then Err.Clear: TextWrite2 = False
End Function
 
'׷��txt�ļ�(�Զ������)
Public Function TextAppend2(TextPath, str, Optional strCharset = "UTF-8") As Boolean
    With CreateObject("Adodb.Stream")
        .Type = 2
        .Charset = strCharset
        .Open
        .LoadFromFile TextPath
'        Do Until .EOS '������β
'           .SkipLine
'        Loop
        .Readtext '������β
        .Writetext str
        .SaveToFile TextPath, 2
        .Close
    End With
    TextAppend2 = True
    If Err Then Err.Clear: TextAppend2 = False
End Function
 
'���ļ�Ϊ�ֽ�����
Public Function FileToByte(strFileName) As Byte()
    With CreateObject("Adodb.Stream")
        .Open
        .Type = 1
        .LoadFromFile strFileName
        FileToByte = .Read
        .Close
    End With
End Function
 
'�ֽ�����ת�ļ�
Public Function ByteToFile(arrByte, strFileName)
    With CreateObject("Adodb.Stream")
        .Type = 1 'adTypeBinary
        .Open
        .Write arrByte
        .SaveToFile strFileName, 2
        .Close
    End With
End Function
 
'�ļ����Ƿ����
Public Function FolderExists(Path) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        FolderExists = .FolderExists(Path)
    End With
    Err.Clear
End Function
 
'ɾ���ļ���
Public Function FolderDelete(Path) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        .DeleteFolder Path
    End With
    FolderDelete = True
    If Err Then Err.Clear: FolderDelete = False
End Function
 
'�����ļ���
Public Function FolderCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
         .CopyFolder Source, Destination, OverWrite
    End With
    FolderCopy = True
    If Err Then Err.Clear: FolderCopy = False
End Function
 
'�����ļ��У����Դ����ϼ������ڵ��ļ��У������༶
Public Function FolderCreate(Path) As Boolean
    On Error Resume Next
    Dim i As Long, f As Object
    With CreateObject("Scripting.FileSystemObject")
        Dim pt As String, col As Collection
        Set col = New Collection
        Do Until Path = ""
            If .FolderExists(Path) Then
                Exit Do
            Else
                col.Add Path
            End If
            Path = .GetParentFolderName(Path)
        Loop
        For i = col.Count To 1 Step -1
            Set f = .CreateFolder(col.Item(i))
        Next
    End With
    FolderCreate = col.Count > 0
    Err.Clear
End Function
 
'�����ļ������ļ���
Public Function FolderSearch(pPath) As Variant
    On Error GoTo ErrFSO
    Dim Folder, f As Object, arr, i As Long
    Set Folder = CreateObject("Scripting.FileSystemObject").GetFolder(pPath).SubFolders
    ReDim arr(0 To Folder.Count - 1)
    i = 0
    For Each f In Folder
        arr(i) = f.Path & "\"
        i = i + 1
    Next
    FolderSearch = arr
    Exit Function
ErrFSO:
    FolderSearch = Array()
End Function
 
'�����ļ���(�����ļ���)
Public Function FolderSearchSub(pPath) As Variant
    Dim DirFile, mf As Long, pPath1 As String, colQueue As Collection, fileNameDic As Variant
    On Error Resume Next
    Set colQueue = New Collection
    Set fileNameDic = CreateObject("scripting.dictionary")
    pPath = Trim(pPath)
    If Right(pPath, 1) <> "\" Then pPath = pPath & "\"
    pPath1 = pPath
    Do Until colQueue Is Nothing
        DirFile = Dir(pPath1, vbDirectory)
        Do While DirFile <> ""
            If DirFile <> "." And DirFile <> ".." Then
                If (GetAttr(pPath1 & DirFile) And vbDirectory) = vbDirectory Then
                    colQueue.Add pPath1 & DirFile & "\", pPath1 & DirFile & "\"
                    fileNameDic.Add pPath1 & DirFile & "\", pPath1 & DirFile & "\"
                End If
            End If
            DirFile = Dir
        Loop
        If colQueue.Count > 0 Then
            pPath1 = colQueue(1)
            colQueue.Remove (1)
        Else
            Set colQueue = Nothing
        End If
    Loop
    FolderSearchSub = fileNameDic.Keys
End Function
 
'�ļ��Ƿ����
Public Function FileExists(Path) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        FileExists = .FileExists(Path)
    End With
    Err.Clear
End Function
 
'ɾ���ļ�
Public Function FileDelete(Path) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        .DeleteFile Path
    End With
    FileDelete = True
    If Err Then Err.Clear: FileDelete = False
End Function
 
'�����ļ�
Public Function FileCopy(Source, Destination, Optional OverWrite As Boolean = True) As Boolean
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
         .CopyFile Source, Destination, OverWrite
    End With
    FileCopy = True
    If Err Then Err.Clear: FileCopy = False
End Function
 
'�����ļ������ļ�
Public Function FileSearch(pPath) As Variant
    On Error GoTo ErrFSO
    Dim Folder, f As Object, arr, i As Long
    Set Folder = CreateObject("Scripting.FileSystemObject").GetFolder(pPath).Files
    ReDim arr(0 To Folder.Count - 1)
    i = 0
    For Each f In Folder
        arr(i) = f.Path
        i = i + 1
    Next
    FileSearch = arr
    Exit Function
ErrFSO:
    FileSearch = Array()
End Function
 
'�����ļ������ļ�(�����ļ���)
'pPath������ʼ·����pMask���Ҫ����д,�ǵ�������д"*.xlsx",���Ǻ�
Public Function FileSearchSub(pPath, Optional pMask As String = "") As Variant
    Dim DirFile, mf As Long, pPath1 As String, colQueue As Collection, fileNameDic As Variant
    On Error Resume Next
    Set colQueue = New Collection
    Set fileNameDic = CreateObject("scripting.dictionary")
    pPath = Trim(pPath)
    If Right(pPath, 1) <> "\" Then pPath = pPath & "\"
    pPath1 = pPath
    Do Until colQueue Is Nothing
        DirFile = Dir(pPath1 & pMask)
        Do While DirFile <> ""
            fileNameDic.Add pPath1 & DirFile, pPath1 & DirFile
            DirFile = Dir
        Loop
        DirFile = Dir(pPath1, vbDirectory)
        Do While DirFile <> ""
            If DirFile <> "." And DirFile <> ".." Then
                If (GetAttr(pPath1 & DirFile) And vbDirectory) = vbDirectory Then
                    colQueue.Add pPath1 & DirFile & "\", pPath1 & DirFile & "\"
                End If
            End If
            DirFile = Dir
        Loop
        If colQueue.Count > 0 Then
            pPath1 = colQueue(1)
            colQueue.Remove (1)
        Else
            Set colQueue = Nothing
        End If
    Loop
    FileSearchSub = fileNameDic.Keys
End Function
 
'·��-------------------------------------------------------------------------------------------------------------------------------------
'������ʱ·��
Public Function PathGetTemp() As String
    On Error Resume Next
    PathGetTemp = VBA.Environ("TEMP")
End Function
 
'�����ĵ�·��
Public Function PathGetMyDocuments() As String
    On Error Resume Next
    With CreateObject("WScript.Shell")
        PathGetMyDocuments = .SpecialFolders("MyDocuments")
    End With
End Function
 
'��������·��
Public Function PathGetDesktop() As String
    On Error Resume Next
    With CreateObject("WScript.Shell")
        PathGetDesktop = .SpecialFolders("Desktop")
    End With
End Function
 
'�����ļ�����������չ��
Public Function PathBaseName(Path) As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        PathBaseName = .GetBaseName(Path)
    End With
    Err.Clear
End Function
 
'�����ļ�����������չ��
Public Function PathFileName(Path) As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        PathFileName = .GetFileName(Path)
    End With
    Err.Clear
End Function
 
'������չ��������.
Public Function PathExtensionName(Path) As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        PathExtensionName = .GetExtensionName(Path)
    End With
    Err.Clear
End Function
 
'����·��,ĩβ����\
Public Function PathParentFolderName(Path) As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        PathParentFolderName = .GetParentFolderName(Path)
    End With
    Err.Clear
End Function
 
'�ж��Ƿ����ļ���
Public Function PathIsFolder(Path) As Boolean
    On Error Resume Next
    PathIsFolder = (GetAttr(Path) And vbDirectory) = vbDirectory
    Err.Clear
End Function
 
'����ļ���
Public Function PathTempName() As String
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        PathTempName = .GetTempName()
    End With
    Err.Clear
End Function

'�����ظ�ʱ�����Ƽ���� Name��ǰ���� DelimiterLeft������ָ��� DelimiterRight����Ҳ�ָ���
Public Function PathNameSerialNumber(Name, Optional DelimiterLeft = "(", Optional DelimiterRight = ")") As String
    On Error Resume Next
    Static stringPattern As String
    Static Regex As Object
    If stringPattern <> DelimiterLeft & DelimiterRight Or Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
        Dim searchResults As Object
        With Regex
            .Global = False
            .ignoreCase = False
            .multiline = False
            .Pattern = "^(.+)" & RegExp_Pattern_Modify_(DelimiterLeft) & "(\d+)" & RegExp_Pattern_Modify_(DelimiterRight) & "$"
        End With
        stringPattern = DelimiterLeft & DelimiterRight
    End If
    Set searchResults = Regex.Execute(Name)
    If searchResults.Count > 0 Then
        PathNameSerialNumber = searchResults(0).SubMatches(0)
        PathNameSerialNumber = PathNameSerialNumber & DelimiterLeft & VBA.Val(searchResults(0).SubMatches(1)) + 1 & DelimiterRight
    Else
        PathNameSerialNumber = Name & DelimiterLeft & 1 & DelimiterRight
    End If
End Function



 
'��Ԫ��-----------------------------------------------------------------------------------------------------------------------------------
'����ת��ĸ
Public Function ColumnChr(ByVal v) As String
    Do
        ColumnChr = Chr((v - 1) Mod 26 + 65) & ColumnChr
        v = (v - 1) \ 26
    Loop Until v = 0
End Function
 
'����ת��ĸArr
Function ColumnChrArr(ParamArray arr()) As Variant
    Dim i As Long, parr
    parr = ArrFlatten(arr)
    For i = LBound(parr) To UBound(parr)
        parr(i) = ColumnChr(parr(i))
    Next
    ColumnChrArr = parr
End Function
 
'��ĸת����
Public Function ColumnI(ByVal s) As Long
    s = Ucase(s)
    Dim i As Long, l As Long: l = Len(s)
    For i = 1 To l
        ColumnI = ColumnI + (Asc(Mid(s, i, 1)) - 64) * 26 ^ (l - i)
    Next
End Function
 
'��ĸת����Arr
Function ColumnIArr(ParamArray arr()) As Variant
    Dim i As Long, parr
    parr = ArrFlatten(arr)
    For i = LBound(parr) To UBound(parr)
        parr(i) = ColumnI(parr(i))
    Next
    ColumnIArr = parr
End Function
 
'��Ԫ�񲢼���չ,���뵥Ԫ������򼯺ϵ�Range���󣬺ϲ���Range
Public Function UnionEx(ByRef Rngs) As Range
    Dim i As Long, s As String, l As Long
    Dim rng As Range, Are As String
    Dim sh As Worksheet
    Call StringBuilder_
    For Each rng In Rngs
        If Not rng Is Nothing Then
            Are = rng.Address(False, False)
            If l + Len(Are) > 255 Then
                s = Left(StringBuilder_(), l - 1)
                If sh Is Nothing Then
                    Set sh = rng.Parent: Set UnionEx = sh.Range(s)
                Else
                    Set UnionEx = Application.Union(UnionEx, sh.Range(s))
                End If
            End If
            l = StringBuilder_(Are & ",")
        End If
    Next
    s = Left(StringBuilder_(), l - 1)
    Set UnionEx = Application.Union(UnionEx, sh.Range(s))
End Function
 
'��Ԫ�񲢼���չ,���뵥Ԫ������򼯺ϵ��ַ�����ַ���ϲ���Range
Public Function UnionEx_Str(ByRef Rngs, sh) As Range
    Dim i As Long, s As String, l As Long
    Dim Are
    Call StringBuilder_
    For Each Are In Rngs
        If l + Len(Are) > 255 Then
            s = Left(StringBuilder_(), l - 1)
            If UnionEx_Str Is Nothing Then
                 Set UnionEx_Str = sh.Range(s)
            Else
                 Set UnionEx_Str = Application.Union(UnionEx_Str, sh.Range(s))
            End If
        End If
        l = StringBuilder_(Are & ",")
    Next
    s = Left(StringBuilder_(), l - 1)
    Set UnionEx_Str = Application.Union(UnionEx_Str, sh.Range(s))
End Function
 
'ĩβ����������
Public Function SheetNew(wb As Workbook, Optional Name As String = "") As Worksheet
    Dim Ash
    With wb
        Set Ash = ActiveSheet
        Set SheetNew = .Worksheets.Add(, .Worksheets(.Worksheets.Count))
        If Name <> "" Then SheetNew.Name = Name
        Ash.Activate
    End With
End Function
 
'���ƹ�����ĩβ
Public Function SheetCopyAfter(sh, Optional Name As String = "") As Worksheet
    Dim Ash
    With sh.Parent
        Set Ash = ActiveSheet
         sh.Copy After:=.Worksheets(.Worksheets.Count)
         Set SheetCopyAfter = .Worksheets(.Worksheets.Count)
        If Name <> "" Then SheetCopyAfter.Name = Name
        Ash.Activate
    End With
End Function
 
'���ƹ������¹�����
Public Function SheetCopyNow(sh, Optional Name As String = "") As Worksheet
    Dim Ash
    With sh.Parent
        Set Ash = ActiveSheet
         sh.Copy
         Set SheetCopyNow = ActiveSheet
        If Name <> "" Then SheetCopyNow.Name = Name
        Ash.Activate
    End With
End Function
 
'��鹤�����Ƿ����
Public Function SheetIsName(wb As Workbook, ByVal Name As String) As Boolean
    Dim sh: SheetIsName = False
    Name = Lcase(Name)
    With wb
        For Each sh In .Sheets
            If Lcase(sh.Name) Like Name Then
                SheetIsName = True
                Exit For
            End If
        Next
    End With
End Function
 
'��鹤�����Ƿ���ڣ�Name��������׺
Public Function WorkbookIsName(ByVal Name As String) As Boolean
    Dim wb: WorkbookIsName = False
    Name = Lcase(Name)
    For Each wb In Application.Workbooks
        If StrGetLeftRev(Lcase(wb.Name), ".") Like Name Then
            WorkbookIsName = True
            Exit For
        End If
    Next
End Function
 
'����д�빤����
Public Function ArrToRange(ByRef arr, ByVal rng)
    Dim rn As Range
    If TypeName(rng) = "String" Then Set rn = Range(rng) Else Set rn = rng
    Select Case ArrDimension(arr)
        Case 0
            rn.Value = arr
        Case 2
            Set rn = rn.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, _
                UBound(arr, 2) - LBound(arr, 2) + 1)
            rn.Value = arr
        Case 1
            Set rn = rn.Cells(1, 1).Resize(1, UBound(arr) - LBound(arr) + 1)
            rn.Value = arr
    End Select
End Function
 
'����д�빤���������
Public Function ArrToRangeUndo(ByRef arr, ByVal rng)
    Dim rn As Range
    If TypeName(rng) = "String" Then Set rn = Range(rng) Else Set rn = rng
    Select Case ArrDimension(arr)
        Case 0
            RangAddUndo rn
            rn.Value = arr
            RangStartUndo
        Case 2
            Set rn = rn.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, _
                UBound(arr, 2) - LBound(arr, 2) + 1)
            RangAddUndo rn
            rn.Value = arr
            RangStartUndo
        Case 1
            Set rn = rn.Cells(1, 1).Resize(1, UBound(arr) - LBound(arr) + 1)
            RangAddUndo rn
            rn.Value = arr
            RangStartUndo
    End Select
End Function
 
'��ӳ�������
Public Function RangAddUndo(ByVal rng)
    If TypeName(rng) = "String" Then Set rng = Range(rng)
    RangeUndoCollection_.Add rng.Address(External:=True)
    RangeUndoCollection_.Add rng.Value(11)
End Function
 
'��������
Public Function RangStartUndo()
    Dim EndIndex As Long
    EndIndex = RangeUndoCollection_.Count
    If EndIndex > 1 Then
        Application.OnUndo Range(RangeUndoCollection_.Item(EndIndex - 1)).Address(False, False), "RangeUndo_"
    End If
End Function
 
'��������д�빤����
Private Sub RangeUndo_()
    Dim EndIndex As Long
    EndIndex = RangeUndoCollection_.Count
    If EndIndex > 1 Then
        Range(RangeUndoCollection_.Item(EndIndex - 1)).Value(11) = RangeUndoCollection_.Item(EndIndex)
        RangeUndoCollection_.Remove EndIndex
        RangeUndoCollection_.Remove EndIndex - 1
        Application.OnTime Now + TimeValue("00:00:01") * 0.5, "RangStartUndo"
    End If
End Sub
 
'��Ԫ����������չ����
Public Function RngResizeDownRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range
    Dim rn As Range
    If TypeName(rng) = "String" Then
        Set rn = Range(rng)
    Else
        Set rn = rng
    End If
     Set RngResizeDownRow = rn.Resize(RngDownRow(rn, FilterShowAllData, CancelHidden) - rn.Row + 1)
End Function
 
'��Ԫ����������չ����
Public Function RngResizeRightColumn(ByRef rng) As Range
    Dim rn As Range
    If TypeName(rng) = "String" Then
        Set rn = Range(rng)
    Else
        Set rn = rng
    End If
    Set RngResizeRightColumn = rn.Resize(, RngRightColumn(rn) - rn.Column + 1)
End Function
 
'��Ԫ�������һ����չ����
Public Function RngResizeEndRow(ByRef rng, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Range
    Dim rn As Range
    If TypeName(rng) = "String" Then
        Set rn = Range(rng)
    Else
        Set rn = rng
    End If
     Set RngResizeEndRow = rn.Resize(RngEndRow(rn, FilterShowAllData, CancelHidden) - rn.Row + 1)
End Function
 
'��Ԫ�������һ����չ����
Public Function RngResizeEndColumn(ByRef rng) As Range
    Dim rn As Range
    If TypeName(rng) = "String" Then
        Set rn = Range(rng)
    Else
        Set rn = rng
    End If
    Set RngResizeEndColumn = rn.Resize(, RngEndColumn(rn) - rn.Column + 1)
End Function
 
'��Ԫ������һ��
Public Function RngDownRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long
    Dim maxRow As Long, rn As Range
    With rng.Parent
        If FilterShowAllData And .FilterMode Then .ShowAllData
        If CancelHidden Then .Rows.Hidden = False
        Dim shend As Long: shend = .Rows.Count
        Dim rnend As Long
        maxRow = rng.Row
        For Each rn In rng.Rows(1).Cells
            rnend = rn.End(xlDown).Row
            If maxRow < rnend And rnend <> shend Then maxRow = rnend
        Next
    End With
    RngDownRow = maxRow
End Function
 
'��Ԫ������һ��
Public Function RngRightColumn(ByRef rng As Range) As Long
    Dim maxColumn As Long, rn As Range
    With rng.Parent
        Dim shend As Long: shend = .Columns.Count
        Dim rnend As Long
        maxColumn = rng.Column
        For Each rn In rng.Columns(1).Cells
            rnend = rn.End(xlToRight).Column
            If maxColumn < rnend And rnend <> shend Then maxColumn = rnend
        Next
    End With
    RngRightColumn = maxColumn
End Function
 
'��Ԫ�����һ��
Public Function RngEndRow(ByRef rng As Range, Optional FilterShowAllData As Boolean = False, Optional CancelHidden As Boolean = False) As Long
    Dim maxRow As Long, rn As Range
    With rng.Parent
        If FilterShowAllData And .FilterMode Then .ShowAllData
        If CancelHidden Then .Rows.Hidden = False
        Dim shend As Long: shend = .Rows.Count
        Dim rnend As Long
        maxRow = rng.Row
        For Each rn In rng.Rows(1).Cells
            rnend = .Cells(shend, rn.Column).End(xlUp).Row
            If maxRow < rnend Then maxRow = rnend
        Next
    End With
    RngEndRow = maxRow
End Function
 
'��Ԫ�����һ��
Public Function RngEndColumn(ByRef rng As Range) As Long
    Dim maxColumn As Long, rn As Range
    With rng.Parent
        Dim shend As Long: shend = .Columns.Count
        Dim rnend As Long
        maxColumn = rng.Column
        For Each rn In rng.Columns(1).Cells
            rnend = .Cells(rn.Row, shend).End(xlToLeft).Column
            If maxColumn < rnend Then maxColumn = rnend
        Next
    End With
    RngEndColumn = maxColumn
End Function
 
'��Ԫ��ֵ������,��֤һ����Ԫ��Ҳ������
Public Function RangeToArr(rng As Range) As Variant
    Dim arr(1 To 1, 1 To 1), i As Long
    RangeToArr = rng.Value
    If Not ArrIsValid(RangeToArr) Then
        arr(1, 1) = RangeToArr
        RangeToArr = arr
    End If
End Function
 
'���ºϲ���ֵ��Ԫ��
Public Sub RngMerge_Empty(MergeRng As Range)
    On Error Resume Next
    Dim rng As Range
    For Each rng In MergeRng
        If rng = "" Then rng.Offset(-1).Resize(2).Merge
    Next
End Sub
 
'�ظ�ֵ�ϲ���Ԫ��
Public Sub RngMerge_Repeat(MergeRng As Range)
    On Error Resume Next
    Dim rng As Range
    Application.DisplayAlerts = False
    For Each rng In MergeRng
        If rng = rng.Offset(-1) Or rng = rng.Offset(-1).MergeArea(1, 1) Then
            rng.Offset(-1).Resize(2).Merge
        End If
    Next
    Application.DisplayAlerts = True
End Sub
 
'�ӿ���
Public Sub RngAddBorders(rng As Range)
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
 
'��Ԫ����ж���
Public Sub RngAlignmentCenter(rng As Range)
    With rng
        .HorizontalAlignment = xlCenter 'ˮƽ���뷽ʽ
        .VerticalAlignment = xlCenter '��ֱ���뷽ʽ
    End With
End Sub
 
'���ܹ����� SelectName�����Ĺ������� RemoveName�ų��Ĺ������� RngAddress��Ԫ������Ĭ��UsedRange  wb������Ĭ�ϵ�ǰ
Public Function SheetsSummary(Optional SelectName = "*", Optional RemoveName = "", Optional RngAddress = "", Optional wb As Workbook = Nothing) As Variant
    Dim sh As Worksheet, rng As Range
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook
    ArrayDynamic_
    For Each sh In wb.Worksheets
        If sh.Name Like SelectName Then
            If Not sh.Name Like RemoveName Then
                If RngAddress = "" Then
                    Set rng = sh.UsedRange
                Else
                    Set rng = sh.Range(RngAddress)
                End If
                ArrayDynamic_ RangeToArr(rng)
            End If
        End If
    Next
    SheetsSummary = ArrMergeRow(ArrayDynamic_)
End Function
 
'��������͸�ӱ� SourceData����Դ��Ԫ�� TableDestination���õ�Ԫ�� TableName͸�ӱ�����
'Sub test()����
'    Dim PC As PivotTable
'    Set PC = UCreatePivotTable(Range("A1:D27"), Range("F6"), "����͸�ӱ�1")
'    USetPivotField PC, "ҵ��Ա", xlRowField, 1
'    USetPivotField PC, "�ͻ�����", xlRowField, 2
'    USetPivotField PC, "�ɽ����", xlDataField, 1, "�ɽ������", xlSum
'    USetPivotField PC, "�ɽ����", xlDataField, 2, "�ɽ�������", xlCount
'End Sub
Public Function UCreatePivotTable(SourceData As Range, TableDestination As Range, TableName) As PivotTable
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim PC As PivotCache
    Set PC = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SourceData, Version:=xlPivotTableVersion14)
    Set UCreatePivotTable = PC.CreatePivotTable(TableDestination:=TableDestination, TableName:=TableName, DefaultVersion:=xlPivotTableVersion14)
    UCreatePivotTable.RowAxisLayout xlTabularRow
    UCreatePivotTable.RepeatAllLabels xlRepeatLabels
    Application.ScreenUpdating = True
End Function
 
'����͸�ӱ��ֶ� PTable͸�ӱ����UCreatePivotTable����ֵ  FieldName������
'Orientation �ֶ�λ������ xlRowField(�б�ǩ) xlColumnField(�б�ǩ) xlDataField(����)
'Position �ֶ�˳��
'Caption  �ֶα���
'Fun   Orientation=xlDataField(����)ʱ ���û��ܷ�ʽ��xlSum  xlCount  xlMin  xlMax
Public Sub USetPivotField(PTable As PivotTable, FieldName As String, Orientation As XlPivotFieldOrientation, _
        Position As Long, Optional Caption As String = "", Optional Fun As XlConsolidationFunction = xlCount)
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim i As Long
    With PTable.PivotFields(FieldName)
        .Orientation = Orientation
        .Position = Position
        If Caption <> "" Then .Caption = Caption
        If Orientation = xlDataField Then .Function = Fun
        If Orientation = xlColumnField Or Orientation = xlRowField Then
            For i = LBound(.Subtotals) To UBound(.Subtotals)
                .Subtotals(i) = False
            Next
        End If
    End With
    Application.ScreenUpdating = True
End Sub



'����������ʽ  Rng������ʽ��Χ  Formula��ʽ  Color��ɫRGBֵ
Public Function FormatConditionAdd(rng As Range, Formula, Color) As FormatCondition
    Dim FC As FormatCondition
    Set FC = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=Formula)
    FC.SetFirstPriority
    With FC.Interior
        .Color = Color
    End With
    FC.StopIfTrue = False
    Set FormatConditionAdd = FC
End Function
 
'����������ʽͼ��  Rng������ʽ��Χ  Formula��ʽ  PatternColor��ɫRGBֵ
Public Function FormatConditionAdd_Pattern(rng As Range, Formula, PatternColor, Optional Pattern As XlPattern = xlPatternGray50) As FormatCondition
    Dim FC As FormatCondition
    Set FC = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=Formula)
    FC.SetFirstPriority
    With FC.Interior
        .PatternColor = PatternColor
        .Pattern = Pattern
    End With
    FC.StopIfTrue = False
    Set FormatConditionAdd_Pattern = FC
End Function
 
'����ʽ����������ʽ
Public Function FormatConditionFind(rng As Range, ByVal Formula) As FormatCondition
    Dim FC As FormatCondition
    Formula = VBA.Ucase(Formula)
    For Each FC In rng.FormatConditions
        If VBA.Ucase(FC.Formula1) Like Formula Then Set FormatConditionFind = FC: Exit For
    Next
End Function
 
'����ɫ����������ʽ
Public Function FormatConditionFind_Color(rng As Range, Color) As FormatCondition
    Dim FC As FormatCondition
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Color = Color Then Set FormatConditionFind_Color = FC: Exit For
        End With
    Next
End Function
 
'��ͼ������������ʽ
Public Function FormatConditionFind_Pattern(rng As Range, Pattern As XlPattern, PatternColor) As FormatCondition
    Dim FC As FormatCondition
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Pattern = Pattern And .PatternColor = PatternColor Then Set FormatConditionFind_Pattern = FC: Exit For
        End With
    Next
End Function
  
'����ʽ����������ʽ����  ע��Formula:="=ROW($A1)=*"�Ǵ���д�� ������A1������A65536 ����Formula:="=ROW($A*)=*"
Public Function FormatConditionFindCount(rng As Range, ByVal Formula) As Long
    Dim FC As FormatCondition, k As Long
    Formula = VBA.Ucase(Formula)
    For Each FC In rng.FormatConditions
        If VBA.Ucase(FC.Formula1) Like Formula Then k = k + 1
    Next
    FormatConditionFindCount = k
End Function

'����ɫ����������ʽ����
Public Function FormatConditionFindCount_Color(rng As Range, Color) As Long
    Dim FC As FormatCondition, k As Long
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Color = Color Then k = k + 1
        End With
    Next
    FormatConditionFindCount_Color = k
End Function

'��ͼ������������ʽ����
Public Function FormatConditionFindCount_Pattern(rng As Range, Pattern As XlPattern, PatternColor) As Long
    Dim FC As FormatCondition, k As Long
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Pattern = Pattern And .PatternColor = PatternColor Then k = k + 1
        End With
    Next
    FormatConditionFindCount_Pattern = k
End Function

'������ʽ�޸Ĺ�ʽ
Public Function FormatConditionModify_Formula(FC As FormatCondition, Formula)
    FC.Modify Type:=xlExpression, Formula1:=Formula
End Function
 
'������ʽ�޸���ɫ
Public Function FormatConditionModify_Color(FC As FormatCondition, Color)
    With FC.Interior
        .Color = Color
    End With
End Function
 
'������ʽ�޸�ͼ����ɫ
Public Function FormatConditionModify_Pattern(FC As FormatCondition, Pattern As XlPattern, PatternColor)
    With FC.Interior
        .Pattern = Pattern
        .PatternColor = PatternColor
    End With
End Function
 
'������ʽ�����ɫ
Public Function FormatConditionModify_ClearColor(FC As FormatCondition)
    With FC.Interior
        .Pattern = xlPatternNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Function
 
'����ʽɾ��������ʽ ע��Formula:="=ROW($A1)=*"�Ǵ���д�� ������A1������A65536 ����Formula:="=ROW($A*)=*"
Public Function FormatConditionDelete(rng As Range, ByVal Formula)
    Dim FC As FormatCondition
    Formula = VBA.Ucase(Formula)
    For Each FC In rng.FormatConditions
        If VBA.Ucase(FC.Formula1) Like Formula Then FC.Delete
    Next
End Function
 
'����ɫɾ��������ʽ
Public Function FormatConditionDelete_Color(rng As Range, Color)
    Dim FC As FormatCondition
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Color = Color Then FC.Delete
        End With
    Next
End Function
 
'��ͼ��ɾ��������ʽ
Public Function FormatConditionDelete_Pattern(rng As Range, Pattern As XlPattern, PatternColor)
    Dim FC As FormatCondition
    For Each FC In rng.FormatConditions
        With FC.Interior
            If .Pattern = Pattern And .PatternColor = PatternColor Then FC.Delete
        End With
    Next
End Function

'������Ч�� rng��Ԫ�� Formula����"a,b,c" ShowError ��ʾ������ʾ���ҽ�ֹ���� AlertStyle������ʾ��ʽ
Public Sub Rng_Validation(rng As Range, Formula, Optional ShowError As Boolean = True, Optional AlertStyle As XlDVAlertStyle = xlValidAlertStop)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=AlertStyle, Operator:= _
            xlBetween, Formula1:=Formula
        'AlertStyle  xlValidAlertInformation 3 ��Ϣͼ��
        '                    xlValidAlertStop 1 ֹͣͼ��
        '                    xlValidAlertWarning 2 ����ͼ��
        .IgnoreBlank = True '�����ֵ
        .InCellDropdown = True '�����б�
        .InputTitle = "" '������ʾ����
        .InputMessage = "" '������ʾ����
        .ErrorTitle = "" '������ʾ�����
        .ErrorMessage = "" '������ʾ������
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False '��ʾ������ʾ InputTitle InputMessage
        .ShowError = True '��ʾ������ʾ ErrorTitle ErrorMessage ����True��ֹ����
    End With
End Sub

'�����ע
Public Function RngAddComment(rng As Range, CommentText, Optional Visible As Boolean = False) As Comment
    Set RngAddComment = rng.Comment
    If rng.Comment Is Nothing Then
       Set RngAddComment = rng.AddComment(CommentText)
    Else
        RngAddComment.Text CommentText
    End If
    RngAddComment.Visible = Visible
End Function

'���ͼƬ PicturePath����·�� rng��Ԫ�� LowerWidth��������� LowerHeight�߶������� OriginalSizeRatio�Ƿ�ԭ��С����
Public Function RngAddPicture(PicturePath, rng As Range, Optional LowerWidth = 0, Optional LowerHeight = 0, Optional OriginalSizeRatio As Boolean = False) As Shape
    Dim ImageWH, ratio, w, h
    If OriginalSizeRatio Then
        ImageWH = ImageSize(PicturePath)
        ratio = ImageWH(0) / ImageWH(1)
        w = rng.Width - LowerWidth * 2
        h = rng.Height - LowerHeight * 2
        If ratio > w / h Then
            Set RngAddPicture = rng.Parent.Shapes.AddPicture(PicturePath, msoFalse, msoTrue, rng.Left + LowerWidth, rng.Top + (rng.Height - w / ratio) / 2, w, w / ratio)
        Else
            Set RngAddPicture = rng.Parent.Shapes.AddPicture(PicturePath, msoFalse, msoTrue, rng.Left + (rng.Width - h * ratio) / 2, rng.Top + LowerHeight, h * ratio, h)
        End If
    Else
        Set RngAddPicture = rng.Parent.Shapes.AddPicture(PicturePath, msoFalse, msoTrue, rng.Left + LowerWidth, rng.Top + LowerHeight, rng.Width - LowerWidth * 2, rng.Height - LowerHeight * 2)
    End If
End Function







'��ѧ-------------------------------------------------------------------------------------------------------------------------------------
 
'�������
Public Function SumParams(ParamArray arr()) As Double
    Dim v
    For Each v In arr
        SumParams = SumParams + Val(v)
    Next
End Function
 
'���������ֵ
Public Function MaxParams(ParamArray arr()) As Double
    Dim v
    MaxParams = -1.79769313486231E+308
    For Each v In arr
        If IsNumeric(v) Then
            If MaxParams < v * 1 Then MaxParams = v
        End If
    Next
End Function

'����ȡ���ֵ Ч�ʸ�
Public Function MaxParams2(Number1, Number2) As Double
    If Number1 < Number2 Then MaxParams2 = Number2 Else MaxParams2 = Number1
End Function

'��������Сֵ
Public Function MinParams(ParamArray arr()) As Double
    Dim v
    MinParams = 1.79769313486231E+308
    For Each v In arr
        If IsNumeric(v) Then
            If MinParams > v * 1 Then MinParams = v
        End If
    Next
End Function

'����ȡ��Сֵ Ч�ʸ�
Public Function MinParams2(Number1, Number2) As Double
    If Number1 > Number2 Then MinParams2 = Number2 Else MinParams2 = Number1
End Function

'������������ı���
Public Function MultiplesUp(Number, Multiples) As Double
    MultiplesUp = IntUp(Number / Multiples) * Multiples
End Function
 
'������������ı���
Public Function MultiplesDown(Number, Multiples) As Double
    MultiplesDown = VBA.Int(Number / Multiples + 0.00000000001) * Multiples
End Function
 
'��������ȡ��
Public Function IntUp(Number) As Double
    IntUp = -IntDown(-Number)
End Function
 
'��������ȡ��
Public Function IntDown(Number) As Double
    IntDown = VBA.Int(Number + 0.00000000001)
End Function
 
'��������
Public Function RoundUp(Number, Optional ByVal NumDigitsAfterDecimal = 0) As Double
    NumDigitsAfterDecimal = 10 ^ NumDigitsAfterDecimal
    RoundUp = IntUp(Number * NumDigitsAfterDecimal) / NumDigitsAfterDecimal
End Function
 
'��������
Public Function RoundDown(Number, Optional ByVal NumDigitsAfterDecimal = 0) As Double
    NumDigitsAfterDecimal = 10 ^ NumDigitsAfterDecimal
    RoundDown = VBA.Int(Number * NumDigitsAfterDecimal + 0.00000000001) / NumDigitsAfterDecimal
End Function

'��������ָ�������ı���
Public Function MultipleUp(Number, Significance) As Double
    MultipleUp = IntUp(Number / Significance) * Significance
End Function

'��������ָ�������ı���
Public Function MultipleDown(Number, Significance) As Double
    MultipleDown = IntDown(Number / Significance) * Significance
End Function

'��������ָ�������ı���
Public Function MultipleRound(Number, Significance) As Double
    MultipleRound = RoundEX(Number / Significance) * Significance
End Function

'������������㵼�µľ���ȱʧ
Public Function Float_Clear(Number) As Double
    Float_Clear = VBA.Round(Number, 10)
End Function

'�����������
Public Function RoundEX(Number, Optional NumDigitsAfterDecimal = 0) As Double
    RoundEX = VBA.Round(Number + 0.000000000001, NumDigitsAfterDecimal)
End Function

'����  ʮ�ڴ������಻����
Function ModNumber(Number1, Number2) As Double
    ModNumber = Number1 - VBA.Int(Number1 / Number2 + 0.00000000001) * Number2
End Function

'��� +Number �� -Number
Public Function RandAddSub(Optional Number As Double = 1) As Double
    RandAddSub = ((Rnd >= 0.5) Or 1) * Number
End Function
 
'����Χ�����
Public Function RandBetween(l, r) As Double
    'Randomize
    RandBetween = IntDown((r - l + 1) * Rnd()) + l
End Function
 
'������� Number��������� interval��ִ�С NumberSplit(5, 2)->[2,2,1]
Public Function NumberSplit(Number, Interval) As Variant
    Dim i As Long
    If Number * Interval > 0 And Interval <> 0 Then
        ArrayDynamic_
        For i = 1 To VBA.Int(Number / Interval + 0.00000000001)
            ArrayDynamic_ Interval
        Next
        If (Number Mod Interval) <> 0 Then ArrayDynamic_ Number Mod Interval
        NumberSplit = ArrayDynamic_
    Else
        NumberSplit = Array()
    End If
End Function

'���ִ�дתСд
Public Function NumberLCase(NumberStr) As Double
    Dim i As Long, n As String, nDW As Double, nDWL As Double, s As Double
    Static dicDX As Object, dicDW As Object, dicDWL As Object
    If dicDX Is Nothing Then
        Set dicDX = CreateObject("scripting.Dictionary")
        dicDX("��") = 0
        dicDX("Ҽ") = 1
        dicDX("��") = 2
        dicDX("��") = 3
        dicDX("��") = 4
        dicDX("��") = 5
        dicDX("½") = 6
        dicDX("��") = 7
        dicDX("��") = 8
        dicDX("��") = 9
        Set dicDW = CreateObject("scripting.Dictionary")
        dicDW("ʰ") = 10 ^ 1
        dicDW("��") = 10 ^ 2
        dicDW("Ǫ") = 10 ^ 3
        Set dicDWL = CreateObject("scripting.Dictionary")
        dicDWL("��") = 10 ^ 4
        dicDWL("��") = 10 ^ 8
        dicDWL("��") = 10 ^ 12
    End If
    nDW = 1
    nDWL = 1
    For i = VBA.Len(NumberStr) To 1 Step -1
        n = Mid(NumberStr, i, 1)
        If dicDWL.Exists(n) Then
            nDWL = dicDWL(n)
            nDW = nDWL
        ElseIf dicDW.Exists(n) Then
            nDW = dicDW(n) * nDWL
        ElseIf dicDX.Exists(n) Then
            s = s + dicDX(n) * nDW
        ElseIf n = "��" Then
            s = -s
        End If
    Next
    NumberLCase = s
End Function

'����ת��д
Public Function NumberUCase(ByVal Number) As String
    Dim i As Long, maxlen As Long, minl As Double, n As Double, np As Double, s As String, sf As String
    Static arrDX, arrDW
    If Not IsArray(arrDX) Then arrDX = Array("��", "Ҽ", "��", "��", "��", "��", "½", "��", "��", "��")
    If Not IsArray(arrDW) Then arrDW = Array("", "��", "��", "��")
    If Number = 0 Then
        NumberUCase = "��": Exit Function
    ElseIf Number < 0 Then
        Number = -Number: sf = "��"
    End If
    maxlen = VBA.Int(Log(Number) / Log(10) + 0.00000000001) + 1
    maxlen = IntUp(maxlen / 4) - 1
    np = 1000
    n = Number
    For i = 0 To maxlen
        minl = 10000 ^ i
        n = ModNumber(VBA.Int(Number / minl + 0.00000000001), 10000)   'Mod������� Int(Number / minl) Mod 10000
        If np < 1000 And np > 0 Then s = "��" & s
        If n > 0 Then
            s = NumberUCaseThousand_(n) & arrDW(i) & s
        End If
        
        np = n
    Next
    NumberUCase = sf & s
End Function

'����ת��д���������� �ڲ�ʹ��
Private Function NumberUCaseThousand_(Number) As String
    Dim i As Long, maxlen As Long, minl As Double, n As Double, np As Double, s As String
    Static arrDX, arrDW
    If Not IsArray(arrDX) Then arrDX = Array("��", "Ҽ", "��", "��", "��", "��", "½", "��", "��", "��")
    If Not IsArray(arrDW) Then arrDW = Array("", "ʰ", "��", "Ǫ")
    If Number = 0 Then NumberUCaseThousand_ = "��": Exit Function
    maxlen = VBA.Int(Log(Number) / Log(10) + 0.00000000001) + 1
    np = 0
    For i = 0 To maxlen - 1
        minl = 10 ^ i
        n = VBA.Int(Number / minl + 0.00000000001) Mod 10
       
        If n > 0 Then
            s = arrDX(n) & arrDW(i) & s
        Else
            If np <> 0 Then
                s = "��" & s
            End If
        End If
        np = n
    Next
    NumberUCaseThousand_ = s
End Function

'�����Сд
Public Function RMBLCase(NumberStr) As Currency
    Dim i As Long, n As String, nDW As Currency, nDWL As Currency, s As Currency
    Static dicDX As Object, dicDW As Object, dicDWL As Object
    If dicDX Is Nothing Then
        Set dicDX = CreateObject("scripting.Dictionary")
        dicDX("��") = 0
        dicDX("Ҽ") = 1
        dicDX("��") = 2
        dicDX("��") = 3
        dicDX("��") = 4
        dicDX("��") = 5
        dicDX("½") = 6
        dicDX("��") = 7
        dicDX("��") = 8
        dicDX("��") = 9
        Set dicDW = CreateObject("scripting.Dictionary")
        dicDW("��") = 10 ^ -2
        dicDW("��") = 10 ^ -1
        dicDW("Ԫ") = 10 ^ 0
        dicDW("ʰ") = 10 ^ 1
        dicDW("��") = 10 ^ 2
        dicDW("Ǫ") = 10 ^ 3
        Set dicDWL = CreateObject("scripting.Dictionary")
        dicDWL("��") = 10 ^ 4
        dicDWL("��") = 10 ^ 8
        dicDWL("��") = 10 ^ 12
    End If
    nDW = 1
    nDWL = 1
    For i = VBA.Len(NumberStr) To 1 Step -1
        n = Mid(NumberStr, i, 1)
        If dicDWL.Exists(n) Then
            nDWL = dicDWL(n)
            nDW = nDWL
        ElseIf dicDW.Exists(n) Then
            nDW = dicDW(n) * nDWL
        ElseIf dicDX.Exists(n) Then
            s = s + dicDX(n) * nDW
        ElseIf n = "��" Then
            s = -s
        End If
    Next
    RMBLCase = s
End Function

'����Ҵ�д
Public Function RMBUCase(curmoney) As String
    Dim curmoney1 As Currency
    Dim i1 As Currency
    Dim i2 As Currency
    Dim i3 As Currency
    Dim s1 As String
    Static arrDX
    If Not IsArray(arrDX) Then arrDX = Array("��", "Ҽ", "��", "��", "��", "��", "½", "��", "��", "��")
    curmoney1 = VBA.Abs(RoundEX(curmoney, 2))
    i1 = VBA.Int(curmoney1 + 0.00000000001)
    i2 = ModNumber(VBA.Int(curmoney1 * 10 + 0.00000000001), 10)
    i3 = ModNumber(VBA.Int(curmoney1 * 100 + 0.00000000001), 10)
    If i1 > 0 Then s1 = NumberUCase(i1) & "Ԫ"
    
    If i3 <> 0 And i2 <> 0 Then
        s1 = s1 & arrDX(i2) & "��" & arrDX(i3) & "��"
    ElseIf i3 = 0 And i2 <> 0 Then
        s1 = s1 & arrDX(i2) & "����"
    ElseIf i3 <> 0 And i2 = 0 Then
        If i1 = 0 Then
            s1 = arrDX(i3) & "��"
        Else
            s1 = s1 & arrDX(i2) & arrDX(i3) & "��"
        End If
    Else
        s1 = s1 & "��"
    End If
    If curmoney < 0 Then
        RMBUCase = "��" & s1
    Else
        RMBUCase = s1
    End If
End Function

'����Ҵ�д �Աȱ���
'Public Function RMBDX(M)
'    RMBDX = Replace(Application.Text(Round(M + 0.00000001, 2), "[DBnum2]"), ".", "Ԫ")
'    RMBDX = IIf(Left(Right(RMBDX, 3), 1) = "Ԫ", Left(RMBDX, Len(RMBDX) - 1) & "��" & Right(RMBDX, 1) & "��", IIf(Left(Right(RMBDX, 2), 1) = "Ԫ", RMBDX & "����", IIf(RMBDX = "��", "", RMBDX & "Ԫ��")))
'    RMBDX = Replace(Replace(Replace(Replace(RMBDX, "��Ԫ���", ""), "��Ԫ", ""), "���", "��"), "-", "��")
'End Function

 
'����Ƚ� �ڲ�
Public Function NumberRangeInside(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean
    Select Case NumberRangeRule
        Case Include_Exclude: NumberRangeInside = Number >= NumberL And Number < NumberR
        Case Exclude_Include: NumberRangeInside = Number > NumberL And Number <= NumberR
        Case Include_Include: NumberRangeInside = Number >= NumberL And Number <= NumberR
        Case Exclude_Exclude: NumberRangeInside = Number > NumberL And Number < NumberR
    End Select
End Function
 
'����Ƚ� �ⲿ
Public Function NumberRangeExternal(Number, NumberL, NumberR, Optional NumberRangeRule As NumberRangeType = Include_Exclude) As Boolean
    Select Case NumberRangeRule
        Case Include_Exclude: NumberRangeExternal = Number <= NumberL Or Number > NumberR
        Case Exclude_Include: NumberRangeExternal = Number < NumberL Or Number >= NumberR
        Case Include_Include: NumberRangeExternal = Number <= NumberL Or Number >= NumberR
        Case Exclude_Exclude: NumberRangeExternal = Number < NumberL Or Number > NumberR
    End Select
End Function

'�ж�ż��
Public Function IsEven(Number) As Boolean
    IsEven = (Number And 1) = 0
End Function

'�ж�����
Public Function IsOdd(Number) As Boolean
    IsOdd = (Number And 1) = 1
End Function

'ѭ����� (i,3)->1,2,3,1,2,3,1,2,3
Public Function Number_Cycle(ByRef Number, ByRef CycleCount) As Long
    Number_Cycle = ((Number - 1) Mod CycleCount) + 1
End Function

'�ظ���� (i,3)->1,1,1,2,2,2,3,3,3
Public Function Number_Repeat(ByRef Number, ByRef RepeatCount) As Long
    Number_Repeat = VBA.Int((Number - 1) / RepeatCount + 0.00000000001) + 1
End Function

'������ (i,3)->1,4,7,10,13,16,19,22,25
Public Function Number_Separated(ByRef Number, ByRef SeparatedCount) As Long
    Number_Separated = (Number - 1) * SeparatedCount + 1
End Function

'Pi��ֵ
Public Property Get vbPi() As Double
    vbPi = Atn(1) * 4
End Property

'�Ƕ�ת����
Public Function AngleToRadian(Angle) As Double
    Dim pi As Double
    pi = Atn(1) * 4
    AngleToRadian = Angle / 180 * pi
End Function

'����ת�Ƕ�
Public Function RadianToAngle(Radian, Optional ByVal NumDigitsAfterDecimal = 3) As Double
    Dim pi As Double
    pi = Atn(1) * 4
    RadianToAngle = RoundEX(Radian / pi * 180, NumDigitsAfterDecimal)
End Function





'����-------------------------------------------------------------------------------------------------------------------------------------

'�⹹
Public Property Get Deconstr(ParamArray DValue() As Variant)
    Deconstr = DValue
End Property

Public Property Let Deconstr(ParamArray DValue() As Variant, ByRef Value As Variant)
    Dim l1, l2, v
    l1 = UBound(DValue)
    l2 = 0
    For Each v In Value
        Cover DValue(l2), v
        l2 = l2 + 1
        If l2 > l1 Then Exit For
    Next
End Property
 
'��ֵ
Public Function Cover(iValue, jValue)
    If VBA.IsObject(jValue) Then
        Set iValue = jValue
    Else
        iValue = jValue
    End If
End Function
 
'����
Public Function Exchange(iValue, jValue)
    Dim kValue As Variant
    Cover kValue, iValue
    Cover iValue, jValue
    Cover jValue, kValue
End Function

'Col����ת����
Public Function ColToArr(ByRef col, Optional Transpose2D As Boolean = False) As Variant
    Dim i, arrRE()
    If Transpose2D Then
        ReDim arrRE(1 To col.Count, 1 To 1)
        For i = 1 To col.Count
            Cover arrRE(i, 1), col(i)
        Next
    Else
        ReDim arrRE(1 To col.Count)
        For i = 1 To col.Count
            Cover arrRE(i), col(i)
        Next
    End If
    ColToArr = arrRE
End Function

'�����ֵ� itemΪ�������� �ظ�ֵ����ȡ��ǰ
Public Function DictionaryCreate(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = StartIndex
    For Each v In arr
        If Not dic.Exists(v) Then
            dic.Add v, i
        End If
        i = i + 1
    Next
    Set DictionaryCreate = dic
End Function

'�����ֵ� itemΪ�������� ���� �ظ�ֵ����ȡ���
Public Function DictionaryCreateRev(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = StartIndex
    For Each v In arr
        dic(v) = i
        i = i + 1
    Next
    Set DictionaryCreateRev = dic
End Function

'�����ֵ� �ظ�ֵ��ӵ���������
Public Function DictionaryCreateIndex_ItemIsCol(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = StartIndex
    For Each v In arr
        If Not dic.Exists(v) Then
            dic.Add v, New Collection
        End If
        dic(v).Add i
        i = i + 1
    Next
    Set DictionaryCreateIndex_ItemIsCol = dic
End Function

'�����ֵ� itemΪ�ֵ���������
Public Function DictionaryCreate_DicIndex(arr, Optional StartIndex As Long = 1, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    For Each v In arr
        If Not dic.Exists(v) Then
            dic.Add v, dic.Count + StartIndex
        End If
    Next
    Set DictionaryCreate_DicIndex = dic
End Function

'�����ֵ� itemΪԪ������
Public Function DictionaryCreate_Count(arr, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    For Each v In arr
        If dic.Exists(v) Then
            dic(v) = dic(v) + 1
        Else
            dic.Add v, 1
        End If
    Next
    Set DictionaryCreate_Count = dic
End Function

'�����ֵ� ˫���鵽�ֵ�
Public Function DictionaryCreate_Items(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long, u As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = LBound(arrItems)
    u = UBound(arrItems)
    For Each v In arrKeys
        If i > u Then Exit For
        If Not dic.Exists(v) Then
            dic.Add v, arrItems(i)
        End If
        i = i + 1
    Next
    Set DictionaryCreate_Items = dic
End Function

'�����ֵ� ˫���鵽�ֵ� ����
Public Function DictionaryCreate_ItemsRev(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long, u As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = LBound(arrItems)
    u = UBound(arrItems)
    For Each v In arrKeys
        If i > u Then Exit For
        If IsObject(arrItems(i)) Then
            Set dic(v) = arrItems(i)
        Else
            dic(v) = arrItems(i)
        End If
        i = i + 1
    Next
    Set DictionaryCreate_ItemsRev = dic
End Function

'�����ֵ� ˫���鵽�ֵ� �ظ�ֵ��ӵ�����
Public Function DictionaryCreate_ItemsIsCol(arrKeys, arrItems, Optional CompareMode As CompareMethod = BinaryCompare) As Object
    Dim dic As Object, v As Variant, i As Long, u As Long
    Set dic = CreateObject("scripting.Dictionary")
    dic.CompareMode = CompareMode
    i = LBound(arrItems)
    u = UBound(arrItems)
    For Each v In arrKeys
        If i > u Then Exit For
        If Not dic.Exists(v) Then
            dic.Add v, New Collection
        End If
        dic(v).Add arrItems(i)
        i = i + 1
    Next
    Set DictionaryCreate_ItemsIsCol = dic
End Function

'�ֵ�������� �ظ������޸�ԭ��ֵ
Public Function DictionaryAdds(dic, arrKeys, arrItems) As Object
    Dim v As Variant, i As Long, u As Long
    i = LBound(arrItems)
    u = UBound(arrItems)
    For Each v In arrKeys
        If i > u Then Exit For
        If Not dic.Exists(v) Then
            dic.Add v, arrItems(i)
        End If
        i = i + 1
    Next
    Set DictionaryAdds = dic
End Function

'�ֵ�������� �ظ��򸲸�ԭ��ֵ
Public Function DictionaryAddsRev(dic, arrKeys, arrItems) As Object
    Dim v As Variant, i As Long, u As Long
    i = LBound(arrItems)
    u = UBound(arrItems)
    For Each v In arrKeys
        If i > u Then Exit For
        If IsObject(arrItems(i)) Then
            Set dic(v) = arrItems(i)
        Else
            dic(v) = arrItems(i)
        End If
        i = i + 1
    Next
    Set DictionaryAddsRev = dic
End Function

'�ֵ�ϲ�
Public Function DictionaryMerge(ParamArray Dics()) As Object
    Dim dic As Object, dicRE, v
    Set dic = CreateObject("scripting.Dictionary")
    For Each dicRE In Dics
        For Each v In dicRE
            If Not dic.Exists(v) Then
                dic.Add v, dicRE(v)
            End If
        Next
    Next
    Set DictionaryMerge = dic
End Function

'�ֵ�ϲ� ���� ���ظ������滻ǰ��
Public Function DictionaryMergeRev(ParamArray Dics()) As Object
    Dim dic As Object, dicRE, v
    Set dic = CreateObject("scripting.Dictionary")
    For Each dicRE In Dics
        For Each v In dicRE
            If IsObject(dicRE(v)) Then
                Set dic(v) = dicRE(v)
            Else
                dic(v) = dicRE(v)
            End If
        Next
    Next
    Set DictionaryMergeRev = dic
End Function

'�ֵ�ȡ���ֵ �����Key
Public Function DictionaryGetValuesParam(dic, ParamArray Keys()) As Variant
    Dim v As Variant, i As Long, j As Long
    Keys = ArrFlatten(Keys)
    For i = LBound(Keys) To UBound(Keys)
        If dic.Exists(Keys(i)) Then
            Cover Keys(i), dic(Keys(i))
        Else
            Cover Keys(i), Empty
        End If
    Next
    DictionaryGetValuesParam = Keys
End Function

'�ֵ�ȡ���ֵ  arrKey������һά��ά���鷵�ض�Ӧ��С��Itemֵ���� NoExistsValue��������ֵ
Public Function DictionaryGetValues(dic, ByVal arrKey, Optional NoExistsValue = Empty) As Variant
    Dim v As Variant, i As Long, j As Long
    Select Case ArrDimension(arrKey)
        Case 1
            For i = LBound(arrKey) To UBound(arrKey)
                If dic.Exists(arrKey(i)) Then
                    Cover arrKey(i), dic(arrKey(i))
                Else
                    Cover arrKey(i), NoExistsValue
                End If
            Next
            DictionaryGetValues = arrKey
        Case 2
            Dim l As Long, u As Long
            l = LBound(arrKey, 2): u = UBound(arrKey, 2)
            For i = LBound(arrKey, 1) To UBound(arrKey, 1)
                For j = l To u
                    If dic.Exists(arrKey(i, j)) Then
                        Cover arrKey(i, j), dic(arrKey(i, j))
                    Else
                        Cover arrKey(i, j), NoExistsValue
                    End If
                Next
            Next
            DictionaryGetValues = arrKey
        Case Else
            If dic.Exists(arrKey) Then
                Cover DictionaryGetValues, dic(arrKey)
            Else
                Cover DictionaryGetValues, NoExistsValue
            End If
    End Select
End Function

'�ֵ��ж϶��ֵ
Public Function DictionaryExists(dic, ByVal arrKey) As Variant
    Dim v As Variant, i As Long, j As Long
    Select Case ArrDimension(arrKey)
        Case 1
            For i = LBound(arrKey) To UBound(arrKey)
                arrKey(i) = dic.Exists(arrKey(i))
            Next
            DictionaryExists = arrKey
        Case 2
            Dim l As Long, u As Long
            l = LBound(arrKey, 2): u = UBound(arrKey, 2)
            For i = LBound(arrKey, 1) To UBound(arrKey, 1)
                For j = l To u
                    arrKey(i, j) = dic.Exists(arrKey(i, j))
                Next
            Next
            DictionaryExists = arrKey
        Case Else
            DictionaryExists = dic.Exists(arrKey)
    End Select
End Function
 
'�ֵ䵽��ά���� 1����Key 2����Item
Public Function DictionaryToArr2D(dic) As Variant
    DictionaryToArr2D = ArrMergeColumnParam(dic.Keys, dic.Items)
End Function
 
'Application_Attribute(False)�ر�һϵ��Ӱ��Ч������
'**ע������������� Application_Attribute(True)
Public Function Application_Attribute(bol As Boolean)
    Application.ScreenUpdating = bol '//��Ļˢ��
    Application.DisplayAlerts = bol '//ϵͳ��ʾ
    Application.EnableEvents = bol  '//���������¼�
    Application.AskToUpdateLinks = bol  '�ⲿ����
    Application.Calculation = IIf(bol, xlAutomatic, xlManual) '//�Զ�����
End Function
 
'������Ĳ�ռCPU�ӳ�,��λ����
Public Function Sleep(PauseTime)
    Dim StartTimer, StartTimer2
    StartTimer = GetTimer
    Do While GetTimer < StartTimer + PauseTime
        WaitMessage
        DoEvents
        Sleep_ 10
    Loop
End Function

'��ӡ���� arg��ӡ���� RowCount��ӡ��������������  DividerLine�Ƿ��зָ���*��ͨ����Ĭ�ϲ���ӡΪFalseʱ�Ŵ�ӡ�ָ��ߣ���������Ĭ�ϴ�ӡΪFalseʱ����ӡ*
Public Function PrintEx(ByRef arg, Optional RowCount = 0, Optional DividerLine As Boolean = True)
    Select Case TypeName(arg)
        Case "Range"
            If DividerLine Then Debug.Print String(150, "-")
            Debug.Print "Range.Address��" & arg.Parent.Name & "!" & arg.Address(False, False)
            ArrPrint_ arg.Value, RowCount, False
        Case "Dictionary"
            DictionaryPrint_ arg, RowCount, DividerLine
        Case "Collection"
            CollectionPrint_ arg, RowCount, DividerLine
        Case Else
            If IsArray(arg) Then
                ArrPrintAll_ arg, RowCount, DividerLine
            Else
                If Not DividerLine Then Debug.Print String(150, "-")
                Debug.Print arg
            End If
    End Select
End Function

'�ڲ����� ��ӡǶ������ arr���� RowCount��ӡ��������������
Private Function ArrPrintAll_(ByRef arr, Optional RowCount = 0, Optional DividerLine As Boolean = True)
    Dim i As Long, j As Long
    Dim l As Long, u As Long
    Select Case ArrDimension(arr)
        Case 1
            If LBound(arr, 1) > UBound(arr, 1) Then
                ArrPrint_ arr, RowCount, DividerLine
            Else
                Select Case ArrDimension(arr(LBound(arr, 1)))
                    Case 1, 2
                        For i = LBound(arr) To UBound(arr)
                            ArrPrintAll_ arr(i), RowCount, DividerLine
                        Next
                    Case 0
                        ArrPrint_ arr, RowCount, DividerLine
                End Select
            End If
        Case 2
            Select Case ArrDimension(arr(LBound(arr, 1), LBound(arr, 2)))
                Case 1, 2
                    l = LBound(arr, 2): u = UBound(arr, 2)
                    For i = LBound(arr, 1) To UBound(arr, 1)
                        For j = l To u
                            ArrPrintAll_ arr(i, j), RowCount, DividerLine
                        Next
                    Next
                Case 0
                    ArrPrint_ arr, RowCount, DividerLine
            End Select
        Case 0
            ArrPrint_ arr, RowCount, DividerLine
    End Select
End Function

'�ڲ����� ��ӡ���� arr���� RowCount��ӡ��������������
Private Function ArrPrint_(ByVal arr, Optional RowCount = 0, Optional DividerLine As Boolean = True)
    Dim st As String
    Dim i As Long, j As Long, ArrPrint11, arrrow, arrrow2
    Dim istart, jstart, iend, jend
    If DividerLine Then Debug.Print String(150, "-")
    Select Case ArrDimension(arr)
        Case 1
            If LBound(arr) > UBound(arr) Then Debug.Print "����Ϊ��": Exit Function
            Dim arrRE
            ReDim arrRE(LBound(arr) To UBound(arr), -1 To -1)
            For i = LBound(arr) To UBound(arr)
                arrRE(i, -1) = arr(i)
            Next
            arr = arrRE
        Case 0
            Debug.Print "��������": Exit Function
    End Select
    If RowCount = 0 Then
        istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    ElseIf RowCount > 0 Then
        istart = LBound(arr, 1)
        iend = IIf(UBound(arr, 1) - LBound(arr, 1) + 1 > RowCount, LBound(arr, 1) + RowCount - 1, UBound(arr, 1))
    ElseIf RowCount < 0 Then
        istart = UBound(arr, 1) + RowCount + 1
        If istart < LBound(arr, 1) Then istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    End If
    jstart = LBound(arr, 2): jend = UBound(arr, 2)
    Dim linshilen As Long
    ReDim ArrPrint1(istart - 1 To iend)
    ReDim arrrow(jstart - 1 To jend)
    For i = istart To iend
        For j = jstart To jend
            linshilen = LenB(StrConv(arr(i, j), vbFromUnicode))
            If arrrow(j) < linshilen Then
                arrrow(j) = linshilen
            End If
        Next
    Next
    For j = jstart To jend '�б����С
        linshilen = LenB(StrConv(CStr(j), vbFromUnicode))
        If arrrow(j) < linshilen Then
            arrrow(j) = linshilen
        End If
    Next
    arrrow(jstart - 1) = LenB(StrConv(CStr(iend - istart + 1), vbFromUnicode))
    linshilen = LenB(StrConv(CStr(istart), vbFromUnicode))
    If arrrow(jstart - 1) < linshilen Then arrrow(jstart - 1) = linshilen
    For i = istart To iend
        ReDim arrrow2(jstart - 1 To jend)
        arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
        RSet arrrow2(jstart - 1) = i
        For j = jstart To jend
            arrrow2(j) = Space(arrrow(j) - LenB(StrConv(arr(i, j), vbFromUnicode)) + Len(arr(i, j)))
            LSet arrrow2(j) = arr(i, j)
        Next
        ArrPrint1(i) = Join(arrrow2, " | ")
    Next
    For j = jstart To jend '�б����ֶ�
        arrrow2(j) = Space(arrrow(j))
        RSet arrrow2(j) = j
    Next
    arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
    RSet arrrow2(jstart - 1) = iend - istart + 1
    ArrPrint1(istart - 1) = Join(arrrow2, " | ")
    Debug.Print Join(ArrPrint1, vbCrLf)
End Function
 
'�ڲ����� ��ӡ�ֵ� dic�ֵ� RowCount��ӡ��������������
Private Function DictionaryPrint_(ByRef dic, Optional RowCount = 0, Optional DividerLine As Boolean = True)
    Dim st As String, arr
    Dim i As Long, j As Long, ArrPrint11, arrrow, arrrow2
    Dim istart, jstart, iend, jend
    If DividerLine Then Debug.Print String(150, "-")
    If TypeName(dic) = "Dictionary" Then
        If dic.Count > 0 Then
            arr = DictionaryToArr2D(dic)
        Else
            Debug.Print "�ֵ�Ϊ��": Exit Function
        End If
    Else
        Debug.Print "�����ֵ�": Exit Function
    End If
    If RowCount = 0 Then
        istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    ElseIf RowCount > 0 Then
        istart = LBound(arr, 1)
        iend = IIf(UBound(arr, 1) - LBound(arr, 1) + 1 > RowCount, LBound(arr, 1) + RowCount - 1, UBound(arr, 1))
    ElseIf RowCount < 0 Then
        istart = UBound(arr, 1) + RowCount + 1
        If istart < LBound(arr, 1) Then istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    End If
    jstart = LBound(arr, 2): jend = UBound(arr, 2)
    Dim linshilen As Long
    ReDim ArrPrint1(istart - 1 To iend)
    ReDim arrrow(jstart - 1 To jend)
    For i = istart To iend
        For j = jstart To jend
            linshilen = LenB(StrConv(arr(i, j), vbFromUnicode))
            If arrrow(j) < linshilen Then
                arrrow(j) = linshilen
            End If
        Next
    Next
    
    linshilen = LenB(StrConv(CStr("Key"), vbFromUnicode))
    arrrow(jstart) = linshilen
    linshilen = LenB(StrConv(CStr("Item"), vbFromUnicode))
    arrrow(jend) = linshilen
    
    arrrow(jstart - 1) = LenB(StrConv(CStr(iend - istart + 1), vbFromUnicode))
    linshilen = LenB(StrConv(CStr(istart), vbFromUnicode))
    If arrrow(jstart - 1) < linshilen Then arrrow(jstart - 1) = linshilen
    For i = istart To iend
        ReDim arrrow2(jstart - 1 To jend)
        arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
        RSet arrrow2(jstart - 1) = i
        For j = jstart To jend
            arrrow2(j) = Space(arrrow(j) - LenB(StrConv(arr(i, j), vbFromUnicode)) + Len(arr(i, j)))
            LSet arrrow2(j) = arr(i, j)
        Next
        ArrPrint1(i) = Join(arrrow2, " | ")
    Next

    arrrow2(jstart) = Space(arrrow(jstart))
    RSet arrrow2(jstart) = "Key"
    arrrow2(jend) = Space(arrrow(jstart))
    RSet arrrow2(jend) = "Item"
    
    arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
    RSet arrrow2(jstart - 1) = iend - istart + 1
    ArrPrint1(istart - 1) = Join(arrrow2, " | ")
    Debug.Print Join(ArrPrint1, vbCrLf)
End Function

'�ڲ����� ��ӡ���� col�ֵ� RowCount��ӡ��������������
Private Function CollectionPrint_(ByRef col, Optional RowCount = 0, Optional DividerLine As Boolean = True)
    Dim st As String, arr
    Dim i As Long, j As Long, ArrPrint11, arrrow, arrrow2
    Dim istart, jstart, iend, jend
    If DividerLine Then Debug.Print String(150, "-")
    If TypeName(col) = "Collection" Then
        If col.Count > 0 Then
            arr = ColToArr(col, True)
        Else
            Debug.Print "����Ϊ��": Exit Function
        End If
    Else
        Debug.Print "���Ǽ���": Exit Function
    End If
    If RowCount = 0 Then
        istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    ElseIf RowCount > 0 Then
        istart = LBound(arr, 1)
        iend = IIf(UBound(arr, 1) - LBound(arr, 1) + 1 > RowCount, LBound(arr, 1) + RowCount - 1, UBound(arr, 1))
    ElseIf RowCount < 0 Then
        istart = UBound(arr, 1) + RowCount + 1
        If istart < LBound(arr, 1) Then istart = LBound(arr, 1)
        iend = UBound(arr, 1)
    End If
    jstart = LBound(arr, 2): jend = UBound(arr, 2)
    Dim linshilen As Long
    ReDim ArrPrint1(istart - 1 To iend)
    ReDim arrrow(jstart - 1 To jend)
    For i = istart To iend
        For j = jstart To jend
            linshilen = LenB(StrConv(arr(i, j), vbFromUnicode))
            If arrrow(j) < linshilen Then
                arrrow(j) = linshilen
            End If
        Next
    Next
    
    linshilen = LenB(StrConv(CStr("Item"), vbFromUnicode))
    arrrow(jstart) = linshilen
    
    arrrow(jstart - 1) = LenB(StrConv(CStr(iend - istart + 1), vbFromUnicode))
    linshilen = LenB(StrConv(CStr(istart), vbFromUnicode))
    If arrrow(jstart - 1) < linshilen Then arrrow(jstart - 1) = linshilen
    For i = istart To iend
        ReDim arrrow2(jstart - 1 To jend)
        arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
        RSet arrrow2(jstart - 1) = i
        For j = jstart To jend
            arrrow2(j) = Space(arrrow(j) - LenB(StrConv(arr(i, j), vbFromUnicode)) + Len(arr(i, j)))
            LSet arrrow2(j) = arr(i, j)
        Next
        ArrPrint1(i) = Join(arrrow2, " | ")
    Next

    arrrow2(jstart) = Space(arrrow(jstart))
    RSet arrrow2(jstart) = "Item"
    
    arrrow2(jstart - 1) = Space(arrrow(jstart - 1))
    RSet arrrow2(jstart - 1) = iend - istart + 1
    ArrPrint1(istart - 1) = Join(arrrow2, " | ")
    Debug.Print Join(ArrPrint1, vbCrLf)
End Function

'����Base64
Public Function encodeBase64(Bytes) As String
    With CreateObject("msxml2.domdocument").createelement("b64")
        .DataType = "bin.base64"
        .nodetypedvalue = Bytes
        encodeBase64 = .Text
    End With
End Function

'����Base64
Public Function decodeBase64(String1) As Byte()
    Dim Dom As Object
    Set Dom = CreateObject("msxml2.domdocument").createelement("b64")
    Dom.DataType = "bin.base64"
    Dom.Text = String1
    decodeBase64 = Dom.nodetypedvalue
End Function

'ͼƬ���ؿ���С
Public Function ImageSize(ImagePath) As Variant
    Dim Img As Object
    Set Img = CreateObject("WIA.ImageFile")
    Img.LoadFile ImagePath
    ImageSize = Array(Img.Width, Img.Height)
End Function

'����LoadPicture ֧�ֶ���ͼƬ��ʽ
Public Function LoadPictureEx(filename) As IPictureDisp
    Dim Img, v
    Set Img = CreateObject("WIA.ImageFile")
    Img.LoadFile filename
    Set v = Img.FileData
    Set LoadPictureEx = v.Picture
End Function


'��չCLng ֧������ת��
Public Function CLngEx(Expression) As Variant
    Dim arrRE() As Long, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Long
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CLngEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Long
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CLngEx = arrRE
        Case Else
            CLngEx = VBA.CLng(Expression)
    End Select
End Function
 
'��չCDate ֧������ת��
Public Function CDateEx(Expression) As Variant
    Dim arrRE() As Date, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Date
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CDateEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Date
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CDateEx = arrRE
        Case Else
            CDateEx = VBA.CDate(Expression)
    End Select
End Function

'��չCDbl ֧������ת��
Public Function CDblEx(Expression) As Variant
    Dim arrRE() As Double, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Double
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CDblEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Double
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CDblEx = arrRE
        Case Else
            CDblEx = VBA.CDbl(Expression)
    End Select
End Function

'��չCCur ֧������ת��
Public Function CCurEx(Expression) As Variant
    Dim arrRE() As Currency, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Currency
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CCurEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Currency
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CCurEx = arrRE
        Case Else
            CCurEx = VBA.CCur(Expression)
    End Select
End Function

'��չCStr ֧������ת��
Public Function CStrEx(Expression) As Variant
    Dim arrRE() As String, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As String
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CStrEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As String
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CStrEx = arrRE
        Case Else
            CStrEx = VBA.CStr(Expression)
    End Select
End Function

'��չCVar ֧������ת��
Public Function CVarEx(Expression) As Variant
    Dim arrRE() As Variant, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Variant
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CVarEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Variant
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CVarEx = arrRE
        Case Else
            CVarEx = VBA.CVar(Expression)
    End Select
End Function

'��չCBool ֧������ת��
Public Function CBoolEx(Expression) As Variant
    Dim arrRE() As Boolean, i As Long, j As Long
    Dim l1 As Long, u1 As Long
    Dim l2 As Long, u2 As Long
    Select Case ArrDimension(Expression)
        Case 1
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            ReDim arrRE(l1 To u1) As Boolean
            For i = l1 To u1
                arrRE(i) = Expression(i)
            Next
            CBoolEx = arrRE
        Case 2
            l1 = LBound(Expression, 1): u1 = UBound(Expression, 1)
            l2 = LBound(Expression, 2): u2 = UBound(Expression, 2)
            ReDim arrRE(l1 To u1, l2 To u2) As Boolean
            For i = l1 To u1
                For j = l2 To u2
                    arrRE(i, j) = Expression(i, j)
                Next
            Next
            CBoolEx = arrRE
        Case Else
            CBoolEx = VBA.CBool(Expression)
    End Select
End Function




'Http-------------------------------------------------------------------------------------------------------------------------------------
'Get����
Public Function HttpGet(Url, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .SetTimeouts 2000, 2000, 2000, 2000
        .Open "GET", Url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        If Not RequestHeaderDic Is Nothing Then
            Dim k
            For Each k In RequestHeaderDic
                .setRequestHeader k, RequestHeaderDic(k)
            Next
        End If
        .Send
        If strCharset = "" Then
            HttpGet = .Responsetext
        Else
            HttpGet = ByteToStr(.ResponseBody, strCharset)
        End If
    End With
End Function
 
'Get�����ļ�
Public Sub HttpDownload(Url, DownloadFileName, Optional RequestHeaderDic = Nothing)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .SetTimeouts 2000, 2000, 2000, 2000
        .Open "GET", Url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        If Not RequestHeaderDic Is Nothing Then
            Dim k
            For Each k In RequestHeaderDic
                .setRequestHeader k, RequestHeaderDic(k)
            Next
        End If
        .Send
        Call ByteToFile(.ResponseBody, DownloadFileName)
    End With
End Sub
 
'Post����
Public Function HttpPost(Url, Optional SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant
    Dim strText As String
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .SetTimeouts 2000, 2000, 2000, 2000
        .Open "POST", Url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        If Not RequestHeaderDic Is Nothing Then
            Dim k
            For Each k In RequestHeaderDic
                .setRequestHeader k, RequestHeaderDic(k)
            Next
        End If
        If IsMissing(SendValue) And IsError(SendValue) Then
            .Send
        Else
            .Send SendValue
        End If
        If strCharset = "" Then
            HttpPost = .Responsetext
        Else
            HttpPost = ByteToStr(.ResponseBody, strCharset)
        End If
    End With
End Function
 
'Post���� ���ͱ�����
Public Function HttpPost_Form(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant
    Dim strText As String
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .SetTimeouts 2000, 2000, 2000, 2000
        .Open "POST", Url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        If Not RequestHeaderDic Is Nothing Then
            Dim k
            For Each k In RequestHeaderDic
                .setRequestHeader k, RequestHeaderDic(k)
            Next
        End If
        .Send SendValue
        If strCharset = "" Then
            HttpPost_Form = .Responsetext
        Else
            HttpPost_Form = ByteToStr(.ResponseBody, strCharset)
        End If
    End With
End Function
 
'Post���� ����Json����
Public Function HttpPost_Json(Url, SendValue, Optional RequestHeaderDic = Nothing, Optional strCharset As String = "UTF-8") As Variant
    Dim strText As String
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .SetTimeouts 2000, 2000, 2000, 2000
        .Open "POST", Url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        If Not RequestHeaderDic Is Nothing Then
            Dim k
            For Each k In RequestHeaderDic
                .setRequestHeader k, RequestHeaderDic(k)
            Next
        End If
        .Send SendValue
        If strCharset = "" Then
            HttpPost_Json = .Responsetext
        Else
            HttpPost_Json = ByteToStr(.ResponseBody, strCharset)
        End If
    End With
End Function
 
'��ȡJSON����
Private Function HttpReadJson(Jsonstr As String, Routestr As String) As Variant
    Dim HTML, Window
    Set HTML = CreateObject("htmlfile")
    Set Window = HTML.parentWindow
    Jsonstr = Replace(Jsonstr, vbCr, "")
    Jsonstr = Replace(Jsonstr, vbLf, "")
    Window.execScript "var js= " & Jsonstr
    HttpReadJson = Window.eval("js." & Routestr)
End Function




