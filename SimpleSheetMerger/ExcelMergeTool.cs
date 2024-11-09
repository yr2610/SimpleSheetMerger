using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

public class ExcelMergeTool : IExcelAddIn
{
    private static ExcelMergeTool _instance;
    public static ExcelMergeTool Instance => _instance ?? (_instance = new ExcelMergeTool());

    private List<string> mergeFilePaths = new List<string>();
    //private List<string> conflictCells = new List<string>();
    //private Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>> sheetRanges = new Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>>();

    class RangeInfo
    {
        public int? IdColumnOffset { get; set; }
        public HashSet<int> IgnoreColumnOffsets { get; set; }
    }

    class RangeData
    {
        public object[,] Values { get; set; }
        public IEnumerable<object> IdValues { get; set; }
        public HashSet<int> IgnoreColumnOffsets { get; set; }
    }

    class SheetAddressInfo
    {
        public string Address { get; set; }
        public Func<object, object[], (bool merged, object value)> Function { get; set; }
        public RangeInfo RangeInfo { get; set; }
    }

    public void AutoOpen()
    {
        // リボンを登録
        // ExcelDnaUtil.Application.RegisterRibbon(new MyRibbon()); // この行は不要
    }

    public void AutoClose() { }

    public void DragDropFiles(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            mergeFilePaths.AddRange(files.Where(file => !mergeFilePaths.Contains(file)));
        }
    }

    public void RemoveFile(string filePath)
    {
        mergeFilePaths.Remove(filePath);
    }

    public void OnMergeButtonClick(IRibbonControl control)
    {
        // マージ処理を呼び出す
        MergeFiles(mergeFilePaths);
    }

    public void OnSelectFilesButtonClick(IRibbonControl control)
    {
        // ファイル選択フォームを表示
        ShowFileSelectionForm();
    }

    static dynamic GetSheetIfExists(Excel.Workbook workbook, string sheetName)
    {
        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            if (sheet.Name == sheetName)
            {
                return sheet;
            }
        }
        return null;
    }

    static Excel.Name GetNamedRange(Excel.Worksheet sheet, string name)
    {
        try
        {
            Excel.Name namedRange = sheet.Names.Item(name);
            return namedRange;
        }
        catch (Exception)
        {
            return null; // エラーが発生した場合は null を返します
        }
    }

    static SheetAddressInfo GetSheetAddressInfo(Excel.Worksheet sheet)
    {
        const string ssSheetRangeName = "SS_SHEET"; // 名前付き範囲の名前
        Excel.Name namedRange = GetNamedRange(sheet, ssSheetRangeName);

        if (namedRange == null)
        {
            return null;
        }

        // 名前付き範囲が存在する場合、その範囲を使用
        string address = namedRange.RefersToRange.Address;
        RangeInfo rangeInfo = null;

        // コメントが存在する場合、それを YAML として解析
        if (namedRange.Comment != null)
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            rangeInfo = deserializer.Deserialize<RangeInfo>(namedRange.Comment);
        }

        return new SheetAddressInfo
        {
            Address = address,
            Function = null,
            RangeInfo = rangeInfo
        };
    }

    static Dictionary<string, List<SheetAddressInfo>> CollectSheetAddresses()
    {
        const string indexSheetName = "index"; // シート名
        const string startCellAddress = "B16"; // 開始セルのアドレス
        const string endMarker = "END"; // 終端を示す文字列
        const string leftColumnAddress = "U"; // 左端の列のアドレス
        const string rightColumnAddress = "AA"; // 右端の列のアドレス
        const string headerRowAddress = "AD"; // ヘッダー行のアドレス
        const string bottomRowAddress = "AE"; // 最下行のアドレス
        string[] ignoreSheetNames = { "無視シート", }; // 無視するシート名のリスト

        var result = new Dictionary<string, List<SheetAddressInfo>>();

        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Worksheet indexSheet = xlApp.Worksheets[indexSheetName];
        Excel.Range startCell = indexSheet.Range[startCellAddress];
        Excel.Range currentCell = startCell;

        // 終端を示す文字列が見つかるまで下方向にたどる
        while (currentCell.Value == null || currentCell.Value.ToString() != endMarker)
        {
            currentCell = currentCell.Offset[1, 0];
        }

        // 範囲を設定
        Excel.Range range = indexSheet.Range[startCell, currentCell.Offset[-1, 0]];

        foreach (Excel.Range cell in range)
        {
            if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
            {
                string sheetName = cell.Value.ToString();

                // 無視リストに含まれるシート名をスキップ
                if (Array.Exists(ignoreSheetNames, name => name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                Excel.Worksheet sheet = xlApp.Worksheets[sheetName];
                var sheetAddressInfo = GetSheetAddressInfo(sheet);

                // 名前付き範囲が存在しない場合、indexSheet の情報からアドレスを作成
                if (sheetAddressInfo == null)
                {
                    string leftColumn = indexSheet.Cells[cell.Row, leftColumnAddress].Value.ToString();
                    string rightColumn = indexSheet.Cells[cell.Row, rightColumnAddress].Value.ToString();
                    int headerRow = (int)indexSheet.Cells[cell.Row, headerRowAddress].Value;
                    int topRow = headerRow + 1;
                    int bottomRow = (int)indexSheet.Cells[cell.Row, bottomRowAddress].Value;
                    string address = $"{leftColumn}{topRow}:{rightColumn}{bottomRow}";

                    sheetAddressInfo = new SheetAddressInfo
                    {
                        Address = address,
                        Function = null,
                        RangeInfo = null,
                    };
                }

                // シート名が辞書に存在しない場合、新しいリストを作成
                if (!result.ContainsKey(sheetName))
                {
                    result[sheetName] = new List<SheetAddressInfo>();
                }

                // アドレスをリストに追加
                result[sheetName].Add(sheetAddressInfo);
            }
        }

        return result;
    }

    static object[,] GetValuesAs2DArray(object range)
    {
        if (range is object[,] array)
        {
            // 既に配列の場合はそのまま返す
            return array;
        }
        else if (range is object singleValue)
        {
            // 1つのセルの場合、1-originのように見える2次元配列として返す
            // 実際の配列のサイズは1x1
            var result = Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
            result.SetValue(singleValue, 1, 1);
            return (object[,])result;
        }

        // 何もない場合は空の1x1の2次元配列を返す
        var emptyResult = Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
        emptyResult.SetValue(null, 1, 1);
        return (object[,])emptyResult;
    }

    static IEnumerable<object> GetColumnWithOffset(Excel.Worksheet worksheet, string address, int columnOffset)
    {
        // 指定されたアドレスの範囲を取得
        var range = worksheet.Range[address];

        // 範囲の開始列を取得
        int startColumn = range.Column;

        // オフセット後の列番号を計算
        int targetColumn = startColumn + columnOffset;

        // 指定された範囲の行を基準にして、対象列を取得
        var offsetColumn = worksheet.Range[worksheet.Cells[range.Row, targetColumn], worksheet.Cells[range.Row + range.Rows.Count - 1, targetColumn]];

        // 2次元配列として範囲を取得
        var values = GetValuesAs2DArray(offsetColumn.Value2);

        // 2次元配列をList<object>に変換
        var result = new List<object>();
        for (int i = 1; i <= values.GetLength(0); i++)
        {
            result.Add(values[i, 1]);
        }

        return result;
    }

    // idValues を key にした行（List<object>）の dictionary を作る
    static Dictionary<string, List<object>> CreateRowDictionaryWithIDKeys(object[,] values, IEnumerable<object> idValues)
    {
#if true
        var dictionary = new Dictionary<string, List<object>>();
        int rowIndex = 1;

        foreach (var idValue in idValues)
        {
            if (idValue == null)
            {
                rowIndex++;
                continue;
            }

            string id = idValue.ToString();
            var rowValues = new List<object>();

            for (int j = 1; j <= values.GetLength(1); j++)
            {
                rowValues.Add(values[rowIndex, j]);
            }

            dictionary[id] = rowValues;
            rowIndex++;
        }

        return dictionary;
#else
        // LINQ駆使した版
        return idValues
            .Zip(Enumerable.Range(1, values.GetLength(0)), (idValue, rowIndex) => (idValue, rowIndex))
            .Where(pair => pair.idValue != null)
            .ToDictionary(
                pair => pair.idValue.ToString(),
                pair => Enumerable.Range(1, values.GetLength(1))
                    .Select(colIndex => values[pair.rowIndex, colIndex])
                    .ToList()
            );
#endif
    }

    static object[,] CopyValuesById(object[,] baseValues, IEnumerable<object> baseIdValues, Dictionary<string, List<object>> valuesDictionary, HashSet<int> ignoreColumnOffsets)
    {
        object[,] result = (object[,])baseValues.Clone();

        int rowIndex = 1; // 1-originのため、1から開始

        foreach (var idValue in baseIdValues)
        {
            if (idValue == null)
            {
                rowIndex++;
                continue;
            }

            string id = idValue.ToString();

            if (valuesDictionary.TryGetValue(id, out var values))
            {
                int colIndex = 1; // 1-originに変換
                foreach (var value in values)
                {
                    if (!ignoreColumnOffsets.Contains(colIndex - 1))
                    {
                        result[rowIndex, colIndex] = value;
                    }
                    colIndex++;
                }
            }
            rowIndex++;
        }

        return result;
    }

    public class ConflictData
    {
        public string SheetName { get; set; }
        public string CellAddress { get; set; }
        public object Base { get; set; }
        public List<object> Values { get; set; }
        public object Merged { get; set; }
        public bool Resolved { get; set; }

        // 動的プロパティを追加
        public object Value1 { get { return Values.Count > 0 ? Values[0] : string.Empty; } }
        public object Value2 { get { return Values.Count > 1 ? Values[1] : string.Empty; } }
        public object Value3 { get { return Values.Count > 2 ? Values[2] : string.Empty; } }
        public object Value4 { get { return Values.Count > 3 ? Values[3] : string.Empty; } }
        public object Value5 { get { return Values.Count > 4 ? Values[4] : string.Empty; } }
        public object Value6 { get { return Values.Count > 5 ? Values[5] : string.Empty; } }
        public object Value7 { get { return Values.Count > 6 ? Values[6] : string.Empty; } }
        public object Value8 { get { return Values.Count > 7 ? Values[7] : string.Empty; } }
    }

    public void MergeFiles(List<string> mergeFilePaths)
    {
        // 現在のアクティブなブックを取得
        var excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var baseWorkbook = excelApp.ActiveWorkbook;

        // アクティブなブックがない場合の処理
        if (baseWorkbook == null)
        {
            MessageBox.Show("アクティブなブックがありません。操作を続行するにはブックを開いてください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return;
        }

        if (mergeFilePaths.Count == 0)
        {
            MessageBox.Show("マージするファイルが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return;
        }

        excelApp.ScreenUpdating = false;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        excelApp.EnableEvents = false;

        var conflictCells = new List<ConflictData>();
        var sheetRanges = CollectSheetAddresses();

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        // 各セルの値を保持する辞書
        var baseValuesDict = new Dictionary<Tuple<string, string>, RangeData>();
        var cellData = new Dictionary<(string sheetName, int row, int col), List<(object value, int sourceFileIndex)>>();

        // ベースシートの値を収集
        foreach (var sheetName in sheetRanges.Keys)
        {
            var baseSheet = baseWorkbook.Sheets[sheetName];

            foreach (var sheetRange in sheetRanges[sheetName])
            {
                var rangeAddress = sheetRange.Address;
                var baseRange = baseSheet.Range[rangeAddress];
                var baseValues = baseRange.Value2 as object[,];
                IEnumerable<object> idValues = null;
                var key = Tuple.Create(sheetName, rangeAddress);

                if (sheetRange.RangeInfo?.IdColumnOffset != null)
                {
                    idValues = GetColumnWithOffset(baseSheet, rangeAddress, sheetRange.RangeInfo.IdColumnOffset.Value);
                }

                var value = new RangeData
                {
                    Values = baseValues,
                    IdValues = idValues,
                };

                baseValuesDict[key] = value;
            }
        }

        foreach (var mergeFilePath in mergeFilePaths)
        {
            var mergeWorkbook = excelApp.Workbooks.Open(mergeFilePath);
            int mergeFileIndex = mergeFilePaths.IndexOf(mergeFilePath);

            foreach (var sheetName in sheetRanges.Keys)
            {
                var mergeSheet = GetSheetIfExists(mergeWorkbook, sheetName);

                if (mergeSheet == null)
                {
                    continue;
                }

                foreach (var sheetRange in sheetRanges[sheetName])
                {
                    var rangeAddress = sheetRange.Address;
                    var key = Tuple.Create(sheetName, rangeAddress);
                    var baseValues = baseValuesDict[key].Values;
                    var baseIdValues = baseValuesDict[key].IdValues;

                    object[,] GetSortedMergeSheetValuesById()
                    {
                        if (baseIdValues == null)
                        {
                            return null;
                        }
                        SheetAddressInfo mergeSheetAddressInfo = GetSheetAddressInfo(mergeSheet);
                        var rangeInfo = mergeSheetAddressInfo?.RangeInfo;
                        if (rangeInfo == null)
                        {
                            return null;
                        }
                        if (!rangeInfo.IdColumnOffset.HasValue)
                        {
                            return null;
                        }
                        var idColumnOffset = rangeInfo.IdColumnOffset.Value;

                        var mergeRangeAddress = mergeSheetAddressInfo.Address;
                        var range = mergeSheet.Range[mergeRangeAddress];
                        var values = range.Value2 as object[,];
                        var idValues = GetColumnWithOffset(mergeSheet, mergeRangeAddress, idColumnOffset);

                        // idValues を key にした行（List<object>）の dictionary を作る
                        var valuesDictionary = CreateRowDictionaryWithIDKeys(values, idValues);

                        // baseValues のコピーを作って、mergeValuesからidを基に上書きコピーする
                        // idが見つからない行、ignoreColumn は何もしないので、baseのものが採用される
                        var result = CopyValuesById(baseValues, baseIdValues, valuesDictionary, rangeInfo.IgnoreColumnOffsets);
                        
                        return result;
                    }

                    // baseSheet に ID が存在する場合、 mergeSheet の値も ID から検索する
                    var mergeValues = GetSortedMergeSheetValuesById();
                    if (mergeValues == null)
                    {
                        var mergeRange = mergeSheet.Range[rangeAddress];
                        mergeValues = mergeRange.Value2 as object[,];
                    }

                    // 各セルの値を収集
                    for (int row = 1; row <= mergeValues.GetLength(0); row++)
                    {
                        for (int col = 1; col <= mergeValues.GetLength(1); col++)
                        {
                            var mergeValue = mergeValues[row, col]?.ToString() ?? "";
                            var baseValue = baseValues[row, col]?.ToString() ?? "";

                            if (mergeValue != baseValue)
                            {
                                var cellKey = (sheetName: sheetName, row: row, col: col);

                                if (!cellData.ContainsKey(cellKey))
                                {
                                    cellData[cellKey] = new List<(object value, int sourceFileIndex)>();
                                }

                                cellData[cellKey].Add((mergeValue, mergeFileIndex));
                            }
                        }
                    }
                }
            }
            mergeWorkbook.Close(false);
        }

        var sheetNames = cellData.Keys.Select(k => k.sheetName).Distinct();

        excelApp.StatusBar = false;
        excelApp.ScreenUpdating = true;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        excelApp.EnableEvents = true;

        var selectedSheets = ShowMergeSheetSelection(sheetNames);

        if (selectedSheets.Count == 0)
        {
            MessageBox.Show("キャンセルされました");

            return;
        }

        excelApp.ScreenUpdating = false;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        excelApp.EnableEvents = false;

        //cellData = FilterCellData(cellData, selectedSheets);

        List<(string sheetName, int numCells)> mergedSheets = new List<(string sheetName, int numCells)>();

        // 競合をチェックしてマージ
        foreach (var sheetName in selectedSheets)
        {
            // sheetName が cellValues に存在しない場合はスキップ
            var relevantKeys = cellData.Keys.Where(key => key.sheetName == sheetName).ToList();
            if (!relevantKeys.Any())
            {
                continue;
            }

            mergedSheets.Add((sheetName: sheetName, numCells: relevantKeys.Count));

            var baseSheet = baseWorkbook.Sheets[sheetName];

            foreach (var rangeTuple in sheetRanges[sheetName])
            {
                var rangeAddress = rangeTuple.Address;
                var mergeFunc = rangeTuple.Function;
                var baseRange = baseSheet.Range[rangeAddress];
                var baseValues = baseValuesDict[Tuple.Create(sheetName, rangeAddress)].Values;

                foreach (var key in relevantKeys)
                {
                    int row = key.row;
                    int col = key.col;
                    var baseValue = baseValues[row, col]?.ToString() ?? "";
                    var values = cellData[key];

                    if (values.Count == 1)
                    {
                        baseValues[row, col] = values[0].value;
                        continue;
                    }

                    var uniqueValues = values.ToLookup(cell => cell.value, cell => cell.sourceFileIndex);

                    if (uniqueValues.Count == 1)
                    {
                        baseValues[row, col] = values[0].value;
                        continue;
                    }

                    if (mergeFunc != null)
                    {
                        var mergeResult = mergeFunc(baseValue, uniqueValues.ToArray());
                        if (mergeResult.merged)
                        {
                            baseValues[row, col] = mergeResult.value;
                            continue;
                        }
                    }

                    var conflictedValues = uniqueValues.Select(cell => $"{cell.First() + 1}: {cell.Key}");

                    // mergeFunc が null またはマージに失敗した場合
                    baseValues[row, col] = $"※競合※\nbase: {baseValue}\n" + string.Join("\n", conflictedValues);

                    // 最大のインデックスを取得
                    int maxIndex = uniqueValues.SelectMany(g => g).Max();

                    // List<object> を初期化し、null で埋める
                    List<object> resultList = Enumerable.Range(0, maxIndex + 1)
                                                        .Select(i => (object)null)
                                                        .ToList();

                    // ILookup<object, int> を List<object> に変換
                    uniqueValues.SelectMany(group => group.Select(index => new { group.Key, index }))
                                .ToList()
                                .ForEach(item => resultList[item.index] = item.Key);

                    var conflictInfo = new ConflictData
                    {
                        SheetName = sheetName,
                        CellAddress = baseRange.Cells[row, col].Address(RowAbsolute: false, ColumnAbsolute: false),
                        Base = baseValue,
                        Merged = null,
                        Resolved = false,
                        Values = resultList,
                    };
                    conflictCells.Add(conflictInfo);
                }

                // 変更をシートに反映
                baseRange.Value2 = baseValues;
            }
        }
        stopwatch.Stop();
        Console.Write($"実行時間: {stopwatch.ElapsedMilliseconds}ミリ秒");

        if (mergedSheets.Count != 0)
        {
            //baseWorkbook.Save();
        }

        excelApp.StatusBar = false;
        excelApp.ScreenUpdating = true;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        excelApp.EnableEvents = true;

        // 競合があった場合にウィンドウを表示
        if (conflictCells.Count > 0)
        {
            ShowConflictWindow(conflictCells);
        }

        ShowResultWindow(mergedSheets);
    }

    private List<string> ShowMergeSheetSelection(IEnumerable<string> sheetNames)
    {
        List<string> selectedSheets = new List<string>();

        Form selectionForm = new Form
        {
            Text = "マージ対象のシート選択",
            Width = 800, // 初期幅を広めに設定
            Height = 400,
            TopMost = true
        };

        DataGridView gridView = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false, // 空の行が追加されないようにする
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // フォントサイズを大きく設定
        };

        // チェックボックス列を追加
        var checkBoxColumn = new DataGridViewCheckBoxColumn
        {
            HeaderText = "選択",
            Name = "Selected",
            TrueValue = true,
            FalseValue = false,
            Width = 50
        };
        gridView.Columns.Add(checkBoxColumn);

        // シート名列を追加
        var sheetNameColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "シート名",
            Name = "SheetName",
            ReadOnly = true
        };
        gridView.Columns.Add(sheetNameColumn);

        // データを追加
        foreach (var sheetName in sheetNames)
        {
            gridView.Rows.Add(true, sheetName);
        }

        // 列幅とヘッダーの高さを自動調整
        gridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        gridView.AutoResizeColumnHeadersHeight();
        gridView.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);

        // ボタンのフォントとサイズを設定
        var buttonFont = new System.Drawing.Font("Microsoft Sans Serif", 14);
        var buttonHeight = 50;

        // 全選択/全解除ボタン
        Button toggleButton = new Button
        {
            Text = "全選択/全解除",
            Dock = DockStyle.Top,
            Font = buttonFont,
            Height = buttonHeight
        };
        toggleButton.Click += (sender, e) =>
        {
            bool allChecked = gridView.Rows.Cast<DataGridViewRow>().All(row => (bool)row.Cells["Selected"].Value);
            foreach (DataGridViewRow row in gridView.Rows)
            {
                row.Cells["Selected"].Value = !allChecked;
            }
        };

        // OKボタン
        Button okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Dock = DockStyle.Bottom,
            Font = buttonFont,
            Height = buttonHeight
        };

        // キャンセルボタン
        Button cancelButton = new Button
        {
            Text = "キャンセル",
            DialogResult = DialogResult.Cancel,
            Dock = DockStyle.Bottom,
            Font = buttonFont,
            Height = buttonHeight
        };

        // シート名をダブルクリックで表示
        gridView.CellDoubleClick += (sender, e) =>
        {
            if (e.RowIndex >= 0)
            {
                var sheetName = gridView.Rows[e.RowIndex].Cells["SheetName"].Value.ToString();
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var sheet = (Excel.Worksheet)excelApp.Sheets[sheetName];
                sheet.Activate();
            }
        };

        selectionForm.Controls.Add(gridView);
        selectionForm.Controls.Add(toggleButton);
        selectionForm.Controls.Add(okButton);
        selectionForm.Controls.Add(cancelButton);

        // フォームのサイズを調整して全体が表示されるようにする
        selectionForm.Load += (sender, e) =>
        {
            gridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells); // 列幅を自動調整
            gridView.AutoResizeColumnHeadersHeight(); // ヘッダーの高さを自動調整
            gridView.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells); // 行の高さを自動調整
            selectionForm.Width = gridView.PreferredSize.Width + 40; // 余白を考慮して調整
        };

        var result = selectionForm.ShowDialog();

        if (result == DialogResult.OK)
        {
            selectedSheets = gridView.Rows.Cast<DataGridViewRow>()
                .Where(row => (bool)row.Cells["Selected"].Value)
                .Select(row => row.Cells["SheetName"].Value.ToString())
                .ToList();
        }

        return selectedSheets;
    }

    private Dictionary<(string sheetName, int row, int col), List<(object value, int sourceFileIndex)>> FilterCellData(
        Dictionary<(string sheetName, int row, int col), List<(object value, int sourceFileIndex)>> cellData,
        List<string> selectedSheets)
    {
        return cellData
            .Where(kvp => selectedSheets.Contains(kvp.Key.sheetName))
            .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
    }

    private void ShowResultWindow(IEnumerable<(string sheetName, int numCells)> mergedSheets)
    {
        if (mergedSheets.Count() == 0)
        {
            MessageBox.Show("変更箇所はありませんでした");
            return;
        }

        Form resultForm = new Form
        {
            Text = "マージ完了",
            Width = 400,
            Height = 300,
            TopMost = true // topmostに設定
        };

        ListBox mergedSheetListBox = new ListBox
        {
            Dock = DockStyle.Fill,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // 文字を大きく設定
        };

        foreach (var cell in mergedSheets)
        {
            mergedSheetListBox.Items.Add($"{cell.sheetName}: {cell.numCells}");
        }

        mergedSheetListBox.DoubleClick += (sender, e) =>
        {
            if (mergedSheetListBox.SelectedItem != null)
            {
                var selectedCell = mergedSheets.ElementAtOrDefault(mergedSheetListBox.SelectedIndex);
                var sheetName = selectedCell.sheetName;
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var sheet = (Excel.Worksheet)excelApp.Sheets[sheetName];
                sheet.Activate();
            }
        };

        resultForm.Controls.Add(mergedSheetListBox);
        resultForm.ShowDialog();

        //MessageBox.Show("マージを確定するにはブックを保存してください。");
    }

    private void SelectExcelCell(string sheetName, string cellAddress)
    {
        var excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        var sheet = excelApp.Worksheets[sheetName];
        var cell = sheet.Range[cellAddress];

        sheet.Activate();
        cell.Select();
        excelApp.ActiveWindow.ScrollRow = cell.Row;
        excelApp.ActiveWindow.ScrollColumn = cell.Column;
    }

    private void ShowConflictWindow(List<ConflictData> conflictData)
    {
        if (conflictData.Count() == 0)
        {
            return;
        }

        Form conflictForm = new Form
        {
            Text = "競合がありました",
            Width = 800,
            Height = 300,
            TopMost = true // topmostに設定
        };

        // 最大要素数を取得
        int maxValuesCount = conflictData.Max(c => c.Values.Count);

        // BindingListに変換
        BindingList<ConflictData> bindingList = new BindingList<ConflictData>(conflictData);

        DataGridView conflictDataGridView;
        Button okButton;
        Button cancelButton;

        void SetupDataGridView()
        {
            conflictDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                //Font = new System.Drawing.Font("Microsoft Sans Serif", 14), // 文字を大きく設定
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            };

            // 固定列の追加
            DataGridViewCheckBoxColumn resolvedColumn = new DataGridViewCheckBoxColumn
            {
                DataPropertyName = "Resolved",
                Name = "Resolved",
                HeaderText = "Resolved",
            };
            conflictDataGridView.Columns.Add(resolvedColumn);

            DataGridViewTextBoxColumn sheetNameColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SheetName",
                Name = "SheetName",
                HeaderText = "Sheet Name",
                ReadOnly = true,
            };
            conflictDataGridView.Columns.Add(sheetNameColumn);

            DataGridViewTextBoxColumn cellAddressColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "CellAddress",
                Name = "CellAddress",
                HeaderText = "Add",
                ReadOnly = true,
            };
            conflictDataGridView.Columns.Add(cellAddressColumn);

            DataGridViewTextBoxColumn mergedColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Merged",
                Name = "Merged",
                HeaderText = "Merged",
            };
            conflictDataGridView.Columns.Add(mergedColumn);
            conflictDataGridView.Columns["Merged"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            DataGridViewTextBoxColumn baseColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Base",
                Name = "Base",
                HeaderText = "Base",
                ReadOnly = true,
            };
            conflictDataGridView.Columns.Add(baseColumn);
            conflictDataGridView.Columns["Base"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            // 動的列の追加
            for (int i = 0; i < maxValuesCount; i++)
            {
                DataGridViewTextBoxColumn valuesColumn = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = $"Value{i + 1}",
                    Name = $"Value{i + 1}",
                    HeaderText = $"Value {i + 1}",
                    ReadOnly = true,
                };
                conflictDataGridView.Columns.Add(valuesColumn);
                conflictDataGridView.Columns[$"Value{i + 1}"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }

            // okButton ボタンを参照しているので SetupButtons 呼び出しより後に
            // Event handlers
            //conflictDataGridView.CellClick += DataGridView_CellClick;
            //conflictDataGridView.CellValueChanged += DataGridView_CellValueChanged;

            conflictForm.Controls.Add(conflictDataGridView);

            // データソースを設定
            conflictDataGridView.DataSource = bindingList;
        }

        void SetupButtons()
        {
            var buttonFont = new System.Drawing.Font("Microsoft Sans Serif", 14);
            var buttonHeight = 50;

            // OKボタン
            okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
                Font = buttonFont,
                Height = buttonHeight,
                Enabled = false,
            };
            okButton.Click += OkButton_Click;
            conflictForm.Controls.Add(okButton);

            // キャンセルボタン
            cancelButton = new Button
            {
                Text = "キャンセル",
                DialogResult = DialogResult.Cancel,
                Dock = DockStyle.Bottom,
                Font = buttonFont,
                Height = buttonHeight,
            };

            cancelButton.Click += CancelButton_Click;
            conflictForm.Controls.Add(cancelButton);
        }

        SetupDataGridView();
        SetupButtons();

        // okButton ボタンを参照しているので SetupButtons 呼び出しより後に
        conflictDataGridView.CellClick += DataGridView_CellClick;
        conflictDataGridView.CellValueChanged += DataGridView_CellValueChanged;

        // checkbox がクリックで変化した時に即座にchangedイベントが呼ばれるために必要な処理
        conflictDataGridView.CurrentCellDirtyStateChanged += (sender, e) =>
        {
            if (conflictDataGridView.CurrentCellAddress.X == 0 &&
                conflictDataGridView.IsCurrentCellDirty)
            {
                conflictDataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        };

        // フォームのサイズを調整して全体が表示されるようにする
        conflictForm.Load += (sender, e) =>
        {
            conflictDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells); // 列幅を自動調整
            conflictDataGridView.AutoResizeColumnHeadersHeight(); // ヘッダーの高さを自動調整
            conflictDataGridView.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells); // 行の高さを自動調整
            conflictForm.Width = conflictDataGridView.PreferredSize.Width + 40; // 余白を考慮して調整

            okButton.Enabled = IsAllResolved();
        };

        void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView dataGridView = sender as DataGridView;

                if (e.ColumnIndex == 4 || e.ColumnIndex >= 5)
                {
                    string value = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    dataGridView.Rows[e.RowIndex].Cells["Merged"].Value = value;
                }

                string sheetName = dataGridView.Rows[e.RowIndex].Cells["SheetName"].Value.ToString();
                string cellAddress = dataGridView.Rows[e.RowIndex].Cells["CellAddress"].Value.ToString();
                SelectExcelCell(sheetName, cellAddress);
            }
        }

        bool IsAllResolved()
        {
            return conflictDataGridView.Rows.Cast<DataGridViewRow>()
                                    .All(row => (bool)row.Cells["Resolved"].Value);
        }
        //void UpdateOkButtonEnabled()
        //{
        //    okButton.Enabled = IsAllResolved();
        //}

        void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;

            if (e.ColumnIndex == 0)
            {
                okButton.Enabled = IsAllResolved();
            }
            else if (e.ColumnIndex == 3)
            {
                var row = dataGridView.Rows[e.RowIndex];
                if (!row.IsNewRow)
                {
                    var excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                    var mergedValue = row.Cells["Merged"].Value.ToString();
                    var sheetName = row.Cells["SheetName"].Value.ToString(); // シート名を取得
                    var cellAddress = row.Cells["CellAddress"].Value.ToString(); // セルアドレスを取得
                    var sheet = (Excel.Worksheet)excelApp.Sheets[sheetName];
                    var range = sheet.Range[cellAddress];
                    range.Value2 = mergedValue;
                }
            }

        }

        void OkButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("競合が解決されました。");
        }

        void CancelButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("マージをキャンセルするにはブックを保存せずに閉じてください。");
        }

        conflictForm.ShowDialog();
    }

    private void ShowFileSelectionForm()
    {
        var excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var baseWorkbook = excelApp.ActiveWorkbook;

        if (baseWorkbook == null)
        {
            MessageBox.Show("先にベースとなるブックを開いてください。");
            return;
        }

        Form fileSelectionForm = new Form
        {
            Text = "ファイル選択",
            Width = 1200, // 幅を2倍に設定
            Height = 530, // 高さを調整
        };

        ListBox fileListBox = new ListBox
        {
            Dock = DockStyle.Top,
            Height = 300,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // 文字を大きく設定
        };

        Button addButton = new Button
        {
            Text = "追加",
            Dock = DockStyle.Top,
            Height = 60,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // 文字を大きく設定
        };

        Button removeButton = new Button
        {
            Text = "削除",
            Dock = DockStyle.Top,
            Height = 60,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // 文字を大きく設定
        };

        Button closeButton = new Button
        {
            Text = "閉じる",
            Dock = DockStyle.Top,
            Height = 60,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 14) // 文字を大きく設定
        };

        fileSelectionForm.Controls.Add(closeButton);
        fileSelectionForm.Controls.Add(removeButton);
        fileSelectionForm.Controls.Add(addButton);
        fileSelectionForm.Controls.Add(fileListBox);

        // 既存のファイルをリストボックスに追加
        UpdateFileListBox(fileListBox);

        addButton.Click += (sender, e) =>
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var baseWorkbookPath = excelApp.ActiveWorkbook.FullName;

                foreach (string file in openFileDialog.FileNames)
                {
                    if (!mergeFilePaths.Contains(file) && file != baseWorkbookPath)
                    {
                        mergeFilePaths.Add(file);
                        UpdateFileListBox(fileListBox);
                    }
                }
            }
        };

        removeButton.Click += (sender, e) =>
        {
            RemoveSelectedItems(fileListBox);
            UpdateFileListBox(fileListBox);
        };

        closeButton.Click += (sender, e) =>
        {
            fileSelectionForm.Close();
        };

        // リストボックスにドラッグアンドドロップを有効にする
        fileListBox.AllowDrop = true;
        fileListBox.DragEnter += new DragEventHandler(Form_DragEnter);
        fileListBox.DragDrop += new DragEventHandler(Form_DragDrop);

        // KeyDownイベントを追加
        fileListBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Delete)
            {
                RemoveSelectedItems(fileListBox);
                UpdateFileListBox(fileListBox);
            }
        };

        // ESCキーでフォームを閉じる
        fileSelectionForm.KeyPreview = true;
        fileSelectionForm.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                fileSelectionForm.Close();
            }
        };

        fileSelectionForm.ShowDialog();
    }

    private void UpdateFileListBox(ListBox fileListBox)
    {
        fileListBox.Items.Clear();
        for (int i = 0; i < mergeFilePaths.Count; i++)
        {
            fileListBox.Items.Add($"{i + 1}. {mergeFilePaths[i]}");
        }
    }

    private void RemoveSelectedItems(ListBox fileListBox)
    {
        var selectedItems = fileListBox.SelectedItems.Cast<string>().ToList();
        foreach (var item in selectedItems)
        {
            var filePath = item.Substring(item.IndexOf(' ') + 1); // インデックスを除去してファイルパスを取得
            mergeFilePaths.Remove(filePath);
            fileListBox.Items.Remove(item);
        }
    }

    private void Form_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effect = DragDropEffects.Copy;
        }
    }

    private void Form_DragDrop(object sender, DragEventArgs e)
    {
        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        var excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var baseWorkbookPath = excelApp.ActiveWorkbook.FullName;
        foreach (var file in files)
        {
            if (!mergeFilePaths.Contains(file) && file != baseWorkbookPath)
            {
                mergeFilePaths.Add(file);
            }
        }
        UpdateFileListBox((ListBox)sender);
    }
}

[ComVisible(true)]
public class MyRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonID)
    {
        string projectName = Assembly.GetExecutingAssembly().GetName().Name;
        return $@"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='customTab' label='{projectName}'>
        <group id='customGroup' label='Merge'>
          <button id='selectFilesButton' label='ファイル選択' size='large' imageMso='FileSave' onAction='OnSelectFilesButtonClick' />
          <button id='mergeButton' label='Merge' size='large' imageMso='TableDrawTable' onAction='OnMergeButtonClick' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnMergeButtonClick(IRibbonControl control)
    {
        // マージ処理を呼び出す
        ExcelMergeTool.Instance.OnMergeButtonClick(control);
    }

    public void OnSelectFilesButtonClick(IRibbonControl control)
    {
        // ファイル選択処理を呼び出す
        ExcelMergeTool.Instance.OnSelectFilesButtonClick(control);
    }
}
