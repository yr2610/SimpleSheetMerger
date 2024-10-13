using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
    private List<string> conflictCells = new List<string>();
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
        public Func<string, string[], Tuple<bool, string>> Function { get; set; }
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

        var sheetRanges = CollectSheetAddresses();

        // 各セルの値を保持する辞書
        var cellValues = new Dictionary<Tuple<string, int, int>, List<string>>();
        var cellSources = new Dictionary<Tuple<string, int, int>, List<string>>();
        var baseValuesDict = new Dictionary<Tuple<string, string>, RangeData>();

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
                                var cellKey = Tuple.Create(sheetName, row, col);

                                if (!cellValues.ContainsKey(cellKey))
                                {
                                    cellValues[cellKey] = new List<string>();
                                    cellSources[cellKey] = new List<string>();
                                }

                                if (!cellValues[cellKey].Contains(mergeValue))
                                {
                                    cellValues[cellKey].Add(mergeValue);
                                    cellSources[cellKey].Add($"{mergeFilePaths.IndexOf(mergeFilePath) + 1}: {mergeValue}");
                                }
                            }
                        }
                    }
                }
            }
            mergeWorkbook.Close(false);
        }

        // 競合をチェックしてマージ
        foreach (var sheetName in sheetRanges.Keys)
        {
            // sheetName が cellValues に存在しない場合はスキップ
            var relevantKeys = cellValues.Keys.Where(key => key.Item1 == sheetName).ToList();
            if (!relevantKeys.Any())
            {
                continue;
            }

            var baseSheet = baseWorkbook.Sheets[sheetName];

            foreach (var rangeTuple in sheetRanges[sheetName])
            {
                var rangeAddress = rangeTuple.Address;
                var mergeFunc = rangeTuple.Function;
                var baseRange = baseSheet.Range[rangeAddress];
                var baseValues = baseValuesDict[Tuple.Create(sheetName, rangeAddress)].Values;

                foreach (var key in relevantKeys)
                {
                    int row = key.Item2;
                    int col = key.Item3;
                    var baseValue = baseValues[row, col]?.ToString() ?? "";
                    var values = cellValues[key];

                    if (values.Count == 1)
                    {
                        baseValues[row, col] = values[0];
                        continue;
                    }

                    var uniqueValues = new HashSet<string>(values);

                    if (uniqueValues.Count == 1)
                    {
                        baseValues[row, col] = uniqueValues.First();
                        continue;
                    }

                    if (mergeFunc != null)
                    {
                        var mergeResult = mergeFunc(baseValue, uniqueValues.ToArray());
                        if (mergeResult.Item1)
                        {
                            baseValues[row, col] = mergeResult.Item2;
                            continue;
                        }
                    }

                    // mergeFunc が null またはマージに失敗した場合
                    baseValues[row, col] = $"※競合※\nbase: {baseValue}\n" + string.Join("\n", cellSources[key]);
                    conflictCells.Add($"{sheetName}: {baseRange.Cells[row, col].Address}");
                }

                // 変更をシートに反映
                baseRange.Value2 = baseValues;
            }
        }

        baseWorkbook.Save();

        excelApp.StatusBar = false;
        excelApp.ScreenUpdating = true;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        excelApp.EnableEvents = true;

        // 競合があった場合にウィンドウを表示
        if (conflictCells.Count > 0)
        {
            ShowConflictWindow();
        }
    }

    private void ShowConflictWindow()
    {
        Form conflictForm = new Form
        {
            Text = "競合がありました",
            Width = 400,
            Height = 300,
            TopMost = true // topmostに設定
        };

        ListBox conflictListBox = new ListBox
        {
            Dock = DockStyle.Fill
        };

        foreach (var cell in conflictCells)
        {
            conflictListBox.Items.Add(cell);
        }

        conflictListBox.DoubleClick += (sender, e) =>
        {
            if (conflictListBox.SelectedItem != null)
            {
                var selectedCell = conflictListBox.SelectedItem.ToString();
                var parts = selectedCell.Split(':');
                var sheetName = parts[0].Trim();
                var cellAddress = parts[1].Trim();

                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var sheet = (Excel.Worksheet)excelApp.Sheets[sheetName];
                var range = sheet.Range[cellAddress];
                sheet.Activate();
                range.Select();
            }
        };

        conflictForm.Controls.Add(conflictListBox);
        conflictForm.Show();
    }

    private void ShowFileSelectionForm()
    {
        var excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var baseWorkbook = excelApp.ActiveWorkbook;

        if (baseWorkbook == null)
        {
            MessageBox.Show("先にブックを開いてください。");
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
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='customTab' label='マージツール'>
        <group id='customGroup' label='操作'>
          <button id='selectFilesButton' label='ファイル選択' onAction='OnSelectFilesButtonClick' />
          <button id='mergeButton' label='マージ' onAction='OnMergeButtonClick' />
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
