﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelMergeTool : IExcelAddIn
{
    private static ExcelMergeTool _instance;
    public static ExcelMergeTool Instance => _instance ?? (_instance = new ExcelMergeTool());

    private List<string> mergeFilePaths = new List<string>();
    private List<string> conflictCells = new List<string>();
    private Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>> sheetRanges = new Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>>();


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

    static Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>> CollectSheetAddresses()
    {
        const string indexSheetName = "index"; // シート名
        const string startCellAddress = "B16"; // 開始セルのアドレス
        const string endMarker = "END"; // 終端を示す文字列
        const string leftColumnAddress = "U"; // 左端の列のアドレス
        const string rightColumnAddress = "AA"; // 右端の列のアドレス
        const string headerRowAddress = "AD"; // ヘッダー行のアドレス
        const string bottomRowAddress = "AE"; // 最下行のアドレス
        string[] ignoreSheetNames = { "無視シート", }; // 無視するシート名のリスト

        var result = new Dictionary<string, List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>>();

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

                string leftColumn = indexSheet.Cells[cell.Row, leftColumnAddress].Value.ToString();
                string rightColumn = indexSheet.Cells[cell.Row, rightColumnAddress].Value.ToString();
                int headerRow = (int)indexSheet.Cells[cell.Row, headerRowAddress].Value;
                int topRow = headerRow + 1;
                int bottomRow = (int)indexSheet.Cells[cell.Row, bottomRowAddress].Value;

                string address = $"{leftColumn}{topRow}:{rightColumn}{bottomRow}";

                // シート名が辞書に存在しない場合、新しいリストを作成
                if (!result.ContainsKey(sheetName))
                {
                    result[sheetName] = new List<Tuple<string, Func<string, string[], Tuple<bool, string>>>>();
                }

                // アドレスをリストに追加
                var rangeTuple = Tuple.Create(address, (Func<string, string[], Tuple<bool, string>>)null);

                result[sheetName].Add(rangeTuple);
            }
        }

        return result;
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

        sheetRanges = CollectSheetAddresses();

        // 各セルの値を保持する辞書
        var cellValues = new Dictionary<Tuple<string, int, int>, List<string>>();
        var cellSources = new Dictionary<Tuple<string, int, int>, List<string>>();
        var baseValuesDict = new Dictionary<Tuple<string, string>, object[,]>();

        // ベースシートの値を収集
        foreach (var sheetName in sheetRanges.Keys)
        {
            var baseSheet = baseWorkbook.Sheets[sheetName];

            foreach (var rangeTuple in sheetRanges[sheetName])
            {
                var rangeAddress = rangeTuple.Item1;
                var baseRange = baseSheet.Range[rangeAddress];
                var baseValues = baseRange.Value2 as object[,];
                var key = Tuple.Create(sheetName, rangeAddress);
                baseValuesDict[key] = baseValues;
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

                foreach (var rangeTuple in sheetRanges[sheetName])
                {
                    var rangeAddress = rangeTuple.Item1;
                    var mergeRange = mergeSheet.Range[rangeAddress];
                    var mergeValues = mergeRange.Value2 as object[,];
                    var key = Tuple.Create(sheetName, rangeAddress);
                    var baseValues = baseValuesDict[key];

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
            var baseSheet = baseWorkbook.Sheets[sheetName];

            foreach (var rangeTuple in sheetRanges[sheetName])
            {
                var rangeAddress = rangeTuple.Item1;
                var mergeFunc = rangeTuple.Item2;
                var baseRange = baseSheet.Range[rangeAddress];
                var baseValues = baseValuesDict[Tuple.Create(sheetName, rangeAddress)];

                for (int row = 1; row <= baseValues.GetLength(0); row++)
                {
                    for (int col = 1; col <= baseValues.GetLength(1); col++)
                    {
                        var baseValue = baseValues[row, col]?.ToString() ?? "";
                        var key = Tuple.Create(sheetName, row, col);

                        if (cellValues.ContainsKey(key))
                        {
                            var values = cellValues[key];
                            if (values.Count == 1)
                            {
                                baseValues[row, col] = values[0];
                            }
                            else if (values.Count > 1)
                            {
                                var uniqueValues = new HashSet<string>(values);
                                if (uniqueValues.Count > 1)
                                {
                                    if (mergeFunc != null)
                                    {
                                        var mergeResult = mergeFunc(baseValue, values.ToArray());
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
                                else
                                {
                                    baseValues[row, col] = values[0];
                                }
                            }
                        }
                    }
                }

                // 変更をシートに反映
                baseRange.Value2 = baseValues;
            }
        }

        baseWorkbook.Save();

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
