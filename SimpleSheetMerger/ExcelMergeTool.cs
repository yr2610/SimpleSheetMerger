using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelMergeTool : IExcelAddIn
{
    private static ExcelMergeTool _instance;
    public static ExcelMergeTool Instance => _instance ?? (_instance = new ExcelMergeTool());

    private List<string> updateFiles = new List<string>();

    public void AutoOpen()
    {
        // リボンのカスタマイズを登録
    }

    public void AutoClose()
    {
        // クリーンアップコード
    }

    public void ShowFileListForm()
    {
        var form = new Form
        {
            Width = 900, // ウィンドウの幅を3倍に広げる
            Height = 600
        };
        var addButton = new Button
        {
            Text = "追加",
            Top = 10,
            Left = 10,
            Width = 200, // ボタンの幅を2倍に
            Height = 60, // ボタンの高さを2倍に
            Font = new System.Drawing.Font("Microsoft Sans Serif", 16) // 文字サイズを2倍に
        };
        var removeButton = new Button
        {
            Text = "削除",
            Top = 80,
            Left = 10,
            Width = 200, // ボタンの幅を2倍に
            Height = 60, // ボタンの高さを2倍に
            Font = new System.Drawing.Font("Microsoft Sans Serif", 16) // 文字サイズを2倍に
        };
        var listBox = new ListBox
        {
            Top = 150,
            Left = 10,
            Width = 860, // リストボックスの幅をウィンドウに合わせる
            Height = 400,
            Font = new System.Drawing.Font("Microsoft Sans Serif", 16) // 文字サイズを2倍に
        };

        addButton.Click += (sender, e) => {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                if (!updateFiles.Contains(filePath))
                {
                    updateFiles.Add(filePath);
                    listBox.Items.Add(filePath);
                }
            }
        };

        removeButton.Click += (sender, e) => {
            if (listBox.SelectedItem != null)
            {
                updateFiles.Remove(listBox.SelectedItem.ToString());
                listBox.Items.Remove(listBox.SelectedItem);
            }
        };

        listBox.AllowDrop = true;
        listBox.DragEnter += (sender, e) => {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        };
        listBox.DragDrop += (sender, e) => {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files)
            {
                if ((Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                     Path.GetExtension(file).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)) &&
                    !updateFiles.Contains(file))
                {
                    updateFiles.Add(file);
                    listBox.Items.Add(file);
                }
            }
        };

        form.Controls.Add(addButton);
        form.Controls.Add(removeButton);
        form.Controls.Add(listBox);
        form.ShowDialog();
    }

    public void MergeExcelFiles()
    {
        var excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var baseWorkbook = excelApp.ActiveWorkbook;
        var baseFilePath = baseWorkbook.FullName;

        if (string.IsNullOrEmpty(baseFilePath))
        {
            MessageBox.Show("ベースブックを保存してください。");
            return;
        }

        var backupFilePath = Path.Combine(Path.GetDirectoryName(baseFilePath),
            $"{Path.GetFileNameWithoutExtension(baseFilePath)}_backup_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        baseWorkbook.SaveCopyAs(backupFilePath);

        var mergedData = new Dictionary<string, object>();

        foreach (var file in updateFiles)
        {
            var updateWorkbook = excelApp.Workbooks.Open(file);
            foreach (Excel.Worksheet sheet in updateWorkbook.Sheets)
            {
                foreach (Excel.Range cell in sheet.UsedRange)
                {
                    var key = $"{sheet.Name}!{cell.Address}";
                    if (!mergedData.ContainsKey(key))
                    {
                        mergedData[key] = cell.Value;
                    }
                    else
                    {
                        // 競合処理
                        if (!mergedData[key].Equals(cell.Value))
                        {
                            MessageBox.Show($"競合検出: {key}");
                        }
                    }
                }
            }
            updateWorkbook.Close(false);
        }

        // マージ結果をベースブックに反映
        foreach (var kvp in mergedData)
        {
            var parts = kvp.Key.Split('!');
            var sheetName = parts[0];
            var cellAddress = parts[1];

            var sheet = baseWorkbook.Sheets[sheetName] as Excel.Worksheet;
            if (sheet == null)
            {
                sheet = (Excel.Worksheet)baseWorkbook.Sheets.Add(After: baseWorkbook.Sheets[baseWorkbook.Sheets.Count]);
                sheet.Name = sheetName;
            }

            sheet.Range[cellAddress].Value = kvp.Value;
        }

        baseWorkbook.Save();
    }
}

[ComVisible(true)]
public class CustomRibbon : ExcelRibbon
{
    private const string ribbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='customTab' label='マージツール'>
        <group id='customGroup' label='操作'>
          <button id='selectFilesButton' label='ファイル選択' onAction='SelectFiles' />
          <button id='mergeFilesButton' label='マージ' onAction='MergeFiles' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public override string GetCustomUI(string ribbonID)
    {
        return ribbonXml;
    }

    public void SelectFiles(IRibbonControl control)
    {
        ExcelMergeTool.Instance.ShowFileListForm();
    }

    public void MergeFiles(IRibbonControl control)
    {
        ExcelMergeTool.Instance.MergeExcelFiles();
    }
}
