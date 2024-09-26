using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleSheetMerger
{
    class Merger
    {

        // string original = "1, 2, 3, 4, 5";
        // string[] updates = { "3, 4, 5, 6, 7", "8, 9, 10" };
        // 
        // string mergedText = MergeText(original, updates);
        // Console.WriteLine(mergedText); // 出力: 3, 4, 5, 6, 7, 8, 9, 10
        public static bool MergeText(string original, string[] updates, out string merged)
        {
            // TODO: コンマ区切りの整数じゃない場合に false 返すとか

            // 更新後の文字列の配列をすべて結合し、整数のリストに変換
            List<int> mergedList = new List<int>();
            foreach (string update in updates)
            {
                List<int> updateList = update.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                .Select(s => int.Parse(s.Trim()))
                                                .ToList();
                mergedList.AddRange(updateList);
            }

            // 重複を排除し、ソート
            mergedList = mergedList.Distinct().OrderBy(x => x).ToList();

            // , 区切りの文字列に変換して返す（, の後ろに半角スペースを追加）
            merged = string.Join(", ", mergedList);

            return true;
        }

#if false
        // string originalText = "Line1\nLine2\nLine3";
        // string[] updatedTexts = {
        //     "Line1\nLine2 modified\nLine3",
        //     "Line1\nLine2\nLine3\nLine4"
        // };
        // 
        // string resultText = DetectAndAppendDifferences(originalText, updatedTexts);
        public static string AppendText(string originalText, string[] updatedTexts)
        {
            var originalLines = originalText.Split('\n').ToList();
            var resultLines = new List<string>(originalLines);
            var differ = new Differ();
            var diffBuilder = new InlineDiffBuilder(differ);

            foreach (var updatedText in updatedTexts)
            {
                var diff = diffBuilder.BuildDiffModel(originalText, updatedText);
                bool hasMiddleChanges = diff.Lines.Any(line => line.Type == ChangeType.Modified || line.Type == ChangeType.Deleted);

                if (hasMiddleChanges)
                {
                    resultLines.Add("---");
                    resultLines.AddRange(updatedText.Split('\n'));
                }
                else
                {
                    var addedLines = updatedText.Split('\n').Skip(originalLines.Count);
                    if (addedLines.Any())
                    {
                        resultLines.Add("---");
                        resultLines.AddRange(addedLines);
                    }
                }
            }

            return string.Join("\n", resultLines);
        }
#endif

    }


    public static class MergeFunctions
    {
        // カスタムデリゲートを定義
        public delegate bool MergeFunctionDelegate(string baseValue, string[] mergeValues, out string result);

        public enum MergeFunctionType
        {
            MergeFunction,
            AnotherMergeFunction
        }

        public static bool MergeFunction(string baseValue, string[] mergeValues, out string result)
        {
            // マージロジックを実装します
            result = mergeValues[0]; // 例として最初の値を使用
            return true;
        }

        public static bool AnotherMergeFunction(string baseValue, string[] mergeValues, out string result)
        {
            // 別のマージロジックを実装します
            result = string.Join(",", mergeValues); // 例として全ての値を結合
            return true;
        }

        public static Func<string, string[], Tuple<bool, string>> GetWrappedFunction(string functionName)
        {
            if (Enum.TryParse(functionName, out MergeFunctionType functionType))
            {
                return GetWrappedFunction(functionType);
            }
            throw new ArgumentException("Invalid function name");
        }

        private static Func<string, string[], Tuple<bool, string>> GetWrappedFunction(MergeFunctionType functionType)
        {
            switch (functionType)
            {
                case MergeFunctionType.MergeFunction:
                    return WrapMergeFunction(MergeFunctions.MergeFunction);
                case MergeFunctionType.AnotherMergeFunction:
                    return WrapMergeFunction(MergeFunctions.AnotherMergeFunction);
                default:
                    throw new ArgumentException("Invalid MergeFunctionType");
            }
        }

        private static Func<string, string[], Tuple<bool, string>> WrapMergeFunction(MergeFunctionDelegate mergeFunc)
        {
            return (baseValue, mergeValues) =>
            {
                if (mergeFunc(baseValue, mergeValues, out var result))
                {
                    return Tuple.Create(true, result);
                }
                return Tuple.Create(false, result);
            };
        }

    }

}
