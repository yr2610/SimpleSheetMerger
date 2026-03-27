using System;
using System.Collections.Generic;
using System.Security.Principal;
using System.Windows.Forms;

namespace SimpleSheetMerger
{
    /// <summary>
    /// 暫定のユーザー認可をまとめるヘルパークラスです。
    /// 将来は別方式に置き換えやすいように、認可ロジックを1か所に集約します。
    /// </summary>
    internal static class AuthorizationHelper
    {
        // 暫定的に全ユーザー許可へ切り替えるためのフラグです。
        // 緊急回避用のため、既定値は必ず false にしておきます。
        internal static bool ALLOW_ALL_USERS = false;

        // 暫定実装のため、許可ユーザーはコード内に固定で保持します。
        // 空のままなら誰も許可しない仕様にし、常に deny by default を維持します。
        private static readonly HashSet<string> AllowedUsers = new HashSet<string>(StringComparer.Ordinal)
        {
            // 例: @"DOMAIN\UserName",
            @"LAPTOP-9S8RJR29\shinn",

        };

        /// <summary>
        /// 現在の Windows ユーザー名を取得します。
        /// まず DOMAIN\UserName 形式を優先し、失敗時だけ Environment.UserName にフォールバックします。
        /// </summary>
        internal static string GetCurrentUserName()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                if (identity != null && !string.IsNullOrWhiteSpace(identity.Name))
                {
                    return identity.Name;
                }
            }
            catch
            {
                // 暫定実装のため詳細な例外処理は入れず、単純なフォールバックだけを行います。
            }

            return Environment.UserName ?? string.Empty;
        }

        /// <summary>
        /// 許可済みユーザーかどうかだけを判定します。
        /// 空リストを allow all と解釈しないため、未登録ユーザーはすべて拒否します。
        /// </summary>
        internal static bool IsAuthorizedUser()
        {
            if (ALLOW_ALL_USERS)
            {
                return true;
            }

            string currentUserName = GetCurrentUserName();
            return AllowedUsers.Contains(currentUserName);
        }

        /// <summary>
        /// 実行直前の認可チェックを行います。
        /// UI の見た目だけでなく処理経路自体を止めるため、各入口でこのメソッドを呼びます。
        /// </summary>
        internal static bool EnsureAuthorizedUser()
        {
            if (IsAuthorizedUser())
            {
                return true;
            }

            string currentUserName = GetCurrentUserName();
            MessageBox.Show(
                "このユーザーには本アドインの使用許可がありません。" + Environment.NewLine +
                "Current user: " + currentUserName,
                "認可エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return false;
        }
    }
}
