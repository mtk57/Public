using System;
using System.Reflection;
using SimpleExcelGrep.Services;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// Excel InteropサービスクラスはExcelファイルの操作を担当
    /// </summary>
    internal class ExcelInteropService
    {
        private readonly LogService _logger;

        /// <summary>
        /// ExcelInteropServiceのコンストラクタ
        /// </summary>
        /// <param name="logger">ログサービス</param>
        public ExcelInteropService(LogService logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Excelファイルを開く
        /// </summary>
        public bool OpenExcelFile(string filePath, string sheetName, string cellPosition)
        {
            _logger.LogMessage($"OpenExcelFile 開始: ファイル={filePath}, シート={sheetName}, セル={cellPosition}");

            try
            {
                // Excel Interop を使用して開く
                _logger.LogMessage("Excel Interop を使用して開こうとしています...");
                bool interopSuccess = OpenExcelWithInterop(filePath, sheetName, cellPosition);
                _logger.LogMessage($"Excel Interop 結果: {(interopSuccess ? "成功" : "失敗")}");

                if (interopSuccess)
                {
                    return true;
                }

                // 通常の方法でファイルを開く
                _logger.LogMessage("通常の方法でファイルを開きます");
                System.Diagnostics.Process.Start(filePath);
                return true;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Excelファイルを開けませんでした: {ex.Message}";
                _logger.LogMessage($"エラー: {errorMsg}\n詳細: {ex.ToString()}", true);
                return false;
            }
        }

        /// <summary>
        /// Excel Interopを使用してファイルを開く実装（Late Binding方式）
        /// </summary>
        private bool OpenExcelWithInterop(string filePath, string sheetName, string cellPosition)
        {
            _logger.LogMessage($"OpenExcelWithInterop 開始 (Late Binding方式)");

            object excelApp = null;
            object workbooks = null;
            object workbook = null;

            try
            {
                // COMの状態を確認
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    _logger.LogMessage("警告: Excel.ApplicationのCOMタイプが見つかりません。Excel Interopが正しくインストールされていない可能性があります。");
                    return false;
                }

                // Excel アプリケーションを起動
                _logger.LogMessage("Excel アプリケーションのインスタンス作成中...");
                excelApp = Activator.CreateInstance(excelType);

                // バージョン情報を取得
                object version = excelType.InvokeMember("Version",
                    BindingFlags.GetProperty, null, excelApp, null);
                _logger.LogMessage($"Excel バージョン: {version}");

                // Visible プロパティを設定
                excelType.InvokeMember("Visible",
                    BindingFlags.SetProperty, null, excelApp, new object[] { true });

                // ファイルを開く
                _logger.LogMessage($"ファイルを開いています: {filePath}");

                // Workbooks コレクションを取得
                workbooks = excelType.InvokeMember("Workbooks",
                    BindingFlags.GetProperty, null, excelApp, null);

                // Open メソッドを呼び出す
                Type workbooksType = workbooks.GetType();
                workbook = workbooksType.InvokeMember("Open",
                    BindingFlags.InvokeMethod, null, workbooks, new object[] {
                        filePath,       // ファイルパス
                        Type.Missing,   // UpdateLinks
                        true,           // ReadOnly
                        Type.Missing,   // Format
                        Type.Missing,   // Password
                        Type.Missing,   // WriteResPassword
                        Type.Missing,   // IgnoreReadOnlyRecommended
                        Type.Missing,   // Origin
                        Type.Missing,   // Delimiter
                        Type.Missing,   // Editable
                        Type.Missing,   // Notify
                        Type.Missing,   // Converter
                        Type.Missing,   // AddToMru
                        Type.Missing,   // Local
                        Type.Missing    // CorruptLoad
                    });
                _logger.LogMessage("ワークブックを開きました");

                if (!string.IsNullOrEmpty(sheetName))
                {
                    if (TryActivateSheet(workbook, sheetName, cellPosition))
                    {
                        _logger.LogMessage("Excel Interopの処理が完了しました (Late Binding方式)");
                        return true;
                    }
                    else if (string.IsNullOrEmpty(cellPosition) || cellPosition == "図形内" || cellPosition == "図形内 (GF)")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                _logger.LogMessage("Excel Interopの処理が完了しました (Late Binding方式)");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"Excel Interopエラー (Late Binding): {ex.GetType().Name}: {ex.Message}");
                _logger.LogMessage($"スタックトレース: {ex.StackTrace}");
                return false;
            }
            finally
            {
                // リソースの解放
                try
                {
                    if (workbook != null)
                    {
                        _logger.LogMessage("COMオブジェクト (workbook) を解放します");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }

                    if (workbooks != null)
                    {
                        _logger.LogMessage("COMオブジェクト (workbooks) を解放します");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                        workbooks = null;
                    }

                    if (excelApp != null)
                    {
                        try
                        {
                            // 明示的にアプリケーションを終了（他のインスタンスに影響しないように）
                            Type excelType = excelApp.GetType();
                            excelType.InvokeMember("Quit",
                                BindingFlags.InvokeMethod, null, excelApp, null);
                            _logger.LogMessage("Excelアプリケーションを終了しました");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogMessage($"Excel終了エラー: {ex.Message}");
                        }

                        _logger.LogMessage("COMオブジェクト (excelApp) を解放します");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }

                    // GCを強制的に実行して、残りのCOMオブジェクトを確実に解放
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    _logger.LogMessage($"COMオブジェクト解放エラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// シートをアクティブ化して、必要に応じてセルを選択
        /// </summary>
        private bool TryActivateSheet(object workbook, string sheetName, string cellPosition)
        {
            object sheets = null;
            object targetSheet = null;
            Type sheetsType = null;
            object count = null; // count 変数をここで宣言
    
            try
            {
                Type workbookType = workbook.GetType();
                sheets = workbookType.InvokeMember("Sheets",
                    BindingFlags.GetProperty, null, workbook, null);
                sheetsType = sheets.GetType();

                // シート名一覧をログに出力
                _logger.LogMessage("利用可能なシート:");
                count = sheetsType.InvokeMember("Count",
                    BindingFlags.GetProperty, null, sheets, null);
        
                for (int i = 1; i <= (int)count; i++)
                {
                    object sheet = null;
                    object name = null;
            
                    try
                    {
                        sheet = sheetsType.InvokeMember("Item",
                            BindingFlags.GetProperty, null, sheets, new object[] { i });
                
                        Type sheetType = sheet.GetType();
                        name = sheetType.InvokeMember("Name",
                            BindingFlags.GetProperty, null, sheet, null);
                        
                        _logger.LogMessage($" - [{name}]");
                    }
                    finally
                    {
                        // 繰り返し処理の中で取得したCOMオブジェクトを解放
                        if (sheet != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                            sheet = null;
                        }
                
                        // 修正: name オブジェクトも解放する
                        if (name != null && System.Runtime.InteropServices.Marshal.IsComObject(name))
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(name);
                            name = null;
                        }
                    }
                }

                // 指定されたシート名のシートを取得
                _logger.LogMessage($"シート '{sheetName}' を検索中...");

                try
                {
                    targetSheet = sheetsType.InvokeMember("Item",
                        BindingFlags.GetProperty, null, sheets, new object[] { sheetName });
                    _logger.LogMessage($"シート '{sheetName}' が見つかりました");
                }
                catch (Exception ex)
                {
                    _logger.LogMessage($"シート取得エラー: {ex.Message}");
                    return false;
                }

                // シートをアクティブにする
                if (targetSheet != null)
                {
                    _logger.LogMessage($"シート '{sheetName}' をアクティブ化します");
                    Type sheetType = targetSheet.GetType();
                    sheetType.InvokeMember("Activate",
                        BindingFlags.InvokeMethod, null, targetSheet, null);
                    _logger.LogMessage("シートのアクティブ化に成功しました");

                    // セル位置が指定されていれば選択
                    if (!string.IsNullOrEmpty(cellPosition) && 
                        cellPosition != "図形内" && cellPosition != "図形内 (GF)")
                    {
                        bool result = TrySelectCell(targetSheet, cellPosition);
                        return result;
                    }
            
                    return true;
                }
                else
                {
                    _logger.LogMessage($"警告: シート '{sheetName}' が見つかりませんでした");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"シートのアクティブ化エラー: {ex.Message}");
                return false;
            }
            finally
            {
                // すべてのCOMオブジェクトを解放
                if (targetSheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetSheet);
                    targetSheet = null;
                }
        
                if (sheets != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                    sheets = null;
                }
        
                // 修正: count オブジェクトも解放する
                if (count != null && System.Runtime.InteropServices.Marshal.IsComObject(count))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(count);
                    count = null;
                }
            }
        }

        /// <summary>
        /// 指定したシート上のセルを選択
        /// </summary>
        private bool TrySelectCell(object sheet, string cellPosition)
        {
            _logger.LogMessage($"セル {cellPosition} を選択します");
            object range = null;
            
            try
            {
                Type sheetType = sheet.GetType();
                
                // Rangeを取得
                range = sheetType.InvokeMember("Range",
                    BindingFlags.GetProperty, null, sheet, new object[] { cellPosition });

                // セルを選択
                Type rangeType = range.GetType();
                rangeType.InvokeMember("Select",
                    BindingFlags.InvokeMethod, null, range, null);
                
                _logger.LogMessage("セルの選択に成功しました");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"セル選択エラー: {ex.Message}");
                return false;
            }
            finally
            {
                // Rangeオブジェクトを解放
                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;
                }
            }
        }
    }
}