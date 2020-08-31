using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ClickCapture {
    /// <summary>
    /// メインクラス
    /// </summary>
    public static class Program {

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        public static void Main() {
            // ミューテックスを使用して多重起動を抑制する
            var eMutex = new System.Threading.Mutex(false, "MYSOFTWARE_001");
            // ミューテックスの所有権を要求
            if (!eMutex.WaitOne(0, false)) {
                return;
            }

            // アプリケーション準備
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 画面キャプチャクラス生成
            var eClickCapture = new ClickCapture();
            // アイコン設定処理
            eClickCapture.SetComponents();

            // アプリケーション実行
            Application.Run();
        }
    }

    /// <summary>
    /// 画面キャプチャクラス
    /// </summary>
    public class ClickCapture : Form {
        /// <summary>
        /// モード用メニュー項目リスト
        /// </summary>
        private List<ToolStripMenuItem> MenuGroupItems { set; get; } = new List<ToolStripMenuItem>();

        /// <summary>
        /// タスクトレイアイコン
        /// </summary>
        private NotifyIcon AppIcon { set; get; } = new NotifyIcon();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ClickCapture() {
            // タスクバーに表示しない
            ShowInTaskbar = false;
            // タイマー監視設定
            var eTimer = new Timer();
            // 50ミリ秒毎にチェック
            eTimer.Interval = 50;
            // イベントハンドラ設定
            eTimer.Tick += new EventHandler(OnMouseOperation);
            // タイマー実行
            eTimer.Start();
        }

        /// <summary>
        /// アイコン設定処理
        /// </summary>
        public void SetComponents() {
            // 自分自身のプロジェクトパス
            var eThisDir = Path.GetDirectoryName(Directory.GetCurrentDirectory());
            // アイコンパス
            var eIcoPath = Path.Combine(eThisDir, "ClickCapture", "ICO", "Copy.ico");

            // アイコン設定
            AppIcon.Icon = new Icon(eIcoPath);
            // タスクバーの通知領域に表示
            AppIcon.Visible = true;
            // アプリケーションの説明
            AppIcon.Text = "左クリックを押しながら右クリックでスクリーンショット";
            // アイコンに追加
            AppIcon.ContextMenuStrip = CreateMenu();
        }

        /// <summary>
        /// アクティブウインドウのスクリーンキャプチャを保存またはクリップボードにコピーする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMouseOperation(object sender, EventArgs e) {
            try {
                // 左クリックと右クリックが押し続けられている場合
                if (IsKeyPress(Keys.LButton) && IsKeyPress(Keys.RButton)) {
                    // WindowsAPI情報
                    var eWindowsAPI = new WindowsAPI();
                    // アクティブウインドウのサイズと位置を取得する
                    var eRect = eWindowsAPI.GetWindowRectangle();
                    // 幅
                    var eWidth = eRect.Right - eRect.Left;
                    // 高さ
                    var eHeight = eRect.Bottom - eRect.Top;
                    // 新規RECT作成
                    var eRectangle = new Rectangle(eRect.Left, eRect.Top, eWidth, eHeight);
                    // ビットマップに変換
                    var eBitmap = new Bitmap(eRectangle.Width, eRectangle.Height, PixelFormat.Format32bppArgb);
                    // スクリーンショット
                    using (var eFromImage = Graphics.FromImage(eBitmap)) {
                        eFromImage.CopyFromScreen(eRectangle.X, eRectangle.Y, 0, 0, eRectangle.Size, CopyPixelOperation.SourceCopy);
                    }
                    // チェックされたメニュー項目を取得
                    var eCheckedItem = MenuGroupItems
                        .Where(x => x.CheckState == CheckState.Checked)
                        .FirstOrDefault();
                    // クリップボードにコピーする場合
                    if (eCheckedItem.Name == "ClipBoard") {
                        Clipboard.SetImage(eBitmap);
                    }
                    // 画像を保存する場合
                    if (eCheckedItem.Name == "SavePic") {
                        // 自分自身のディレクトリパス
                        var eAppDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        // 保存ディレクトリパス
                        var eSaveDir = Path.Combine(eAppDir, "Pictures");
                        // フォルダ作成
                        Directory.CreateDirectory(eSaveDir);
                        // ファイル作成日
                        var eCreateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                        // 保存ファイルパス生成
                        var eSaveFilePath = Path.Combine(eSaveDir, $"Pict_{eCreateTime}.png");
                        // すでに存在する場合
                        if (File.Exists(eSaveFilePath)) {
                            // ファイルを削除
                            SafeDeleteFile(eSaveFilePath);
                        }
                        // 画像ファイルをPNGで保存する
                        eBitmap.Save(eSaveFilePath, ImageFormat.Png);
                    }
                    // 後片付け
                    eBitmap.Dispose();
                }
            }
            catch (Exception eException) {
                // デバッグログ出力
                Debug.WriteLine(eException.Message);
            }
        }

        /// <summary>
        /// モード切替メニュークリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMode_Click(object sender, EventArgs e) {
            foreach (var eMenuItem in MenuGroupItems) {
                // クリックされた項目とリストの項目が同じ場合
                if (object.ReferenceEquals(eMenuItem, sender)) {
                    // メニュー項目にチェックを入れる
                    eMenuItem.CheckState = CheckState.Checked;
                }
                else {
                    // 上記以外はチェックを外す
                    eMenuItem.CheckState = CheckState.Unchecked;
                }
            }
        }

        /// <summary>
        /// 保存フォルダを開くメニュークリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnDirOpen_Click(object sender, EventArgs e) {
            // 自分自身のディレクトリパス
            var eAppDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            // 保存ディレクトリパス
            var eSaveDir = Path.Combine(eAppDir, "Pictures");
            // フォルダ作成
            Directory.CreateDirectory(eSaveDir);
            // 保存フォルダを開く
            Process.Start(eSaveDir);
        }

        /// <summary>
        /// 終了メニュークリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnClose_Click(object sender, EventArgs e) {
            // アイコンを明示的に終了
            AppIcon.Dispose();
            // アプリケーションを終了
            Application.Exit();
        }

        /// <summary>
        /// 対象キーが押し続けられている場合trueを返却する
        /// </summary>
        /// <param name="pKeyCode"></param>
        /// <returns></returns>
        private bool IsKeyPress(Keys pKeyCode) {
            var eResult = false;
            // Keyを数値に変換
            var eKeyCode = (int)pKeyCode;
            // 最上位ビットが1か否かでキー押下の有無を取得
            var eState = WindowsAPI.GetKeyState((int)pKeyCode);
            // 判定結果
            eResult = eState < 0;
            return eResult;
        }

        /// <summary>
        /// メニュー項目作成
        /// </summary>
        /// <returns></returns>
        private ContextMenuStrip CreateMenu() {
            // 右クリックメニュー
            var eResult = new ContextMenuStrip();

            // メニュー項目 クリップボードにコピー
            var eMenuItemClipBoard = new ToolStripMenuItem();
            eMenuItemClipBoard.Name = "ClipBoard";
            eMenuItemClipBoard.Text = "&クリップボードにコピーする";
            eMenuItemClipBoard.Checked = true;
            eMenuItemClipBoard.Click += new EventHandler(OnMode_Click);
            MenuGroupItems.Add(eMenuItemClipBoard);

            // メニュー項目 画像を保存する
            var eMenuItemSavePic = new ToolStripMenuItem();
            eMenuItemSavePic.Name = "SavePic";
            eMenuItemSavePic.Text = "&画像を保存する";
            eMenuItemSavePic.Checked = false;
            eMenuItemSavePic.Click += new EventHandler(OnMode_Click);
            MenuGroupItems.Add(eMenuItemSavePic);

            // メニュー項目 保存フォルダを開く
            var eMenuItemDirOpen = new ToolStripMenuItem();
            eMenuItemDirOpen.Text = "&保存フォルダを開く";
            eMenuItemDirOpen.Click += new EventHandler(OnDirOpen_Click);

            // メニュー項目 終了
            var eMenuItemClose = new ToolStripMenuItem();
            eMenuItemClose.Text = "&終了";
            eMenuItemClose.Click += new EventHandler(OnClose_Click);

            // 右クリックメニューに追加
            eResult.Items.Add(eMenuItemClipBoard);
            eResult.Items.Add(eMenuItemSavePic);
            eResult.Items.Add(eMenuItemDirOpen);
            eResult.Items.Add(eMenuItemClose);

            return eResult;
        }

        /// <summary>
        /// ファイル削除処理
        /// </summary>
        /// <param name="pFilePath"></param>
        private void SafeDeleteFile(string pFilePath) {
            // 引数がnullの場合
            if (string.IsNullOrEmpty(pFilePath)) {
                return;
            }
            // ファイルパスからFileInfoを生成
            var eFileInfo = new FileInfo(pFilePath);
            // ファイルが存在する場合
            if (eFileInfo.Exists) {
                // 読み取り専用の場合
                if ((eFileInfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly) {
                    // 読み取り専用を解除
                    eFileInfo.Attributes = FileAttributes.Normal;
                }
            }
            // ファイル削除
            eFileInfo.Delete();
        }
    }
}
