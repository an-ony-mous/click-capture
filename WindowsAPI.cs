using System;
using System.Runtime.InteropServices;

namespace ClickCapture {
    /// <summary>
    /// WindowsAPI情報
    /// https://www.manongdao.com/q-1253949.html
    /// </summary>
    public class WindowsAPI {
        /// <summary>
        /// WindowsAPI 指定された仮想キーの状態を取得
        /// </summary>
        /// <param name="pVirtKey"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern short GetKeyState(int pVirtKey);

        /// <summary>
        /// 構造体定義
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// WindowsAPI ウインドウの位置と大きさを取得する
        /// </summary>
        /// <param name="pHandle"></param>
        /// <param name="pRect"></param>
        /// <returns></returns>
        [DllImport("User32.Dll")]
        public static extern int GetWindowRect(IntPtr pHandle, ref RECT pRect);

        /// <summary>
        /// WindowsAPI 現在ユーザーが作業しているウィンドウのハンドルを返す
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        /// <summary>
        /// WindowsAPI 影なしウインドウ取得用
        /// </summary>
        /// <param name="pHandle"></param>
        /// <param name="pAttribute"></param>
        /// <param name="pRect"></param>
        /// <param name="pSize"></param>
        /// <returns></returns>
        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr pHandle, int pAttribute, out RECT pRect, int pSize);

        /// <summary>
        /// 拡張フレーム境界
        /// </summary>
        private readonly int EXTENDED_FRAME_BOUNDS = 9;

        /// <summary>
        /// アクティブウインドウのサイズと位置を取得する
        /// </summary>
        /// <param name="hWnd"></param>
        /// <returns></returns>
        public RECT GetWindowRectangle() {
            // アクティブウインドウを取得
            var eActive = GetForegroundWindow();
            // 出力用サイズ
            var eSize = Marshal.SizeOf(typeof(RECT));
            // アクティブウインドウからRECT構造体を取得
            DwmGetWindowAttribute(eActive, EXTENDED_FRAME_BOUNDS, out RECT eResult, eSize);
            return eResult;
        }
    }
}
