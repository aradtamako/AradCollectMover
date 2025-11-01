#pragma warning disable CA1416
using System.Diagnostics;
using WindowsInput;

class Program
{
    private static int ScreenWidth
    {
        get
        {
            if (field == 0)
            {
                field = Windows.Win32.PInvoke.GetSystemMetrics(Windows.Win32.UI.WindowsAndMessaging.SYSTEM_METRICS_INDEX.SM_CXSCREEN);
            }

            return field;
        }
    }

    private static int ScreenHeight
    {
        get
        {
            if (field == 0)
            {
                field = Windows.Win32.PInvoke.GetSystemMetrics(Windows.Win32.UI.WindowsAndMessaging.SYSTEM_METRICS_INDEX.SM_CYSCREEN);
            }

            return field;
        }
    }

    static double GetAbsoluteX(int x) => x * 65535 / (double)ScreenWidth;
    static double GetAbsoluteY(int y) => y * 65535 / (double)ScreenHeight;

    private static InputSimulator Simulator = new();
    private static Process AradProcess { get; set; } = default!;

    private const int VK_A = 'A';
    private const int VK_LSHIFT = 0xA0;
    private const int VK_LCONTROL = 0xA2;
    private const int VK_LBUTTON = 0x01; // マウスの左ボタン
    private const int VK_RBUTTON = 0x02; // マウスの右ボタン
    private const int VK_XBUTTON1 = 0x05; // マウスの戻るボタン
    private const int VK_SPACE = 0x20;

    private const uint BaseWidth = 1280;
    private const uint BaseHeight = 720;

    private const uint WarehouseBaseX = 290;
    private const uint InventoryBaseX = 740;
    private const uint CellBaseWidth = 35;
    private const uint CellBaseHeight = 35;
    private const uint CollectItemCount = 10;
    private const uint InventoryHorizontalCount = 8;

    private static readonly TimeSpan HoldTime = TimeSpan.FromMilliseconds(50);
    private static readonly TimeSpan Delay = TimeSpan.FromMilliseconds(100);

    private enum ClickType
    {
        Left = 0,
        Right = 1
    }

    /// <summary>
    /// マウスクリックする
    /// </summary>
    private static void MouseClick(TimeSpan holdTime, ClickType clickType)
    {
        if (clickType == ClickType.Left) { Simulator.Mouse.LeftButtonDown(); } else { Simulator.Mouse.RightButtonDown(); }
        Thread.Sleep(holdTime);
        if (clickType == ClickType.Left) { Simulator.Mouse.LeftButtonUp(); } else { Simulator.Mouse.RightButtonUp(); }
        Thread.Sleep(Delay);
    }

    /// <summary>
    /// マウスカーソルを指定座標に移動する
    /// </summary>
    private static void MouseMove(int x, int y) => Simulator.Mouse.MoveMouseTo(GetAbsoluteX(x), GetAbsoluteY(y));

    /// <summary>
    /// ウィンドウ内のマウス座標を取得する
    /// </summary>
    private static System.Drawing.Point? GetClientAreaMousePosition()
    {
        Windows.Win32.PInvoke.GetCursorPos(out var point);
        var hWnd = AradProcess.MainWindowHandle;
        Windows.Win32.PInvoke.ScreenToClient(new Windows.Win32.Foundation.HWND(hWnd), ref point);

        return point;
    }

    private static void InitializeAradProcess()
    {
        var process = Process.GetProcessesByName("ARAD").FirstOrDefault();
        if (process == null)
        {
            throw new Exception("Arad not found");
        }

        AradProcess = process;
    }

    /// <summary>
    /// コレクトボックスからコレクトを取り出す
    /// </summary>
    private static void TakeoutCollect()
    {
        Windows.Win32.PInvoke.GetCursorPos(out var point);
        Windows.Win32.PInvoke.GetClientRect(new Windows.Win32.Foundation.HWND(AradProcess.MainWindowHandle), out var rect);
        var widthRate = (double)rect.Width / BaseWidth;
        var heightRate = (double)rect.Height / BaseHeight;

        // 1段目
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((100 * widthRate)), point.Y + (int)Math.Floor((80 * 0 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((200 * widthRate)), point.Y + (int)Math.Floor((80 * 0 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        // 2段目
        MouseMove(point.X + (int)Math.Floor((50 * widthRate)), point.Y + (int)Math.Floor((80 * 1 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((150 * widthRate)), point.Y + (int)Math.Floor((80 * 1 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        // 3段目
        MouseMove(point.X + (int)Math.Floor((50 * widthRate)), point.Y + (int)Math.Floor((80 * 2 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((150 * widthRate)), point.Y + (int)Math.Floor((80 * 2 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        // 4段目
        MouseMove(point.X, point.Y + (int)Math.Floor((80 * 3 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((100 * widthRate)), point.Y + (int)Math.Floor((80 * 3 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);

        MouseMove(point.X + (int)Math.Floor((200 * widthRate)), point.Y + (int)Math.Floor((80 * 3 * heightRate)));
        MouseClick(HoldTime, ClickType.Right);
    }

    /// <summary>
    /// コレクトを倉庫からインベントリ、または、コレクトをインベントリから倉庫に移動する処理
    /// </summary>
    static void TransferCollect(ClickType clickType)
    {
        var point = GetClientAreaMousePosition();
        if (point != null)
        {
            Windows.Win32.PInvoke.GetClientRect(new Windows.Win32.Foundation.HWND(AradProcess.MainWindowHandle), out var rect);
            var widthRate = (double)rect.Width / BaseWidth;
            var heightRate = (double)rect.Height / BaseHeight;

            var baseX = 0u;
            if (point.Value.X > WarehouseBaseX && point.Value.X < InventoryBaseX)
            {
                // コレクトを倉庫からインベントリに移動する処理
                baseX = WarehouseBaseX;
            }
            else if (point.Value.X > WarehouseBaseX)
            {
                // コレクトをインベントリから倉庫に移動する処理
                baseX = InventoryBaseX;
            }
            else
            {
                return;
            }

            var index = Math.Floor((point.Value.X - (int)Math.Floor(baseX * widthRate)) / (double)Math.Floor(CellBaseWidth * widthRate));

            // 1行目
            var line1Count = (InventoryHorizontalCount - index);
            for (var i = 0; i < line1Count; i++)
            {
                MouseMove(
                    point.Value.X + (int)Math.Floor((CellBaseWidth * i * widthRate)),
                    point.Value.Y
                    );
                MouseClick(HoldTime, clickType);
            }

            // 2行目
            var line2Count = (CollectItemCount - line1Count);
            for (var i = 0; i < line2Count; i++)
            {
                MouseMove(
                    point.Value.X - (int)Math.Floor((CellBaseWidth * index * widthRate)) + (int)Math.Floor((CellBaseWidth * i * widthRate)),
                    point.Value.Y + (int)Math.Floor((CellBaseHeight * heightRate))
                    );
                MouseClick(HoldTime, clickType);
            }
        }
    }

    static void Main(string[] args)
    {
        InitializeAradProcess();

        var wasPressedZ = false;
        var wasPressedX = false;

        while (true)
        {
            var isZPressed = (Windows.Win32.PInvoke.GetAsyncKeyState('Z') & 0x8000) != 0;
            var isXPressed = (Windows.Win32.PInvoke.GetAsyncKeyState('X') & 0x8000) != 0;
            var isXButton1Pressed = (Windows.Win32.PInvoke.GetAsyncKeyState(VK_XBUTTON1) & 0x8000) != 0;
            var isMouseLeftPressed = (Windows.Win32.PInvoke.GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
            var isMouseRightPressed = (Windows.Win32.PInvoke.GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;

            var isPressedZ = isZPressed && isXButton1Pressed;
            var isPressedX = isXPressed && isXButton1Pressed;

            if (isPressedZ && !wasPressedZ)
            {
                TakeoutCollect();
            }

            if (isPressedX && !wasPressedX)
            {
                if (isMouseLeftPressed)
                {
                    TransferCollect(ClickType.Left);
                }
                else if (isMouseRightPressed)
                {
                    TransferCollect(ClickType.Right);
                }
            }

            wasPressedZ = isPressedZ;
            wasPressedX = isPressedX;

            Thread.Sleep(10);
        }
    }
}