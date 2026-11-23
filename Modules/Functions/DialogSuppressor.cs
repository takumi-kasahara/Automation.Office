using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class NUIDialogSuppressor
{
  private const int WM_SYSCOMMAND = 0x0112;
  private const int SC_CLOSE = 0xF060;
  public static void Run(string stopFilePath, string targetExe)
  {
    while (!File.Exists(stopFilePath))
    {
      var handled = false;

      NativeMethods.EnumWindows(new NativeMethods.EnumWindowsProc((hWnd, lParam) =>
      {
        if (handled) { return false; }

        if (targetExe != null)
        {
          uint processId;
          NativeMethods.GetWindowThreadProcessId(hWnd, out processId);
          if (!IsProcessMatch(processId, targetExe)) { return true; }
        }
        var classBuilder = new StringBuilder(64);
        NativeMethods.GetClassName(hWnd, classBuilder, classBuilder.Capacity);
        if (!IsTargetDialogClass(classBuilder.ToString())) { return true; }
        TrySuppressDialog(hWnd);
        handled = true;
        return false;
      }), IntPtr.Zero);

      Thread.Sleep(100);
    }
  }
  private static bool IsTargetDialogClass(string className)
  {
    if (string.IsNullOrEmpty(className)) { return false; }
    if (string.Equals(className, "NUIDialog", StringComparison.OrdinalIgnoreCase)) { return true; }
    if (className.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase)) { return true; }
    return false;
  }
  private static bool IsProcessMatch(uint processId, string targetExe)
  {
    if (processId == 0) { return false; }
    try
    {
      using (var process = Process.GetProcessById((int)processId))
      {
        var processName = process.ProcessName;
        var targetName = Path.GetFileNameWithoutExtension(targetExe);
        return string.Equals(processName, targetName, StringComparison.OrdinalIgnoreCase);
      }
    }
    catch { return false; }
  }
  private static void TrySuppressDialog(IntPtr hWnd)
  {
    NativeMethods.PostMessage(hWnd, WM_SYSCOMMAND, new IntPtr(SC_CLOSE), IntPtr.Zero);
    if (WaitForWindowClose(hWnd, 100)) { return; }
  }
  private static bool WaitForWindowClose(IntPtr hWnd, int timeoutMs)
  {
    var stopwatch = Stopwatch.StartNew();
    while (NativeMethods.IsWindow(hWnd) && stopwatch.ElapsedMilliseconds < timeoutMs) { Thread.Sleep(10); }
    return !NativeMethods.IsWindow(hWnd);
  }
  private static class NativeMethods
  {
    internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);
    [DllImport("user32.dll")]
    internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    internal static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    internal static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  }
}
public static class VBProjectDialogSuppressor
{
  private const int BM_CLICK = 0x00F5;
  public static void Run(string stopFilePath, string targetExe)
  {
    while (!File.Exists(stopFilePath))
    {
      var clicked = false;

      NativeMethods.EnumWindows(new NativeMethods.EnumWindowsProc((hWnd, lParam) =>
      {
        if (clicked) { return false; }

        if (targetExe != null)
        {
          uint processId = 0;
          NativeMethods.GetWindowThreadProcessId(hWnd, out processId);
          if (!IsProcessMatch(processId, targetExe)) { return true; }
        }
        var classBuilder = new StringBuilder(64);
        NativeMethods.GetClassName(hWnd, classBuilder, classBuilder.Capacity);
        if (!string.Equals(classBuilder.ToString(), "#32770", StringComparison.OrdinalIgnoreCase)) { return true; }

        NativeMethods.EnumChildWindows(hWnd, new NativeMethods.EnumWindowsProc((child, childParam) =>
        {
          if (clicked) { return false; }
          var childClassBuilder = new StringBuilder(64);
          NativeMethods.GetClassName(child, childClassBuilder, childClassBuilder.Capacity);
          if (!string.Equals(childClassBuilder.ToString(), "Button", StringComparison.OrdinalIgnoreCase)) { return true; }
          if (NativeMethods.GetDlgCtrlID(child) != 1) { return true; }
          NativeMethods.SendMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
          clicked = true;
          return false;
        }), IntPtr.Zero);

        if (clicked) { WaitForWindowClose(hWnd, 100); }
        return !clicked;
      }), IntPtr.Zero);

      Thread.Sleep(100);
    }
  }
  private static bool IsProcessMatch(uint processId, string targetExe)
  {
    if (processId == 0) { return false; }
    try
    {
      using (var process = Process.GetProcessById((int)processId))
      {
        var processName = process.ProcessName;
        var targetName = Path.GetFileNameWithoutExtension(targetExe);
        return string.Equals(processName, targetName, StringComparison.OrdinalIgnoreCase);
      }
    }
    catch { return false; }
  }
  private static bool WaitForWindowClose(IntPtr hWnd, int timeoutMs)
  {
    var stopwatch = Stopwatch.StartNew();
    while (NativeMethods.IsWindow(hWnd) && stopwatch.ElapsedMilliseconds < timeoutMs) { Thread.Sleep(10); }
    return !NativeMethods.IsWindow(hWnd);
  }
  private static class NativeMethods
  {
    internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);
    [DllImport("user32.dll")]
    internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    internal static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    internal static extern int GetDlgCtrlID(IntPtr hWnd);
    [DllImport("user32.dll")]
    internal static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    internal static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  }
}
