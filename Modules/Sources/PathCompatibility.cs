using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public static class PathCompatibility
{
  private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
  private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
  private const int MAX_PATH = 32767;
  public static string GetRelativePath(string relativeTo, string path)
  {
    if (string.IsNullOrWhiteSpace(relativeTo))
    {
      throw new ArgumentException("relativeTo must not be null or empty.", "relativeTo");
    }
    if (string.IsNullOrWhiteSpace(path))
    {
      throw new ArgumentException("path must not be null or empty.", "path");
    }
    var fromPath = Path.GetFullPath(relativeTo);
    var toPath = Path.GetFullPath(path);
    var fromAttr = NativeMethods.GetPathAttributes(fromPath);
    var toAttr = NativeMethods.GetPathAttributes(toPath);
    var resultBuffer = new StringBuilder(MAX_PATH);
    if (!NativeMethods.PathRelativePathTo(resultBuffer, fromPath, fromAttr, toPath, toAttr))
    {
      var error = Marshal.GetLastWin32Error();
      throw new InvalidOperationException(string.Format("PathRelativePathTo failed from '{0}' to '{1}', Win32 error {2}.", fromPath, toPath, error));
    }
    return resultBuffer.ToString();
  }
  private static class NativeMethods
  {
    [DllImport("Shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern bool PathRelativePathTo(
      StringBuilder pszPath,
      string pszFrom,
      uint dwAttrFrom,
      string pszTo,
      uint dwAttrTo);
    internal static uint GetPathAttributes(string path)
    {
      if (string.IsNullOrEmpty(path))
      {
        return FILE_ATTRIBUTE_NORMAL;
      }
      if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
          path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
      {
        return FILE_ATTRIBUTE_DIRECTORY;
      }
      if (Directory.Exists(path))
      {
        return FILE_ATTRIBUTE_DIRECTORY;
      }
      if (File.Exists(path))
      {
        return FILE_ATTRIBUTE_NORMAL;
      }
      return string.IsNullOrEmpty(Path.GetExtension(path))
        ? FILE_ATTRIBUTE_DIRECTORY
        : FILE_ATTRIBUTE_NORMAL;
    }
  }
}
