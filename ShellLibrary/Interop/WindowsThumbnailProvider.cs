using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Intrinsics.X86;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using BExplorer.Shell;
using BExplorer.Shell._Plugin_Interfaces;
using BExplorer.Shell.Interop;
using ShellLibrary.Interop;
using Windows.Storage;
using Windows.Storage.FileProperties;
using Windows.Storage.Streams;
using static BExplorer.Shell.Interop.Gdi32;
using static ThumbnailGenerator.WindowsThumbnailProvider;
using SystemProperties = BExplorer.Shell.SystemProperties;

namespace ThumbnailGenerator {
  [Flags]
  public enum ThumbnailOptions {
    None = 0x00,
    BiggerSizeOk = 0x01,
    InMemoryOnly = 0x02,
    IconOnly = 0x04,
    ThumbnailOnly = 0x08,
    InCacheOnly = 0x10,
  }

  public class WindowsThumbnailProvider {
    private const String IShellItem2Guid = "7E9FB0D3-919F-4307-AB2E-9B1860310C93";
    private const String IShellImageFactoryGuid = "bcc18b79-ba16-442f-80c4-8a59c30c463b";

    /*
      [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
      internal static extern int SHCreateItemFromParsingName(
          [MarshalAs(UnmanagedType.LPWStr)] string path,
        // The following parameter is not used - binding context.
          IntPtr pbc,
          ref Guid riid,
          [MarshalAs(UnmanagedType.Interface)] out IShellItem shellItem);
    */
    [DllImport("shell32.dll", PreserveSig = false)]
    internal static extern Int32 SHCreateItemFromIDList(
        IntPtr pidl,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItem shellItem);

    [DllImport("shell32.dll", PreserveSig = false)]
    internal static extern Int32 SHCreateItemFromIDList(
      IntPtr pidl,
      ref Guid riid,
      [MarshalAs(UnmanagedType.Interface)] out Object shellItem);

    /*
      [DllImport("gdi32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DeleteObject(IntPtr hObject);
    */

    //[ComImport]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    //internal interface IShellItem {
    //  void BindToHandler(IntPtr pbc,
    //      [MarshalAs(UnmanagedType.LPStruct)]Guid bhid,
    //      [MarshalAs(UnmanagedType.LPStruct)]Guid riid,
    //      out IntPtr ppv);

    //  void GetParent(out IShellItem ppsi);
    //  void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
    //  void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    //  void Compare(IShellItem psi, uint hint, out int piOrder);
    //};

    internal enum SIGDN : UInt32 {
      NORMALDISPLAY = 0,
      PARENTRELATIVEPARSING = 0x80018001,
      PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
      DESKTOPABSOLUTEPARSING = 0x80028000,
      PARENTRELATIVEEDITING = 0x80031001,
      DESKTOPABSOLUTEEDITING = 0x8004c000,
      FILESYSPATH = 0x80058000,
      URL = 0x80068000
    }

    public enum HResult {
      Ok = 0x0000,
      False = 0x0001,
      InvalidArguments = unchecked((Int32)0x80070057),
      OutOfMemory = unchecked((Int32)0x8007000E),
      NoInterface = unchecked((Int32)0x80004002),
      Fail = unchecked((Int32)0x80004005),
      ElementNotFound = unchecked((Int32)0x80070490),
      TypeElementNotFound = unchecked((Int32)0x8002802B),
      NoObject = unchecked((Int32)0x800401E5),
      Win32ErrorCanceled = 1223,
      Canceled = unchecked((Int32)0x800704C7),
      ResourceInUse = unchecked((Int32)0x800700AA),
      AccessDenied = unchecked((Int32)0x80030005),
      WTS_E_FAILEDEXTRACTION = unchecked((int)0x8004b200),
      WTS_E_EXTRACTIONTIMEDOUT = unchecked((int)0x8004b201),
      WTS_E_SURROGATEUNAVAILABLE = unchecked((int)0x8004b202),
      WTS_E_FASTEXTRACTIONNOTSUPPORTED = unchecked((int)0x8004b203),
      WTS_E_DATAFILEUNAVAILABLE = unchecked((int)0x8004b204),
      STG_E_FILENOTFOUND = unchecked((int)0x80030002),
      WTS_E_EXTRACTIONPENDING = unchecked((int)0x8004B205),
    }

    [ComImport]
    [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItemImageFactory {
      [PreserveSig]
      HResult GetImage(
      [In, MarshalAs(UnmanagedType.Struct)] NativeSize size,
      [In] ThumbnailOptions flags,
      [Out] out IntPtr phbm);
    }

    public enum WTS_ALPHATYPE {
      /// <summary>The bitmap is an unknown format. The Shell tries nonetheless to detect whether the image has an alpha channel.</summary>
      WTSAT_UNKNOWN = 0x0,

      /// <summary>The bitmap is an RGB image without alpha. The alpha channel is invalid and the Shell ignores it.</summary>
      WTSAT_RGB = 0x1,

      /// <summary>The bitmap is an ARGB image with a valid alpha channel.</summary>
      WTSAT_ARGB = 0x2
    }

    [ComImport, Guid("e357fccd-a995-4576-b01f-234630154e96"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IThumbnailProvider {
      /// <summary>Gets a thumbnail image and alpha type.</summary>
      /// <param name="cx">
      /// The maximum thumbnail size, in pixels. The Shell draws the returned bitmap at this size or smaller. The returned bitmap
      /// should fit into a square of width and height cx, though it does not need to be a square image. The Shell scales the bitmap to
      /// render at lower sizes. For example, if the image has a 6:4 aspect ratio, then the returned bitmap should also have a 6:4
      /// aspect ratio.
      /// </param>
      /// <param name="phbmp">
      /// When this method returns, contains a pointer to the thumbnail image handle. The image must be a DIB section and 32 bits per
      /// pixel. The Shell scales down the bitmap if its width or height is larger than the size specified by cx. The Shell always
      /// respects the aspect ratio and never scales a bitmap larger than its original size.
      /// </param>
      /// <param name="pdwAlpha">
      /// When this method returns, contains a pointer to one of the following values from the WTS_ALPHATYPE enumeration.
      /// </param>
      /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
      [PreserveSig]
      HResult GetThumbnail(UInt32 cx, out IntPtr phbmp, out WTS_ALPHATYPE pdwAlpha);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeSize {
      private Int32 width;
      private Int32 height;

      public Int32 Width { set { width = value; } }
      public Int32 Height { set { height = value; } }
    };

    /*
      [StructLayout(LayoutKind.Sequential)]
      public struct RGBQUAD {
        public byte rgbBlue;
        public byte rgbGreen;
        public byte rgbRed;
        public byte rgbReserved;
      }
    */
    /*
      public static Bitmap GetThumbnail(string fileName, int width, int height, ThumbnailOptions options) {
        IntPtr hBitmap = GetHBitmap(Path.GetFullPath(fileName), width, height, options);

        try {
          // return a System.Drawing.Bitmap from the hBitmap
          return GetBitmapFromHBitmap(hBitmap);
        } finally {
          // delete HBitmap to avoid memory leaks
          DeleteObject(hBitmap);
        }
      }
    */
    public static IntPtr GetThumbnail(IListItemEx item, Int32 width, Int32 height, ThumbnailOptions options, Boolean isForThumbnailSource, out HResult hr) {
      IntPtr hBitmap = GetHBitmap(item, width, height, options, out hr, isForThumbnailSource);

      return hBitmap;
    }

    public static IntPtr GetThumbnail(IShellItem nativeShellItem, Int32 width, Int32 height, ThumbnailOptions options) {
      IntPtr hBitmap = GetHBitmap(nativeShellItem, width, height, options);

      return hBitmap;
    }

    /*
      public static Bitmap GetBitmapFromHBitmap(IntPtr nativeHBitmap) {
        Bitmap bmp = Bitmap.FromHbitmap(nativeHBitmap);

        if (Bitmap.GetPixelFormatSize(bmp.PixelFormat) < 32)
          return bmp;

        return CreateAlphaBitmap(bmp, PixelFormat.Format32bppArgb);
      }
    */

    /*
      public static Bitmap CreateAlphaBitmap(Bitmap srcBitmap, PixelFormat targetPixelFormat) {
        Bitmap result = new Bitmap(srcBitmap.Width, srcBitmap.Height, targetPixelFormat);

        Rectangle bmpBounds = new Rectangle(0, 0, srcBitmap.Width, srcBitmap.Height);

        BitmapData srcData = srcBitmap.LockBits(bmpBounds, ImageLockMode.ReadOnly, srcBitmap.PixelFormat);

        bool isAlplaBitmap = false;

        try {
          for (int y = 0; y <= srcData.Height - 1; y++) {
            for (int x = 0; x <= srcData.Width - 1; x++) {
              Color pixelColor = Color.FromArgb(
                  Marshal.ReadInt32(srcData.Scan0, (srcData.Stride * y) + (4 * x)));

              if (pixelColor.A > 0 & pixelColor.A < 255) {
                isAlplaBitmap = true;
              }

              result.SetPixel(x, y, pixelColor);
            }
          }
        } finally {
          srcBitmap.UnlockBits(srcData);
        }

        if (isAlplaBitmap) {
          return result;
        } else {
          return srcBitmap;
        }
      }
    */

    private static IntPtr GetHBitmap(IListItemEx item, Int32 width, Int32 height, ThumbnailOptions options, out HResult hr,  Boolean isForThumbnailSource = false) {
      Object nativeShellItem;
      var shellItem2Guid = new Guid(IShellImageFactoryGuid);
      var retCode = SHCreateItemFromIDList(item.PIDL, ref shellItem2Guid, out nativeShellItem);

      if (retCode != 0) {
        hr = HResult.Fail;
        return IntPtr.Zero;
      }

      var perceivedType = PerceivedType.Unspecified;
      var percTypeVal = item.GetPropertyValue(SystemProperties.PerceivedType, typeof(PerceivedType))?.Value;
      if (percTypeVal != null) {
        perceivedType = (PerceivedType)percTypeVal;
      }

      //var syncStatus = StorageSyncStatus.Unknown;
      var syncStatusVal = item.GetPropertyValue(SystemProperties.StorageSyncStatus, typeof(StorageSyncStatus))?.Value;
      //if (syncStatusVal != null) {
      //  syncStatus = (StorageSyncStatus)syncStatusVal;
      //}

      if (syncStatusVal != null && !item.IsFolder && (options & ThumbnailOptions.IconOnly) == 0) {
        item.IsNeedLoadFromStorage = true;
      }

      hr = HResult.NoObject;
      IntPtr hBitmap = IntPtr.Zero;

      //if (!((perceivedType == PerceivedType.Image || perceivedType == PerceivedType.Video) && !item.IsFolder && (options & ThumbnailOptions.IconOnly) == 0)) {
      var nativeSize = default(NativeSize);
      nativeSize.Width = width;
      nativeSize.Height = height;

      //IThumbnailCache thumbCache = null;

      //if (item.ComInterface != null) {

      //  var IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
      //  var CLSID_LocalThumbnailCache = new Guid("50EF4544-AC9F-4A8E-B21B-8A26180DB13F");

      //  IntPtr cachePointer;
      //  Ole32.CoCreateInstance(ref CLSID_LocalThumbnailCache, IntPtr.Zero, Ole32.CLSCTX.ALL, ref IID_IUnknown, out cachePointer);

      //  thumbCache = (IThumbnailCache)Marshal.GetObjectForIUnknown(cachePointer);
      //}

      //var tcacheFlags = WTS_FLAGS.WTS_SCALETOREQUESTEDSIZE;

      //if ((options & ThumbnailOptions.InCacheOnly) != 0 && !item.IsNeedLoadFromStorage) {
      //  tcacheFlags |= WTS_FLAGS.WTS_INCACHEONLY;
      //} else {
      //  tcacheFlags |= WTS_FLAGS.WTS_EXTRACT;
      //}
      //ISharedBitmap bmp1 = null;
      //var flags_1 = WTS_CACHEFLAGS.WTS_DEFAULT;
      //var thumbId = default(WTS_THUMBNAILID);
      //var res = BExplorer.Shell.Interop.HResult.S_OK;
      //if ((options & ThumbnailOptions.IconOnly) == 0) {
      //  try {
      //    res = thumbCache.GetThumbnail(item.ComInterface, (UInt32)width, tcacheFlags, out bmp1, flags_1, thumbId);
      //  } finally {
      //    if (bmp1 != null) {
      //      bmp1.Detach(out hBitmap);
      //      //Gdi32.ConvertPixelByPixel(hBitmap, out var width1, out var height1);
      //      Marshal.ReleaseComObject(bmp1);
      //      hr = HResult.Ok;
      //    }
      //  }
      //}

      //if (res == BExplorer.Shell.Interop.HResult.WTS_E_FAILEDEXTRACTION || (options & ThumbnailOptions.IconOnly) != 0) {
      //  hr = ((IShellItemImageFactory)nativeShellItem).GetImage(nativeSize, options, out hBitmap);
      //}

      if (!item.IsNeedLoadFromStorage || (options & ThumbnailOptions.IconOnly) != 0 || (item.IsFolder && isForThumbnailSource)) {
        hr = ((IShellItemImageFactory)nativeShellItem).GetImage(nativeSize, options, out hBitmap);
        //} 

        //if (hr == HResult.Ok && !item.IsNeedLoadFromStorage) {
        //  Gdi32.ConvertPixelByPixel(hBitmap, out var resWidth, out var resHeight);
        //  item.IsNeedRefreshing = ((resWidth > resHeight && resWidth != width) || (resWidth < resHeight && resHeight != width) || (resWidth == resHeight && resWidth != width)) && !item.IsOnlyLowQuality;
        //}
      }



      if (((((hr == HResult.NoObject) || isForThumbnailSource) && (perceivedType == PerceivedType.Image || perceivedType == PerceivedType.Video))) && (options & ThumbnailOptions.IconOnly) == 0) {
        //var iShell = (IShellItem)nativeShellItem;
        //iShell.BindToHandler(IntPtr.Zero, BHID.SThumbnailhandler, typeof(IThumbnailProvider).GUID, out IntPtr ress);
        //var iThumbnailProvider = (IThumbnailProvider)Marshal.GetTypedObjectForIUnknown(ress, typeof(IThumbnailProvider));

        //var resTP = iThumbnailProvider.GetThumbnail((UInt32)width, out hBitmap, out var alphaType);

        var flags = Windows.Storage.FileProperties.ThumbnailOptions.ResizeThumbnail;

        try {
          if (item.IsFileSystem && isForThumbnailSource) {
            IStorageItemProperties storageItem = item.IsFolder ? Windows.Storage.StorageFolder.GetFolderFromPathAsync(item.ParsingName).GetAwaiter().GetResult() : Windows.Storage.StorageFile.GetFileFromPathAsync(item.ParsingName).GetAwaiter().GetResult();
            var thumb = storageItem?.GetThumbnailAsync(ThumbnailMode.SingleItem, (UInt32)width, flags).GetAwaiter().GetResult();
            if (thumb != null) {
              Debug.WriteLine($"Returned thumb for path: {item.ParsingName}!!!!");
              var retry = 0;
              while ((isForThumbnailSource) && thumb != null && thumb.Type == ThumbnailType.Icon && retry < 10) {
                Thread.Sleep(100);
                thumb = item.StorageItem?.GetThumbnailAsync(ThumbnailMode.SingleItem, (UInt32)width, flags).GetAwaiter().GetResult();
                retry++;
              }
              IBuffer buf;
              var inputBuffer = new Windows.Storage.Streams.Buffer(512);
              var destFileStream = new MemoryStream();
              while ((buf = thumb.ReadAsync(inputBuffer, inputBuffer.Capacity, Windows.Storage.Streams.InputStreamOptions.None).GetAwaiter().GetResult()).Length > 0) {
                destFileStream.WriteAsync(buf.ToArray()).GetAwaiter().GetResult();
              }

              if (thumb != null && thumb.Type == ThumbnailType.Image) {
                using (var thumbNailStream = destFileStream) {
                  var bmp = (Bitmap)Image.FromStream(thumbNailStream);

                  if (thumb.ContentType == "image/bmp") {
                    var bmp_fixed = Gdi32.UnpremultiplyAlpha(bmp);
                    hBitmap = bmp_fixed.GetHbitmap();
                  } else {
                    hBitmap = bmp.GetHbitmap();
                  }

                  bmp.Dispose();
                  item.IsThumbnailLoaded = true;
                  item.IsNeedRefreshing = thumb.ReturnedSmallerCachedSize;
                  item.IsIconLoaded = true;
                  hr = HResult.Ok;
                }
              }
              thumb.Dispose();
            }

          }

          if (!isForThumbnailSource) {
            System.Windows.Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)(() => { hBitmap = item.hThumbnail(width, out var isLowQuality); }));
            //hBitmap = item.hThumbnail(width, out var isLowQuality);
            if (hBitmap != IntPtr.Zero) {
              //Gdi32.ConvertPixelByPixel(hBitmap, out var resWidth, out var resHeight);
              item.IsThumbnailLoaded = true;
              item.IsNeedRefreshing = false; //((resWidth > resHeight && resWidth != width) || (resWidth < resHeight && resHeight != width) || (resWidth == resHeight && resWidth != width)) && !item.IsOnlyLowQuality;
              item.IsIconLoaded = true;
              hr = HResult.Ok;
            }
          }
        } catch {
          //item.IsNeedRefreshing = true;
          Marshal.ReleaseComObject(nativeShellItem);
          return IntPtr.Zero;
          //hBitmap = IntPtr.Zero;
          //item.IsNeedRefreshing = true;
        }

      }

      Marshal.ReleaseComObject(nativeShellItem);

      if (hr == HResult.Ok) {
        return hBitmap;
      }

      item.IsThumbnailLoaded = true;
      item.IsNeedRefreshing = true;
      item.IsOnlyLowQuality = false;
      item.IsIconLoaded = true;
      //item.IsNeedLoadFromStorage = true;

      return IntPtr.Zero;
    }

    private static IntPtr GetHBitmap(IShellItem nativeShellItem, Int32 width, Int32 height, ThumbnailOptions options) {

      NativeSize nativeSize = default;
      nativeSize.Width = width;
      nativeSize.Height = height;

      IntPtr hBitmap;
      var hr = ((IShellItemImageFactory)nativeShellItem).GetImage(nativeSize, options, out hBitmap);

      Marshal.ReleaseComObject(nativeShellItem);

      if (hr == HResult.Ok) {
        return hBitmap;
      }

      return IntPtr.Zero;
    }
  }
}