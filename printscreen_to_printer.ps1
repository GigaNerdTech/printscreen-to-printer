# Takes a screenshot of selected process and sends to default printer on thick client
# Written by Joshua Woleben 1/3/19

$process_name = "notepad.exe"
[string] $adm

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAp {
      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool SetForegroundWindow(IntPtr hWnd);

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@

        $p = Get-Process -name $process_name
        $h = $p.MainWindowHandle
        [void] [WinAp]::SetForegroundWindow($h)
        [void] [WinAp]::ShowWindow($h, 3)

[Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#        Add-Type -AssemblyName System.Drawing
        $jpegCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
            Where-Object { $_.FormatDescription -eq "JPEG" }


        Start-Sleep -Milliseconds 250
        [Windows.Forms.Sendkeys]::SendWait("%{PrtSc}")
        Start-Sleep -Milliseconds 250
        $bitmap = [Windows.Forms.Clipboard]::GetImage()
        $ep = New-Object Drawing.Imaging.EncoderParameters
        $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)
        $screenCapturePathBase = "$pwd\ScreenCapture"
        $c = 0
        while (Test-Path "${screenCapturePathBase}${c}.jpg") {
            $c++
        }
        $bitmap.Save("${screenCapturePathBase}${c}.jpg",$jpegCodec, $ep)
    
    $bitmap = $null
    $doc = New-Object System.Drawing.Printing.PrintDocument
    $doc.DocumentName = "${screenCapturePathBase}${c}.jpg"

    $doc.add_PrintPage({
        $g = $_.Graphics
        $pageBounds = $_.MarginBounds
        $img = New-Object Drawing.Bitmap("${screenCapturePathBase}${c}.jpg")

        $adjustedImageSize = $img.Size
        $ratio = [double] 1;

        $fitWidth = [bool] ($img.Size.Width > $img.Size.Height)
        if (($img.Size.Width -le $_.MarginBounds.Width) -and
            ($img.Size.Height -le $_.MarginBounds.Height)) {
            $adjustedImageSize = New-Object System.Drawing.SizeF($img.Size.Width, $img.Size.Height)
        } else {
            if ($fitWidth) {
                $ratio = [double] ($_.MarginBounds.Width / $img.Size.Width)
            } else {
                $ratio = [double] ($_.MarginBounds.Height / $img.Size.Height)
            }
            $adjustedImageSize = New-Object System.Drawing.SizeF($_.MarginBounds.Width, [float]($img.Size.Height * $ratio))
        }
    $recDest = New-Object Drawing.RectangleF($pageBounds.Location, $adjustedImageSize)
    $recSrc = New-Object Drawing.RectangleF(0,0, $img.Width, $img.Height)

    $_.Graphics.DrawImage($img, $recDest, $recSrc, [Drawing.GraphicsUnit]"Pixel")

    $_.HasMorePages = $false
})
 $doc.Print()

