OpenXML SDKをオフラインでNuGetする手順

1.nuget.exeのダウンロード
  https://www.nuget.org/downloads
  Windows x86 Command Line
  nuget.exe v6.13.1

2.nuget.exe install DocumentFormat.OpenXml -OutputDirectory C:\Temp\OpenXMLSDK

3.C:\Temp\OpenXMLSDK から以下2つのDLLを参照追加する。
  DocumentFormat.OpenXml.dll
  DocumentFormat.OpenXml.Framework.dll

