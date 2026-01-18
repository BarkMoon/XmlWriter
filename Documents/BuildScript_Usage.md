# BuildScript.ps1 使用ガイド

`BuildScript.ps1` は、`XmlWriter` プロジェクトをビルドするためのPowerShellスクリプトです。
通常の `dotnet build` コマンドでは発生する可能性のある、MSBuildのパス解決や.NET Frameworkバージョンの問題を回避するために用意されています。

## 概要

このスクリプトは以下の手順でビルドを行います。

1.  **MSBuildの検索**
    *   `vswhere.exe` を使用して、インストールされているVisual StudioまたはBuild Toolsに含まれる最新の `MSBuild.exe` を探します。
    *   見つからない場合、フォールバックとして .NET Framework (v4.0.30319) のMSBuildを探します。
2.  **ビルド実行**
    *   特定した `MSBuild.exe` を使用して、`XmlWriter.csproj` を `Release` 構成でビルドします。

## 使用方法

PowerShellターミナルで、スクリプトのあるディレクトリ (`XmlWriter` フォルダ) に移動し、以下を実行します。

```powershell
./BuildScript.ps1
```

※ 実行ポリシーのエラーが出る場合は、以下のように `Bypass` オプションを付けて実行してください。

```powershell
powershell -ExecutionPolicy Bypass -File BuildScript.ps1
```

## 前提条件

*   Windows OS
*   Visual Studio または Visual Studio Build Tools がインストールされていること (推奨)
*   あるいは、.NET Framework 4.7.2 ターゲットパックがインストールされていること

## 出力

成功すると以下のパスに実行ファイルが生成されます。

`bin/Release/XmlWriter.exe`
