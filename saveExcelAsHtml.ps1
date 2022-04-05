###
# エクセル(.xlsx)をHTML(.html)形式に変換
#
# 実行コマンド
# PowerShell -ExecutionPolicy RemoteSigned ".\saveExcelAsHtml.ps1"
###

# ダイアログを出して、ファイルを選択する
# @return fileList ファイルリスト
function fileSelect() {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excelファイル形式|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xlam;*.xla;*.xlw;*.xlr;"

    # 起動時のディレクトリPath
    $dialog.InitialDirectory = Convert-Path .

    # ダイアログウインドウタイトル
    $dialog.Title = "ファイル選択"

    # 複数選択
    $dialog.Multiselect = $true

    # ダイアログ表示
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $dialog.FileNames
    } else {
        return $null
    }
}

###
# メイン処理
###


# エクセル操作初期化
$excel = New-Object -ComObject Excel.Application

# エクセル可視化
$excel.Visible = $False

# 変数にセット
$targetDir = [System.IO.Directory]::GetCurrentDirectory()
$savaDir = $targetDir+"\html"

#保存ディレクトリの作成
New-Item $savaDir -ItemType Directory

$itemList = fileSelect
foreach($item in $itemList) {

    $filename = [System.IO.Path]::GetFileName("$item")
    $saveFile = Join-Path $targetDir "html" | Join-Path -ChildPath $filename

    # エクセルを開く
    $book = $excel.Workbooks.Open($item)

    # ファイルをhtml形式で保存
    # 第二引数は、保存形式でコード値は下記URL参照
    # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
    $book.SaveAs([System.IO.Path]::ChangeExtension($saveFile,".html"),44)

    # エクセルを閉じる
    $excel.Quit()

    Write-Host $saveFile
}

# 後始末
$excel.Quit()
$excel = $null
[GC]::Collect()

echo "complete!!"
