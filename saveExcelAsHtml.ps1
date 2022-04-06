###
# �G�N�Z��(.xlsx)��HTML(.html)�`���ɕϊ�
#
# ���s�R�}���h
# PowerShell -ExecutionPolicy RemoteSigned ".\saveExcelAsHtml.ps1"
###

# �_�C�A���O���o���āA�t�@�C����I������
# @return fileList �t�@�C�����X�g
function fileSelect() {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel�t�@�C���`��|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xlam;*.xla;*.xlw;*.xlr;"

    # �N�����̃f�B���N�g��Path
    $dialog.InitialDirectory = Convert-Path .

    # �_�C�A���O�E�C���h�E�^�C�g��
    $dialog.Title = "�t�@�C���I��"

    # �����I��
    $dialog.Multiselect = $true

    # �_�C�A���O�\��
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $dialog.FileNames
    } else {
        return $null
    }
}

###
# ���C������
###


# �G�N�Z�����쏉����
$excel = New-Object -ComObject Excel.Application

# �G�N�Z������
$excel.Visible = $False

# �ϐ��ɃZ�b�g
$targetDir = [System.IO.Directory]::GetCurrentDirectory()
$savaDir = $targetDir+"\html"

#�ۑ��f�B���N�g���̍쐬
New-Item $savaDir -ItemType Directory

$itemList = fileSelect
foreach($item in $itemList) {

    $filename = [System.IO.Path]::GetFileName("$item")
    $saveFile = Join-Path $targetDir "html" | Join-Path -ChildPath $filename

    # �G�N�Z�����J��
    $book = $excel.Workbooks.Open($item)

    # �t�@�C����html�`���ŕۑ�
    # �������́A�ۑ��`���ŃR�[�h�l�͉��LURL�Q��
    # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
    $book.SaveAs([System.IO.Path]::ChangeExtension($saveFile,".html"),44)

    # �G�N�Z�������
    $excel.Quit()

    Write-Host $saveFile
}

# ��n��
$excel.Quit()
$excel = $null
[GC]::Collect()

echo "complete!!"
