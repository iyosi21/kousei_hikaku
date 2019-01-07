Attribute VB_Name = "Module1"
Option Explicit

Sub bookhikaku()
Dim wbk As Workbook: Set wbk = ThisWorkbook
Dim wbk1 As Worksheet: Set wbk1 = wbk.Worksheets(1) 'wbk1�ɓ]�L��̃V�[�g������
Dim wbk2 As Workbook
Dim wbk3 As Workbook
Dim wsn As String: wsn = "�T�[�o���X�g"

Dim WBK_Col As Long: WBK_Col = 1 '��r�u�b�N�ɓ]�@����ŏ��̃A�h���X
Dim WBK_Row As Long: WBK_Row = 2 '��r�u�b�N�ɓ]�@����ŏ��̃A�h���X

Dim D1, D2, D3, D4 As Long '���[�v�p�̕ϐ�
Dim OutCAL As Integer: OutCAL = 16 '��r�\�̉E���̃X�^�[�g�J����

Dim RefFr As Range: Set RefFr = wbk1.Range("C1")
Dim RefTo As Range: Set RefTo = wbk1.Range("S1")

'�z��n�̕ϐ�
Dim f_test1 As Variant
Dim l_test1 As Variant
Dim f_test2 As Variant
Dim l_test2 As Variant

Dim Judg_Aray1 As Variant
Dim Lost_Aray1 As Variant
Dim Add_Aray As Variant


D3 = 0
wbk1.Range(Cells(3, 1), Cells(Rows.Count, Columns.Count)).ClearContents
wbk1.Range(Cells(3, 1), Cells(Rows.Count, Columns.Count)).Interior.ColorIndex = 0



'�Q�ƌ��t�@�C�����J��
Dim OpenFileName As String
MsgBox "�Q�ƌ��̃t�@�C����I�����Ă�������"
  OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
  If OpenFileName <> "False" Then
    Workbooks.Open OpenFileName
  Else
    MsgBox "�L�����Z������܂���"
    Exit Sub
  End If

  Dim FileName: FileName = Dir(OpenFileName)
  Set wbk2 = Workbooks(FileName)
  
  With wbk2.Worksheets(wsn)
  If .AutoFilterMode Then             ''�I�[�g�t�B���^���ݒ肳��Ă�����
      If .AutoFilter.FilterMode Then  ''�i�荞�݂�����Ă�����
             .Range("A1").AutoFilter     ''�I�[�g�t�B���^����������
      End If
  End If
  End With
  
RefFr.Value = wbk2.Name

  
'�Q�Ɛ�t�@�C�����J��
MsgBox "�Q�Ɛ�̃t�@�C����I�����Ă�������"
  OpenFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
  If OpenFileName <> "False" Then
    Workbooks.Open OpenFileName
  Else
    MsgBox "�L�����Z������܂���"
    Exit Sub
  End If
  FileName = Dir(OpenFileName)
  Set wbk3 = Workbooks(FileName)
  
  With wbk3.Worksheets(wsn)
  If .AutoFilterMode Then             ''�I�[�g�t�B���^���ݒ肳��Ă�����
      If .AutoFilter.FilterMode Then  ''�i�荞�݂�����Ă�����
             .Range("A1").AutoFilter     ''�I�[�g�t�B���^����������
      End If
  End If
  End With
    
RefTo.Value = wbk3.Name

If wbk2.Name = wbk3.Name Then
    MsgBox "�����t�@�C����I�����Ă��܂�", vbExclamation
    Exit Sub
End If
  
MsgBox "�������J�n���܂�"


'excelfeez�X�^�[�g
Call FreezeExcel
  
'��r���V�[�g�̍ŏI�s���
Dim Copy_WBR As Integer: Copy_WBR = wbk2.Worksheets(wsn).Cells(Rows.Count, 12).End(xlUp).Row

'��r��V�[�g�̍ŏI�s���
Dim Reve_WBR As Integer: Reve_WBR = wbk3.Worksheets(wsn).Cells(Rows.Count, 12).End(xlUp).Row

'IP�A�h���X�̔z�񂩂�A�����Ă�����̂����𒊏o����B
f_test1 = Create_First_Aray(wbk2.Worksheets(wsn), Copy_WBR)
l_test1 = isLive(f_test1)

f_test2 = Create_First_Aray(wbk3.Worksheets(wsn), Reve_WBR)
l_test2 = isLive(f_test2)

Judg_Aray1 = JudgAndCreate_Aray(l_test1, l_test2)
'judg_aray�ŁA��r���Ĉ�������̂����̂������z�񂪂ł����B

Lost_Aray1 = FoundLostIP(l_test1, l_test2)
Add_Aray = FoundLostIP(l_test2, l_test1)

'Judg_Aray1��]�L���鏈���B
'Judg_Aray1�ōŏ��̗v�f��empty���Ə��������
If Not IsEmpty(Judg_Aray1(1, 1)) Then
    For D1 = 1 To UBound(Judg_Aray1, 1)
        For D2 = 1 To UBound(Judg_Aray1, 2)
            If D2 < 15 Then
                wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Value = Judg_Aray1(D1, D2)
                If wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Value <> Judg_Aray1(D1, D2 + 14) Then
                    wbk1.Cells(WBK_Row + D1, WBK_Col + D2).Interior.ColorIndex = 6
                
                End If
                    
            Else
                wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Value = Judg_Aray1(D1, D2)
                If wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Value <> Judg_Aray1(D1, D2 - 14) Then
                    wbk1.Cells(WBK_Row + D1, WBK_Col + 1 + D2).Interior.ColorIndex = 6
                    
                End If
                
            End If
    
        Next D2
    
    Next D1
Else
    D1 = 0
End If

If Not IsEmpty(Lost_Aray1(1, 1)) Then
    For D3 = 1 To UBound(Lost_Aray1, 1)
        wbk1.Cells(WBK_Row + D1 + D3, WBK_Col).Value = "�폜"
        wbk1.Cells(WBK_Row + D1 + D3, WBK_Col).Interior.ColorIndex = 3
        
        For D4 = 1 To UBound(Lost_Aray1, 2)
            wbk1.Cells(WBK_Row + D1 + D3, WBK_Col + D4).Value = Lost_Aray1(D3, D4)
        
        Next D4
    
    Next D3
End If

If Not IsEmpty(Add_Aray(1, 1)) Then
    For D3 = 1 To UBound(Add_Aray, 1)
        wbk1.Cells(WBK_Row + D1 + D3, OutCAL).Value = "�ǉ�"
        wbk1.Cells(WBK_Row + D1 + D3, OutCAL).Interior.ColorIndex = 6
        
        For D4 = 1 To UBound(Add_Aray, 2)
            wbk1.Cells(WBK_Row + D1 + D3, OutCAL + D4).Value = Add_Aray(D3, D4)
        
        Next D4
    
    Next D3
End If


'�J���Ă���u�b�N����鏈��
wbk2.Close
wbk3.Close

MsgBox "�����I��"

Call MeltExcel

End Sub



'�n���ꂽ��̔z��̒��ŁA������Ȃ�����IP�̔z���n���B
Function FoundLostIP(Aray1 As Variant, Aray2 As Variant) As Variant
Dim FoundAray() As Variant
Dim i As Integer
Dim k As Integer
Dim m As Integer
Dim F_Found As Boolean
Dim Lost_Cnt As Integer
Dim Lost_Ins As Integer
Lost_Cnt = 0

F_Found = False

'��v���Ȃ�IP�̗v�f���𒲂ׂ�B
For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                Exit For
                
            End If
            
            If k = UBound(Aray2, 1) Then
                Lost_Cnt = Lost_Cnt + 1
            
            End If
        
        Next k
            
    End If
 
Next i

If Lost_Cnt = 0 Then
    ReDim FoundAray(1 To 1, 1 To 14)
    FoundLostIP = FoundAray
    Exit Function
    
End If

'��v���Ȃ�IP�̗v�f����񎟌��z��Ő錾����B
ReDim FoundAray(1 To Lost_Cnt, 1 To 14)
F_Found = False

Lost_Ins = 1
For i = 1 To UBound(Aray1, 1)

    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                Exit For
                
            End If
            
            If k = UBound(Aray2, 1) Then
                For m = 1 To 14
                    FoundAray(Lost_Ins, m) = Aray1(i, m + 10)
        
                Next m
                Lost_Ins = Lost_Ins + 1
                
            End If
        
        Next k
        
    End If
 
Next i

FoundLostIP = FoundAray

End Function


'�n���ꂽ��̔z��̒��ŁA��v���Ȃ����̂𒊏o�����z������֐�
Function JudgAndCreate_Aray(Aray1 As Variant, Aray2 As Variant) As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer

Dim Judg As Boolean
Dim Judg_Count As Integer
Dim C_Aray1() As Variant

Judg = False
Judg_Count = 0
Dim Aray_Count As Integer
Aray_Count = 1

'�v�f���ω����Ă���IP�A�h���X�̗v�f���𒲂ׂ�B
For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
       For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                For j = 1 To (UBound(Aray1, 2) - 11)
                    If Aray1(i, 11 + j) <> Aray2(k, 11 + j) Then
                        Judg_Count = Judg_Count + 1
                        Exit For
                
                    End If
                
                Next j
            
                Exit For
            End If 'Aray1 = Aray2
       
       Next k
    End If

Next i

If Judg_Count = 0 Then
    ReDim C_aray(1 To 1, 1 To 28)
    JudgAndCreate_Aray = C_aray
    Exit Function
End If

'�v�f�ɕω���������IP�̐���񎟌��z��Ɋi�[���邽�߂̐錾���s���B
ReDim C_aray(1 To Judg_Count, 1 To 28)

'C_Aray�ɂ͔�r���̔z��Ɣ�r��̔z�񂪗��������Ă���B
'�񎟔z���15�ԂŐ܂�Ԃ��i14�Ԃ܂ł��Q�ƌ��̔z��j

For i = 1 To UBound(Aray1, 1)
    If Aray1(i, 11) Like "*.*" And Not IsEmpty(Aray1(i, 11)) Then
        For k = 1 To UBound(Aray2, 1)
            If Aray1(i, 11) = Aray2(k, 11) Then
                For j = 1 To (UBound(Aray1, 2) - 11)
                    If Aray1(i, 11 + j) <> Aray2(k, 11 + j) Then
                        For m = 1 To (UBound(C_aray, 2) / 2)
'�����ł΂�����
                            C_aray(Aray_Count, m) = Aray1(i, 10 + m)
                            C_aray(Aray_Count, m + 14) = Aray2(k, 10 + m)
                    
                        Next m
                        
                        Aray_Count = Aray_Count + 1
                        Exit For
            
                    End If
                
                Next j
            
                Exit For
        
            End If
    
        Next k
    
    End If

Next i

JudgAndCreate_Aray = C_aray

End Function


'�ŏI�s�̔ԍ�������ƁAB�񂩂�Y��܂ł̒l��z��ɂ��ĕԂ�
Function Create_First_Aray(WS As Worksheet, LastRow As Integer) As Variant
'B�񂩂�Y��܂Œl���Ƃ�
'B���2�AY���25
'�s��4����X�^�[�g
Dim Parent_Range As Range
Set Parent_Range = WS.Range(WS.Cells(4, 2), WS.Cells(LastRow, 25))
Create_First_Aray = Parent_Range

End Function

'�����Ă����z�񂩂琶���Ă�����̂����𒊏o����B
'�z��̂P��ڂɁu�\��v�̕����������Ă���z��������ĕԂ��B
Function isLive(Aray As Variant) As Variant
Dim Live_Cnt As Integer
Dim Live_Aray As Variant
'�񎟌��z��̗v�f������鎞�ɁAUbound������Ȃ������̂ŁA���̕ϐ���Ubound�����B
Dim Aray_Clm As Integer: Aray_Clm = UBound(Aray, 2)
Dim i, j As Integer
Dim NotIsLive_Cnt As Integer

Live_Cnt = 0
For i = 1 To UBound(Aray, 1)
    If Not Aray(i, 1) Like "*�\��*" Then
        Live_Cnt = Live_Cnt + 1
    End If
Next i

ReDim Live_Aray(1 To Live_Cnt, 1 To Aray_Clm)

NotIsLive_Cnt = 1
For i = 1 To UBound(Aray, 1)
    If Not Aray(i, 1) Like "*�\��*" Then
        For j = 1 To Aray_Clm
            Live_Aray(NotIsLive_Cnt, j) = Aray(i, j)
        Next j
        NotIsLive_Cnt = NotIsLive_Cnt + 1
    Else

    End If
Next i

isLive = Live_Aray

End Function


Private Function FreezeExcel()
    With Application
        '.Visible = False '�S�̂̕\�����~
        .DisplayAlerts = False '�A���[�g�̕\�����~
        .StatusBar = False '�X�e�[�^�X�o�[�̕\���X�V���~
        .ScreenUpdating = False '�X�N���[���̕`����~
        .EnableEvents = False '�C�x���g���ꎞ��~
        .Calculation = xlManual '�v�Z���蓮���[�h�ɂ���
    End With
End Function
   

Private Function MeltExcel()
    With Application
        .Calculation = xlAutomatic '�v�Z���������[�h�ɖ߂�
        .EnableEvents = True '�C�x���g���ĊJ
        .ScreenUpdating = True '�X�N���[���̕`����ĊJ
        .StatusBar = True '�X�e�[�^�X�o�[�̕\�����ĊJ
        .DisplayAlerts = True '�A���[�g�̕\�����ĊJ
        '.Visible = True '�S�̂̕\������
    End With
End Function
