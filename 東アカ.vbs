Sub abc()
    ' �e���`���̖�萔�~���_��萔�Ƃ��Ē�`
    Const HISSYU_MAX As Integer = (25 - 1 + 1) * 1
    Const IPPAN_MAX As Integer = ((89 - 26 + 1) + (122 - 121 + 1)) * 1
    Const ZYOUKYOU_MAX As Integer = (120 - 91 + 1) * 2
    
    ' For���p�̃J�E���^�ϐ���錾
    Dim i, j As Integer
    ' ���_�L�^�p�̕ϐ���錾
    Dim hissyu, ippan, zyoukyou As Integer
    
    ' �ߑO�̐���
    For i = 1 To 50
        ' �ϐ���������
        hissyu = 0
        ippan = 0
        zyoukyou = 0
        
        ' �u�V�K�v�V�[�g��I��
        With ThisWorkbook.Sheets("�f�[�^")
            ' �u�K�C�v�̓��_���v�Z����
            For j = 1 To 25
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    hissyu = hissyu + 1
                End If
            Next j

            
            
            ' �u��ʁv�̓��_���v�Z����
            For j = 26 To 89
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    ippan = ippan + 1
                End If
            Next j
                        For j = 121 To 122
            If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                     ippan = ippan + 1
                End If
            Next j
            
            ' �u�󋵁v�̓��_���v�Z����
            For j = 91 To 120
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    zyoukyou = zyoukyou + 2
                End If
            Next j
        End With
        
        ' �u���_�v�V�[�g��I��
        With ThisWorkbook.Sheets("���_")
            ' �u�K�C�v�����
            .Cells(3 * i, 3) = hissyu
            .Cells(3 * i, 4) = hissyu / HISSYU_MAX
            
            ' �u��ʁv�����
            .Cells(3 * i, 5) = ippan
            .Cells(3 * i, 6) = ippan / IPPAN_MAX
            
            ' �u�󋵁v�����
            .Cells(3 * i, 7) = zyoukyou
            .Cells(3 * i, 8) = zyoukyou / ZYOUKYOU_MAX
            
            '�u��ʂƏ󋵁v�̍��v
            .Cells(3 * i, 9) = .Cells(3 * i, 5) + .Cells(3 * i, 7)
            .Cells(3 * i, 10) = .Cells(3 * i, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
        End With
    Next i
    
    
    ' �ߌ�̐���
    For i = 1 To 50
        ' �ϐ���������
        hissyu = 0
        ippan = 0
        zyoukyou = 0
        
        ' �u�V�K�v�V�[�g��I��
        With Workbooks("�R�s�[���A�J��2��ߌ�.xlsx").Sheets("���A�J��2��ߌ�")
            ' �u�K�C�v�̓��_���v�Z����
            For j = 1 To 25
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    hissyu = hissyu + 1
                End If
            Next j

            
            
            ' �u��ʁv�̓��_���v�Z����
            For j = 26 To 89
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    ippan = ippan + 1
                End If
            Next j
                        For j = 121 To 122
            If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                     ippan = ippan + 1
                End If
            Next j
            
            ' �u�󋵁v�̓��_���v�Z����
            For j = 91 To 120
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    zyoukyou = zyoukyou + 2
                End If
            Next j
        End With
        
        ' �u���_�v�V�[�g��I��
        With ThisWorkbook.Sheets("���_")
            ' �u�K�C�v�����
            .Cells(3 * i + 1, 3) = hissyu
            .Cells(3 * i + 1, 4) = hissyu / HISSYU_MAX
            
            ' �u��ʁv�����
            .Cells(3 * i + 1, 5) = ippan
            .Cells(3 * i + 1, 6) = ippan / IPPAN_MAX
            
            ' �u�󋵁v�����
            .Cells(3 * i + 1, 7) = zyoukyou
            .Cells(3 * i + 1, 8) = zyoukyou / ZYOUKYOU_MAX
            
            '�u��ʂƏ󋵁v�̍��v
            .Cells(3 * i + 1, 9) = .Cells(3 * i + 1, 5) + .Cells(3 * i + 1, 7)
            .Cells(3 * i + 1, 10) = .Cells(3 * i + 1, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
        End With
    Next i
    
    
    ' ��������
    For i = 1 To 50
        ' �u���_�v�V�[�g��I��
        With ThisWorkbook.Sheets("���_")
            ' �u�K�C�v�����
            .Cells(3 * i + 2, 3) = .Cells(3 * i, 3) + .Cells(3 * i + 1, 3)
            .Cells(3 * i + 2, 4) = .Cells(3 * i + 2, 3) / HISSYU_MAX / 2
            
            ' �u��ʁv�����
            .Cells(3 * i + 2, 5) = .Cells(3 * i, 5) + .Cells(3 * i + 1, 5)
            .Cells(3 * i + 2, 6) = .Cells(3 * i + 2, 5) / IPPAN_MAX / 2
            
            ' �u�󋵁v�����
            .Cells(3 * i + 2, 7) = .Cells(3 * i, 7) + .Cells(3 * i + 1, 7)
            .Cells(3 * i + 2, 8) = .Cells(3 * i + 2, 7) / ZYOUKYOU_MAX / 2
            
            '�@�u��ʂƏ󋵁v�̍��v
            .Cells(3 * i + 2, 9) = .Cells(3 * i + 2, 5) + .Cells(3 * i + 2, 7)
            .Cells(3 * i + 2, 10) = .Cells(3 * i + 2, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
            
            
        End With
    Next i
    
End Sub
