Sub abc()
    ' 各問題形式の問題数×得点を定数として定義
    Const HISSYU_MAX As Integer = (25 - 1 + 1) * 1
    Const IPPAN_MAX As Integer = ((89 - 26 + 1) + (122 - 121 + 1)) * 1
    Const ZYOUKYOU_MAX As Integer = (120 - 91 + 1) * 2
    
    ' For文用のカウンタ変数を宣言
    Dim i, j As Integer
    ' 得点記録用の変数を宣言
    Dim hissyu, ippan, zyoukyou As Integer
    
    ' 午前の成績
    For i = 1 To 50
        ' 変数を初期化
        hissyu = 0
        ippan = 0
        zyoukyou = 0
        
        ' 「新規」シートを選択
        With ThisWorkbook.Sheets("データ")
            ' 「必修」の得点を計算する
            For j = 1 To 25
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    hissyu = hissyu + 1
                End If
            Next j

            
            
            ' 「一般」の得点を計算する
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
            
            ' 「状況」の得点を計算する
            For j = 91 To 120
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    zyoukyou = zyoukyou + 2
                End If
            Next j
        End With
        
        ' 「得点」シートを選択
        With ThisWorkbook.Sheets("得点")
            ' 「必修」を入力
            .Cells(3 * i, 3) = hissyu
            .Cells(3 * i, 4) = hissyu / HISSYU_MAX
            
            ' 「一般」を入力
            .Cells(3 * i, 5) = ippan
            .Cells(3 * i, 6) = ippan / IPPAN_MAX
            
            ' 「状況」を入力
            .Cells(3 * i, 7) = zyoukyou
            .Cells(3 * i, 8) = zyoukyou / ZYOUKYOU_MAX
            
            '「一般と状況」の合計
            .Cells(3 * i, 9) = .Cells(3 * i, 5) + .Cells(3 * i, 7)
            .Cells(3 * i, 10) = .Cells(3 * i, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
        End With
    Next i
    
    
    ' 午後の成績
    For i = 1 To 50
        ' 変数を初期化
        hissyu = 0
        ippan = 0
        zyoukyou = 0
        
        ' 「新規」シートを選択
        With Workbooks("コピー東アカ第2回午後.xlsx").Sheets("東アカ第2回午後")
            ' 「必修」の得点を計算する
            For j = 1 To 25
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    hissyu = hissyu + 1
                End If
            Next j

            
            
            ' 「一般」の得点を計算する
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
            
            ' 「状況」の得点を計算する
            For j = 91 To 120
                If .Cells(1 + i, 6 + j).Value = .Cells(2 + i, 6 + j).Value Then
                    zyoukyou = zyoukyou + 2
                End If
            Next j
        End With
        
        ' 「得点」シートを選択
        With ThisWorkbook.Sheets("得点")
            ' 「必修」を入力
            .Cells(3 * i + 1, 3) = hissyu
            .Cells(3 * i + 1, 4) = hissyu / HISSYU_MAX
            
            ' 「一般」を入力
            .Cells(3 * i + 1, 5) = ippan
            .Cells(3 * i + 1, 6) = ippan / IPPAN_MAX
            
            ' 「状況」を入力
            .Cells(3 * i + 1, 7) = zyoukyou
            .Cells(3 * i + 1, 8) = zyoukyou / ZYOUKYOU_MAX
            
            '「一般と状況」の合計
            .Cells(3 * i + 1, 9) = .Cells(3 * i + 1, 5) + .Cells(3 * i + 1, 7)
            .Cells(3 * i + 1, 10) = .Cells(3 * i + 1, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
        End With
    Next i
    
    
    ' 総合成績
    For i = 1 To 50
        ' 「得点」シートを選択
        With ThisWorkbook.Sheets("得点")
            ' 「必修」を入力
            .Cells(3 * i + 2, 3) = .Cells(3 * i, 3) + .Cells(3 * i + 1, 3)
            .Cells(3 * i + 2, 4) = .Cells(3 * i + 2, 3) / HISSYU_MAX / 2
            
            ' 「一般」を入力
            .Cells(3 * i + 2, 5) = .Cells(3 * i, 5) + .Cells(3 * i + 1, 5)
            .Cells(3 * i + 2, 6) = .Cells(3 * i + 2, 5) / IPPAN_MAX / 2
            
            ' 「状況」を入力
            .Cells(3 * i + 2, 7) = .Cells(3 * i, 7) + .Cells(3 * i + 1, 7)
            .Cells(3 * i + 2, 8) = .Cells(3 * i + 2, 7) / ZYOUKYOU_MAX / 2
            
            '　「一般と状況」の合計
            .Cells(3 * i + 2, 9) = .Cells(3 * i + 2, 5) + .Cells(3 * i + 2, 7)
            .Cells(3 * i + 2, 10) = .Cells(3 * i + 2, 9) / (IPPAN_MAX + ZYOUKYOU_MAX) / 2
            
            
        End With
    Next i
    
End Sub
