Public Class Form1
    Dim Controler_angle, Motor_angle As Single '控制點及馬達角度紀錄
    Dim ControlToMotorF_angle As Single '控制點與馬達正轉差
    Dim ControlToMotorR_angle As Single '控制點與馬達逆轉差
    Dim g As Graphics '簡化程式用
    Dim Controler_Move As Boolean '紀錄滑鼠點擊狀況
    Dim Controler_Click1 As Boolean '紀錄是否有進入範圍圈
    Dim MotorToCon_Turn As Single  '比較角位大小
    Dim angle_d As Point  '紀錄圖形上次移動量
    Dim O_Point As Point '紀錄原點
    Dim Controler_Point As Point '紀錄控制點點位
    Dim Motor_Step As Single '馬達前進角度量
    Dim RadioButton_Now As Integer '紀錄選擇開關的現狀(1：追隨，2：手動，3：復歸)
    Dim MotorManual_Turn As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '載入初始設定
        Controler_Move = False
        O_Point.X = 150
        O_Point.Y = 150
        Controler_angle = 0F
        MotorToCon_Turn = 0F
        Motor_angle = 0F
        Controler_Point.X = O_Point.X + 100
        Controler_Point.Y = O_Point.Y
        Timer1.Enabled = True
        ControlToMotorF_angle = 0F
        ControlToMotorR_angle = 0F
        Motor_Step = 0F
        RadioButton_Now = 1
        TextBox1.Text = 0
        MotorManual_Turn = True
    End Sub
    Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        '畫圖用函數
        g = e.Graphics '簡化程式

        g.DrawString("原點→", New Font("標楷體", 16, FontStyle.Bold), Brushes.Blue, (O_Point.X + 5), (O_Point.Y - 10))
        g.TranslateTransform(O_Point.X - 75, O_Point.Y - 75) '移動中心點
        g.DrawEllipse(Pens.Black, 0, 0, 150, 150) '劃一圓外框
        g.TranslateTransform(75, 75) '移動中心點
        g.RotateTransform(Motor_angle)
        g.DrawLine(Pens.Black, 0, 0, 75, 0) '製作結果線
        g.RotateTransform(-Motor_angle)
        g.RotateTransform(Controler_angle)
        g.TranslateTransform(100, 0) '移動至畫控制點位
        g.DrawLine(Pens.Red, 0, 0, -25, 0) '製作結果線
        g.TranslateTransform(-15, -15) '移動至畫控制點位
        g.DrawEllipse(Pens.Red, 0, 0, 30, 30) '製作控制點
    End Sub
    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        '當滑鼠按下時需再確認控制點是否可移動
        If Controler_Click1 Then
            Controler_Move = True '控制點可開啟
        End If
    End Sub
    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        '滑鼠放開時關閉控制點移動模式aa
        Controler_Move = False
    End Sub
    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If RadioButton_Now = 1 Then '如為追隨模式時啟動
            Dim Controler_XL, Controler_XR, Controler_YL, Controler_YR As Integer '計算控制點範圍
            '計算控制點的座標
            Controler_Point.X = O_Point.X + 100 * Math.Cos(Controler_angle * Math.PI / 180)
            Controler_Point.Y = O_Point.Y + 100 * Math.Sin(Controler_angle * Math.PI / 180)
            '計算控制圈可點擊範圍
            Controler_XL = Controler_Point.X - 15
            Controler_XR = Controler_Point.X + 15
            Controler_YL = Controler_Point.Y - 15
            Controler_YR = Controler_Point.Y + 15
            '計算滑鼠點位與原點的差
            angle_d.X = e.X - O_Point.X
            angle_d.Y = e.Y - O_Point.Y
            '判斷是否可移動控制點並在滑鼠DOWN加以判斷
            Controler_Click1 = (e.X >= Controler_XL) AndAlso (Controler_XR >= e.X) AndAlso (e.Y >= Controler_YL) AndAlso (Controler_YR >= e.Y)
            '滑鼠點擊判斷完後進行控制點移動(傳座標給畫圖用函數
            If Controler_Move Then
                Controler_Move_Sub() '控制點移動函數aa
            End If
            Me.Invalidate()  '重新畫圖
        End If
    End Sub
    Private Sub Controler_Move_Sub()
        Controler_angle = Math.Acos(angle_d.X / (angle_d.X ^ 2 + angle_d.Y ^ 2) ^ 0.5) / Math.PI * 180 '計算移動後角度
        Controler_angle = Controler_angle.ToString("000.0") '限制精度
        '因用COS計算，所以到COS無法計算的-相位時需協助
        If angle_d.Y < 0 Then
            Controler_angle = Controler_angle * -1 + 360
        End If

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        '此模擬步進馬達追隨移動模式
        Select Case MotorToCon_FollowMove() '呼叫馬達的追隨移動函數
            Case -1
                Motor_angle += Motor_Step
                If Motor_angle >= 360 Then
                    Motor_angle -= 360
                End If
            Case 1
                Motor_angle -= Motor_Step
                If Motor_angle < 0 Then
                    Motor_angle += 360
                End If
        End Select
        Motor_angle = Motor_angle.ToString("000.0")
        Me.Invalidate() '重新畫圖
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If Motor_angle = Controler_angle Then '判斷是否已經到位 如果是就跳出及關閉TIMER
            Timer2.Enabled = False
            Return
        End If
        Select Case MotorManual_Turn '判斷轉向
            Case 1
                If Math.Abs(Controler_angle - Motor_angle) > 1 Then '設定步進量
                    Motor_angle += 1
                Else
                    Motor_angle += 0.1
                End If
                If Motor_angle >= 360 Then '超過360就砍360
                    Motor_angle -= 360
                End If
            Case 0
                If Math.Abs(Controler_angle - Motor_angle) > 1 Then '設定步進量
                    Motor_angle -= 1
                Else
                    Motor_angle -= 0.1
                End If
                If Motor_angle < 0 Then '低於0補360
                    Motor_angle += 360
                End If
        End Select
        Motor_angle = Motor_angle.ToString("000.0")
        Me.Invalidate()
    End Sub
    Function MotorToCon_FollowMove() As Integer
        '模擬馬達追隨移動函數
        MotorToCon_Turn = Controler_angle - Motor_angle '計算控制點與馬達角位差
        MotorToCon_Turn = MotorToCon_Turn.ToString("000.0")
        ControlToMotorF_angle = Math.Abs(Controler_angle - Motor_angle) '計算控制點與馬達角位差的正轉量
        ControlToMotorF_angle = ControlToMotorF_angle.ToString("000.0")
        ControlToMotorR_angle = 360 - ControlToMotorF_angle '計算控制點與馬達角位差的逆轉量
        If ControlToMotorR_angle >= 360 Then '因計算模式可能會炸出360故所訂至360以內
            ControlToMotorR_angle -= 360
        End If
        ControlToMotorR_angle = ControlToMotorR_angle.ToString("000.0")

        If ControlToMotorF_angle > 1 Then '設定步進量
            Motor_Step = 1
        Else
            Motor_Step = 0.1
        End If

        '判斷正逆轉 1為正轉 -1為逆轉 0為停止
        If ((MotorToCon_Turn > 0) AndAlso (ControlToMotorF_angle > ControlToMotorR_angle)) OrElse ((MotorToCon_Turn < 0) AndAlso (ControlToMotorR_angle > ControlToMotorF_angle)) Then
            Return (1)
        ElseIf ((MotorToCon_Turn < 0) AndAlso (ControlToMotorF_angle > ControlToMotorR_angle)) OrElse ((MotorToCon_Turn > 0) AndAlso (ControlToMotorR_angle > ControlToMotorF_angle)) Then
            Return (-1)
        Else
            Return (0)
        End If
    End Function

    Private Sub RadioButton_Follow_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Follow.CheckedChanged
        '追隨模式的按鈕狀態設定
        Button_F.Enabled = False
        Button_R.Enabled = False
        Button_S.Enabled = False
        TrackBar1.Enabled = False
        TextBox1.Enabled = False
        Timer1.Enabled = True
        Timer2.Enabled = False
        RadioButton_Now = 1
    End Sub

    Private Sub RadioButton_Manual_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Manual.CheckedChanged
        '手動模式的按鈕狀態設定
        Button_F.Enabled = True
        Button_R.Enabled = True
        Button_S.Enabled = False
        TrackBar1.Enabled = True
        TextBox1.Enabled = True
        Timer1.Enabled = False
        Timer2.Enabled = False
        RadioButton_Now = 2
    End Sub

    Private Sub RadioButton_Reply_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Reply.CheckedChanged
        '復歸模式的按鈕狀態設定
        Button_F.Enabled = True
        Button_R.Enabled = True
        Button_S.Enabled = True
        TrackBar1.Enabled = False
        TextBox1.Enabled = False
        Timer1.Enabled = False
        Timer2.Enabled = False
        RadioButton_Now = 3
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Select Case TrackBar1.Value '設定拉桿後呈現數值
            Case 0
                TextBox1.Text = 0
            Case 1
                TextBox1.Text = 0.1
            Case 2
                TextBox1.Text = 0.5
            Case 3
                TextBox1.Text = 1
            Case 4
                TextBox1.Text = 5
            Case 5
                TextBox1.Text = 15
        End Select
    End Sub
    Private Sub Button_F_Click(sender As Object, e As EventArgs) Handles Button_F.Click
        '正轉按鈕按下後函數
        Select Case RadioButton_Now '判斷模式
            Case 2 '手動模式
                Controler_angle += TextBox1.Text '控制點+輸入格的數字
                If Controler_angle >= 360 Then '超過360就砍360
                    Controler_angle -= 360
                End If
                Controler_angle = Controler_angle.ToString("000.0")
            Case 3 '自動模式
                Controler_angle = 0F '控制點歸0
                Controler_angle = Controler_angle.ToString("000.0")
        End Select
        Timer2.Enabled = True '開啟手動模擬用
        MotorManual_Turn = True '計入馬達轉向
    End Sub
    Private Sub Button_R_Click(sender As Object, e As EventArgs) Handles Button_R.Click
        '逆轉按鈕按下後函數
        Select Case RadioButton_Now '判斷模式
            Case 2 '手動模式
                Controler_angle -= TextBox1.Text '控制點-輸入格的數字
                If Controler_angle < 0 Then '低於0時補360
                    Controler_angle += 360
                End If
                Controler_angle = Controler_angle.ToString("000.0")

            Case 3 '自動模式
                Controler_angle = 0F '控制點歸0
                Controler_angle = Controler_angle.ToString("000.0")
        End Select
        Timer2.Enabled = True '開啟手動模擬用
        MotorManual_Turn = False '計入馬達轉向
    End Sub
    Private Sub Button_S_Click(sender As Object, e As EventArgs) Handles Button_S.Click
        Controler_angle = 0F '控制點歸0
        Controler_angle = Controler_angle.ToString("000.0")
        Timer1.Enabled = True '追隨模式開啟
    End Sub
End Class