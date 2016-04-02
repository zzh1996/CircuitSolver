Attribute VB_Name = "ModuleMatrix"
Option Explicit

'  功能：  使用全选主元高斯消去法求解线性方程组
'  参数    n     - Integer型变量，线性方程组的阶数
'         dblA   - Double型 n x n 二维数组，线性方程组的系数矩阵
'         dblB   - Double型长度为 n 的一维数组，线性方程组的常数向量，返回方程组的解向量
'  返回值：Boolean型，求解成功为True，无解或求解失败为False

Public Function LEGauss(n As Integer, dblA() As Double, dblB() As Double) As Boolean
    ' 局部变量
    Dim i As Integer, j As Integer, k As Integer
    Dim nIs As Integer
    ReDim nJs(n) As Integer
    Dim d As Double, t As Double
    
    ' 开始求解
    For k = 1 To n - 1
        d = 0#
        
        ' 归一
        For i = k To n
            For j = k To n
                t = Abs(dblA(i, j))
                If t > d Then
                    d = t
                    nJs(k) = j
                    nIs = i
                End If
            Next j
        Next i
        
        ' 无解，返回
        If d + 1# = 1# Then
            LEGauss = False
            Exit Function
        End If
        
        ' 消元
        If nJs(k) <> k Then
            For i = 1 To n
                t = dblA(i, k)
                dblA(i, k) = dblA(i, nJs(k))
                dblA(i, nJs(k)) = t
            Next i
        End If
        
        If nIs <> k Then
            For j = k To n
                t = dblA(k, j)
                dblA(k, j) = dblA(nIs, j)
                dblA(nIs, j) = t
            Next j
            t = dblB(k)
            dblB(k) = dblB(nIs)
            dblB(nIs) = t
        End If
        
        d = dblA(k, k)
        For j = k + 1 To n
            dblA(k, j) = dblA(k, j) / d
        Next j
        
        dblB(k) = dblB(k) / d
        For i = k + 1 To n
            For j = k + 1 To n
                dblA(i, j) = dblA(i, j) - dblA(i, k) * dblA(k, j)
            Next j
            dblB(i) = dblB(i) - dblA(i, k) * dblB(k)
        Next i
    Next k
    
    d = dblA(n, n)
    
    ' 无解，返回
    If Abs(d) + 1# = 1# Then
        LEGauss = False
        Exit Function
    End If
    
    ' 回代
    dblB(n) = dblB(n) / d
    For i = n - 1 To 1 Step -1
        t = 0#
        For j = i + 1 To n
          t = t + dblA(i, j) * dblB(j)
        Next j
        dblB(i) = dblB(i) - t
    Next i
    
    ' 调整解的次序
    nJs(n) = n
    For k = n To 1 Step -1
        If nJs(k) <> k Then
            t = dblB(k)
            dblB(k) = dblB(nJs(k))
            dblB(nJs(k)) = t
        End If
    Next k
    
    ' 求解成功
    LEGauss = True
    
End Function
