Attribute VB_Name = "ModuleMatrix"
Option Explicit

'  ���ܣ�  ʹ��ȫѡ��Ԫ��˹��ȥ��������Է�����
'  ����    n     - Integer�ͱ��������Է�����Ľ���
'         dblA   - Double�� n x n ��ά���飬���Է������ϵ������
'         dblB   - Double�ͳ���Ϊ n ��һά���飬���Է�����ĳ������������ط�����Ľ�����
'  ����ֵ��Boolean�ͣ����ɹ�ΪTrue���޽�����ʧ��ΪFalse

Public Function LEGauss(n As Integer, dblA() As Double, dblB() As Double) As Boolean
    ' �ֲ�����
    Dim i As Integer, j As Integer, k As Integer
    Dim nIs As Integer
    ReDim nJs(n) As Integer
    Dim d As Double, t As Double
    
    ' ��ʼ���
    For k = 1 To n - 1
        d = 0#
        
        ' ��һ
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
        
        ' �޽⣬����
        If d + 1# = 1# Then
            LEGauss = False
            Exit Function
        End If
        
        ' ��Ԫ
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
    
    ' �޽⣬����
    If Abs(d) + 1# = 1# Then
        LEGauss = False
        Exit Function
    End If
    
    ' �ش�
    dblB(n) = dblB(n) / d
    For i = n - 1 To 1 Step -1
        t = 0#
        For j = i + 1 To n
          t = t + dblA(i, j) * dblB(j)
        Next j
        dblB(i) = dblB(i) - t
    Next i
    
    ' ������Ĵ���
    nJs(n) = n
    For k = n To 1 Step -1
        If nJs(k) <> k Then
            t = dblB(k)
            dblB(k) = dblB(nJs(k))
            dblB(nJs(k)) = t
        End If
    Next k
    
    ' ���ɹ�
    LEGauss = True
    
End Function
