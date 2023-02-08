VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "R76101120_Hw3"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   ScaleHeight     =   6555
   ScaleWidth      =   15510
   Begin VB.ListBox List9 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   7920
      TabIndex        =   9
      Top             =   1320
      Width           =   3500
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   11760
      TabIndex        =   6
      Top             =   1320
      Width           =   3500
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   3500
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3500
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "glass.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Entropy_Based Interval"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "Equal_Frequency Interval"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "select naive Bayes result:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Equal_Width Interval"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Partition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String, out_file As String, nstr As String
Dim out_rec As String
Dim att(214, 10) As Double
Dim origin_att(214, 10), dummy As Double
Dim equal_width(10, 10) As Integer
Dim equal_frequency(10, 10) As Integer
Dim entropy_based(10, 10) As Integer
Dim max(10, 1), min(10, 1), bigM, bound, Interval, f_sort(214, 2) As Double
Dim i, j, num, pos, bin, Sorted, ii, jj, cut_len, cut_size, p As Integer
Dim val, key, a, b As Integer
Dim selected(9) As Boolean
Dim classcount(7) As Integer
Dim cut_array(214) As Double

Sub Entropy(EA() As Double)
    Dim cut As Integer
    Dim size, l_num, r_num, cut_index, NL, NR As Integer
    Dim T_Ent, Le, Re, cut_point, prob, Ent, Min_Ent, Min_cut_point, delta As Double
    Dim countEA(7), countL(7), countR(7) As Integer
    Min_Ent = bigM
    T_Ent = 0
    cut = 0

    size = UBound(EA) - LBound(EA)
    If size > 1 Then
        For i = 1 To size - 1
            If EA(i, 1) <> EA(i + 1, 1) Then
                cut = cut + 1
            End If
        Next
    End If
    
    
    If cut = 0 Then
        'List8.AddItem "stop"
        'List8.AddItem "-----------------------"
    Else
        For j = 1 To 7
            countEA(j) = 0
        Next
        For i = 1 To size
            countEA(EA(i, 1)) = countEA(EA(i, 1)) + 1
        Next
       
        For i = 1 To 7
            prob = countEA(i) / size
            If prob <> 0 Then
                T_Ent = T_Ent - prob * Math.Log(prob) / Math.Log(2)
            End If
        Next
        
        For i = 1 To size - 1
            If EA(i, 1) <> EA(i + 1, 1) Then
                cut_point = (EA(i, 0) + EA(i + 1, 0)) / 2
                For j = 1 To 7
                    countL(j) = 0
                    countR(j) = 0
                Next
                For j = 1 To size
                    'If j < i + 1 Then
                    If EA(j, 0) <= cut_point Then
                        countL(EA(j, 1)) = countL(EA(j, 1)) + 1
                    Else
                        countR(EA(j, 1)) = countR(EA(j, 1)) + 1
                    End If
                Next
                NL = 0
                NR = 0
                For j = 1 To 7
                    NL = NL + countL(j)
                    NR = NR + countR(j)
                Next
                If NL = 0 Or NR = 0 Then
                    Exit For
                End If

                Le = 0
                Re = 0
                For p = 1 To 7
                    'prob = countL(p) / i
                    prob = countL(p) / NL
                    If prob <> 0 Then
                        Le = Le - prob * Math.Log(prob) / Math.Log(2)
                    End If
                Next
                For p = 1 To 7
                    'prob = countR(p) / (size - i)
                    prob = countR(p) / NR
                    If prob <> 0 Then
                        Re = Re - prob * Math.Log(prob) / Math.Log(2)
                    End If
                Next
                
                Ent = Le * NL / size + Re * NR / size
                'Ent = Le * i / size + Re * (size - i) / size
                Dim k, k1, k2 As Integer
                Dim reject As Double
                'reject condition
                k = 0
                k1 = 0
                k2 = 0
                For p = 1 To 7
                    If countEA(p) <> 0 Then
                        k = k + 1
                    End If
                    If countL(p) <> 0 Then
                        k1 = k1 + 1
                    End If
                    If countR(p) <> 0 Then
                        k2 = k2 + 1
                    End If
                Next

                delta = (Math.Log(3 ^ k - 2) / Math.Log(2)) - k * T_Ent + k1 * Le + k2 * Re
                reject = T_Ent - Ent - ((Math.Log(size - 1) / Math.Log(2)) + delta) / size
                If Min_Ent >= Ent And reject > 0 Then
                    Min_Ent = Ent
                    Min_cut_point = cut_point
                    'cut_index = i + 1
                    cut_index = NL + 1
                End If
            End If
        Next
        If Min_Ent <> bigM Then
            cut_len = cut_len + 1

            cut_array(cut_len) = Min_cut_point

            l_num = cut_index - 1
            r_num = size - l_num
            Dim left() As Double
            Dim right() As Double
            ReDim left(l_num, 2) As Double
            ReDim right(r_num, 2) As Double
            'List8.AddItem l_num & " " & r_num
            Dim l, r As Integer
            l = 1
            r = 1
            For i = 1 To size
                If i < cut_index Then
                    left(l, 0) = EA(i, 0)
                    left(l, 1) = EA(i, 1)
                    l = l + 1
                Else
                    right(r, 0) = EA(i, 0)
                    right(r, 1) = EA(i, 1)
                    r = r + 1
                End If
            Next
            Call Entropy(left)
            Call Entropy(right)
        End If
    End If
End Sub

Sub BAY(b_att() As Double)
    Dim ran(214), size, num_bay, temp, a, b, c, d As Integer
    Dim train_data() As Double
    Dim test_data() As Double
    Dim Cmax, bestC, pc(7), pxc(9, 7), total_pxc, correct, v(9) As Double
    Dim predict() As Double
    Dim test_num, train_num As Integer
    Randomize Timer
    
    Dim pcjx() As Double
    Dim att_val_count(9, 9) As Double
    Dim col As Integer
    

    
    
    classcount(0) = 0
    classcount(1) = 70
    classcount(2) = 76
    classcount(3) = 17
    classcount(4) = 0
    classcount(5) = 13
    classcount(6) = 9
    classcount(7) = 29
    
    '計算各attribute有幾種值
    Dim att_type(9) As Boolean
    Dim value As Integer
    For i = 0 To 9
        att_type(i) = 0
    Next
    For col = 1 To 9
'        'value = 0
'        'While i < 10
        For value = 0 To 9
            For i = 1 To 214
                If b_att(i, col) = value Then
                    att_type(value) = 1
                    'value = value + 1
'                    Exit For
                End If
            Next
'        'end while
        Next
    Next
    'For i = 0 To 9
    '    List10.AddItem att_type(i)
    'Next
    
    '計算各attribute的值各有幾個
    For i = 0 To 9
        For j = 0 To 9
            att_val_count(i, j) = 0
        Next
    Next
    For col = 1 To 9
        For j = 1 To 214
            att_val_count(col, b_att(j, col)) = att_val_count(col, b_att(j, col)) + 1
        Next
    Next
    
    '三圍矩陣存p(Xi|Cj)with laplace
    ReDim pcjx(9, 7, 9) As Double
    For col = 1 To 9
    'col = 1
        For i = 1 To 214
            pcjx(col, b_att(i, 10), b_att(i, col)) = pcjx(col, b_att(i, 10), b_att(i, col)) + 1
            
        Next
         '& " " & pcjx(col, 1, 1) & " " & pcjx(col, 1, 2)
        'For i = 1 To 7
        '    For j = 0 To 9
        '        pcjx(col, i, j) = (pcjx(col, i, j) + 1) / (classcount(i) + 10)
        '    Next
            
        'Next
        'List5.AddItem pcjx(col, 1, 0)
    Next
    'List5.AddItem "----------------------------"
'--------------------------------------------------------------
    Dim classval As Integer
    Dim maxp As Double
    Dim p As Double
    Dim maxpclass As Integer
    Dim accucount As Integer
    Dim att_accu As Double
    Dim max_att_accu As Double
    Dim maxcol As Integer
    maxcol = -1
    max_att_accu = -10000
    maxp = -10000
    maxpclass = -10000
    classval = 1
    'col = 2
    For col = 1 To 9
        accucount = 0
        For i = 1 To 214
            For classval = 1 To 7
                p = (classcount(classval) / 214) * (pcjx(col, classval, b_att(i, col)) + 1 / (classcount(classval) + 10))
                If p > maxp Then
                    maxp = p
                    maxpclass = classval
                End If
            Next
            If b_att(i, 10) = maxpclass Then
                accucount = accucount + 1
            End If
        Next
        att_accu = accucount / 214
        If att_accu > max_att_accu Then
            max_att_accu = att_accu
            maxcol = col
        End If
        'List8.AddItem max_att_accu & " " & maxcol
    Next
    
    List6.AddItem "Attribute chosen: A" & maxcol
    List6.AddItem "Accuracy:" & " " & max_att_accu
'--------------------------------------------------------------
    Dim maxAccu As Double
    Dim index As Integer
    maxAccu = 0
        '一開始全不選
    For i = 0 To 9
        selected(i) = 0
    Next
    selected(maxcol) = 1
    'forward
    For i = 1 To 9
        index = -1
        For j = 1 To 9
            If Not selected(j) Then
                selected(j) = 1
                If Accuracy > maxAccu Then
                    maxAccu = Accuracy
                    index = j
                End If
                selected(j) = 0
            End If
        Next
        If index = -1 Then Exit For
        selected(index) = 1

        List6.AddItem ("Attribute Selected：A" & index)
        List6.AddItem ("Accuracy：" & Math.Round(maxAccu, 5))
    Next
    
    
        
    List6.AddItem ("The Attribute Subset :")
    For i = 1 To 9
        If selected(i) Then
            List6.AddItem ("A" & i)
        End If
    Next i
    List6.AddItem ("--------------------------------------------------")
    
End Sub
Private Function Accuracy()
    Dim classval As Integer
    classval = 1
    For i = 1 To 214
        For classval = 1 To 7
            
        Next
    Next
    'Accuracy = 1
End Function

Private Sub Partition_click()
    bigM = 100000000
    For i = 1 To 10
        max(i, 0) = -bigM
        min(i, 0) = bigM
        For j = 0 To 10
            equal_width(i, j) = 0
            equal_frequency(i, j) = 0
            entropy_based(i, j) = 0
        Next
    Next
    'List1.Clear
    List2.Clear
    'List3.Clear
    List4.Clear
    List6.Clear
    'List7.Clear
    List9.Clear
    'check whether the file name is empty
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        'check whether the data file exists
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            Open in_file For Input As #1
            num = 1
            Do While Not EOF(1)
                For i = 1 To 11
                    If i = 1 Then
                        Input #1, dummy
                    Else
                        Input #1, origin_att(num, i - 1)
                        att(num, i - 1) = origin_att(num, i - 1)
                        If (i > 1 And i < 12) Then
                            If max(i - 1, 0) < att(num, i - 1) Then
                                max(i - 1, 0) = att(num, i - 1)
                                max(i - 1, 1) = num
                            End If
                            If min(i - 1, 0) > att(num, i - 1) Then
                                min(i - 1, 0) = att(num, i - 1)
                                min(i - 1, 1) = num
                            End If
                        End If
                    End If
                    
                Next
                
                'List1.AddItem att(num, 1) & " " & att(num, 2) & " " & att(num, 3) & " " & att(num, 4) & " " & att(num, 5) & " " & att(num, 6) & " " & att(num, 7) & " " & att(num, 8) & " " & att(num, 9) & " " & att(num, 10)
                num = num + 1
            Loop
            max(10, 0) = -bigM
            min(10, 0) = bigM
            Close #1
        End If
    End If
    bin = 10
    
    'Equal_Width
    For i = 1 To 214
        For j = 1 To 10
            If max(j, 0) <> -bigM Then
                Interval = (max(j, 0) - min(j, 0)) / bin
                pos = (att(i, j) - min(j, 0)) / Interval
                pos = -Fix(-(pos - 0.00001))
                'List2.AddItem att(i, j) & " " & min(j, 0) & " " & Interval & " " & pos
                att(i, j) = pos
                equal_width(j, att(i, j)) = equal_width(j, att(i, j)) + 1
            Else
                equal_width(j, att(i, j)) = equal_width(j, att(i, j)) + 1
            End If
        Next
        'List1.AddItem att(i, 1) & " " & att(i, 2) & " " & att(i, 3) & " " & att(i, 4) & " " & att(i, 5) & " " & att(i, 6) & " " & att(i, 7) & " " & att(i, 8) & " " & att(i, 9) & " " & att(i, 10)
    Next
    For i = 1 To 9
        If max(i, 0) <> -bigM Then
            Interval = (max(i, 0) - min(i, 0)) / bin
            List2.AddItem "A" & i
            List2.AddItem 0 & ":[" & min(i, 0) & " , " & min(i, 0) + Interval & "]"
            min(i, 0) = min(i, 0) + Interval
            For j = 2 To 10
                List2.AddItem j - 1 & ":(" & min(i, 0) & " , " & min(i, 0) + Interval & "]"
                min(i, 0) = min(i, 0) + Interval
            Next
        End If
        List2.AddItem "-------------------------------------------------------"
    Next
    
    List6.AddItem "[Equal Width]"
    
    Call BAY(att)
    
    For i = 1 To 214
        For j = 1 To 10
            att(i, j) = origin_att(i, j)
        Next
    Next
    
    Dim Interval_A(11) As Double
    
    'Equal_Frequency
    For j = 1 To 9
        
        For i = 1 To 214
            If max(j, 0) <> -bigM Then
                If Sorted = 0 Then
                    For a = 1 To 214
                        f_sort(a, 0) = att(a, j)
                        f_sort(a, 1) = a
                    Next
                    Sorted = 1
                    For a = 2 To 214
                        val = f_sort(a, 0)
                        key = f_sort(a, 1)
                        b = a - 1
                        While b > 0 And f_sort(b, 0) >= val
                            f_sort(b + 1, 0) = f_sort(b, 0)
                            f_sort(b + 1, 1) = f_sort(b, 1)
                            b = b - 1
                        Wend
                        f_sort(b + 1, 0) = val
                        f_sort(b + 1, 1) = key
                    Next
                    Interval_A(1) = f_sort(1, 0)
                    For a = 2 To 5
                        Interval_A(a) = (f_sort(22 * (a - 1), 0) + f_sort(22 * (a - 1) + 1, 0)) / 2
                    Next
                    For a = 6 To 10
                        Interval_A(a) = (f_sort(22 * 4 + 21 * (a - 5), 0) + f_sort(22 * 4 + 21 * (a - 5) + 1, 0)) / 2
                    Next
                    Interval_A(11) = f_sort(214, 0)
                End If
                
                For a = 2 To 11
                    If a = 2 Then
                        If Interval_A(2) >= f_sort(i, 0) Then
                            pos = 0
                        End If
                    Else
                        If f_sort(i, 0) <= Interval_A(a) And f_sort(i, 0) > Interval_A(a - 1) Then
                            pos = a - 2
                        End If
                    End If
                Next
                att(f_sort(i, 1), j) = pos
            End If
        Next
        Sorted = 0
        If max(j, 0) <> -bigM Then
            List4.AddItem "A" & j
            List4.AddItem 0 & ":[" & f_sort(1, 0) & " , " & (f_sort(22, 0) + f_sort(23, 0)) / 2 & "]"
            For a = 2 To 4
                List4.AddItem a - 1 & ":(" & (f_sort(22 * (a - 1), 0) + f_sort(22 * (a - 1) + 1, 0)) / 2 & " , " & (f_sort(22 * a, 0) + f_sort(22 * a + 1, 0)) / 2 & "]"
            Next
            List4.AddItem 4 & ":(" & (f_sort(22 * 4, 0) + f_sort(22 * 4 + 1, 0)) / 2 & " , " & (f_sort(22 * 4 + 21, 0) + f_sort(22 * 4 + 22, 0)) / 2 & "]"
            For a = 6 To 9
                List4.AddItem a - 1 & ":(" & (f_sort(22 * 4 + 21 * (a - 5), 0) + f_sort(22 * 4 + 21 * (a - 5) + 1, 0)) / 2 & " , " & (f_sort(22 * 4 + 21 * (a - 4), 0) + f_sort(22 * 4 + 21 * (a - 4) + 1, 0)) / 2 & "]"
            Next
            List4.AddItem 9 & ":(" & (f_sort(22 * 4 + 21 * 5, 0) + f_sort(22 * 4 + 21 * 5 + 1, 0)) / 2 & " , " & f_sort(214, 0) & "]"
        End If
        List4.AddItem "-------------------------------------------------------"
    Next
    For i = 1 To 214
        'List3.AddItem att(i, 1) & " " & att(i, 2) & " " & att(i, 3) & " " & att(i, 4) & " " & att(i, 5) & " " & att(i, 6) & " " & att(i, 7) & " " & att(i, 8) & " " & att(i, 9) & " " & att(i, 10)
    Next

    List6.AddItem "[Equal frequency]"
    Call BAY(att)
    
    For i = 1 To 214
        For j = 1 To 10
            att(i, j) = origin_att(i, j)
        Next
    Next
    
    For ii = 1 To 10
        If (ii <> 10) Then
            Dim ent_array(214, 2) As Double
            For jj = 1 To 214
                ent_array(jj, 0) = att(jj, ii)
                ent_array(jj, 1) = att(jj, 10)
            Next
            For a = 2 To 214
                val = ent_array(a, 0)
                key = ent_array(a, 1)
                b = a - 1
                While b > 0 And ent_array(b, 0) > val
                    ent_array(b + 1, 0) = ent_array(b, 0)
                    ent_array(b + 1, 1) = ent_array(b, 1)
                    b = b - 1
                Wend
                ent_array(b + 1, 0) = val
                ent_array(b + 1, 1) = key
            Next
            
            cut_len = 0
            cut_size = UBound(cut_array) - LBound(cut_array)
            For a = 1 To cut_size
                cut_array(a) = -1
            Next
            Call Entropy(ent_array)
            'List8.AddItem "====================================="
            'List9.AddItem UBound(cut_array) & " " & LBound(cut_array)
            For a = 2 To cut_size
                val = cut_array(a)
                b = a - 1
                While b > 0 And cut_array(b) > val
                    cut_array(b + 1) = cut_array(b)
                    b = b - 1
                Wend
                cut_array(b + 1) = val
            Next
            
            Dim print_cut_point As New Collection
            print_cut_point.Add ent_array(1, 0)
            For a = 1 To cut_size
                If cut_array(a) <> -1 Then
                    print_cut_point.Add cut_array(a)
                End If
            Next
            print_cut_point.Add ent_array(214, 0)
            List9.AddItem "A" & ii & ":"
            For a = 1 To print_cut_point.Count - 1
                If a = 1 Then
                    List9.AddItem a - 1 & ":[" & print_cut_point(a) & " , " & print_cut_point(a + 1) & "]"
                Else
                    List9.AddItem a - 1 & ":(" & print_cut_point(a) & " , " & print_cut_point(a + 1) & "]"
                End If
            Next
            List9.AddItem "-------------------------------------------------------"
            For jj = 1 To 214
                If att(jj, ii) <= print_cut_point(2) Then
                    att(jj, ii) = 0
                Else
                    For a = 2 To print_cut_point.Count - 1
                        If att(jj, ii) > print_cut_point(a) And att(jj, ii) <= print_cut_point(a + 1) Then
                            att(jj, ii) = a - 1
                            Exit For
                        End If
                    Next
                End If
            Next
            For a = 1 To print_cut_point.Count
                print_cut_point.Remove (1)
            Next
        End If
    Next
    
    For i = 1 To 214
        'List7.AddItem att(i, 1) & " " & att(i, 2) & " " & att(i, 3) & " " & att(i, 4) & " " & att(i, 5) & " " & att(i, 6) & " " & att(i, 7) & " " & att(i, 8) & " " & att(i, 9) & " " & att(i, 10)
    Next
    List6.AddItem "[Entropy Based]"
    Call BAY(att)
End Sub


