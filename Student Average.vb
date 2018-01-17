Module Module1
    'Program:   Student Grade Average
    'Programmer:    Ryan Roque
    'Date:  03/13/2017
    'Purpose:   Program requests midterm score and final score and computes average for each student and class.

    Sub Main()
        'Declerations
        Dim intMidterm As Integer = 0
        Dim intFinal As Integer = 0
        Dim strName As String = ""
        Dim dblGradeAverage As Double = 0
        Dim strEvaluation As String = ""
        Dim intClassSize As Integer = 0
        Dim dblClassAverage As Double = 0
        Dim intCount As Integer = 1

        'Get grades
        getClassSize(intClassSize)
        Do
            getStudentInfo(strName)
            getMidterm(intMidterm)
            getFinal(intFinal)

            'Calculate Average
            calculateAverage(intMidterm, intFinal, dblGradeAverage, dblClassAverage)
            standardPassFail(dblGradeAverage, strEvaluation)

            'Display Average
            displayAverage(strName, dblGradeAverage, strEvaluation)


            'Clear for next entry
            If intCount < intClassSize Then
                Console.Write("Press the enter key when ready to enter the next student")
                Console.ReadKey(True)
                Console.Clear()
            Else
                Console.Write("Press the enter key when ready to for the class average.")
                Console.ReadKey(True)
            End If

            'Intcrement Count
            intCount = intCount + 1

        Loop Until intCount > intClassSize

        calculateClassAverage(dblClassAverage, intClassSize)
        standardPassFail(dblClassAverage, strEvaluation)
        displayClassAverage(dblClassAverage, strEvaluation)
        terminateProgram()

    End Sub

    Private Sub getClassSize(ByRef intSize As Integer)

        Console.Write("Enter the number of students in the class.")
        intSize = CInt(Console.ReadLine())
        Console.Clear()

    End Sub

    Private Sub getStudentinfo(ByRef strStudent As String)

        Console.Write("Enter the name of the student.")
        strStudent = Console.ReadLine()

    End Sub

    Private Sub getMidTerm(ByRef intGrade As Integer)

        Console.Write("What is the student's midterm score?")
        intGrade = CInt(Console.ReadLine())

    End Sub

    Private Sub getFinal(ByRef intGrade As Integer)

        Console.Write("What is the student's final score?")
        intGrade = CInt(Console.ReadLine())

    End Sub

    Private Sub calculateAverage(ByVal intGrade1 As Integer, ByVal intGrade2 As Integer, ByRef dblAverage As Double, ByRef dblClassTotal As Double)

        dblAverage = CDbl((intGrade1 + intGrade2) / 2)
        dblClassTotal = dblClassTotal + dblAverage

    End Sub

    Private Sub standardPassFail(ByVal dblAverage As Double, ByRef strEval As String)

        If dblAverage >= 90 Then
            strEval = "A"
        ElseIf dblAverage > 80 Then
            strEval = "B"
        ElseIf dblAverage > 70 Then
            strEval = "C"
        ElseIf dblAverage > 60 Then
            strEval = "D"
        Else
            strEval = "F"
        End If

    End Sub

    Private Sub displayAverage(ByVal strStudent As String, ByVal dblAverage As Double, ByVal strEval As String)

        Console.Clear()
        Console.WriteLine("Student: " & strStudent)
        Console.WriteLine("The student's average is " & dblAverage & "%")
        Console.WriteLine("Evaluation: " & strEval)

    End Sub
    Private Sub calculateClassAverage(ByRef dblClassTotal As Double, ByVal intSize As Integer)

        dblClassTotal = dblClassTotal / intSize

    End Sub

    Private Sub displayClassAverage(ByVal dblAverage As Double, ByVal strEval As String)

        Console.Clear()
        Console.WriteLine("Press the enter key to terminate the program.")
        Console.WriteLine("The class evaluation is " & strEval)

    End Sub

    Private Sub terminateProgram()

        Console.WriteLine()
        Console.Write("Press the enter key to terminate the program")
        Console.Read()

    End Sub
End Module
