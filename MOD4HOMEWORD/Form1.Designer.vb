<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtFood = New System.Windows.Forms.TextBox()
        Me.txtCalsServe = New System.Windows.Forms.TextBox()
        Me.txtGramsServe = New System.Windows.Forms.TextBox()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnCalc
        '
        Me.btnCalc.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.btnCalc.Location = New System.Drawing.Point(27, 236)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(445, 48)
        Me.btnCalc.TabIndex = 0
        Me.btnCalc.Text = "Compute % Calories from Fat"
        Me.btnCalc.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label1.Location = New System.Drawing.Point(134, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(114, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Name of Food:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label2.Location = New System.Drawing.Point(151, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Calories per serving:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label3.Location = New System.Drawing.Point(122, 170)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(186, 20)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Grams of fat per serviing:"
        '
        'txtFood
        '
        Me.txtFood.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtFood.Location = New System.Drawing.Point(253, 37)
        Me.txtFood.Name = "txtFood"
        Me.txtFood.Size = New System.Drawing.Size(129, 26)
        Me.txtFood.TabIndex = 4
        '
        'txtCalsServe
        '
        Me.txtCalsServe.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtCalsServe.Location = New System.Drawing.Point(307, 104)
        Me.txtCalsServe.Name = "txtCalsServe"
        Me.txtCalsServe.Size = New System.Drawing.Size(75, 26)
        Me.txtCalsServe.TabIndex = 5
        '
        'txtGramsServe
        '
        Me.txtGramsServe.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtGramsServe.Location = New System.Drawing.Point(307, 168)
        Me.txtGramsServe.Name = "txtGramsServe"
        Me.txtGramsServe.Size = New System.Drawing.Size(75, 26)
        Me.txtGramsServe.TabIndex = 6
        '
        'txtMessage
        '
        Me.txtMessage.Enabled = False
        Me.txtMessage.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtMessage.Location = New System.Drawing.Point(27, 318)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(445, 59)
        Me.txtMessage.TabIndex = 7
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(506, 411)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.txtGramsServe)
        Me.Controls.Add(Me.txtCalsServe)
        Me.Controls.Add(Me.txtFood)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnCalc)
        Me.Name = "Form1"
        Me.Text = "Nutrition"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCalc As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtFood As TextBox
    Friend WithEvents txtCalsServe As TextBox
    Friend WithEvents txtGramsServe As TextBox
    Friend WithEvents txtMessage As TextBox

    '' Author: Francisco Bumanglag
    '' Project: Nutrition 
    '' Assignment: Module 4 Homework
    '' Course: Visual Basic, Santa Ana College
    '' Class: CMPR105 Derendal Huynh 
    '' Date: September 16, 2022

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        'DECLARATIONS
        Const fatCalsPerServe = 9
        Dim food As String = txtFood.Text
        Dim approved As String
        Dim calsPerServe, gramsPerServe, totalCalsFat As Integer
        Dim percentCals As Double

        'CHECK FOR VALID ENTRIES
        If food <> "" And IsNumeric(txtCalsServe.Text) And IsNumeric(txtGramsServe.Text) Then

            'CALCULATIONS
            calsPerServe = txtCalsServe.Text
            gramsPerServe = txtGramsServe.Text

            totalCalsFat = gramsPerServe * fatCalsPerServe
            percentCals = totalCalsFat / calsPerServe

            'DETERMINE AHA RECOMMENDATIONS
            If percentCals < 0.3 Then
                approved = "does not exceed AHA recommendation"
            Else approved = "exceeds AHA recommendation"
            End If

            'PRESENT OUTCOME AND FORMAT TO PERCENT
            txtMessage.Text = food & " contains " & FormatPercent(percentCals, 2) &
                    " calories from fat, which " & approved

        Else
            'ERROR MESSAGE FOR EMPTY OR INVALID INPUTS
            MessageBox.Show("PLEASE ENTER THE CORRECT VALUE(S)")

        End If

    End Sub

End Class


