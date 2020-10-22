Option Explicit On
Option Strict On
Option Compare Binary
'Lane Coleman
'RCET 0265
'Fall 2020
'Car Rental Form.
Public Class RentalForm
    Dim customersTotal, milesTotal As Integer
    Dim chargesTotal As Decimal

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim goodData As Boolean
        goodData = False
        goodData = Validate()
        If goodData = False Then
            Exit Sub
        End If
        customersTotal += 1
        Calculate()
    End Sub
    Function Validate() As Boolean
        Dim eval As Boolean
        Dim zipCheck, bOCheck, eOCheck, daysCheck As Integer

        If NameTextBox.Text = Nothing Then
            ActiveControl = NameTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf AddressTextBox.Text = Nothing Then
            ActiveControl = AddressTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf CityTextBox.Text = Nothing Then
            ActiveControl = CityTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf StateTextBox.Text = Nothing Then
            ActiveControl = StateTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf ZipCodeTextBox.Text = Nothing Then
            ActiveControl = ZipCodeTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf BeginOdometerTextBox.Text = Nothing Then
            ActiveControl = BeginOdometerTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf EndOdometerTextBox.Text = Nothing Then
            ActiveControl = EndOdometerTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        ElseIf DaysTextBox.Text = Nothing Then
            ActiveControl = DaysTextBox
            MsgBox("Please Fill Out All Fields")
            Exit Function

        End If

        Try
            zipCheck = CInt(ZipCodeTextBox.Text)
        Catch ex As Exception
            ActiveControl = ZipCodeTextBox
            ZipCodeTextBox.Text = Nothing
            MsgBox("Zipcode must be a whole number")
            Exit Function
        End Try

        Try
            bOCheck = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            ActiveControl = BeginOdometerTextBox
            BeginOdometerTextBox.Text = Nothing
            MsgBox("Beginning Odometer Reading must be a whole number")
            Exit Function
        End Try

        Try
            eOCheck = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            ActiveControl = EndOdometerTextBox
            EndOdometerTextBox.Text = Nothing
            MsgBox("End Odometer Reading must be a whole number.")
            Exit Function
        End Try

        Try
            daysCheck = CInt(DaysTextBox.Text)
        Catch ex As Exception
            ActiveControl = DaysTextBox
            DaysTextBox.Text = Nothing
            MsgBox("Number of Days must be a whole number")
            Exit Function
        End Try

        If bOCheck > eOCheck Then
            ActiveControl = BeginOdometerTextBox
            BeginOdometerTextBox.Text = Nothing
            EndOdometerTextBox.Text = Nothing
            MsgBox("Beginning Odometer Reading can't be higher than end odometer reading")
            Exit Function
        End If

        If daysCheck > 45 Or daysCheck <= 0 Then
            ActiveControl = DaysTextBox
            DaysTextBox.Text = Nothing
            MsgBox("Number of days must be above zero and no more than 45")
            Exit Function
        End If
        eval = True
        Return eval
    End Function

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MsgBox("Number of Customers: " & customersTotal & vbNewLine & "Total Miles Driven: " & milesTotal & vbNewLine & "Total Charges: $" & chargesTotal)
        Clear()
    End Sub

    Sub Calculate()
        Dim numDays, dayCharge As Integer
        Dim milesCharge, totalMiles, discount, roundDiscount As Decimal

        numDays = CInt(DaysTextBox.Text)
        dayCharge = numDays * 15
        DayChargeTextBox.Text = "$" & CStr(dayCharge) & ".00"

        If KilometersradioButton.Checked = True Then
            totalMiles = CInt((0.62 * CInt(EndOdometerTextBox.Text)) - (0.62 * CInt(BeginOdometerTextBox.Text)))
            TotalMilesTextBox.Text = CStr((0.62 * CInt(EndOdometerTextBox.Text)) - (0.62 * CInt(BeginOdometerTextBox.Text))) & "mi"
        Else
            totalMiles = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
            TotalMilesTextBox.Text = CStr(CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) & "mi"
        End If

        If totalMiles > 500 Then
            milesCharge += 300 * 0.12D
            milesCharge += (totalMiles - 500) * 0.1D
        ElseIf totalMiles < 501 And totalMiles > 201 Then
            milesCharge = (totalMiles - 200) * 0.12D
        End If
        MileageChargeTextBox.Text = "$" & CStr(milesCharge)

        If AAAcheckbox.Checked = True And Seniorcheckbox.Checked = True Then
            discount = ((dayCharge + milesCharge) * 0.03D) + ((dayCharge + milesCharge) * 0.05D)
        ElseIf AAAcheckbox.Checked = True And Seniorcheckbox.Checked = False Then
            discount = (dayCharge + milesCharge) * 0.05D
        ElseIf Seniorcheckbox.Checked = True And AAAcheckbox.Checked = False Then
            discount = (dayCharge + milesCharge) * 0.03D
        End If

        roundDiscount = Math.Round(discount, 2, MidpointRounding.AwayFromZero)
        TotalDiscountTextBox.Text = "$" & CStr(roundDiscount)
        TotalChargeTextBox.Text = CStr((dayCharge + milesCharge) - roundDiscount)
        milesTotal += CInt(totalMiles)
        chargesTotal += (dayCharge + milesCharge) - roundDiscount
        SummaryButton.Enabled = True
    End Sub



    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim result As MsgBoxResult
        result = MsgBox("Are you sure you want to exit?", MsgBoxStyle.YesNo)
        If result = vbYes Then
            Me.Close()
        End If

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Clear()
    End Sub

    Sub Clear()
        NameTextBox.Text = Nothing
        AddressTextBox.Text = Nothing
        CityTextBox.Text = Nothing
        StateTextBox.Text = Nothing
        ZipCodeTextBox.Text = Nothing
        BeginOdometerTextBox.Text = Nothing
        EndOdometerTextBox.Text = Nothing
        DaysTextBox.Text = Nothing
        TotalMilesTextBox.Text = Nothing
        MileageChargeTextBox.Text = Nothing
        DayChargeTextBox.Text = Nothing
        TotalDiscountTextBox.Text = Nothing
        TotalChargeTextBox.Text = Nothing
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub
End Class
