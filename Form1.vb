Public Class frmTGAutomotive
    'Input Validation'
    Function ValidateInputs() As Boolean

        If IsNumeric(txtParts.Text) Then
            If CDec(txtParts.Text) < 1 Then
                MessageBox.Show("The cost of parts is too low.", "Parts Cost",
                MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtParts.Focus()
                Return False
            End If
        Else
            MessageBox.Show("The cost of parts is not a valid number.", "Parts Cost",
            MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            txtParts.Focus()
            Return False
        End If

        If IsNumeric(txtLabor.Text) Then
            If CDec(txtLabor.Text) < 1 Then
                MessageBox.Show("The cost of labor is too low.", "Parts Cost",
                MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                txtLabor.Focus()
                Return False
            End If
        Else
            MessageBox.Show("The cost of labor is not a valid number.", "Parts Cost",
            MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            txtLabor.Focus()
            Return False
        End If

        Return True
    End Function

    'Check which oil and lube services were selected'
    Function CalcOilLubeCharges() As Decimal
        Const oilChangeCost As Decimal = 36
        Const lubeJobCost As Decimal = 28
        Dim oilLubeTotal As Decimal

        If chkOilChange.Checked = True Then
            oilLubeTotal += oilChangeCost
        End If

        If chkLubeJob.Checked = True Then
            oilLubeTotal += lubeJobCost
        End If

        Return oilLubeTotal
    End Function

    'Check which flush services were selected'
    Function CalcFlushCharges() As Decimal
        Const radiatorCost As Decimal = 50
        Const transmissionCost As Decimal = 120
        Dim flushTotal As Decimal

        If chkRadiator.Checked = True Then
            flushTotal += radiatorCost
        End If

        If chkTransmission.Checked = True Then
            flushTotal += transmissionCost
        End If

        Return flushTotal
    End Function

    'Check which miscellaneous services were selected'
    Function CalcMiscCharges() As Decimal
        Const inspectionCost As Decimal = 15
        Const mufflerCost As Decimal = 200
        Const rotationCost As Decimal = 20
        Dim miscTotal As Decimal

        If chkInspection.Checked = True Then
            miscTotal += inspectionCost
        End If

        If chkReplaceMuffler.Checked = True Then
            miscTotal += mufflerCost
        End If

        If chkTireRotation.Checked = True Then
            miscTotal += rotationCost
        End If

        Return miscTotal
    End Function

    'Calculate total cost for everything'
    Sub CalculateTotalCharges()
        Const laborCharge As Decimal = 60
        Const taxRate As Decimal = 0.06D
        Dim parts As Decimal
        Dim labor As Decimal
        Dim laborTotal As Decimal
        Dim servicesAndLabor As Decimal
        Dim total As Decimal
        Dim taxTotal As Decimal

        If ValidateInputs() Then
            parts = CDec(txtParts.Text)
            labor = CDec(txtLabor.Text) / 60
        End If

        laborTotal = labor * laborCharge

        servicesAndLabor = CalcOilLubeCharges() + CalcFlushCharges() + CalcMiscCharges() + laborTotal
        taxTotal = taxRate * parts
        total = servicesAndLabor + parts + taxTotal

        txtSumServicesLabor.Text = servicesAndLabor.ToString("C")
        txtSumParts.Text = parts.ToString("C")
        txtSumTax.Text = taxTotal.ToString("C")
        txtSumTotal.Text = total.ToString("C")

    End Sub

    'Clear oil and lube'
    Sub ClearOilLube()
        chkOilChange.Checked = False
        chkLubeJob.Checked = False

    End Sub

    'Clear flushes'
    Sub ClearFlushes()
        chkRadiator.Checked = False
        chkTransmission.Checked = False

    End Sub

    'Clear miscellaneous'
    Sub ClearMisc()
        chkInspection.Checked = False
        chkReplaceMuffler.Checked = False
        chkTireRotation.Checked = False

    End Sub

    'Clear textboxes'
    Sub ClearOther()
        txtParts.Text = ""
        txtLabor.Text = ""
        txtSumServicesLabor.Text = ""
        txtSumParts.Text = ""
        txtSumTax.Text = ""
        txtSumTotal.Text = ""

    End Sub

    'Calculate button clicked'
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        CalculateTotalCharges()

    End Sub

    'Clear button clicked'
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearOilLube()
        ClearFlushes()
        ClearMisc()
        ClearOther()

    End Sub

    'Exit button clicked'
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()

    End Sub

End Class