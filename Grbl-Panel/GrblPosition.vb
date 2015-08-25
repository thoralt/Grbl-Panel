Imports GrblPanel.GrblIF

Partial Class GrblGui
    Private _IsProbing As Boolean = False
    Private _ProbeAxis As String = ""
    Private _ProbeDirection As String = ""

    Public Class GrblPosition
        Private _gui As GrblGui

        Public Sub New(ByRef gui As GrblGui)
            _gui = gui
            ' For Connected events
            AddHandler (GrblGui.Connected), AddressOf GrblConnected
            AddHandler (_gui.settings.GrblSettingsRetrieved), AddressOf GrblSettingsRetrieved
        End Sub

        Public Sub enablePosition(ByVal action As Boolean)
            _gui.gbPosition.Enabled = action
            If action = True Then
                _gui.grblPort.addRcvDelegate(AddressOf _gui.showGrblPositions)
            Else
                _gui.grblPort.deleteRcvDelegate(AddressOf _gui.showGrblPositions)
            End If
        End Sub

        Public Sub shutdown()
            ' Close up shop
            enablePosition(False)
        End Sub

        Private Sub GrblConnected(ByVal msg As String)     ' Handles GrblGui.Connected Event
            If msg = "Connected" Then

                ' We are connected to Grbl so highlight need for Homing Cycle
                _gui.btnHome.BackColor = Color.Crimson
            End If
        End Sub

        Private Sub GrblSettingsRetrieved()  ' Handles the named event
            ' Settings from Grbl are now available to query
            If _gui.settings.IsHomingEnabled = 1 Then
                ' Enable the Home Cycle button
                _gui.btnHome.Visible = True
            End If

        End Sub
    End Class


    Public Sub showGrblPositions(ByVal data As String)
        ' We come here from the recv_data thread so have to do this trick to cross threads
        ' (http://msdn.microsoft.com/en-ca/library/ms171728(v=vs.85).aspx)

        'Console.WriteLine("showGrblPosition: " + data)
        If Me.tbWorkX.InvokeRequired Then
            ' we need to cross thread this callback
            Dim ncb As New grblDataReceived(AddressOf Me.showGrblPositions)
            Me.BeginInvoke(ncb, New Object() {data})
            Return
        Else
            ' Show data in the Positions group (from our own thread)
            If (data.Contains("MPos:")) Then
                ' Lets display the values
                data = data.Remove(data.Length - 2, 2)   ' Remove the "> " at end
                Dim positions = Split(data, ":")
                Dim machPos = Split(positions(1), ",")
                Dim workPos = Split(positions(2), ",")

                tbMachX.Text = machPos(0).ToString
                tbMachY.Text = machPos(1).ToString
                tbMachZ.Text = machPos(2).ToString

                'Set same values into the repeater view on Offsets page
                tbOffSetsMachX.Text = machPos(0).ToString
                tbOffSetsMachY.Text = machPos(1).ToString
                tbOffSetsMachZ.Text = machPos(2).ToString

                tbWorkX.Text = workPos(0).ToString
                tbWorkY.Text = workPos(1).ToString
                tbWorkZ.Text = workPos(2).ToString

            ElseIf data.Contains("[PRB:") And _IsProbing Then
                ' handle incoming probe cycle position
                _IsProbing = False
                If data.Contains(":1]") Then ' only continue on successful probing
                    Dim toolDiameter As Double
                    Dim pullOff As Double
                    Dim plateThickness As Double

                    ' fetch and convert a few settings to double using fixed locale setting (force "." as decimal separator)
                    Try
                        toolDiameter = Double.Parse(My.Settings.ProbeToolDiameter, Globalization.CultureInfo.InvariantCulture)
                    Catch ex As Exception
                        MessageBox.Show("The setting for probe tool diameter is not a valid floating point number.",
                                        "Probing failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                    Try
                        pullOff = Double.Parse(My.Settings.ProbePullOff, Globalization.CultureInfo.InvariantCulture)
                    Catch ex As Exception
                        MessageBox.Show("The setting for probe pull off is not a valid floating point number.",
                                        "Probing failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                    Try
                        plateThickness = Double.Parse(My.Settings.ProbePlateThickness, Globalization.CultureInfo.InvariantCulture)
                    Catch ex As Exception
                        MessageBox.Show("The setting for probe plate thickness is not a valid floating point number.",
                                        "Probing failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try

                    ' calculate tool radius for X/Y, set it to 0 for Z
                    Dim toolRadius As Double = 0
                    If _ProbeAxis = "X" Or _ProbeAxis = "Y" Then
                        toolRadius = toolDiameter / 2.0
                    End If

                    ' calculate positive/negative pull off distance and tool center correction values
                    Dim pullOffDistance As Double = pullOff + plateThickness + toolRadius
                    Dim centerCorrection As Double = plateThickness + toolRadius
                    If _ProbeDirection = "+" Then
                        centerCorrection = -centerCorrection
                        pullOffDistance = -pullOffDistance
                    End If

                    ' probing was successful, set new location and pull off
                    gcode.sendGCodeLine("G90G10L20P0" + _ProbeAxis + centerCorrection.ToString())
                    gcode.sendGCodeLine("G0" + _ProbeAxis + pullOffDistance.ToString())
                Else
                    ' probing failed
                End If
            End If
        End If
    End Sub

    Private Sub btnPosition_Click(sender As Object, e As EventArgs) Handles btnWork0.Click, btnHome.Click, btnWorkSoftHome.Click, btnWorkSpclPosition.Click
        Dim b As Button = sender
        Select Case b.Tag
            Case "HomeCycle"
                ' Send Home command string
                gcode.sendGCodeLine("$H")
                btnHome.BackColor = Color.Transparent       ' In case it was crimson for inital connect
                tabCtlPosition.SelectedTab = tpWork         ' And show them the Work tab
            Case "Spcl Posn1"
                gcode.sendGCodeLine(tbSettingsSpclPosition1.Text)
            Case "Spcl Posn2"
                gcode.sendGCodeLine(tbSettingsSpclPosition2.Text)
            Case "ZeroXYZ"
                gcode.sendGCodeLine(tbSettingsZeroXYZCmd.Text)
        End Select

    End Sub

    ''' <summary>
    ''' Starts a probe cycle by switching to absolute coordinates, setting position 
    ''' of current axis to zero and sending the probe command for given axis and direction.
    ''' </summary>
    ''' <param name="axisAndDirection">Axis and direction, allowed values: X+, X-, Y+, Y-, Z+, Z-</param>
    ''' <remarks></remarks>
    Private Sub probeAxis(axisAndDirection As String)
        ' sanity checks
        axisAndDirection = axisAndDirection.ToUpper()
        If axisAndDirection.Length <> 2 Then Return
        If Not (axisAndDirection.EndsWith("-") Or
                axisAndDirection.EndsWith("+")) Or
           Not (axisAndDirection.StartsWith("X") Or
                axisAndDirection.StartsWith("Y") Or
                axisAndDirection.StartsWith("Z")) Then Return

        _ProbeAxis = axisAndDirection.Substring(0, 1)
        _ProbeDirection = axisAndDirection.Substring(1, 1)
        _IsProbing = True

        ' switch to absolute coordinates and set current axis position to zero
        gcode.sendGCodeLine("G90G10L20P0" + _ProbeAxis + "0")

        ' start probe cycle
        gcode.sendGCodeLine("G38.2" + _ProbeAxis + _ProbeDirection + My.Settings.ProbeDistance + "F" + My.Settings.ProbeFeed)

        ' probing ends when GRBL responds with a line starting with "[PRB:"

    End Sub

    Private Sub btnWorkXYZ0_Click(sender As Object, e As EventArgs) Handles btnWorkX0.Click, btnWorkY0.Click, btnWorkZ0.Click,
            btnProbeZn.Click, btnProbeYn.Click, btnProbeXn.Click, btnProbeYp.Click, btnProbeXp.Click
        Dim btn As Button = sender
        Select Case btn.Tag
            Case "X"
                gcode.sendGCodeLine(My.Settings.WorkX0Cmd)
            Case "Y"
                gcode.sendGCodeLine(My.Settings.WorkY0Cmd)
            Case "Z"
                gcode.sendGCodeLine(My.Settings.WorkZ0Cmd)
            Case "PX+"
                probeAxis("X+")
            Case "PX-"
                probeAxis("X-")
            Case "PY+"
                probeAxis("Y+")
            Case "PY-"
                probeAxis("Y-")
            Case "PZ-"
                probeAxis("Z-")
        End Select

    End Sub

End Class
