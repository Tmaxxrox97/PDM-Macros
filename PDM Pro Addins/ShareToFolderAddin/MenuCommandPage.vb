Imports EPDM.Interop.epdm
Public Class MenuCommandPage
    Public Sub LoadData(ByRef poCmd As EdmCmd)

        'Get the property interface used to access the framework
        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        'Populate the edit box from a variable
        TextBox1.Text = props.GetValEx("MenuCommand")
        TextBox2.Text = props.GetValEx("HelpText")
        TextBox3.Text = props.GetValEx("SharedPath")
        TextBox4.Text = props.GetValEx("SharedExtension")
        TextBox5.Text = props.GetValEx("SubFolder")
        If props.GetValEx("Checked") = "True" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

    End Sub

    Public Sub StoreData(ByRef poCmd As EdmCmd)

        'Get the property interface used to access the framework
        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        'Make sure the user has typed a value in the edit box
        If Not CheckBox1.Checked Then
            props.SetValEx("MenuCommand", "")
            props.SetValEx("HelpText", "")
            props.SetValEx("Checked", "False")
        Else
            props.SetValEx("MenuCommand", TextBox1.Text)
            props.SetValEx("HelpText", TextBox2.Text)
            props.SetValEx("Checked", "True")
        End If

        props.SetValEx("SharedPath", TextBox3.Text)
        props.SetValEx("SharedExtension", TextBox4.Text)
        props.SetValEx("SubFolder", TextBox5.Text)

    End Sub

End Class
