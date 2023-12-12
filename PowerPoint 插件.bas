Sub AddCustomButton()
    ' 在 PowerPoint 中添加自定义按钮
    Dim ribbonXml As String
    ribbonXml = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & _
                "  <mso:ribbon>" & _
                "    <mso:qat>" & _
                "      <mso:sharedControls>" & _
                "        <mso:control idQ='mso:SlideShowStart' visible='false' />" & _
                "      </mso:sharedControls>" & _
                "    </mso:qat>" & _
                "    <mso:tabs>" & _
                "      <mso:tab idQ='mso:TabAddIns'>" & _
                "        <mso:group idQ='mso:Group1' label='Custom Group'>" & _
                "          <mso:button idQ='mso:ButtonID' label='Click Me' imageMso='HappyFace' onAction='ShowMessageBox' />" & _
                "        </mso:group>" & _
                "      </mso:tab>" & _
                "    </mso:tabs>" & _
                "  </mso:ribbon>" & _
                "</mso:customUI>"

    ' 加载自定义 UI
    Application.AddCustomUI ("MyRibbonUI"), ribbonXml
End Sub

Sub ShowMessageBox(control As IRibbonControl)
    ' 按钮点击时显示消息框
    MsgBox "Hello, PowerPoint!"
End Sub
