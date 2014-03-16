
Imports Contensive.BaseClasses
Imports Newtonsoft.Json

Namespace Contensive.addons.themeManager
    '
    Public Class ajaxJSONClass
        Inherits BaseClasses.AddonBaseClass
        '

        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Dim returnJSON As String = ""
            '
            Try
                Dim method As String = CP.Doc.GetProperty("method")
                Dim frameRqs As String = CP.Doc.GetProperty("frameRqs")
                Dim rightNow As Date = getRightNow(CP)
                Dim orderId As Integer
                Dim cs As CPCSBaseClass
                Dim csOptions As CPCSBaseClass
                'Dim jsonConvert As JsonConvert
                Dim shipOption As Dictionary(Of String, String)
                Dim optionList As New List(Of Dictionary(Of String, String))
                Dim orderShipWeight As Double = 0
                Dim orderShipZip As String = ""
                Dim orderShipCountry As String = ""
                Dim orderItemCharge As Double = 0
                Dim shipperErrorMessage As String = ""
                Dim shipCharge As Double
                '
                ' get form
                '
                Select Case method.ToLower()
                    Case "getshipoptions"
                        '
                        ' typical example of a method that returns data to the client
                        '
                        orderId = CP.Doc.GetInteger("orderId")
                        cs = CP.CSNew()
                        If cs.Open("orders", "id=" & orderId) Then
                            orderShipWeight = cs.GetNumber("shipweight")
                            orderShipZip = cs.GetText("shipZip")
                            orderShipCountry = cs.GetText("shipCountry")
                            orderItemCharge = cs.GetNumber("itemCharge")
                        End If
                        Call cs.Close()
                        '
                        csOptions = CP.CSNew()
                        Call csOptions.Open("order ship methods")
                        Do While csOptions.OK
                            shipperErrorMessage = ""
                            shipCharge = 9.99 'getShipCharge(CP, csOptions.GetInteger("id"), orderShipWeight, orderShipZip, orderShipCountry, orderItemCharge, False, 0, shipperErrorMessage)
                            If shipCharge >= 0 Then
                                shipOption = New Dictionary(Of String, String)
                                shipOption.Add("name", csOptions.GetText("name"))
                                shipOption.Add("charge", shipCharge.ToString())
                                optionList.Add(shipOption)
                            End If
                            Call csOptions.GoNext()
                        Loop
                        Call csOptions.Close()
                        '
                        returnJSON = JsonConvert.SerializeObject(optionList)
                    Case Else
                        '
                        '
                        '
                        returnJSON = "{}"
                End Select
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.aoAccountBilling.ajaxJSONClass.execute")
            End Try
            '
            ' assemble body
            '
            Return returnJSON
        End Function
    End Class
End Namespace
