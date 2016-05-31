'Option Explicit On
'Option Strict On
Imports AFRAppWCFHost.AFRService
Imports System.ServiceModel
Imports InterfaceMapper


Namespace InterfaceQB

    Public Module InterfaceQB


        Dim connString As String = "Computer Name=Home;Company Data=Ralph's Pretty Good"
        Dim filePath As String
        Dim Msg As New Mesgbox
        Const cAppID = "1"
        Const cAppName = "Ad Free Receipts Capture"
        Const cRequestID = "1"
        Dim serverName As String
        Dim companyName As String
        Dim versionName As String
        Dim xmlMajorVersion As Short
        Dim xmlMinorVersion As Short
        Dim SaveTime As New DateTime
        Dim requestXML As String
        Dim responseXML As String
        Dim CurrentTime As DateTime
        <STAThread()>
        Public Sub Main()

            StartProg()

        End Sub
        Sub StartProg()

            SaveTime = Now()
            Do
                CurrentTime = GetTime()
                QSR()
            Loop

        End Sub
        Private Sub QSR()           'Quickbooks Sales Receipts
            SaveTime = Now()
            Dim sessionBegun As Boolean = False
            Dim connectionOpen As Boolean = False
            Dim SessionManager As New QBPOSSessionManager
            SessionManager = Nothing
            Dim RequestMsgSet As IMsgSetRequest
            Dim isInSession As Boolean

            'Dim CurrentTime As String = GetTime()
            Try
                isInSession = False

                If (sessionBegun) Then
                    SessionManager.EndSession()
                    sessionBegun = False
                End If
                If (connectionOpen) Then
                    SessionManager.CloseConnection()
                    connectionOpen = False
                End If

                SessionManager = New QBPOSSessionManager
                SessionManager.OpenConnection(cAppID, cAppName)
                connectionOpen = True

                'Step-C: BeginSession to QBPOS
                SessionManager.BeginSession(connString)
                sessionBegun = True

                xmlMajorVersion = 3
                xmlMinorVersion = 0
                RequestMsgSet = SessionManager.CreateMsgSetRequest(xmlMajorVersion, xmlMinorVersion)
                RequestMsgSet.Attributes.OnError = ENRqOnError.roeContinue
                If (RequestMsgSet Is Nothing) Then
                    Mesgbox.MesgBox("Invalid request message set object",, "AFR QB Capture")
                    SessionManager.EndSession()
                    SessionManager.CloseConnection()
                    Exit Sub
                End If
                Dim SRQuery As ISalesReceiptQuery
                SRQuery = RequestMsgSet.AppendSalesReceiptQueryRq
                SRQuery.ORTimeCreatedFilters.TimeCreatedRangeFilter.FromTimeCreated.SetValue(DateTime.Parse("2016-05-01 01:00:00"), False)
                'SRQuery.ORTimeCreatedFilters.TimeCreatedRangeFilter.FromTimeCreated.SetValue(DateTime.Parse(CurrentTime), False)
                SaveTime = Now()
                SRQuery.ORTimeCreatedFilters.TimeCreatedRangeFilter.ToTimeCreated.SetValue(DateTime.Parse(Now().ToString), False)
                'Connect to QuickBooks and begin a session

                'SessionManager.OpenConnection("", "Ad Free Receipts Capture")
                'connectionOpen = True
                'SessionManager.BeginSession("")       ',ENOpenMode.omDontCare
                'sessionBegun = True

                'Send the request and get the response from QuickBooks
                Dim responseMsgSet As IMsgSetResponse
                responseMsgSet = SessionManager.DoRequests(RequestMsgSet)
                'End the session and close the connection to QuickBooks
                SessionManager.EndSession()
                sessionBegun = False
                SessionManager.CloseConnection()
                connectionOpen = False
                WalkSalesReceiptQueryRs(responseMsgSet)

            Catch e As Exception
                If Err.Number = &H80040416 Then
                    Mesgbox.MesgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.",, "AFR QB Capture")
                    SessionManager.CloseConnection()

                ElseIf Err.Number = &H80040422 Then
                    Mesgbox.MesgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.",, "AFR QB Capture")
                    SessionManager.CloseConnection()
                Else
                    Mesgbox.MesgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, "AFR QB Capture")
                End If

                Mesgbox.MesgBox(e.Message,, "AFR QB Capture")
                If (sessionBegun) Then
                    SessionManager.EndSession()
                    sessionBegun = False
                End If
                If (connectionOpen) Then
                    SessionManager.CloseConnection()
                    connectionOpen = False
                End If
            End Try
        End Sub
        Private Function GetTime() As DateTime

            System.Threading.Thread.Sleep(6000)
            'System.Threading.Thread.Sleep(60000)
            Return SaveTime
        End Function

        Private Sub parseConnString()
            Dim pos1, pos2, pos3 As Integer
            'Example of a connection string:
            'connString = "Computer Name=mtvl04a200337;Company Data=My Company;Version=4"
            pos1 = InStr(connString, "Computer Name=")
            pos2 = InStr(connString, "Company Data=")
            pos3 = InStr(connString, "Version=")
            If pos1 > 0 Then
                If pos2 > 0 Then
                    serverName = Mid(connString, pos1 + 14, pos2 - pos1 - 15)
                Else
                    serverName = Mid(connString, pos1 + 14)
                End If
            End If
            If pos2 > 0 Then
                If pos3 > 0 Then
                    companyName = Mid(connString, pos2 + 13, pos3 - pos2 - 14)
                Else
                    companyName = Mid(connString, pos2 + 13)
                End If
            End If
            If pos3 > 0 Then
                versionName = Mid(connString, pos3 + 8)
            End If
        End Sub

        Public Sub WalkSalesReceiptQueryRs(responseMsgSet As IMsgSetResponse)

            If (responseMsgSet Is Nothing) Then

                Exit Sub

            End If

            Dim responseList As IResponseList
            responseList = responseMsgSet.ResponseList
            If (responseList Is Nothing) Then

                Exit Sub

            End If

            'if we sent only one request, there is only one response, we'll walk the list for this sample
            For j = 0 To responseList.Count - 1

                Dim response As IResponse
                response = responseList.GetAt(j)        'was i
                'check the status code of the response, 0=ok, >0 is warning
                If (response.StatusCode >= 0) Then

                    'the request-specific response Is in the details, make sure we have some
                    If (Not response.Detail Is Nothing) Then

                        'make sure the response Is the type we're expecting
                        Dim responseType As ENResponseType
                        responseType = CType(response.Type.GetValue(), ENResponseType)
                        If (responseType = ENResponseType.rtSalesReceiptQueryRs) Then

                            '//upcast to more specific type here, this is safe because we checked with response.Type check above
                            Dim SalesReceiptRet As Interop.qbposfc3.ISalesReceiptRet
                            'IID B4BF0585-E451-4C3A-85B8-1 6F0 B6C54C

                            SalesReceiptRet = CType(response.Detail, Interop.qbposfc3.ISalesReceiptRet)

                            WalkSalesReceiptRet(SalesReceiptRet)

                        End If
                    End If
                End If
            Next j
        End Sub



        Public Sub WalkSalesReceiptRet(SalesReceiptRet As Interop.qbposfc3.ISalesReceiptRet)

            If (SalesReceiptRet Is Nothing) Then

                Exit Sub

            End If

            ' Send the whole thing



            Dim myBinding As New BasicHttpBinding
            Dim myEndpoint As New EndpointAddress("http://localhost:8733/InterfaceMapper.Mapper.svc")

            Dim myChannelFactory As ChannelFactory(Of IMapper) = New ChannelFactory(Of IMapper)(myBinding, myEndpoint)
            Dim wcfClient As IMapper = myChannelFactory.CreateChannel()
            Dim SendTrans As String = ""
            'Get value of TxnID
            If (Not SalesReceiptRet.TxnID Is Nothing) Then

                Dim TxnID22 As String
                TxnID22 = SalesReceiptRet.TxnID.GetValue()
                SendTrans = TxnID22 + "|"
            End If
            'Get value of TimeCreated
            If (Not SalesReceiptRet.TimeCreated Is Nothing) Then

                Dim TimeCreated23 As DateTime
                TimeCreated23 = SalesReceiptRet.TimeCreated.GetValue()
                SendTrans = SendTrans + TimeCreated23 + "|"
            End If
            'Get value of TimeModified
            If (Not SalesReceiptRet.TimeModified Is Nothing) Then

                Dim TimeModified24 As DateTime
                TimeModified24 = SalesReceiptRet.TimeModified.GetValue()
                SendTrans = SendTrans + TimeModified24 + "|"
            End If
            'Get value of Associate
            If (Not SalesReceiptRet.Associate Is Nothing) Then

                Dim Associate25 As String
                Associate25 = SalesReceiptRet.Associate.GetValue()
                SendTrans = SendTrans + Associate25 + "|"
            End If
            'Get value of Cashier
            If (Not SalesReceiptRet.Cashier Is Nothing) Then

                Dim Cashier26 As String
                Cashier26 = SalesReceiptRet.Cashier.GetValue()
                SendTrans = SendTrans + Cashier26 + "|"
            End If

            'Dim wcfreturn As Boolean = wcfClient.SendData(SalesReceiptRet)
            Dim wcfreturn As Boolean = wcfClient.SendData(SendTrans)
            'wcfClient.SendData(SalesReceiptRet)

        End Sub





        '    'Go through all the elements of ISalesReceiptRetList

        '    'Get value of Comments
        '    If (Not SalesReceiptRet.Comments Is Nothing) Then

        '        Dim Comments27 As String
        '        Comments27 = SalesReceiptRet.Comments.GetValue()

        '    End If
        '    'Get value of CustomerListID
        '    If (Not SalesReceiptRet.CustomerListID Is Nothing) Then

        '        Dim CustomerListID28 As String
        '        CustomerListID28 = SalesReceiptRet.CustomerListID.GetValue()

        '    End If
        '    'Get value of Discount
        '    If (Not SalesReceiptRet.Discount Is Nothing) Then

        '        Dim Discount29 As Double
        '        Discount29 = SalesReceiptRet.Discount.GetValue()

        '    End If
        '    'Get value of DiscountPercent
        '    If (Not SalesReceiptRet.DiscountPercent Is Nothing) Then

        '        Dim DiscountPercent30 As IQBFloatType
        '        DiscountPercent30 = SalesReceiptRet.DiscountPercent.GetValue()

        '    End If
        '    'Get value of HistoryDocStatus
        '    If (Not SalesReceiptRet.HistoryDocStatus Is Nothing) Then

        '        Dim HistoryDocStatus31 As ENHistoryDocStatus
        '        HistoryDocStatus31 = SalesReceiptRet.HistoryDocStatus.GetValue()

        '    End If
        '    'Get value of ItemsCount
        '    If (Not SalesReceiptRet.ItemsCount Is Nothing) Then

        '        Dim ItemsCount32 As Integer
        '        ItemsCount32 = SalesReceiptRet.ItemsCount.GetValue()

        '    End If
        '    'Get value of PriceLevelNumber
        '    If (Not SalesReceiptRet.PriceLevelNumber Is Nothing) Then

        '        Dim PriceLevelNumber33 As ENPriceLevelNumber
        '        PriceLevelNumber33 = SalesReceiptRet.PriceLevelNumber.GetValue()

        '    End If
        '    'Get value of PromoCode
        '    If (Not SalesReceiptRet.PromoCode Is Nothing) Then

        '        Dim PromoCode34 As String
        '        PromoCode34 = SalesReceiptRet.PromoCode.GetValue()

        '    End If
        '    'Get value of QuickBooksFlag
        '    If (Not SalesReceiptRet.QuickBooksFlag Is Nothing) Then

        '        Dim QuickBooksFlag35 As ENQuickBooksFlag
        '        QuickBooksFlag35 = SalesReceiptRet.QuickBooksFlag.GetValue()

        '    End If
        '    'Get value of SalesOrderTxnID
        '    If (Not SalesReceiptRet.SalesOrderTxnID Is Nothing) Then

        '        Dim SalesOrderTxnID36 As String
        '        SalesOrderTxnID36 = SalesReceiptRet.SalesOrderTxnID.GetValue()

        '    End If
        '    'Get value of SalesReceiptNumber
        '    If (Not SalesReceiptRet.SalesReceiptNumber Is Nothing) Then

        '        Dim SalesReceiptNumber37 As Integer
        '        SalesReceiptNumber37 = SalesReceiptRet.SalesReceiptNumber.GetValue()

        '    End If
        '    'Get value of SalesReceiptType
        '    If (Not SalesReceiptRet.SalesReceiptType Is Nothing) Then

        '        Dim SalesReceiptType38 As ENSalesReceiptType
        '        SalesReceiptType38 = SalesReceiptRet.SalesReceiptType.GetValue()

        '    End If
        '    'Get value of ShipDate
        '    If (Not SalesReceiptRet.ShipDate Is Nothing) Then

        '        Dim ShipDate39 As DateTime
        '        ShipDate39 = SalesReceiptRet.ShipDate.GetValue()

        '    End If
        '    'Get value of StoreExchangeStatus
        '    If (Not SalesReceiptRet.StoreExchangeStatus Is Nothing) Then

        '        Dim StoreExchangeStatus40 As ENStoreExchangeStatus
        '        StoreExchangeStatus40 = SalesReceiptRet.StoreExchangeStatus.GetValue()

        '    End If
        '    'Get value of StoreNumber
        '    If (Not SalesReceiptRet.StoreNumber Is Nothing) Then

        '        Dim StoreNumber41 As Integer
        '        StoreNumber41 = SalesReceiptRet.StoreNumber.GetValue()

        '    End If
        '    'Get value of Subtotal
        '    If (Not SalesReceiptRet.Subtotal Is Nothing) Then

        '        Dim Subtotal42 As Double
        '        Subtotal42 = SalesReceiptRet.Subtotal.GetValue()

        '    End If
        '    'Get value of TaxAmount
        '    If (Not SalesReceiptRet.TaxAmount Is Nothing) Then

        '        Dim TaxAmount43 As Double
        '        TaxAmount43 = SalesReceiptRet.TaxAmount.GetValue()

        '    End If
        '    'Get value of TaxCategory
        '    If (Not SalesReceiptRet.TaxCategory Is Nothing) Then

        '        Dim TaxCategory44 As String
        '        TaxCategory44 = SalesReceiptRet.TaxCategory.GetValue()

        '    End If
        '    'Get value of TaxPercentage
        '    If (Not SalesReceiptRet.TaxPercentage Is Nothing) Then

        '        Dim TaxPercentage45 As IQBFloatType
        '        TaxPercentage45 = SalesReceiptRet.TaxPercentage.GetValue()

        '    End If
        '    'Get value of TenderType
        '    If (Not SalesReceiptRet.TenderType Is Nothing) Then

        '        Dim TenderType46 As ENTenderType
        '        TenderType46 = SalesReceiptRet.TenderType.GetValue()

        '    End If
        '    'Get value of TipReceiver
        '    If (Not SalesReceiptRet.TipReceiver Is Nothing) Then

        '        Dim TipReceiver47 As String
        '        TipReceiver47 = SalesReceiptRet.TipReceiver.GetValue()

        '    End If
        '    'Get value of Total
        '    If (Not SalesReceiptRet.Total Is Nothing) Then

        '        Dim Total48 As Double
        '        Total48 = SalesReceiptRet.Total.GetValue()

        '    End If
        '    'Get value of TrackingNumber
        '    If (Not SalesReceiptRet.TrackingNumber Is Nothing) Then

        '        Dim TrackingNumber49 As String
        '        TrackingNumber49 = SalesReceiptRet.TrackingNumber.GetValue()

        '    End If
        '    'Get value of TxnDate
        '    If (Not SalesReceiptRet.TxnDate Is Nothing) Then

        '        Dim TxnDate50 As DateTime
        '        TxnDate50 = SalesReceiptRet.TxnDate.GetValue()

        '    End If
        '    'Get value of TxnState
        '    If (Not SalesReceiptRet.TxnState Is Nothing) Then

        '        Dim TxnState51 As ENTxnState
        '        TxnState51 = SalesReceiptRet.TxnState.GetValue()

        '    End If
        '    'Get value of Workstation
        '    If (Not SalesReceiptRet.Workstation Is Nothing) Then

        '        Dim Workstation52 As Integer
        '        Workstation52 = SalesReceiptRet.Workstation.GetValue()

        '    End If
        '    If (Not SalesReceiptRet.BillingInformation Is Nothing) Then

        '        'Get value of City
        '        If (Not SalesReceiptRet.BillingInformation.City Is Nothing) Then

        '            Dim City53 As String
        '            City53 = SalesReceiptRet.BillingInformation.City.GetValue()

        '        End If
        '        'Get value of CompanyName
        '        If (Not SalesReceiptRet.BillingInformation.CompanyName Is Nothing) Then

        '            Dim CompanyName54 As String
        '            CompanyName54 = SalesReceiptRet.BillingInformation.CompanyName.GetValue()

        '        End If
        '        'Get value of Country
        '        If (Not SalesReceiptRet.BillingInformation.Country Is Nothing) Then

        '            Dim Country55 As String
        '            Country55 = SalesReceiptRet.BillingInformation.Country.GetValue()

        '        End If
        '        'Get value of FirstName
        '        If (Not SalesReceiptRet.BillingInformation.FirstName Is Nothing) Then

        '            Dim FirstName56 As String
        '            FirstName56 = SalesReceiptRet.BillingInformation.FirstName.GetValue()

        '        End If
        '        'Get value of LastName
        '        If (Not SalesReceiptRet.BillingInformation.LastName Is Nothing) Then

        '            Dim LastName57 As String
        '            LastName57 = SalesReceiptRet.BillingInformation.LastName.GetValue()

        '        End If
        '        'Get value of Phone
        '        If (Not SalesReceiptRet.BillingInformation.Phone Is Nothing) Then

        '            Dim Phone58 As String
        '            Phone58 = SalesReceiptRet.BillingInformation.Phone.GetValue()

        '        End If
        '        'Get value of Phone2
        '        If (Not SalesReceiptRet.BillingInformation.Phone2 Is Nothing) Then

        '            Dim Phone259 As String
        '            Phone259 = SalesReceiptRet.BillingInformation.Phone2.GetValue()

        '        End If
        '        'Get value of Phone3
        '        If (Not SalesReceiptRet.BillingInformation.Phone3 Is Nothing) Then

        '            Dim Phone360 As String
        '            Phone360 = SalesReceiptRet.BillingInformation.Phone3.GetValue()

        '        End If
        '        'Get value of Phone4
        '        If (Not SalesReceiptRet.BillingInformation.Phone4 Is Nothing) Then

        '            Dim Phone461 As String
        '            Phone461 = SalesReceiptRet.BillingInformation.Phone4.GetValue()

        '        End If
        '        'Get value of PostalCode
        '        If (Not SalesReceiptRet.BillingInformation.PostalCode Is Nothing) Then

        '            Dim PostalCode62 As String
        '            PostalCode62 = SalesReceiptRet.BillingInformation.PostalCode.GetValue()

        '        End If
        '        'Get value of Salutation
        '        If (Not SalesReceiptRet.BillingInformation.Salutation Is Nothing) Then

        '            Dim Salutation63 As String
        '            Salutation63 = SalesReceiptRet.BillingInformation.Salutation.GetValue()

        '        End If
        '        'Get value of State
        '        If (Not SalesReceiptRet.BillingInformation.State Is Nothing) Then

        '            Dim State64 As String
        '            State64 = SalesReceiptRet.BillingInformation.State.GetValue()

        '        End If
        '        'Get value of Street
        '        If (Not SalesReceiptRet.BillingInformation.Street Is Nothing) Then

        '            Dim Street65 As String
        '            Street65 = SalesReceiptRet.BillingInformation.Street.GetValue()

        '        End If
        '        'Get value of Street2
        '        If (Not SalesReceiptRet.BillingInformation.Street2 Is Nothing) Then

        '            Dim Street266 As String
        '            Street266 = SalesReceiptRet.BillingInformation.Street2.GetValue()

        '        End If
        '        'Get value of WebNumber
        '        If (Not SalesReceiptRet.BillingInformation.WebNumber Is Nothing) Then

        '            Dim WebNumber67 As String
        '            WebNumber67 = SalesReceiptRet.BillingInformation.WebNumber.GetValue()

        '        End If
        '    End If
        '    If (Not SalesReceiptRet.ShippingInformation Is Nothing) Then

        '        'Get value of AddressName
        '        If (Not SalesReceiptRet.ShippingInformation.AddressName Is Nothing) Then

        '            Dim AddressName68 As String
        '            AddressName68 = SalesReceiptRet.ShippingInformation.AddressName.GetValue()

        '        End If
        '        'Get value of City
        '        If (Not SalesReceiptRet.ShippingInformation.City Is Nothing) Then

        '            Dim City69 As String
        '            City69 = SalesReceiptRet.ShippingInformation.City.GetValue()

        '        End If
        '        'Get value of CompanyName
        '        If (Not SalesReceiptRet.ShippingInformation.CompanyName Is Nothing) Then

        '            Dim CompanyName70 As String
        '            CompanyName70 = SalesReceiptRet.ShippingInformation.CompanyName.GetValue()

        '        End If
        '        'Get value of Country
        '        If (Not SalesReceiptRet.ShippingInformation.Country Is Nothing) Then

        '            Dim Country71 As String
        '            Country71 = SalesReceiptRet.ShippingInformation.Country.GetValue()

        '        End If
        '        'Get value of FullName
        '        If (Not SalesReceiptRet.ShippingInformation.FullName Is Nothing) Then

        '            Dim FullName72 As String
        '            FullName72 = SalesReceiptRet.ShippingInformation.FullName.GetValue()

        '        End If
        '        'Get value of Phone
        '        If (Not SalesReceiptRet.ShippingInformation.Phone Is Nothing) Then

        '            Dim Phone73 As String
        '            Phone73 = SalesReceiptRet.ShippingInformation.Phone.GetValue()

        '        End If
        '        'Get value of Phone2
        '        If (Not SalesReceiptRet.ShippingInformation.Phone2 Is Nothing) Then

        '            Dim Phone274 As String
        '            Phone274 = SalesReceiptRet.ShippingInformation.Phone2.GetValue()

        '        End If
        '        'Get value of Phone3
        '        If (Not SalesReceiptRet.ShippingInformation.Phone3 Is Nothing) Then

        '            Dim Phone375 As String
        '            Phone375 = SalesReceiptRet.ShippingInformation.Phone3.GetValue()

        '        End If
        '        'Get value of Phone4
        '        If (Not SalesReceiptRet.ShippingInformation.Phone4 Is Nothing) Then

        '            Dim Phone476 As String
        '            Phone476 = SalesReceiptRet.ShippingInformation.Phone4.GetValue()

        '        End If
        '        'Get value of PostalCode
        '        If (Not SalesReceiptRet.ShippingInformation.PostalCode Is Nothing) Then

        '            Dim PostalCode77 As String
        '            PostalCode77 = SalesReceiptRet.ShippingInformation.PostalCode.GetValue()

        '        End If
        '        'Get value of ShipBy
        '        If (Not SalesReceiptRet.ShippingInformation.ShipBy Is Nothing) Then

        '            Dim ShipBy78 As String
        '            ShipBy78 = SalesReceiptRet.ShippingInformation.ShipBy.GetValue()

        '        End If
        '        'Get value of Shipping
        '        If (Not SalesReceiptRet.ShippingInformation.Shipping Is Nothing) Then

        '            Dim Shipping79 As Double
        '            Shipping79 = SalesReceiptRet.ShippingInformation.Shipping.GetValue()

        '        End If
        '        'Get value of State
        '        If (Not SalesReceiptRet.ShippingInformation.State Is Nothing) Then

        '            Dim State80 As String
        '            State80 = SalesReceiptRet.ShippingInformation.State.GetValue()

        '        End If
        '        'Get value of Street
        '        If (Not SalesReceiptRet.ShippingInformation.Street Is Nothing) Then

        '            Dim Street81 As String
        '            Street81 = SalesReceiptRet.ShippingInformation.Street.GetValue()

        '        End If
        '        'Get value of Street2
        '        If (Not SalesReceiptRet.ShippingInformation.Street2 Is Nothing) Then

        '            Dim Street282 As String
        '            Street282 = SalesReceiptRet.ShippingInformation.Street2.GetValue()

        '        End If
        '    End If
        '    If (Not SalesReceiptRet.SalesReceiptItemRetList Is Nothing) Then

        '        Dim i83 As Integer
        '            For i83 = 0 To SalesReceiptRet.SalesReceiptItemRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim SalesReceiptItemRet As ISalesReceiptItemRet
        '                SalesReceiptItemRet = .SalesReceiptItemRetList.GetAt(i83)
        '                'Get value of ListID
        '                If (Not .SalesReceiptItemRetListID Is Nothing) Then

        '                    Dim ListID84 As String
        '                    ListID84 = .SalesReceiptItemRetListID.GetValue()

        '                End If
        '                'Get value of ALU
        '                If (Not .SalesReceiptItemRetALU Is Nothing) Then

        '                    Dim ALU85 As String
        '                    ALU85 = .SalesReceiptItemRetALU.GetValue()

        '                End If
        '                'Get value of Associate
        '                If (Not .SalesReceiptItemRetAssociate Is Nothing) Then

        '                    Dim Associate86 As String
        '                    Associate86 = .SalesReceiptItemRetAssociate.GetValue()

        '                End If
        '                'Get value of Attribute
        '                If (Not .SalesReceiptItemRetAttribute Is Nothing) Then

        '                    Dim Attribute87 As String
        '                    Attribute87 = .SalesReceiptItemRetAttribute.GetValue()

        '                End If
        '                'Get value of Commission
        '                If (Not .SalesReceiptItemRetCommission Is Nothing) Then

        '                    Dim Commission88 As Double
        '                    Commission88 = .SalesReceiptItemRetCommission.GetValue()

        '                End If
        '                'Get value of Cost
        '                If (Not SalesReceiptRet.SalesReceiptItemRetCost Is Nothing) Then

        '                    Dim Cost89 As Double
        '                    Cost89 = .SalesReceiptItemRetCost.GetValue()

        '                End If
        '                'Get value of Desc1
        '                If (Not .SalesReceiptItemRetDesc1 Is Nothing) Then

        '                    Dim Desc190 As String
        '                    Desc190 = .SalesReceiptItemRetDesc1.GetValue()

        '                End If
        '                'Get value of Desc2
        '                If (Not .SalesReceiptItemRetDesc2 Is Nothing) Then

        '                    Dim Desc291 As String
        '                    Desc291 = .SalesReceiptItemRetDesc2.GetValue()

        '                End If
        '                'Get value of Discount
        '                If (Not .SalesReceiptItemRetDiscount Is Nothing) Then

        '                    Dim Discount92 As Double
        '                    Discount92 = .SalesReceiptItemRetDiscount.GetValue()

        '                End If
        '                'Get value of DiscountPercent
        '                If (Not .SalesReceiptItemRetDiscountPercent Is Nothing) Then

        '                    Dim DiscountPercent93 As IQBFloatType
        '                    DiscountPercent93 = .SalesReceiptItemRetDiscountPercent.GetValue()

        '                End If
        '                'Get value of DiscountType
        '                If (Not .SalesReceiptItemRetDiscountType Is Nothing) Then

        '                    Dim DiscountType94 As String
        '                    DiscountType94 = .SalesReceiptItemRetDiscountType.GetValue()

        '                End If
        '                'Get value of DiscountSource
        '                If (Not .SalesReceiptItemRetDiscountSource Is Nothing) Then

        '                    Dim DiscountSource95 As ENDiscountSource
        '                    DiscountSource95 = .SalesReceiptItemRetDiscountSource.GetValue()

        '                End If
        '                'Get value of ExtendedPrice
        '                If (Not .SalesReceiptItemRetExtendedPrice Is Nothing) Then

        '                    Dim ExtendedPrice96 As Double
        '                    ExtendedPrice96 = .SalesReceiptItemRetExtendedPrice.GetValue()

        '                End If
        '                'Get value of ExtendedTax
        '                If (Not .SalesReceiptItemRetExtendedTax Is Nothing) Then

        '                    Dim ExtendedTax97 As Double
        '                    ExtendedTax97 = .SalesReceiptItemRetExtendedTax.GetValue()

        '                End If
        '                'Get value of ItemNumber
        '                If (Not .SalesReceiptItemRetItemNumber Is Nothing) Then

        '                    Dim ItemNumber98 As Integer
        '                    ItemNumber98 = .SalesReceiptItemRetItemNumber.GetValue()

        '                End If
        '                'Get value of NumberOfBaseUnits
        '                If (Not .SalesReceiptItemRetNumberOfBaseUnits Is Nothing) Then

        '                    Dim NumberOfBaseUnits99 As Integer
        '                    NumberOfBaseUnits99 = .SalesReceiptItemRetNumberOfBaseUnits.GetValue()

        '                End If
        '                'Get value of Price
        '                If (Not .SalesReceiptItemRetPrice Is Nothing) Then

        '                    Dim Price100 As Double
        '                    Price100 = .SalesReceiptItemRetPrice.GetValue()

        '                End If
        '                'Get value of PriceLevelNumber
        '                If (Not .SalesReceiptItemRetPriceLevelNumber Is Nothing) Then

        '                    Dim PriceLevelNumber101 As ENPriceLevelNumber
        '                    PriceLevelNumber101 = .SalesReceiptItemRetPriceLevelNumber.GetValue()

        '                End If
        '                'Get value of Qty
        '                If (Not .SalesReceiptItemRetQty Is Nothing) Then

        '                    Dim Qty102 As Integer
        '                    Qty102 = .SalesReceiptItemRetQty.GetValue()

        '                End If
        '                'Get value of SerialNumber
        '                If (Not .SalesReceiptItemRetSerialNumber Is Nothing) Then

        '                    Dim SerialNumber103 As String
        '                    SerialNumber103 = .SalesReceiptItemRetSerialNumber.GetValue()

        '                End If
        '                'Get value of Size
        '                If (Not .SalesReceiptItemRetSize Is Nothing) Then

        '                    Dim Size104 As String
        '                    Size104 = .SalesReceiptItemRetSize.GetValue()

        '                End If
        '                'Get value of TaxAmount
        '                If (Not .SalesReceiptItemRetTaxAmount Is Nothing) Then

        '                    Dim TaxAmount105 As Double
        '                    TaxAmount105 = .SalesReceiptItemRetTaxAmount.GetValue()

        '                End If
        '                'Get value of TaxCode
        '                If (Not .SalesReceiptItemRetTaxCode Is Nothing) Then

        '                    Dim TaxCode106 As String
        '                    TaxCode106 = .SalesReceiptItemRetTaxCode.GetValue()

        '                End If
        '                'Get value of TaxPercentage
        '                If (Not .SalesReceiptItemRetTaxPercentage Is Nothing) Then

        '                    Dim TaxPercentage107 As IQBFloatType
        '                    TaxPercentage107 = .SalesReceiptItemRetTaxPercentage.GetValue()

        '                End If
        '                'Get value of UnitOfMeasure
        '                If (Not .SalesReceiptItemRetUnitOfMeasure Is Nothing) Then

        '                    Dim UnitOfMeasure108 As String
        '                    UnitOfMeasure108 = .SalesReceiptItemRetUnitOfMeasure.GetValue()

        '                End If
        '                'Get value of UPC
        '                If (Not .SalesReceiptItemRetUPC Is Nothing) Then

        '                    Dim UPC109 As String
        '                    UPC109 = .SalesReceiptItemRetUPC.GetValue()

        '                End If
        '                'Get value of WebDesc
        '                If (Not .SalesReceiptItemRetWebDesc Is Nothing) Then

        '                    Dim WebDesc110 As String
        '                    WebDesc110 = .SalesReceiptItemRetWebDesc.GetValue()

        '                End If
        '                'Get value of Manufacturer
        '                If (Not .SalesReceiptItemRetManufacturer Is Nothing) Then

        '                    Dim Manufacturer111 As String
        '                    Manufacturer111 = .SalesReceiptItemRetManufacturer.GetValue()

        '                End If
        '                'Get value of Weight
        '                If (Not .SalesReceiptItemRetWeight Is Nothing) Then

        '                    Dim Weight112 As IQBFloatType
        '                    Weight112 = .SalesReceiptItemRetWeight.GetValue()

        '                End If
        '                'Get value of WebSKU
        '                If (Not .SalesReceiptItemRetWebSKU Is Nothing) Then

        '                    Dim WebSKU113 As String
        '                    WebSKU113 = .SalesReceiptItemRetWebSKU.GetValue()

        '                End If
        '            End With
        '        Next i83
        '    End If
        '    If (Not SalesReceiptRet.TenderAccountRetList Is Nothing) Then

        '        Dim i114 As Integer
        '        For i114 = 0 To SalesReceiptRet.TenderAccountRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderAccountRet As ITenderAccountRet
        '                TenderAccountRet = SalesReceiptRet.TenderAccountRetList.GetAt(i114)
        '                'Get value of TenderAmount
        '                If (Not .TenderAccountRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount115 As Double
        '                    TenderAmount115 = .TenderAccountRetTenderAmount.GetValue()

        '                End If
        '                'Get value of TipAmount
        '                If (Not .TenderAccountRetTipAmount Is Nothing) Then

        '                    Dim TipAmount116 As Double
        '                    TipAmount116 = .TenderAccountRetTipAmount.GetValue()

        '                End If
        '            End With
        '        Next i114
        '    End If
        '    If (Not SalesReceiptRet.TenderCashRetList Is Nothing) Then

        '        Dim i117 As Integer
        '        For i117 = 0 To SalesReceiptRet.TenderCashRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderCashRet As ITenderCashRet
        '                TenderCashRet = SalesReceiptRet.TenderCashRetList.GetAt(i117)
        '                'Get value of TenderAmount
        '                If (Not .TenderCashRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount118 As Double
        '                    TenderAmount118 = .TenderCashRetTenderAmount.GetValue()

        '                End If
        '            End With
        '        Next i117
        '    End If
        '    If (Not SalesReceiptRet.TenderCheckRetList Is Nothing) Then

        '        Dim i119 As Integer
        '        For i119 = 0 To SalesReceiptRet.TenderCheckRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderCheckRet As ITenderCheckRet
        '                TenderCheckRet = SalesReceiptRet.TenderCheckRetList.GetAt(i119)
        '                'Get value of CheckNumber
        '                If (Not .TenderCheckRetCheckNumber Is Nothing) Then

        '                    Dim CheckNumber120 As String
        '                    CheckNumber120 = .TenderCheckRetCheckNumber.GetValue()

        '                End If
        '                'Get value of TenderAmount
        '                If (Not .TenderCheckRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount121 As Double
        '                    TenderAmount121 = .TenderCheckRetTenderAmount.GetValue()

        '                End If
        '            End With
        '        Next i119
        '    End If
        '    If (Not SalesReceiptRet.TenderCreditCardRetList Is Nothing) Then

        '        Dim i122 As Integer
        '        For i122 = 0 To SalesReceiptRet.TenderCreditCardRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderCreditCardRet As ITenderCreditCardRet
        '                TenderCreditCardRet = SalesReceiptRet.TenderCreditCardRetList.GetAt(i122)
        '                'Get value of CardName
        '                If (Not .TenderCreditCardRetCardName Is Nothing) Then

        '                    Dim CardName123 As String
        '                    CardName123 = .TenderCreditCardRetCardName.GetValue()

        '                End If
        '                'Get value of TenderAmount
        '                If (Not .TenderCreditCardRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount124 As Double
        '                    TenderAmount124 = .TenderCreditCardRetTenderAmount.GetValue()

        '                End If
        '                'Get value of TipAmount
        '                If (Not .TenderCreditCardRetTipAmount Is Nothing) Then

        '                    Dim TipAmount125 As Double
        '                    TipAmount125 = .TenderCreditCardRetTipAmount.GetValue()

        '                End If
        '            End With
        '        Next i122
        '    End If
        '    If (Not SalesReceiptRet.TenderDebitCardRetList Is Nothing) Then

        '        Dim i126 As Integer
        '        For i126 = 0 To SalesReceiptRet.TenderDebitCardRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderDebitCardRet As ITenderDebitCardRet
        '                TenderDebitCardRet = SalesReceiptRet.TenderDebitCardRetList.GetAt(i126)
        '                'Get value of Cashback
        '                If (Not .TenderDebitCardRetCashback Is Nothing) Then

        '                    Dim Cashback127 As Double
        '                    Cashback127 = .TenderDebitCardRetCashback.GetValue()

        '                End If
        '                'Get value of TenderAmount
        '                If (Not .TenderDebitCardRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount128 As Double
        '                    TenderAmount128 = .TenderDebitCardRetTenderAmount.GetValue()

        '                End If
        '            End With
        '        Next i126
        '    End If
        '    If (Not SalesReceiptRet.TenderDepositRetList Is Nothing) Then

        '        Dim i129 As Integer
        '        For i129 = 0 To SalesReceiptRet.TenderDepositRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderDepositRet As ITenderDepositRet
        '                TenderDepositRet = SalesReceiptRet.TenderDepositRetList.GetAt(i129)
        '                'Get value of TenderAmount
        '                If (Not .TenderDepositRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount130 As Double
        '                    TenderAmount130 = .TenderDepositRetTenderAmount.GetValue()

        '                End If
        '            End With
        '        Next i129
        '    End If
        '    If (Not SalesReceiptRet.TenderGiftRetList Is Nothing) Then

        '        Dim i131 As Integer
        '        For i131 = 0 To SalesReceiptRet.TenderGiftRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderGiftRet As ITenderGiftRet
        '                TenderGiftRet = SalesReceiptRet.TenderGiftRetList.GetAt(i131)
        '                'Get value of GiftCertificateNumber
        '                If (Not .TenderGiftRetGiftCertificateNumber Is Nothing) Then

        '                    Dim GiftCertificateNumber132 As String
        '                    GiftCertificateNumber132 = .TenderGiftRetGiftCertificateNumber.GetValue()

        '                End If
        '                'Get value of TenderAmount
        '                If (Not .TenderGiftRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount133 As Double
        '                    TenderAmount133 = .TenderGiftRetTenderAmount.GetValue()

        '                End If
        '            End With
        '        Next i131
        '    End If
        '    If (Not SalesReceiptRet.TenderGiftCardRetList Is Nothing) Then

        '        Dim i134 As Integer
        '        For i134 = 0 To SalesReceiptRet.TenderGiftCardRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim TenderGiftCardRet As ITenderGiftCardRet
        '                TenderGiftCardRet = SalesReceiptRet.TenderGiftCardRetList.GetAt(i134)
        '                'Get value of TenderAmount
        '                If (Not .TenderGiftCardRetTenderAmount Is Nothing) Then

        '                    Dim TenderAmount135 As Double
        '                    TenderAmount135 = .TenderGiftCardRetTenderAmount.GetValue()

        '                End If
        '                'Get value of TipAmount
        '                If (Not .TenderGiftCardRetTipAmount Is Nothing) Then

        '                    Dim TipAmount136 As Double
        '                    TipAmount136 = .TenderGiftCardRetTipAmount.GetValue()

        '                End If
        '            End With
        '        Next i134
        '    End If
        '    If (Not SalesReceiptRet.DataExtRetList Is Nothing) Then

        '        Dim i137 As Integer
        '        For i137 = 0 To SalesReceiptRet.DataExtRetList.Count - 1
        '            With SalesReceiptRet
        '                Dim DataExtRet As IDataExtRet
        '                DataExtRet = SalesReceiptRet.DataExtRetList.GetAt(i137)
        '                'Get value of OwnerID
        '                Dim OwnerID138 As String
        '                OwnerID138 = .DataExtRetOwnerID.GetValue()
        '                'Get value of DataExtName
        '                Dim DataExtName139 As String
        '                DataExtName139 = .DataExtRetDataExtName.GetValue()
        '                'Get value of DataExtType
        '                Dim DataExtType140 As ENDataExtType
        '                DataExtType140 = .DataExtRetDataExtType.GetValue()
        '                'Get value of DataExtValue
        '                Dim DataExtValue141 As String
        '                DataExtValue141 = .DataExtRetDataExtValue.GetValue()
        '            End With
        '        Next i137
        '    End If
        'End Sub



    End Module
End Namespace



























