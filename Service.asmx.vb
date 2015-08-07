Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Web.Script.Services

' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
<System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://192.192.157.81/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function login(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As Integer
        If Tools.CheckApikey(Apikey) Then
            If Tools.Login(Id, Pwd) Then
                ' 登入成功
                Tools.SaveLog(Apikey, Id, USER_LOGIN_SUCCESS, HttpContext.Current.Request.UserHostAddress)
                Return 1
            Else
                Tools.SaveLog(Apikey, Id, USER_LOGIN_ERROR, HttpContext.Current.Request.UserHostAddress)
                Return 0
            End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="apikey"></param>
    ''' <param name="id"></param>
    ''' <param name="pwd"></param>
    ''' <param name="Name"></param>
    ''' <param name="MobilePhone"></param>
    ''' <param name="UDID"></param>
    ''' <param name="Email"></param>
    ''' <param name="Address"></param>
    ''' <returns>1: 註冊成功。-1: 帳號重複。0: 註冊失敗。-999: apikey錯誤</returns>
    ''' <remarks></remarks>
    ''' 

    '會員註冊
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function register(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal Name As String, ByVal MobilePhone As String, ByVal UDID As String, ByVal Email As String, ByVal Address As String, ByVal Type As Integer, ByVal PhoneID As String) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As Integer = Tools.Register(Id, Pwd, Name, MobilePhone, UDID, Email, Address, Type, PhoneID)
            Select Case result
                Case 1
                    Tools.SaveLog(Apikey, Id, USER_NEW_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, USER_NEW_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -1
                    Tools.SaveLog(Apikey, Id, USER_ID_DUPLICATE, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -2
                    Tools.SaveLog(Apikey, Id, USER_PhoneID_DUPLICATE, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    '修改使用者資訊
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function updateUser(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal Name As String, ByVal MobilePhone As String, ByVal UDID As String, ByVal Email As String, ByVal Address As String, ByVal PhoneID As String) As Integer
        If Tools.CheckApikey(Apikey) Then
            If Tools.UpdateUser(Id, Pwd, Name, MobilePhone, UDID, Email, Address, PhoneID) Then
                ' 修改成功
                Tools.SaveLog(Apikey, Id, USER_EDIT_SUCCESS, HttpContext.Current.Request.UserHostAddress)
                Return 1
            Else
                Tools.SaveLog(Apikey, Id, USER_EDIT_ERROR, HttpContext.Current.Request.UserHostAddress)
                Return 0
            End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function

    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getUser(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getUser(Id, Pwd)
            dt.TableName = "User"
            Return dt
            '    ' 修改成功
            '    Tools.SaveLog(Apikey, Id, USER_EDIT_SUCCESS, HttpContext.Current.Request.UserHostAddress)
            '    Return 1
            'Else
            '    Tools.SaveLog(Apikey, Id, USER_EDIT_ERROR, HttpContext.Current.Request.UserHostAddress)
            '    Return 0
            'End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得所有慈善機構資訊
    <WebMethod()> _
  <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getAllCharity(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getAllCharity(Id, Pwd)
            dt.TableName = "AllCharity"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '新增商品
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function addProduct(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal PName As String, ByVal Price As Integer, ByVal Cost As Integer, ByVal Stock As Integer, ByVal PicturePath As String, ByVal ClassID As String) As Integer
        If Tools.CheckApikey(Apikey) Then
            If Not Tools.CheckProductAndUserID(PID, Id) Then '確認商品編號不重複
                Dim result As Integer = Tools.addProduct(Id, Pwd, PID, PName, Price, Cost, 0, PicturePath, ClassID)
                Tools.Purchase(Id, Pwd, PID, Cost, Stock, Cost * Stock)
                Select Case result
                    Case 1
                        Tools.SaveLog(Apikey, Id, PRODUCT_NEW_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                    Case 0
                        Tools.SaveLog(Apikey, Id, PRODUCT_NEW_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))

                    Case -2
                        Tools.SaveLog(Apikey, Id, PRODUCT_NEW_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                End Select
                Return result
            Else
                Tools.SaveLog(Apikey, Id, PRODUCT_ID_DUPLICATE, HttpContext.Current.Request("REMOTE_ADDR"))
                Return -1 '商品編號重複

            End If

        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    '修改商品
    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function updateProduct(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal PName As String, ByVal Price As Integer, ByVal Stock As Integer, ByVal PicturePath As String, ByVal ClassID As String) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As Integer = Tools.UpdateProduct(Id, Pwd, PID, PName, Price, Stock, PicturePath, ClassID)
            Select Case result
                Case 1
                    Tools.SaveLog(Apikey, Id, PRODUCT_EDIT_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, PRODUCT_EDIT_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -1
                    Tools.SaveLog(Apikey, Id, PRODUCT_EDIT_ID_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -2
                    Tools.SaveLog(Apikey, Id, PRODUCT_EDIT_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function

    '商品進貨
    <WebMethod()> _
  <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function Purchase(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal CurrentCost As Double, ByVal Quantity As Integer, ByVal Total As Double) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As Integer = Tools.Purchase(Id, Pwd, PID, CurrentCost, Quantity, Total) '商品進貨
            Select Case result
                Case 1
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -1
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_NOT, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -2
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function Salse(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal CurrentPrice As Double, ByVal Quantity As Integer, ByVal Discount As Double, ByVal Total As Double) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As Integer = Tools.Salse(Id, Pwd, PID, CurrentPrice, Quantity, Discount, Total) '商品銷貨
            Select Case result
                Case 1
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -1
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_NOT, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -2
                    Tools.SaveLog(Apikey, Id, PRODUCT_PURCHASE_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    '取得商品資訊
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getProduct(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PId As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getProduct(Id, Pwd, PId)
            dt.TableName = "Product"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得自己所有商品(未分類)
    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getAllProduct(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getAllProduct(Id, Pwd)
            dt.TableName = "Product"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得分類商品(有分類)
    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getProductClassItems(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal ClassID As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getProductClassItems(Id, Pwd, ClassID)
            dt.TableName = "ProductClassItems"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    <WebMethod()> _
  <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getOrderDetail(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal OId As String, ByVal Seller As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getOrderDetail(Id, Pwd, OId, Seller)
            dt.TableName = "OrderDetail"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getOpenOrders(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getOpenOrders(Id, Pwd)
            dt.TableName = "OpenOders"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得營業稅率
    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getBusinessTaxRate(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As Double
        If Tools.CheckApikey(Apikey) Then

            Dim result As Double = Tools.getBusinessTaxRate(Id, Pwd)

            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function


    '取得系統參數
    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getSetting(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then

            Dim dt As DataTable = Tools.getSetting(Id, Pwd)
            dt.TableName = "Setting"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得紅利率
    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getBonusRate(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As Double
        If Tools.CheckApikey(Apikey) Then

            Dim result As Double = Tools.getBonusRate(Id, Pwd)

            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function

    '取得屬於自己的發票
    <WebMethod()> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getMyOrders(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getMyOrders(Id, Pwd)
            dt.TableName = "ReceiveOders"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '所接收過的發票，不一定屬於自己
    <WebMethod()> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getReceiveOrders(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getReceiveOrders(Id, Pwd)
            dt.TableName = "ReceiveOders"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '取得銷貨進貨紀錄
    <WebMethod()> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getProductTransaction(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal StartDate As String, ByVal EndDate As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getProductTransaction(Id, Pwd, StartDate, EndDate)
            dt.TableName = "ReceiveOders"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '儲值
    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function SaveMoney(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal Money As Double) As Integer
        If Tools.CheckApikey(Apikey) Then
            If Tools.SaveMoney(Id, Pwd, Money) Then
                ' 儲值成功
                Return 1
            Else
                Return 0
            End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            Return -999
        End If
    End Function

    '取得商品類別
    <WebMethod()> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getProductClass(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As DataTable
        If Tools.CheckApikey(Apikey) Then
            Dim dt As DataTable = Tools.getProductClass(Id, Pwd)
            dt.TableName = "ProductClass"
            Return dt
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    <WebMethod()> _
    <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function getVirtualBalance(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As Double '取得使用者虛擬帳戶餘額
        If Tools.CheckApikey(Apikey) Then
            Dim result As Double = Tools.getVirtualBalance(Id, Pwd)

            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function addOrder(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As String
        If Tools.CheckApikey(Apikey) Then
            If Tools.CheckStatus(Id) Then '確認此帳號是否停權
                Dim result As String = Tools.addOrder(Id, Pwd) '新增交易紀錄 新增成功即回傳發票編號   
                Select Case result
                    Case 0 '新增失敗
                        Tools.SaveLog(Apikey, Id, ORDER_NEW_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                    Case -1  '發票編號重複
                        Tools.SaveLog(Apikey, Id, ORDER_ID_DUPLICATE, HttpContext.Current.Request("REMOTE_ADDR"))
                    Case -2   '帳密錯誤
                        Tools.SaveLog(Apikey, Id, ORDER_NEW_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                    Case Else
                        Tools.SaveLog(Apikey, Id, ORDER_NEW_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                End Select
                Return result '回傳值為發票編號 
            Else
                Return -3 '此帳號已被停權
            End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If

    End Function

    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function addOrderDetail(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal OID As String, ByVal PID As String, ByVal Purchases As Integer, ByVal CurrentPrice As Integer, ByVal Discount As Double) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As String = Tools.addOrderDetail(Id, Pwd, OID, PID, Purchases, CurrentPrice, Discount) '新增交易商品明細  

            Tools.Salse(Id, Pwd, PID, CurrentPrice, Purchases, Discount, Purchases * Discount * CurrentPrice * 0.01) '商品銷貨紀錄
            Select Case result
                Case 1
                    Tools.SaveLog(Apikey, Id, ORDERDETAIL_NEW_SUCCESS, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, ORDERDETAIL_NEW_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -1
                    Tools.SaveLog(Apikey, Id, ORDERDETAIL_ID_DUPLICATE, HttpContext.Current.Request("REMOTE_ADDR"))
                Case -2
                    Tools.SaveLog(Apikey, Id, ORDERDETAIL_NEW_USER_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    <WebMethod()> _
<ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function orderFinish(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal OID As String, ByVal Buyer As String, ByVal Total As Integer, ByVal Owner As String, ByVal Bonus As Double, ByVal Tax As Double) As Integer
        If Tools.CheckApikey(Apikey) Then
            Dim result As String

            If Id <> Buyer Then
                result = Tools.orderFinish(Id, Pwd, OID, Buyer, Total, Owner, Bonus, Tax) '完成交易憑證

                '發簡訊!!!!!!!!!!!!!
                Dim sms As New SMS.SMS
                Dim phone As String
                Dim dt As DataTable = Tools.getBuyerData(Buyer)
                phone = dt.Rows(0).Item("MobilePhone")
                If phone <> "" Then


                    sms.sendSMS(phone, "感謝您使用雲端電子憑證服務，您已交易成功。本次交易金額共 " & CStr(Total) & " 元 ,稅金" + CStr(Tax) + "元,獲得紅利" + CStr(Bonus) + "點")
                End If

            Else

                Return 2 '不可以賣給自己

            End If
            Select Case result

                Case 1
                    Tools.SaveLog(Apikey, Id, ORDER_FINISH, HttpContext.Current.Request("REMOTE_ADDR"))
                Case 0
                    Tools.SaveLog(Apikey, Id, ORDER_FINISH_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            End Select
            Return result
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request("REMOTE_ADDR"))
            Return -999
        End If
    End Function

    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function deleteOrder(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal OID As String) As Integer

        If Tools.CheckApikey(Apikey) Then
            If Tools.deleteOrder(Id, Pwd, OID) Then
                ' 刪除發票成功
                Tools.SaveLog(Apikey, Id, ORDER_DELETE_SUCCESS, HttpContext.Current.Request.UserHostAddress)
                Return 1
            Else
                '刪除發票失敗
                Tools.SaveLog(Apikey, Id, ORDER_DELETE_ERROR, HttpContext.Current.Request.UserHostAddress)
                Return 0
            End If
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    <WebMethod()> _
   <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function sendEmail(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String) As Integer

        If Tools.CheckApikey(Apikey) Then
            Tools.sendEmail(Id, Pwd)
            Return 1
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

    '手機條碼 轉換成 身份證字號
    <WebMethod()> _
 <ScriptMethod(ResponseFormat:=ResponseFormat.Xml)> _
    Public Function PhoneIDTransfer(ByVal Apikey As String, ByVal Id As String, ByVal Pwd As String, ByVal PhoneID As String) As String
        If Tools.CheckApikey(Apikey) Then
            Dim BuyerID As String = Tools.PhoneIDTransfer(Id, Pwd, PhoneID) '取得對應的身分證字號
            If BuyerID = "" Then
                Return "0"
            End If
            Return BuyerID
        Else
            Tools.SaveLog(Apikey, "", APIKEY_ERROR, HttpContext.Current.Request.UserHostAddress)
            'Return -999
        End If
    End Function

End Class