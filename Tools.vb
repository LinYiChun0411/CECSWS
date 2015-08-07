Imports System.Data.SqlClient
Imports DBClass.DBClass
Imports System.Net.Mail


Public Class Tools

    Shared Function Login(ByVal id As String, ByVal pwd As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Users WHERE ID=@ID AND Pwd=@Pwd"
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = id
        sqlcmd.Parameters.Add("@Pwd", SqlDbType.VarChar).Value = pwd
        Return ExecuteScalar(sqlcmd)
    End Function

    Shared Function CheckBuyer(ByVal Buyer As String) As Boolean '檢查買家是否存在
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Users WHERE ID=@ID"
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = Buyer

        Return ExecuteScalar(sqlcmd)
    End Function

    Shared Function getUser(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT * FROM Users WHERE ID=@ID AND Pwd=@Pwd"
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = id
        sqlcmd.Parameters.Add("@Pwd", SqlDbType.VarChar).Value = pwd
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function getProduct(ByVal id As String, ByVal pwd As String, ByVal pid As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT * FROM Product WHERE PID=@PID And UserID=@UserID"
            sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = pid
            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = id
        End If
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function getAllProduct(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT * FROM Product WHERE UserID=@UserID"
            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = id
        End If
        Return GetDataTable(sqlcmd)
    End Function

    '取得買家資訊
    Shared Function getBuyerData(ByVal id As String) As DataTable
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT * FROM Users WHERE ID=@ID"
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = id
        Return GetDataTable(sqlcmd)
    End Function


    Shared Function getProductClassItems(ByVal id As String, ByVal pwd As String, ByVal ClassID As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT * FROM Product Join ProductClass On Product.ClassID=ProductClass.ID WHERE UserID=@UserID And ClassID=@ClassID"
            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = id
            sqlcmd.Parameters.Add("@ClassID", SqlDbType.VarChar).Value = ClassID
        End If
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function getOrderDetail(ByVal id As String, ByVal pwd As String, ByVal OID As String, ByVal Seller As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT OrderDetail.OID, OrderDetail.PID, OrderDetail.Purchases, OrderDetail.CurrentPrice, OrderDetail.Discount, Product.PName, Product.PicturePath, Product.CreateDate FROM OrderDetail LEFT OUTER JOIN Product ON OrderDetail.PID = Product.PID WHERE (OrderDetail.OID = @OID) And UserID=@UserID"
            sqlcmd.Parameters.Add("@OID", SqlDbType.VarChar).Value = OID
            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Seller
        End If
        Return GetDataTable(sqlcmd)
    End Function

    ' 取得開立出去的發票
    Shared Function getOpenOrders(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Orders.OID, Orders.Seller, Orders.Buyer,(Select Name From Users Where ID=Seller) As SellerName,(Select Name From Users Where ID=Buyer) As BuyerName,Orders.Tax, Orders.Total,Orders.Owner,(Select Name From Users Where ID=Orders.Owner)As OwnerName, CONVERT(VARCHAR(10),Orders.Date, 111 ) As Date  FROM Orders Join Users On Users.ID=Orders.Seller WHERE (Orders.Seller = @Seller) And Orders.Status = 1 ORDER By Orders.Date DESC"

            sqlcmd.Parameters.Add("@Seller", SqlDbType.VarChar).Value = id
        End If
        Return GetDataTable(sqlcmd)
    End Function
    ' 取得屬於自己的發票
    Shared Function getMyOrders(ByVal id As String, ByVal pwd As String) As DataTable '真正所屬
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Orders.OID, Orders.Seller, Orders.Buyer,(select Name From Users where ID=Orders.Seller) As SellerName,(select Name From Users where ID=Orders.Buyer) As BuyerName, CONVERT(VARCHAR(10),Orders.Date, 111 ) As Date,Orders.Tax, Orders.Total,Orders.Owner,Users.Name As OwnerName FROM Orders Join Users On Users.ID=Orders.Owner WHERE (Orders.Owner = @Owner) And Orders.Status = 1 ORDER By Orders.Date DESC"

            sqlcmd.Parameters.Add("@Owner", SqlDbType.VarChar).Value = id
        End If
        Return GetDataTable(sqlcmd)
    End Function


    Shared Function getProductTransaction(ByVal id As String, ByVal pwd As String, ByVal StartDate As String, ByVal EndDate As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Product.PName,ProductTransaction.PID,ProductTransaction.CurrentCost,ProductTransaction.CurrentPrice,ProductTransaction.Quantity,ProductTransaction.Event,CONVERT(VARCHAR(10) ,ProductTransaction.Date, 111 ) As Date FROM ProductTransaction left join Product On Product.PID=ProductTransaction.PID where Product.UserID = @UserID And ProductTransaction.Date > @StartDate And ProductTransaction.Date < @EndDate Order By Date"

            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = id
            sqlcmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
            sqlcmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        End If
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function getReceiveOrders(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Orders.OID, Orders.Seller, Orders.Buyer,(select Name From Users where ID=Orders.Seller) As SellerName,(select Name From Users where ID=Orders.Buyer) As BuyerName,CONVERT(VARCHAR(10),Orders.Date, 111 ) As Date, Orders.Total,Orders.Tax ,Orders.Owner,Users.Name As OwnerName,Users.Type FROM Orders LEFT JOIN Users ON Users.ID = Orders.Owner  WHERE Orders.Buyer=@Buyer And Orders.Owner != @Buyer And Orders.Status = 1 ORDER By Orders.Date DESC"

            sqlcmd.Parameters.Add("@Buyer", SqlDbType.VarChar).Value = id
        End If
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function getVirtualBalance(ByVal id As String, ByVal pwd As String) As Double
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT VirtualBalance FROM Users WHERE Users.ID = @Users"
            sqlcmd.Parameters.Add("@Users", SqlDbType.VarChar).Value = id
        End If
        Return ExecuteScalar(sqlcmd)
    End Function

    '取得營業稅率
    Shared Function getBusinessTaxRate(ByVal id As String, ByVal pwd As String) As Double
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Value FROM Setting Where Name='BusinessTaxRate'"
        End If
        Return ExecuteScalar(sqlcmd)
    End Function


    '取得系統參數
    Shared Function getSetting(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Name,Value FROM Setting"
        End If
        Return GetDataTable(sqlcmd)
    End Function

    '取得紅利率
    Shared Function getBonusRate(ByVal id As String, ByVal pwd As String) As Double
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Value FROM Setting Where Name='Bonus'"
        End If
        Return ExecuteScalar(sqlcmd)
    End Function
    '取得商品類別
    Shared Function getProductClass(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT * FROM ProductClass"
        End If
        Return GetDataTable(sqlcmd)
    End Function

    '取得所有慈善機構資訊
    Shared Function getAllCharity(ByVal id As String, ByVal pwd As String) As DataTable
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT Name As CharityName,ID As CharityID FROM Users where Type=2"
        End If
        Return GetDataTable(sqlcmd)
    End Function

    Shared Function PhoneIDTransfer(ByVal id As String, ByVal pwd As String, ByVal PhoneID As String) As String
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(id, pwd) Then
            sqlcmd.CommandText = "SELECT ID FROM Users where PhoneID=@PhoneID"
            sqlcmd.Parameters.Add("@PhoneID", SqlDbType.VarChar).Value = PhoneID
        End If
        Return ExecuteScalar(sqlcmd)
    End Function

    '修改使用者資訊
    Shared Function UpdateUser(ByVal Id As String, ByVal Pwd As String, ByVal Name As String, ByVal MobilePhone As String, ByVal UDID As String, ByVal Email As String, ByVal Address As String, ByVal PhoneID As String) As Integer
        Dim sqlcmd As New SqlCommand

        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            sqlcmd.CommandText = "Update Users Set Pwd=@Pwd,Name=@Name,MobilePhone=@MobilePhone,UDID=@UDID,Email=@Email,Address=@Address,PhoneID=@PhoneID WHERE Id=@Id"
            sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id
            sqlcmd.Parameters.Add("@Pwd", SqlDbType.VarChar).Value = Pwd
            sqlcmd.Parameters.Add("@Name", SqlDbType.VarChar).Value = Name
            sqlcmd.Parameters.Add("@MobilePhone", SqlDbType.VarChar).Value = MobilePhone
            sqlcmd.Parameters.Add("@UDID", SqlDbType.VarChar).Value = UDID
            sqlcmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Email
            sqlcmd.Parameters.Add("@Address", SqlDbType.VarChar).Value = Address
            sqlcmd.Parameters.Add("@PhoneID", SqlDbType.VarChar).Value = PhoneID

            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return 1
            Else
                Return 0
            End If
        Else
            Return -1   ' 帳號密碼驗證錯誤
        End If

    End Function

    '會員註冊
    Shared Function Register(ByVal Id As String, ByVal Pwd As String, ByVal Name As String, ByVal MobilePhone As String, ByVal UDID As String, ByVal Email As String, ByVal Address As String, ByVal Type As Integer, ByVal PhoneID As String) As Integer
        Dim CreateDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand

        If Not CheckUserId(Id) Then '檢查ID是否唯一
            If Not CheckPhoneID(PhoneID) Then '檢查PhoneID是否唯一
                sqlcmd.CommandText = "INSERT INTO Users(Id,Pwd,Name,MobilePhone,UDID,Email,Address,CreateDate,Type,PhoneID) " & _
                                 "VALUES(@Id,@Pwd,@Name,@MobilePhone,@UDID,@Email,@Address,@CreateDate,@Type,@PhoneID)"
                sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id
                sqlcmd.Parameters.Add("@Pwd", SqlDbType.VarChar).Value = Pwd
                sqlcmd.Parameters.Add("@Name", SqlDbType.VarChar).Value = Name
                sqlcmd.Parameters.Add("@MobilePhone", SqlDbType.VarChar).Value = MobilePhone
                sqlcmd.Parameters.Add("@UDID", SqlDbType.VarChar).Value = UDID
                sqlcmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Email
                sqlcmd.Parameters.Add("@Address", SqlDbType.VarChar).Value = Address
                sqlcmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = CreateDate
                sqlcmd.Parameters.Add("@Type", SqlDbType.Int).Value = Type
                sqlcmd.Parameters.Add("@PhoneID", SqlDbType.VarChar).Value = PhoneID
                Dim strResult As String = ExecuteNonQuery(sqlcmd)
                If strResult = 1 Then
                    Return 1
                Else
                    Return 0
                End If

            Else
                Return -2 ' PhoneID重複

            End If
        Else
            Return -1 ' 帳號重複
        End If

    End Function

    '新增商品
    Shared Function addProduct(ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal PName As String, ByVal Price As Integer, ByVal Cost As Integer, ByVal Stock As Integer, ByVal PicturePath As String, ByVal ClassID As String) As Integer
        Dim CreateDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand

        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then

            sqlcmd.CommandText = "INSERT INTO Product(UserID,PID,PName,Price,Cost,Stock,PicturePath,CreateDate,ClassID) " & _
                             "VALUES(@UserID,@PID,@PName,@Price,@Cost,@Stock,@PicturePath,@CreateDate,@ClassID)"
            sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id
            sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
            sqlcmd.Parameters.Add("@PName", SqlDbType.VarChar).Value = PName
            sqlcmd.Parameters.Add("@Price", SqlDbType.Int).Value = Price
            sqlcmd.Parameters.Add("@Cost", SqlDbType.Int).Value = Cost
            sqlcmd.Parameters.Add("@Stock", SqlDbType.Int).Value = Stock
            sqlcmd.Parameters.Add("@PicturePath", SqlDbType.VarChar).Value = PicturePath
            sqlcmd.Parameters.Add("@CreateDate", SqlDbType.DateTime).Value = CreateDate
            sqlcmd.Parameters.Add("@ClassID", SqlDbType.VarChar).Value = ClassID
            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return 1
            Else
                Return 0
            End If
     
        Else
            Return -2   ' 使用者帳號密碼錯誤
        End If
    End Function

    '商品進貨
    Shared Function Purchase(ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal CurrentCost As Double, ByVal Quantity As Integer, ByVal Total As Double) As Integer
        Dim PurchaseDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            If CheckProductId(PID) Then '檢查是否有這項商品
                sqlcmd.CommandText = "INSERT INTO ProductTransaction(PID,UserID,CurrentCost,Quantity,Date,Total,Event) " & _
                                 "VALUES(@PID,@UserID,@CurrentCost,@Quantity,@Date,@Total,@Event)"
                sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id
                sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
                sqlcmd.Parameters.Add("@CurrentCost", SqlDbType.Float).Value = CurrentCost
                sqlcmd.Parameters.Add("@Quantity", SqlDbType.Int).Value = Quantity

                sqlcmd.Parameters.Add("@Total", SqlDbType.Float).Value = Total '進貨成本費用小計
                sqlcmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = PurchaseDate
                sqlcmd.Parameters.Add("@Event", SqlDbType.NVarChar).Value = "進貨"


                Dim strResult As String = ExecuteNonQuery(sqlcmd)
                If strResult = 1 Then
                    addStock(Id, PID, Quantity) '增加商品庫存量'
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return -1   ' 使用者沒有這項商品
            End If
        Else
            Return -2   ' 使用者帳號密碼錯誤
        End If

    End Function

    '銷貨
    Shared Function Salse(ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal CurrentPrice As Double, ByVal Quantity As Integer, ByVal Discount As Double, ByVal Total As Double) As Integer
        Dim PurchaseDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            If CheckProductId(PID) Then '檢查是否有這項商品
                sqlcmd.CommandText = "INSERT INTO ProductTransaction(PID,UserID,CurrentPrice,Quantity,Date,Discount,Total,Event) " & _
                                 "VALUES(@PID,@UserID,@CurrentPrice,@Quantity,@Date,@Discount,@Total,@Event)"
                sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id
                sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
                sqlcmd.Parameters.Add("@CurrentPrice", SqlDbType.Float).Value = CurrentPrice
                sqlcmd.Parameters.Add("@Quantity", SqlDbType.Int).Value = Quantity
                sqlcmd.Parameters.Add("@Discount", SqlDbType.Float).Value = Discount
                sqlcmd.Parameters.Add("@Total", SqlDbType.Float).Value = Total '銷貨金額小計
                sqlcmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = PurchaseDate
                sqlcmd.Parameters.Add("@Event", SqlDbType.NVarChar).Value = "銷貨"


                Dim strResult As String = ExecuteNonQuery(sqlcmd)
                If strResult = 1 Then
                    deleteStock(Id, PID, Quantity) '減少商品庫存量'
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return -1   ' 使用者沒有這項商品
            End If
        Else
            Return -2   ' 使用者帳號密碼錯誤
        End If

    End Function

    '增加商品庫存量
    Shared Function addStock(ByVal Id As String, ByVal PID As String, ByVal Quantity As Integer) As String
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "Update Product Set Stock=Stock+@Quantity WHERE PID=@PID And UserID=@UserID"
        sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
        sqlcmd.Parameters.Add("@Quantity", SqlDbType.Int).Value = Quantity
        sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id

        Dim strResult As String = ExecuteNonQuery(sqlcmd)
        If strResult = 1 Then
            Return 1
        Else
            Return 0
        End If
       

    End Function

    '減少商品庫存量
    Shared Function deleteStock(ByVal Id As String, ByVal PID As String, ByVal Quantity As Integer) As String
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "Update Product Set Stock=Stock-@Quantity WHERE PID=@PID And UserID=@UserID"
        sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
        sqlcmd.Parameters.Add("@Quantity", SqlDbType.Int).Value = Quantity
        sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id

        Dim strResult As String = ExecuteNonQuery(sqlcmd)
        If strResult = 1 Then
            Return 1
        Else
            Return 0
        End If


    End Function

    '新增發票
    Shared Function addOrder(ByVal Id As String, ByVal Pwd As String) As String

        Dim CreateDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand
        Dim OID As String = createOID()
        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            sqlcmd.CommandText = "INSERT INTO Orders(OID,Seller,Date) " & _
                             "VALUES(@OID,@Seller,@Date)"
            sqlcmd.Parameters.Add("@OID", SqlDbType.VarChar).Value = OID
            sqlcmd.Parameters.Add("@Seller", SqlDbType.VarChar).Value = Id
            sqlcmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = Date.Now
            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return OID '新增交易憑證成功即回傳交易憑證編號
            Else
                Return 0 '新增交易憑證失敗
            End If

        Else
            Return -2   ' 使用者帳號密碼錯誤

        End If
    End Function

    '完成發票開立
    Shared Function orderFinish(ByVal Id As String, ByVal Pwd As String, ByVal OID As String, ByVal Buyer As String, ByVal Total As Integer, ByVal Owner As String, ByVal Bonus As Double, ByVal Tax As Double) As Integer
        Dim sqlcmd As New SqlCommand
        Dim strResult As String = 0
        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            If CheckBuyer(Buyer) And Buyer <> Id Then '檢查買家是否存在 、 且不能賣給自己
                sqlcmd.CommandText = "Update Orders Set Buyer=@Buyer,Total=@Total,Tax=@Tax,Owner=@Owner,Bonus=@Bonus,Status=1 WHERE OID=@OID"
                sqlcmd.Parameters.Add("@OID", SqlDbType.VarChar).Value = OID '發票編號
                sqlcmd.Parameters.Add("@Buyer", SqlDbType.VarChar).Value = Buyer '買方
                sqlcmd.Parameters.Add("@Total", SqlDbType.Int).Value = Total '交易總額
                sqlcmd.Parameters.Add("@Tax", SqlDbType.Float).Value = Tax '稅金
                sqlcmd.Parameters.Add("@Owner", SqlDbType.VarChar).Value = Owner '買方
                sqlcmd.Parameters.Add("@Bonus", SqlDbType.Float).Value = Bonus '紅利

                strResult = ExecuteNonQuery(sqlcmd)
            End If

            If strResult = 1 Then
                '設定紅利及稅金
                SetTax(Id, Tax) '賣家扣除虛擬餘額
                SetBonus(Buyer, Bonus) '買家增加紅利

                Return 1 '交易憑證完成確認
            Else
                Return 0 '交易憑證確認失敗
            End If

        Else
            Return -2   ' 使用者帳號密碼錯誤

        End If
    End Function

    Shared Function SetTax(ByVal Id As String, ByVal Tax As Double) As String
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "Update Users Set VirtualBalance=VirtualBalance-@Tax WHERE ID=@ID"
        sqlcmd.Parameters.Add("@Tax", SqlDbType.Float).Value = Tax
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = Id
        Dim strResult As String = ExecuteNonQuery(sqlcmd)
        If strResult = 1 Then
            Return 1
        Else
            Return 0
        End If


    End Function


    Shared Function SaveMoney(ByVal Id As String, ByVal Pwd As String, ByVal money As Double) As String
        If Login(Id, Pwd) Then
            Dim sqlcmd As New SqlCommand
            sqlcmd.CommandText = "Update Users Set VirtualBalance=VirtualBalance+@money WHERE ID=@ID"
            sqlcmd.Parameters.Add("@money", SqlDbType.Float).Value = money
            sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = Id
            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function


    Shared Function SetBonus(ByVal Id As String, ByVal Bonus As Double) As Integer
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "Update Users Set Bonus=Bonus+@Bonus WHERE ID=@ID"
        sqlcmd.Parameters.Add("@Bonus", SqlDbType.Float).Value = Bonus
        sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = Id
        Dim strResult As String = ExecuteNonQuery(sqlcmd)
        If strResult = 1 Then
            Return 1
        Else
            Return 0
        End If


    End Function

    '發票作廢
    Shared Function deleteOrder(ByVal Id As String, ByVal Pwd As String, ByVal OID As String) As Integer
        Dim sqlcmd As New SqlCommand
        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            sqlcmd.CommandText = "Update Orders Set Status=0 WHERE OID=@OID And Seller=@Id"
            sqlcmd.Parameters.Add("@OID", SqlDbType.VarChar).Value = OID
            sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id

            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return 1 '發票刪除成功
            Else
                Return 0 '發票刪除失敗
            End If

        Else
            Return -2   ' 使用者帳號密碼錯誤

        End If
    End Function

    Shared Function addOrderDetail(ByVal Id As String, ByVal Pwd As String, ByVal OID As String, ByVal PID As String, ByVal Purchases As Integer, ByVal CurrentPrice As Integer, ByVal Discount As Double) As Integer

        Dim sqlcmd As New SqlCommand

        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            sqlcmd.CommandText = "INSERT INTO OrderDetail(OID,PID,Purchases,CurrentPrice,Discount) " & _
                             "VALUES(@OID,@PID,@Purchases,@CurrentPrice,@Discount)"
            sqlcmd.Parameters.Add("@OID", SqlDbType.VarChar).Value = OID
            sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
            sqlcmd.Parameters.Add("@Purchases", SqlDbType.Int).Value = Purchases
            sqlcmd.Parameters.Add("@CurrentPrice", SqlDbType.Int).Value = CurrentPrice
            sqlcmd.Parameters.Add("@Discount", SqlDbType.Float).Value = Discount

            Dim strResult As String = ExecuteNonQuery(sqlcmd)
            If strResult = 1 Then
                Return 1 '新增交易商品明細成功
            Else
                Return 0 '新增交易商品明細失敗
            End If

        Else
            Return -2   ' 使用者帳號密碼錯誤

        End If
    End Function

    '修改商品
    Shared Function UpdateProduct(ByVal Id As String, ByVal Pwd As String, ByVal PID As String, ByVal PName As String, ByVal Price As Integer, ByVal Stock As Integer, ByVal PicturePath As String, ByVal ClassID As String) As Integer
        Dim sqlcmd As New SqlCommand

        ' 檢查帳號密碼是否正確
        If Login(Id, Pwd) Then
            If CheckProductId(PID) Then
                sqlcmd.CommandText = "Update Product Set PName=@PName,Price=@Price,Stock=@Stock,PicturePath=@PicturePath,ClassID=@ClassID WHERE PID=@PID And UserID=@UserID"
                sqlcmd.Parameters.Add("@PID", SqlDbType.VarChar).Value = PID
                sqlcmd.Parameters.Add("@PName", SqlDbType.VarChar).Value = PName
                sqlcmd.Parameters.Add("@Price", SqlDbType.Int).Value = Price
                sqlcmd.Parameters.Add("@Stock", SqlDbType.Int).Value = Stock
                sqlcmd.Parameters.Add("@PicturePath", SqlDbType.VarChar).Value = PicturePath
                sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id
                sqlcmd.Parameters.Add("@ClassID", SqlDbType.VarChar).Value = ClassID

                Dim strResult As String = ExecuteNonQuery(sqlcmd)
                If strResult = 1 Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return -1   ' 產品編號錯誤
            End If
        Else
            Return -2   ' 使用者帳號密碼錯誤
        End If

    End Function

    Shared Sub sendEmail(ByVal id As String, ByVal pwd As String)
        Dim sqlcmd As New SqlCommand
        If Login(id, pwd) Then '確認帳密
            sqlcmd.CommandText = "SELECT Email FROM Users WHERE Users.ID = @ID"
            sqlcmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = id
            Dim address As String = ExecuteScalar(sqlcmd)
            Dim msg As String = "您的餘額不足無法開立憑證"
            Dim mysubject As String = "憑證開立失敗"
            Dim message As New MailMessage("abc@gmail.com", address) 'MailMessage(寄信者, 收信者)
            message.IsBodyHtml = True
            message.BodyEncoding = System.Text.Encoding.UTF8 'E-mail編碼
            message.Subject = mysubject 'E-mail主旨
            message.Body = msg 'E-mail內容
            Dim smtpClient As New SmtpClient("smtp.gmail.com", 25) '設定E-mail Server和port
            smtpClient.Send(message)
        End If    
    End Sub

    '檢查ID是否存在
    Shared Function CheckUserId(ByVal Id As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Id=@Id"
        sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id
        Return ExecuteScalar(sqlcmd)
    End Function


    '檢查PhoneID是否存在
    Shared Function CheckPhoneID(ByVal PhoneID As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Users WHERE PhoneID=@PhoneID"
        sqlcmd.Parameters.Add("@PhoneID", SqlDbType.VarChar).Value = PhoneID
        Return ExecuteScalar(sqlcmd)
    End Function


    '檢查此帳號是否被停權
    Shared Function CheckStatus(ByVal Id As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Status=1 And Id=@Id"
        sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id
        Return ExecuteScalar(sqlcmd)
    End Function

    '檢查賣家的商品編號是否重複
    Shared Function CheckProductAndUserID(ByVal PId As String, ByVal Id As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Product WHERE PId=@PId And UserID=@UserID"
        sqlcmd.Parameters.Add("@PId", SqlDbType.VarChar).Value = PId
        sqlcmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Id
        Return ExecuteScalar(sqlcmd)
    End Function
    '檢查是否有此商品
    Shared Function CheckProductId(ByVal PId As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Product WHERE PId=@PId"
        sqlcmd.Parameters.Add("@PId", SqlDbType.VarChar).Value = PId

        Return ExecuteScalar(sqlcmd)
    End Function

    Shared Function CheckApikey(ByVal apikey As String) As Boolean
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM APIKey WHERE APIKey=@APIKey AND Enable=1"
        sqlcmd.Parameters.Add("@APIKey", SqlDbType.VarChar).Value = apikey

        'Dim result As Integer = DBClass.ExecuteScalar(sqlcmd)
        'If result <> 1 Then
        '    Tools.SaveLog(apikey, "", INTERFACE_APP, APIKEY_ERROR, "")
        '    Return False
        'Else
        '    Return True
        'End If
        Return ExecuteScalar(sqlcmd)
    End Function

    Shared Function SaveLog(ByVal apikey As String, ByVal Id As String, ByVal LogEvent As String, ByVal IP As String) As Integer
        Dim LogDate As Date = Date.Now
        Dim sqlcmd As New SqlCommand
        Dim UseInterface As String = getUseInterface(apikey)
        If UseInterface Is Nothing Then
            UseInterface = ""
        End If
        sqlcmd.CommandText = "INSERT INTO Log(Id,UseInterface,LogEvent,IP,LogDate,APIKey) " & _
                         "VALUES(@Id,@UseInterface,@LogEvent,@IP,@LogDate,@APIKey)"
        sqlcmd.Parameters.Add("@Id", SqlDbType.VarChar).Value = Id
        sqlcmd.Parameters.Add("@UseInterface", SqlDbType.VarChar).Value = UseInterface
        sqlcmd.Parameters.Add("@LogEvent", SqlDbType.VarChar).Value = LogEvent
        sqlcmd.Parameters.Add("@IP", SqlDbType.VarChar).Value = IP
        sqlcmd.Parameters.Add("@LogDate", SqlDbType.DateTime).Value = LogDate
        sqlcmd.Parameters.Add("@APIKey", SqlDbType.VarChar).Value = apikey
        Return ExecuteScalar(sqlcmd)

    End Function

    Shared Function getUseInterface(ByVal apikey As String) As String
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT Description FROM APIKey WHERE APIKey=@APIKey AND Enable=1"
        sqlcmd.Parameters.Add("@APIKey", SqlDbType.VarChar).Value = apikey
        Return ExecuteScalar(sqlcmd)
    End Function

    Shared Function SQLString(ByVal text As String) As String
        Return "'" & text & "'"
    End Function

    Shared Function createOID() As String
        Dim sqlcmd As New SqlCommand
        sqlcmd.CommandText = "SELECT COUNT(*) FROM Orders"

        If ExecuteScalar(sqlcmd) = 0 Then
            Return "BK10000001"
        Else
            Dim sqlcmd2 As New SqlCommand
            sqlcmd2.CommandText = "SELECT TOP 1 OID FROM Orders ORDER BY OID DESC"
            Dim s As String = ExecuteScalar(sqlcmd2) '抓取最後開出的交易憑證編號
            s = s.Substring(2, 8)
            s = Convert.ToInt32(s) + 1 '交易憑證編號加一號
            Dim OID As String = "BK"
            OID = OID & s
            Return OID
        End If




           
    End Function
End Class
