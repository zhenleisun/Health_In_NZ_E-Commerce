<!--#include file="conn.asp"-->
<%

set rs=server.CreateObject("adodb.recordset")
rs.Open "select payid from webinfo ",conn,1,3
MerchantID_User=rs("westid")'支付ID(客户ID)
rs.close
set rs=nothing
''''''''''''''''''''''''''''''''''''''''''
' 该示范程序接收西部支付发来的支付成功通知，验证其真实性，并作相应的自动处理
''''''''''''''''''''''''''''''''''''''''''

dim MerchantOrderNumber, WestPayOrderNumber, PaidAmount, MerchantID

MerchantOrderNumber=request("MerchantOrderNumber") '和商户支付命令中的订单号相同
WestPayOrderNumber=request("WestPayOrderNumber")
PaidAmount=ccur("0"&request("PaidAmount"))  'WestPay传回的实际支付金额。用CCUR转为数字型。
'注：商户必须根据我们传回商户原始订单号找到原始订单，比较实付金额和原始订单金额，相同才是支付成功。
MerchantID=request("MerchantID")
'注：商户必须判断此商户ID是不是您的商户ID

Dim objHttp, str

Dim PaymentCompleted
PaymentCompleted=false

' 准备回传支付通知表单
str = Request.Form & "&cmd=validate"

'set objHttp = Server.CreateObject("Msxml21.ServerXMLHTTP")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
 
'把WestPay传来的通知内容再传回到WestPay作验证以确保通知信息的真实性 
 
objHttp.open "POST", "http://www.WestPay.com.cn/pay/ISPN.asp", false
'ISPN: Instant Secure Payment Notification

objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

if (objHttp.status <> 200 ) then
	    'HTTP 错误处理
		response.Write("Status="&objHttp.status)
elseif (objHttp.responseText = "VERIFIED") then
	
		'支付通知验证成功
		if MerchantID=""&MerchantID_User&"" then '判断此订单是不是商户的订单。
			PaymentCompleted=true
		end if
	
elseif (objHttp.responseText = "INVALID") then
		'支付通知验证失败
		response.Write("Invalid")
else
		'支付通知验证过程中出现错误
		response.Write("Error")
end if
set objHttp = nothing

if PaymentCompleted then
	'支付成功的处理...（以下程序仅供参考）

	'注：商户必须根据我们传回商户原始订单号找到原始订单，比较实付金额和原始订单金额，相同才是支付成功。
	'此处为在线充值,不用金额比对是否够!
	
	'if 您的原始订单金额=PaidAmount then
	reseponse.write "恭喜，支付成功!"
	'else
		'失败
	'end if
else

	'支付不成功的处理...
response.redirect "logout3.asp"end if
%>�