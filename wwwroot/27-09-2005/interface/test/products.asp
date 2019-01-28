<%@LANGUAGE=JAVASCRIPT%>
<%
//Getting the ASP page to return a js file
Response.ContentType="text/js"

//Getting the id
catid = Request.queryString("id")

//Pointer to the database
database="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Server.MapPath("example.mdb") + ";"

//opening db
q="Select * from tblProducts where category = "+catid
rs=Server.CreateObject("ADODB.Recordset")

rs.open(q,database)

//looping database records
while(!rs.EOF){
  txt=rs("ProductName").value + " " + rs("ProductInfo").value
  id=rs("dbid").value
  
  Response.write('addElements("'+txt+'","","show_product_details.asp?id='+id+'")\n')
  rs.moveNext()
}
rs.close()
rs=null

Response.write("loaded()")
%>