<div align="center">

## ADO With JavaScript


</div>

### Description

I have not found any sample that shows database access with javaScript. Well, I created one.

It is very simple since Javascript has the activeXObject that allow us to call an external object in the same way that vbscript uses CreateObject.
 
### More Info
 
You need SQL Server, a valid login, and know the name of a table and at least one column to test the sample.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcio Coelho](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcio-coelho.md)
**Level**          |Intermediate
**User Rating**    |4.5 (94 globes from 21 users)
**Compatibility**  |
**Category**       |[Databases/ JDBC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-jdbc__2-61.md)
**World**          |[Java](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/java.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcio-coelho-ado-with-javascript__2-2804/archive/master.zip)





### Source Code

```
<SCRIPT LANGUAGE=javascript>
	<!--
	var conn = new ActiveXObject("ADODB.Connection") ;
	var connectionstring = "Provider=sqloledb; Data Source=itdev; Initial Catalog=pubs; User ID=sa;Password=yourpassword"
	conn.Open(connectionstring)
	var rs = new ActiveXObject("ADODB.Recordset")
	rs.Open("SELECT gif, description, page, location FROM menu1", conn)
// Be aware that I am selecting from this table, but you need to pick your own table
	while(!rs.eof)
	{
		alert(rs(0));
		rs.movenext
	}
	 rs.close
	conn.close
	//-->
	</SCRIPT>
```

