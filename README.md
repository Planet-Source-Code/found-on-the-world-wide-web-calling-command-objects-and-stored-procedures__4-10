<div align="center">

## Calling Command objects and stored procedures


</div>

### Description

You can design stored procedures to hide complex business rules and logic, leaving a more concise interface available for application development. Here are serveral example on how. Found at:http://www.ieighty.net/~davepamn/command.html by David Nishimoto.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-calling-command-objects-and-stored-procedures__4-10/archive/master.zip)





### Source Code

```
Create a Command
	set cm = Server.CreateObject("ADODB.Command")
	Connecting the Command
	Method 1
	cm.ActiveConnection = cn
	Method 2
	cm.ActiveConnection = "DSN=Karate; UID=Dave; PWD=519;"
	Specifying the Query
	Method 1
	set cm.CommandText = "Select * from schools"
	Method 2 (Table)
		set cm.CommandText = "schools"
		cm.CommandType = adCmdTable
	Method 3 (Stored Procedure)
		set cm.CommandText = "add_school"
		cmdCommandType = adCmdStoredProc
	Method 4 (Stored Procedure with Parameters)
		set cm.CommandText = "add_school"
		cm.cmdCommandType = adCmdStoredProc
		set p = cm.Parameters
		p.Append cm.CreateParameter("@style",adChar,adParamInput,50)
		p.Append cm.CreateParameter("@school", adChar, adParamInput,50)
		p.Append cm.CreateParameter("@id",adInteger,adParamInput)
		cm("@style") = "Kempo"
		cm("@school") = "WSU"
		cm(Id) = 1
		cm.execute
		Method 5 ( Return the results to a recordset)
		rs.Open cm, cn
		Method 6 ( Recordset, type, and locking method)
		rs.Open cm, cn, adOpenKeyset, adLockOptimistic
		(Properties of the Command Object)
		ActiveConnection	The associated Connection Object
		CommandText		The query String
		ComandTimeout	The amout of time before the
					execution is aborted
					Default is 30 seconds
		CommandType		A hint at the type of
					query string
					adCmdText		1
					adCmdTable		2
					adCmdStoredProc	4
					adCmdUnknown	8
		Prepared		Indicate whether the
					command should be
					precompiled
		(Command Object Methods)
		CreateParameter
			set p = Command.CreateParameter(n,t,d,s,v)
			n = Name of the parameter
			t = Type of Parameter
			d= The direction of the parameter
				adParamInput		1
				adParamOutput		2
				adParamInputOut	3
				adParamReturnValue	4
			s= The Maximum size of the parameter
			v= The value of the parameter
		Execute
		Set rs = command.Execute(count, parameters, options)
		count		The number of records affected by the query
		parameters	Array of parameter values
		options		A CommandType constant
	(Parameter Collection Properties)
	Count
	(Parameter Collection Methods)
	Append		Add a Parameter object
	Delete		Remove a Parameter object
				Index the name or ordinal value
	Item		Retrieve a particular Parameter object
			set parameter = Parameters.Item(index)
			index	the name or ordinal value
	Refresh		Reconstruct the collection
	(Parameter Properties)
	Attributes
			adParamLong		128
			adParamNullable	64
			adParamSigned		16
	Direction		Used for input, output, or both
			adParamInput		1
			adParamOutput		2
			adParamInputOutput	3
			adParamReturnValue	4
	Name			The Name of the parameter
	NumericScale		Decimal places after the dot
	Precision		The total number of decimal places
	Size			Size of variable data in bytes
	Type			Type of data being sent
			adBigInt
			adBinary
			adBoolean
			adBSTR
			adChar
			adCurrency
			adDate
			adDBDate		YYYYMMDD
			adDBTime		HHMMSS
			adDBTimeStamp
			adDecimal
			adDouble
			adError
			adGUID
			adIDispatch
			adInteger
			adIUnknown
			adLongVarBinary
			adLongVarChar
			adNumeric
			adSingle
			adSmallInt
			adUnsignedBigInt
			adUnsignedTinyInt
			adUserDefined
			adVariant
			adVarBinary
			adVarChar
			adVarWChar
			adWChar
	Value			Current value of the parameter
	Parameter methods
		AppendChunk	Add data to Parameter value
		GetChunk	Get a portion of the parameter value
	Refreshing Parameters
	The query string must first be examined before you can
	determine the number of parameters and their
	individual data types.
		Save yourself time by declaring the parameter objects
	manually instead of calling the refresh method.
	(Using Prepared Commands)
	* Before queries are actually executed by the data provider
	on the database server, they are examined, optimized,
	and compiled into a pseudo-code that's later
	used to drive the data-retrieval system.
	* To Prepare a Command Object, set the Prepared property
	to true.
	Example
	set cm.CommandText = "Update school set school_name = ?
	where id = ?"
	cm.CommandType = adCmdText
	cm.Prepared =true
	cm.Parameters.append cm.CreateParameter("name",adChar,adParamInput,50)
	cm.Parameters.append cm.CreateParameter("school_id, adInteger, adParamInput)
	cm("name")="Golden Lion"
	cm("id") = 1
	cm.execute
	cm("name")="Dragon kenpo"
	cm("id")=2
	cm.execute
	Stored Procedures
	* To call a stored procedure, the Parameter collection must be
	set to precisely match the number and type of
	parameters defined on the server.
```

