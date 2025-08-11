<cffunction name="SpreadsheetSetCellValue" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="value" type="string" required="true" />
	<cfargument name="row" type="numeric" required="true" />
	<cfargument name="column" type="numeric" required="true" />
	<cfargument name="datatype" type="string" required="false" default="string" />
	
	<cfset arguments.spreadsheet.setCellValue(arguments.value, arguments.row, arguments.column, arguments.datatype) />
</cffunction>