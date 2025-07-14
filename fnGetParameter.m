//Retrieves value from parameter table
//Assumes table is called "Parameters"
//Assumes table has columns "Parameter" and "Value"
//https://excelguru.ca/building-a-parameter-table-for-power-query/
(ParameterName as text) =>
let
	ParamSource = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
	ParamRow = Table.SelectRows(ParamSource, each ([Parameter] = ParameterName)),
	Value=
		if Table.IsEmpty(ParamRow)=true
		then null
		else Record.Field(ParamRow{0},"Value")
in
	Value
