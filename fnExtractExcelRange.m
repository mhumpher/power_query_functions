let
    ExtractExcelRange = (filePath as text, sheetName as text, range as text) as table =>
    let
        // Load the Excel file
        Source = Excel.Workbook(File.Contents(filePath), null, true),
        
        // Get the specified sheet
        Sheet = Table.SelectRows(Source, each [Item] = sheetName and [Kind] = "Sheet"){0}[Data],

        // Convert column letters to index
        ColumnLetterToIndex = (col as text) as number =>
            let
                letters = Text.ToList(Text.Upper(col)),
                reversed = List.Reverse(letters),
                positions = List.Transform({0..List.Count(reversed)-1}, (i) =>
                    (Character.ToNumber(reversed{i}) - Character.ToNumber("A") + 1) * Number.Power(26, i)
                )
            in
                List.Sum(positions),

        // Parse cell reference like "B2"
        ParseCell = (cell as text) as record =>
            let
                colLetters = Text.Select(cell, {"A".."Z", "a".."z"}),
                rowNumbers = Text.Select(cell, {"0".."9"}),
                colIndex = ColumnLetterToIndex(colLetters),
                rowIndex = Number.FromText(rowNumbers)
            in
                [Col = colIndex, Row = rowIndex],

        // Split range like "B2:D10"
        RangeParts = Text.Split(range, ":"),
        StartCell = ParseCell(RangeParts{0}),
        EndCell = ParseCell(RangeParts{1}),

        // Extract rows
        RowCount = EndCell[Row] - StartCell[Row] + 1,
        RawRows = Table.Range(Sheet, StartCell[Row] - 1, RowCount),

        // Extract columns
        ColCount = EndCell[Col] - StartCell[Col] + 1,
        ColumnNames = List.Range(Table.ColumnNames(RawRows), StartCell[Col] - 1, ColCount),
        Final = Table.SelectColumns(RawRows, ColumnNames)
    in
        Final
in
    ExtractExcelRange
