/*
 *  Copyright (c) 2013 James Buncle
 * 
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 * 
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 * 
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 * 
 */
package com.jbuncle.exceltomysql;

import java.sql.Connection;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Map.Entry;
import java.util.*;
import org.apache.poi.ss.usermodel.*;

/**
 * TO-DO: PropertyEditor to change values
 *
 */
public class ExcelToMySQL {

    private final SheetFilter filter;

    public ExcelToMySQL() {
        filter = new SheetPathFilter();
    }

    public ExcelToMySQL(String... allowedPaths) {
        filter = new SheetPathFilter(allowedPaths);
    }

    public void addWorkbook(Connection conn, Workbook workbook) throws SQLException {
        final int numberOfSheets = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
            final Sheet sheet = workbook.getSheetAt(sheetIndex);
            if (filter.accept(sheet)) {
                addTableFromSheet(conn, sheet);
            }
        }
    }

    public void addTableFromSheet(final Connection conn, final Sheet sheet) throws SQLException {
        final int numRows = sheet.getPhysicalNumberOfRows();
        if (numRows < 2) {
            //Not enough or can't determine
        }
        final List<Entry<String, ExcelType>> types = extractTypes(sheet);
        final String tableName = Utils.cleanUp(sheet.getSheetName());
        {
            final String dropStatement = "DROP TABLE IF EXISTS `" + tableName + "`;";
            conn.createStatement().execute(dropStatement);
            System.out.println(dropStatement);
        }
        {
            final String createStatement = getCreateTable(tableName, types);
            System.out.println(createStatement);
            conn.createStatement().execute(createStatement);
        }

        int rowCount = 0;
        for (final Row row : new IteratorWrapper<Row>(sheet.iterator())) {
            if (rowCount > 0) {
                final String insert = createInsert(tableName, types, row);
                System.out.println(insert);
                if (insert != null) {
                    conn.createStatement().execute(insert);
                }
            }
            rowCount++;
        }
    }

    private static String createInsert(final String tableName, final List<Entry<String, ExcelType>> types, final Row row) {
        //Iterate
        final StringBuilder columns = new StringBuilder();
        final StringBuilder values = new StringBuilder();
        final FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();

        int nullCount = 0;
        int columnCount = 0;
        for (Entry<String, ExcelType> sourceType : types) {
            if (isSet(sourceType)) {
                columns.append("`").append(sourceType.getKey()).append("`").append(",");
                Cell cell = row.getCell(columnCount);

                if (cell == null) {
                    values.append("null").append(",");
                } else {
                    cell = evaluator.evaluateInCell(cell);

                    switch (sourceType.getValue()) {
                        case DATE:
                            values.append("'").append(new SimpleDateFormat("yyyy-MM-dd HH:mm").format(cell.getDateCellValue())).append("'").append(",");
                            break;
                        case NUMERIC:
                            values.append(cell.getNumericCellValue()).append(",");
                            break;
                        case BOOLEAN:
                            values.append(cell.getBooleanCellValue()).append(",");
                            break;
                        case STRING:
                            values.append("'").append(cell.getStringCellValue().replaceAll("'", "\\\\'")).append("'").append(",");
                            break;
                        default:
                            values.append("null").append(",");
                            nullCount++;
                            break;
                    }
                }

            }
            columnCount++;
        }
        columns.deleteCharAt(columns.length() - 1);
        values.deleteCharAt(values.length() - 1);

        if (nullCount >= columnCount) {
            return null;
        }
        return "INSERT INTO `" + tableName + "` (" + columns + ") VALUES (" + values + ");";
    }

    private static boolean isSet(Entry<String, ExcelType> entry) {
        return entry != null && entry.getKey() != null && entry.getValue() != null;
    }

    private void addColumnName(List<Entry<String, ExcelType>> columns, int cellCount, String columnName, Sheet sheet) {
        if (filter.accept(sheet, columnName)) {
            //Doesn't exist yet
            if (!Utils.typesContain(columns, columnName)) {
                columns.add(cellCount, new AbstractMap.SimpleEntry<String, ExcelType>(columnName, null));

            } else {
                //Already contains column
                int tempIndex = 1;

                String tempCol = columnName + tempIndex;
                while (Utils.typesContain(columns, tempCol)) {
                    //Contains suggested value, increments
                    tempCol = columnName + tempIndex;
                    tempIndex++;
                }
                addColumnName(columns, cellCount, tempCol, sheet);
            }
        } else {
            columns.add(cellCount, null);
        }
    }

    private List<Entry<String, ExcelType>> extractTypes(Sheet sheet) {

        final ArrayList<Entry<String, ExcelType>> columns = new ArrayList<Entry<String, ExcelType>>();
        int rowCount = 0;
        for (final Row row : new IteratorWrapper<Row>(sheet.iterator())) {
            if (rowCount < 1) {
                int cellCount = 0;
                //First row - get column names
                final Iterator<Cell> cellIterator = row.cellIterator();
                for (final Cell cell : new IteratorWrapper<Cell>(cellIterator)) {
                    final String columnName = Utils.cleanUp(cell.getStringCellValue());
                    addColumnName(columns, cellCount, columnName, sheet);
                    cellCount++;
                }
            } else if (rowCount < 2) {
                int cellCount = 0;
                //Second row - work out column type based on these values
                final Iterator<Cell> cellIterator = row.cellIterator();
                for (final Cell cell : new IteratorWrapper<Cell>(cellIterator)) {
                    Entry<String, ExcelType> type = columns.get(cellCount);
                    if (type != null) {
                        type.setValue(Utils.excelTypeToMySql(cell));
                    }
                    cellCount++;
                }
            }
            rowCount++;
        }
        return columns;

    }

    private static String getCreateTable(final String tableName, final List<Entry<String, ExcelType>> types) {

        if (types.size() < 1) {
            return null;
        }
        StringBuilder create = new StringBuilder();
        create.append("CREATE TABLE IF NOT EXISTS `").append(tableName).append("` (\n");
        //auto add a primary key
        create.append("\t`").append(tableName).append("ID` int(11) NOT NULL AUTO_INCREMENT, \n");

        for (Entry<String, ExcelType> entry : types) {
            if (isSet(entry)) {
                create.append("\t`").append(entry.getKey()).append("` ").append(entry.getValue().getMySqlType()).append(" DEFAULT NULL, \n");
            }
        }
        create.append("\tPRIMARY KEY (`").append(tableName).append("ID`)\n");
        create.append(");\n");
        return create.toString();
    }
}
