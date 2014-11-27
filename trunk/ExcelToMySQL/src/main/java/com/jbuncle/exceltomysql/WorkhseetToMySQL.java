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

import static com.jbuncle.exceltomysql.ExcelType.BOOLEAN;
import static com.jbuncle.exceltomysql.ExcelType.DATE;
import static com.jbuncle.exceltomysql.ExcelType.NUMERIC;
import static com.jbuncle.exceltomysql.ExcelType.STRING;
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
public class WorkhseetToMySQL {

    private final Sheet sheet;
    private final List<Entry<String, ExcelType>> types;
    private final String tableName;
    private int columnOffset;
    /**
     * Tables start (index of columns headings)
     */
    private int rowOffset;

    public WorkhseetToMySQL(Sheet sheet) {
        this.sheet = sheet;
        this.types = new ArrayList<Entry<String, ExcelType>>();
        extractTypes(sheet);
        this.tableName = Utils.cleanUp(sheet.getSheetName());
        this.columnOffset = 0;
    }

    /**
     * Creates a drop statement for the worksheet
     *
     * @return
     */
    public String getDropStatement() {
        return "DROP TABLE IF EXISTS `" + tableName + "`;";
    }

    public void dropExistingTable(final Connection conn) throws SQLException {
        Utils.executeStatements(conn, getDropStatement());
    }

    public void createTable(final Connection conn) throws SQLException {
        Utils.executeStatements(conn, getCreateStatement());
    }

    public void addDataToDatabase(final Connection conn) throws SQLException {
        Utils.executeStatements(conn, getInserts());
    }

    private void addTableFromSheet(final Connection conn) throws SQLException {
        final int numRows = sheet.getPhysicalNumberOfRows();
        if (numRows < 2) {
            //Not enough or can't determine
        }
        {
            final String dropStatement = getDropStatement();
            conn.createStatement().execute(dropStatement);
            System.out.println(dropStatement);
        }
        {
            final String createStatement = getCreateStatement();
            System.out.println(createStatement);
            conn.createStatement().execute(createStatement);
        }

        for (final String insert : getInserts()) {
            System.out.println(insert);
            if (insert != null) {
                conn.createStatement().execute(insert);
            }
        }
    }

    /**
     * Generate and return MySQL Insert statements from Worksheet
     *
     * @return a list of MySQL insert commands generated from worksheet
     */
    public List<String> getInserts() {
        final List<String> updates = new LinkedList<String>();
        int rowCount = 0;
        for (final Row row : new IteratorWrapper<Row>(sheet.iterator())) {
            if (rowCount > rowOffset) {
                //Data rows
                final String insert = createInsertStatement(row);
                if (insert != null) {
                    updates.add(insert);
                }
            }
            rowCount++;
        }
        return updates;
    }

    private String createInsertStatement(final Row row) {
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

                    final String stringValue = getStringValue(sourceType.getValue(), cell);
                    if (stringValue == null) {
                        nullCount++;
                    }
                    values.append(stringValue).append(",");
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

    private String getStringValue(ExcelType type, Cell cell) {
        switch (type) {
            case DATE:
                return "'" + new SimpleDateFormat("yyyy-MM-dd HH:mm").format(cell.getDateCellValue()) + "'";
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case STRING:
                return "'" + cell.getStringCellValue().replaceAll("'", "\\\\'") + "'";
            default:
                return null;
        }
    }

    private static boolean isSet(Entry<String, ExcelType> entry) {
        return entry != null && entry.getKey() != null && entry.getValue() != null;
    }

    private void addColumnName(int cellCount, String columnName, Sheet sheet) {
        //Doesn't exist yet
        if (!Utils.typesContain(types, columnName)) {
            types.add(cellCount, new AbstractMap.SimpleEntry<String, ExcelType>(columnName, null));

        } else {
            //Already contains column
            int tempIndex = 1;

            String tempCol = columnName + tempIndex;
            while (Utils.typesContain(types, tempCol)) {
                //Contains suggested value, increments
                tempCol = columnName + tempIndex;
                tempIndex++;
            }
            addColumnName(cellCount, tempCol, sheet);
        }
    }

    private void extractTypes(Sheet sheet) {


        {
            int cellCount = columnOffset;
            //First row - get column names
            final Iterator<Cell> cellIterator = sheet.getRow(rowOffset).cellIterator();

            for (final Cell cell : new IteratorWrapper<Cell>(cellIterator)) {
                final String columnName = Utils.cleanUp(cell.getStringCellValue());
                addColumnName(cellCount, columnName, sheet);
                cellCount++;
            }
        }
        {
            int cellCount = columnOffset;
            //Second row - work out column type based on these values
            final Iterator<Cell> cellIterator = sheet.getRow(rowOffset + 1).cellIterator();
            for (final Cell cell : new IteratorWrapper<Cell>(cellIterator)) {
                final Entry<String, ExcelType> type = types.get(cellCount);
                if (type != null) {
                    type.setValue(Utils.excelTypeToMySql(cell));
                }
                cellCount++;
            }
        }
    }

    /**
     * Generates a MySQL table create statement which is capable of holding the
     * worksheet data
     *
     * @return the MySQL Table create statement, or null if unable to make the
     * table create statement
     */
    public String getCreateStatement() {

        if (types.size() < 1) {
            return null;
        }
        final StringBuilder create = new StringBuilder();
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
