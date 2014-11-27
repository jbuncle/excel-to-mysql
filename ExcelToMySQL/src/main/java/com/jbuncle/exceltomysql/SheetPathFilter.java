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

import java.util.Arrays;
import java.util.Set;
import java.util.TreeSet;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author James Buncle
 */
public class SheetPathFilter implements SheetFilter {

    final Set<String> acceptedPaths;

    public SheetPathFilter() {
        this.acceptedPaths = new TreeSet<String>();
        this.acceptedPaths.add("*.*");
    }

    public SheetPathFilter(String... allowedPaths) {
        this.acceptedPaths = new TreeSet<String>();
        this.acceptedPaths.addAll(Arrays.asList(allowedPaths));
    }

    @Override
    public boolean accept(Sheet sheet) {
        final String sheetName = Utils.cleanUp(sheet.getSheetName());
        for (String str : this.acceptedPaths) {
            if (str.startsWith("*.")
                    || str.equals(sheetName)
                    || str.startsWith(sheetName + ".")) {
                return true;
            }
        }
        return false;
    }

    @Override
    public boolean accept(Sheet sheet, String column) {
        final String sheetName = Utils.cleanUp(sheet.getSheetName());
        //Column name path {sheet}.{column} or {sheet}.*
        for (String str : acceptedPaths) {
            if (str.equals("*.*")) {
                return true;
            }
            if (str.startsWith(sheetName)) {
                final String columnPath = str.substring(sheetName.length() + 1);
                if (columnPath.equals("*") || columnPath.equals(column)) {
                    return true;
                }
            }
        }
        return false;
    }
}
