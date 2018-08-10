package com.evo;

import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.*;

import java.io.File;

public class ExcelReport {

    public void generate(String targetPath, Object[][] tableData) {

        try {

            SpreadsheetMLPackage pkg = SpreadsheetMLPackage.createPackage();
            WorksheetPart sheet = pkg.createWorksheetPart(new PartName("/xl/worksheets/sheet1.xml"), "Sheet1", 1);

            SheetData sheetData = sheet.getJaxbElement().getSheetData();

            for (int rowIndex = 0; rowIndex < tableData.length; ++rowIndex) {

                // Create a new row
                Row row = Context.getsmlObjectFactory().createRow();
                row.setR((long) rowIndex + 1);

                Object[] tableRow = tableData[rowIndex];
                for (int columnIndex = 0; columnIndex < tableRow.length; ++columnIndex) {
                    Object value = tableRow[columnIndex];
                    if (value instanceof String)
                        row.getC().add(newCellWithInlineString((String) value));
                    else if (value instanceof Integer)
                        row.getC().add(newIntCell(value.toString()));
                }

                // Add the row to our sheet
                sheetData.getRow().add(row);
            }

            pkg.save(new File(targetPath));

        } catch (Exception ex) {

        }
    }

    private Cell newIntCell(String content)
    {
        Cell cell = Context.getsmlObjectFactory().createCell();
        cell.setV(content);
        return cell;
    }

    private Cell newCellWithInlineString(String content) {

        CTXstringWhitespace ctx = Context.getsmlObjectFactory().createCTXstringWhitespace();
        ctx.setValue(content);

        CTRst ctrst = new CTRst();
        ctrst.setT(ctx);

        Cell newCell = Context.getsmlObjectFactory().createCell();
        newCell.setIs(ctrst);
        newCell.setT(STCellType.INLINE_STR);

        return newCell;
    }
}
