package com.evo;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.JaxbValidationEventHandler;
import org.docx4j.model.properties.run.FontSize;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.Styles;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.wml.Fonts;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.*;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import javax.xml.transform.stream.StreamSource;
import java.io.File;
import java.io.StringReader;

public class ExcelReport {

    private static final String MAIN_SCHEMA = "xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    public void generate(String targetPath, Object[][] tableData, double[] colWidths) {

        try {

            SpreadsheetMLPackage pkg = SpreadsheetMLPackage.createPackage();

            Styles styles = new Styles(new PartName("/xl/styles.xml"));
            styles.setJaxbElement(createStylesheet());
            //pkg.addTargetPart(styles);
            pkg.getWorkbookPart().addTargetPart(styles);

            WorksheetPart sheet = pkg.createWorksheetPart(new PartName("/xl/worksheets/sheet1.xml"), "Sheet1", 1);


            Cols columns = new Cols();
            for (int i = 0; i < colWidths.length; ++i)
                columns.getCol().add(createColumn(i + 1, colWidths[i]));
            sheet.getJaxbElement().getCols().add(columns);

            SheetData sheetData = sheet.getJaxbElement().getSheetData();

            for (int rowIndex = 0; rowIndex < tableData.length; ++rowIndex) {

                // Create a new row
                Row row = Context.getsmlObjectFactory().createRow();
                row.setR((long) rowIndex + 1);

                Object[] tableRow = tableData[rowIndex];
                for (int columnIndex = 0; columnIndex < tableRow.length; ++columnIndex) {
                    Object value = tableRow[columnIndex];
                    if (value instanceof Integer)
                        row.getC().add(newNumberCell(value.toString()));
                    else if (value instanceof Double)
                        row.getC().add(newNumberCell(value.toString()));
                    else // for string special case
                        row.getC().add(newCellWithInlineString(value.toString()));
                }

                // Add the row to our sheet
                sheetData.getRow().add(row);
            }

            pkg.save(new File(targetPath));

        } catch (Exception ex) {
            ex = ex;
        }
    }

    private Cell newNumberCell(String content)
    {
        Cell cell = Context.getsmlObjectFactory().createCell();
        cell.setV(content);
        cell.setS((long)1);
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
        newCell.setS((long)1);
        return newCell;
    }

    private Col createColumn(int index, double width) {
        // width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}] / {Maximum Digit Width} * 256) / 256
        Col column = new Col();
        column.setMin(index);
        column.setMax(index);
        column.setWidth(width);
        column.setCustomWidth(true);
        return column;
    }

    private static CTStylesheet createStylesheet() throws JAXBException {
        String openXML =
               "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<numFmts count=\"1\">"
                + "<numFmt formatCode=\"General\" numFmtId=\"164\"/>"
                +"</numFmts>"
                + "<fonts count=\"4\">"
                + "<font>"
                + "<sz val=\"10.0\"/>"
                + "<name val=\"Arial\"/>"
                + "<family val=\"2\"/>"
                +"</font>"
                + "<font>"
                + "<sz val=\"10.0\"/>"
                + "<name val=\"Arial\"/>"
                + "<family val=\"0\"/>"
                +"</font>"
                + "<font>"
                + "<sz val=\"10.0\"/>"
                + "<name val=\"Arial\"/>"
                + "<family val=\"0\"/>"
                +"</font>"
                + "<font>"
                + "<sz val=\"10.0\"/>"
                + "<name val=\"Arial\"/>"
                + "<family val=\"0\"/>"
                +"</font>"
                +"</fonts>"
                + "<fills count=\"2\">"
                + "<fill>"
                + "<patternFill patternType=\"none\"/>"
                +"</fill>"
                + "<fill>"
                + "<patternFill patternType=\"gray125\"/>"
                +"</fill>"
                +"</fills>"
                + "<borders count=\"2\">"
                + "<border diagonalDown=\"false\" diagonalUp=\"false\">"
                + "<left/>"
                + "<right/>"
                + "<top/>"
                + "<bottom/>"
                + "<diagonal/>"
                +"</border>"
                + "<border diagonalDown=\"false\" diagonalUp=\"false\">"
                + "<left style=\"hair\"/>"
                + "<right style=\"hair\"/>"
                + "<top style=\"hair\"/>"
                + "<bottom style=\"hair\"/>"
                + "<diagonal/>"
                +"</border>"
                +"</borders>"
                + "<cellStyleXfs count=\"20\">"
                + "<xf applyAlignment=\"true\" applyBorder=\"true\" applyFont=\"true\" applyProtection=\"true\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\">"
                + "<alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/>"
                + "<protection hidden=\"false\" locked=\"true\"/>"
                +"</xf>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"43\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"41\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"44\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"42\"/>"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"9\"/>"
                +"</cellStyleXfs>"
                + "<cellXfs count=\"2\">"
                + "<xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\" xfId=\"0\">"
                + "<alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/>"
                + "<protection hidden=\"false\" locked=\"true\"/>"
                +"</xf>"
                + "<xf applyAlignment=\"false\" applyBorder=\"true\" applyFont=\"false\" applyProtection=\"false\" borderId=\"1\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\" xfId=\"0\">"
                + "<alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/>"
                + "<protection hidden=\"false\" locked=\"true\"/>"
                +"</xf>"
                +"</cellXfs>"
                + "<cellStyles count=\"6\">"
                + "<cellStyle builtinId=\"0\" customBuiltin=\"false\" name=\"Normal\" xfId=\"0\"/>"
                + "<cellStyle builtinId=\"3\" customBuiltin=\"false\" name=\"Comma\" xfId=\"15\"/>"
                + "<cellStyle builtinId=\"6\" customBuiltin=\"false\" name=\"Comma [0]\" xfId=\"16\"/>"
                + "<cellStyle builtinId=\"4\" customBuiltin=\"false\" name=\"Currency\" xfId=\"17\"/>"
                + "<cellStyle builtinId=\"7\" customBuiltin=\"false\" name=\"Currency [0]\" xfId=\"18\"/>"
                + "<cellStyle builtinId=\"5\" customBuiltin=\"false\" name=\"Percent\" xfId=\"19\"/>"
                +"</cellStyles>"
                +"</styleSheet>";


        Unmarshaller u = Context.jcSML.createUnmarshaller();
        u.setEventHandler(new JaxbValidationEventHandler());
        JAXBElement o = (JAXBElement)u.unmarshal(new StreamSource(new StringReader(openXML)));
        CTStylesheet stylesheet = (CTStylesheet)o.getValue();// XmlUtils.unmarshalString(openXML);
        return stylesheet;
    }
}
