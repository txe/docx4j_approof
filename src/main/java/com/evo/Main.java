package com.evo;

import java.util.*;

public class Main {

    public static void main(String[] args) {

        //testRealWordReport();
        //testSimpleWordReport();
        testExcelReport();
    }

    private static void testRealWordReport() {

        CandidateReportCtx ctx = new CandidateReportCtx();
        ctx.CandidateName = "John Doe";
        ctx.CurrentDate = new Date().toString();
        ctx.MaritalStatus = "kids";

        ctx.CareerList = new ArrayList<CandidateReportCtx.RowData>();
        ctx.CareerList.add(ctx.createRowData("1992", "Moscow", "some note"));
        ctx.CareerList.add(ctx.createRowData("1995", "Moscow", "second note"));

        ctx.CareerExpList = new ArrayList<CandidateReportCtx.RowData>();
        ctx.CareerExpList.add(ctx.createRowData("1992", "Manager", ""));
        ctx.CareerExpList.add(ctx.createRowData("1995", "Programmer", ""));

        ctx.EduList = new ArrayList<CandidateReportCtx.RowData>();
        ctx.EduList.add(ctx.createRowData("1992-1997", "MGU", "some note"));

        ctx.EduExpList = new ArrayList<CandidateReportCtx.RowData>();
        ctx.EduExpList.add(ctx.createRowData("1998", "Course - Manager", "sp1", "more notes"));
        ctx.EduExpList.add(ctx.createRowData("1999", "Course - Analize", "sp2", "second note"));

        new WordReport().generate2("F:/candidate_template.docx", "F:/candidate.docx", ctx);
    }

    private static void testSimpleWordReport() {

        // template uses $(myField), but we pass it as myField
        HashMap<String, String> placeholders = new HashMap<String, String>() {
            {
                this.put("Name", "Company Name here...");
                this.put("colour", "green");
                this.put("placeholder", "Hmmm lemme see");
            }
        };

        new WordReport().generate("F:/template_1.docx", "F:/result_1.docx", placeholders);
    }

    private static void testExcelReport() {

        // make sure rows have the same amount of cells
        Object[] row1 = {"cell", 123, 234.5 };
        Object[] row2 = {"hi", "world", "" };
        Object[] row3 = {"", "", ""};
        Object[] row4 = {456, "789", "" };
        Object[][] tableData = {row1, row2, row3, row4 };

        double[] widths = {10, 20, 30};

        new ExcelReport().generate("F:/result_2.xlsx", tableData, widths);
    }

}
