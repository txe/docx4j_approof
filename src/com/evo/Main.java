package com.evo;

import java.util.HashMap;

public class Main {

    public static void main(String[] args) {

        //testWordReport();
        //testExcelReport();
    }

    private static void testWordReport() {

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

        Object[] row1 = {"cell", 123 };
        Object[] row2 = {"hi", "world" };
        Object[] row3 = {};
        Object[] row4 = {456, "789" };
        Object[][] tableData = {row1, row2, row3, row4 };

        new ExcelReport().generate("F:/result_2.xlsx", tableData);
    }
}
