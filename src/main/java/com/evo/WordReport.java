package com.evo;

import java.io.*;
import java.util.HashMap;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import org.wickedsource.docxstamper.DocxStamper;
import org.wickedsource.docxstamper.DocxStamperConfiguration;

public class WordReport {


    public void generate(String templatePath, String targetPath, HashMap<String, String> placeholders) {

        try {
            WordprocessingMLPackage template = WordprocessingMLPackage
                    .load(new FileInputStream(new File(templatePath)));

            VariablePrepare.prepare(template);
            MainDocumentPart documentPart = template.getMainDocumentPart();
            documentPart.variableReplace(placeholders);

            template.save(new File(targetPath));
        }
        catch (Exception ex)
        {
        }
    }

    public void generate2(String templatePath, String targetPath, Object context) {

        try {
            DocxStamper stamper = new DocxStamper(new DocxStamperConfiguration());

            InputStream template = new FileInputStream(new File(templatePath));
            OutputStream out = new FileOutputStream(new File(targetPath));
            stamper.stamp(template, context, out);

        } catch (Exception ex) {

        }
    }

}
