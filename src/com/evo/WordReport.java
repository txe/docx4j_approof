package com.evo;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

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
}
