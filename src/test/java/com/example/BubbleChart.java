package com.example;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.help.Utils;
import com.linq.Common;
import org.testng.annotations.Test;

/**
 * @Description
 * @createDate 2019/1/10
 **/
public class BubbleChart {

    @Test
    public void test() throws Exception {
        //ExStart:BubbleChart
        // The path to the documents directory.

        String dataDir = Utils.getDataDir(BubbleChart.class);

        String fileName = "BubbleChart.docx";
        // Load the template document.
        Document doc = new Document();
//         Document doc = new Document(dataDir + fileName);

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        engine.buildReport(doc, Common.GetContracts(), "contracts");

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);
        //ExEnd:BubbleChart

        System.out.println("\nBubble chart template document is populated with the data about managers.\nFile saved at " + dataDir);
    }
}
