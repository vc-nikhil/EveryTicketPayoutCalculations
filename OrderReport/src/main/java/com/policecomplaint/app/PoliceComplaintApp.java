package com.policecomplaint.app;

import org.apache.poi.xwpf.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

public class PoliceComplaintApp {

    public static void main(String[] args) {
        XWPFDocument document = new XWPFDocument();

        try {
            // === Header: To, Police Station ===
            addParagraph(document, "To,", ParagraphAlignment.LEFT, true, 12);
            addParagraph(document, "Police Station In-Charge,", ParagraphAlignment.LEFT, true, 12);
            addParagraph(document, "Police Station Name: PS / Connaught Place", ParagraphAlignment.LEFT, false, 12);
            addParagraph(document, "District: New Delhi", ParagraphAlignment.LEFT, false, 12);
            addParagraph(document, "State: Delhi", ParagraphAlignment.LEFT, false, 12);
            addParagraph(document, "Pin Code: 110001", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            // === Date ===
            addParagraph(document, "Date: 11-11-25, 11:40 AM", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            // === Subject ===
            addParagraph(document, "Subject: Theft of iPhone 16 at Rajiv Chowk on 10th November 2025", 
                         ParagraphAlignment.LEFT, true, 12);
            addEmptyLine(document);

            // === Main Complaint Body ===
            String intro = "I, [Please Insert Your Name], son/daughter/wife of [Please Insert Applicable Name], " +
                           "with contact address at [Please Insert Your Contact Address], and contact mobile number of " +
                           "[Please Insert Your Contact Mobile Number] hereby submit this police complaint to bring to your " +
                           "attention a matter that requires your investigation and appropriate action.";
            addParagraph(document, intro, ParagraphAlignment.BOTH, false, 12);
            addEmptyLine(document);

            addParagraph(document, "The details of the incident are as follows:", ParagraphAlignment.LEFT, true, 12);
            addEmptyLine(document);

            // === Section I: Incident Details ===
            addHeading(document, "Section I: Incident Details:", true);
            addBulletList(document, Arrays.asList(
                "Date: 10-11-2025",
                "Time: 03:00 PM",
                "Location: rajiv chowk",
                "Description: On 10th November 2025, at approximately 3:00 PM, my iPhone 16, with IMEI 353838107706281, " +
                "was stolen from a cab at Rajiv Chowk. There were no witnesses to this incident, nor is there any video " +
                "recording, photograph, screenshot, or other supporting evidence available. I am reporting this incident " +
                "to initiate legal proceedings."
            ));

            addEmptyLine(document);
            addLabelValue(document, "Details of Victim(s):", "[Please Enter Name of Victim]");
            addLabelValue(document, "Address of Victim:", "[Not written/blank]");
            addEmptyLine(document);

            addLabelValue(document, "Details of the Accused (if known):", "[Not written/blank]");
            addLabelValue(document, "Address of Victim:", "[Not written/blank]");
            addLabelValue(document, "Mobile Number of Accused:", "[Not written/blank]");
            addEmptyLine(document);

            addLabelValue(document, "Witnesses (if any):", "[Not written/blank]");
            addLabelValue(document, "Address of Witnesses :", "[Not written/blank]");
            addLabelValue(document, "Mobile Number of Witnesses :", "[Not written/blank]");
            addEmptyLine(document);

            // === Section II ===
            addHeading(document, "Section II: Details of Loss, Damage, or Injury:", true);
            addLabelValue(document, "Description of Property (if any):", "[Not written/blank]");
            addLabelValue(document, "Physical Injury (if any):", "[Not written/blank]");
            addLabelValue(document, "Type:", "[Not written/blank]");
            addLabelValue(document, "Medical Documentation:", "[Not written/blank]");
            addLabelValue(document, "When Visited Doctor/Hospital after Incident:", "[Not written/blank] (To be customised to complaint)");
            addEmptyLine(document);

            // === Section III ===
            addHeading(document, "Section III: Was the incident captured on the following(Yes/No):", true);
            addLabelValue(document, "", "[Not written/blank]");
            addLabelValue(document, "CCTV:", "[Not written/blank]");
            addParagraph(document, "(PLEASE MAKE LIST)", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            // === Section IV ===
            addHeading(document, "Section IV: Previous Communication to Police or any relevant body (if any):", true);
            addLabelValue(document, "", "[Not written/blank]");
            addEmptyLine(document);

            // === Section V ===
            addHeading(document, "Section V: Any Other Important Information:", true);
            addLabelValue(document, "", "[Not written/blank]");
            addEmptyLine(document);

            // === Section VI: Supporting Evidence ===
            addHeading(document, "Section VI: Supporting Evidence:", true);
            addParagraph(document, "I am submitting the following documents/evidence that are attached in the Annexure(s) in support of my complaint:", 
                         ParagraphAlignment.LEFT, false, 12);

            // Table for Evidence
            XWPFTable table = document.createTable(3, 3);
            table.setWidth("100%");
            setTableRow(table.getRow(0), new String[]{"S.No / Annexure Number", "Type of Evidence", "Date of Evidence"}, true);
            setTableRow(table.getRow(1), new String[]{"1", "[Not written/blank]", "[Not written/blank]"}, false);
            setTableRow(table.getRow(2), new String[]{"2", "[Not written/blank]", "[Not written/blank]"}, false);
            addEmptyLine(document);

            // === Contact Preference ===
            addParagraph(document, "You can contact me at the mobile number listed above and my preferred time to connect is from " +
                         "[Please Enter Start Time] to [Please Enter End Time].", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            // === Declaration ===
            addParagraph(document, "I declare that the facts stated in this complaint together with its annexures are true and correct " +
                         "to the best of my knowledge and belief. I am willing to provide further information, documentation, " +
                         "and cooperate fully with the investigation as and when required.", 
                         ParagraphAlignment.BOTH, false, 12);
            addEmptyLine(document);

            addParagraph(document, "Thank you for your attention and assistance.", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            addParagraph(document, "Yours sincerely,", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);
            addEmptyLine(document);
            addEmptyLine(document);

            addParagraph(document, "[Signature]", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            addLabelValue(document, "Full Name:", "[Please Enter Your Full Name]");
            addLabelValue(document, "Address:", "[Please Enter Your Address]");
            addLabelValue(document, "Pincode:", "[Please Enter Your PinCode]");
            addLabelValue(document, "Email:", "[Please Enter Your Email]");
            addLabelValue(document, "Mobile Contact:", "[Please Enter Your Mobile Number]");
            addEmptyLine(document);

            addParagraph(document, "Date: 11/11/2025", ParagraphAlignment.LEFT, false, 12);
            addEmptyLine(document);

            // === Annexure ===
            addHeading(document, "Annexure – Detailed Narrative of Events", true);
            String annexure = "On 10th November 2025, at approximately 3:00 PM, I experienced the theft of my mobile phone while traveling in a cab. " +
                              "The incident specifically occurred at Rajiv Chowk, a prominent location, where my valuable device was illicitly taken. " +
                              "I became aware of the theft shortly after, realizing that my iPhone 16 was no longer in my possession or within my immediate belongings. " +
                              "This unforeseen and distressing event has caused considerable inconvenience and concern regarding the security of my personal property.\n\n" +
                              "The stolen item is identified as an iPhone 16, a high-end smartphone essential for my daily communication and various personal and professional tasks. " +
                              "This device carries a unique International Mobile Equipment Identity (IMEI) number, which is 353838107706281. " +
                              "This specific IMEI is crucial for its identification and for any potential recovery efforts by law enforcement agencies. " +
                              "The loss of this device represents a significant financial setback and disruption to my routine.\n\n" +
                              "Regarding the circumstances of the theft, I must report that there were no direct witnesses present who could provide an account of the incident " +
                              "or identify the perpetrator. Furthermore, I do not possess any video recordings, photographs, screenshots, or other forms of supporting evidence " +
                              "such as messages, emails, or documents that might have captured the relevant details of this theft. Despite a thorough review of the situation, " +
                              "no additional evidence has come to light that could assist in the immediate investigation.\n\n" +
                              "In light of this unfortunate incident, I am formally lodging this complaint to initiate a thorough investigation into the theft of my iPhone 16. " +
                              "My primary objective is to seek the recovery of my stolen property and to ensure that the individual responsible for this criminal act is identified " +
                              "and brought to justice. I request the police authorities, specifically those operating within the 110001 pincode jurisdiction, to take prompt and " +
                              "necessary action based on the details provided in this report.";

            addParagraph(document, annexure, ParagraphAlignment.BOTH, false, 12);
            addEmptyLine(document);

            // === Probable Sections ===
            addHeading(document, "Probable Sections", true);
            addBulletList(document, Arrays.asList(
                "BNS SECTION: 303. Theft.",
                "BNS SECTION: 305. Theft in dwelling house or means of transportation or place of worship."
            ));

            addEmptyLine(document);
            addParagraph(document, "Disclaimer: The sections provided are for informational purposes only and are intended as general suggestions. " +
                         "For specific guidance, please consult the appropriate professionals or refer to the relevant official sources.", 
                         ParagraphAlignment.LEFT, false, 11);

            // === Save File ===
            String outputPath = "Police_Complaint_iPhone_Theft.docx";
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }
            System.out.println("Document generated successfully: " + outputPath);

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                document.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // === Helper Methods ===
    private static void addParagraph(XWPFDocument doc, String text, ParagraphAlignment alignment, boolean bold, int fontSize) {
        XWPFParagraph para = doc.createParagraph();
        para.setAlignment(alignment);
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setBold(bold);
        run.setFontSize(fontSize);
        run.setFontFamily("Times New Roman");
    }

    private static void addHeading(XWPFDocument doc, String text, boolean bold) {
        addParagraph(doc, text, ParagraphAlignment.LEFT, bold, 12);
    }

    private static void addEmptyLine(XWPFDocument doc) {
        XWPFParagraph para = doc.createParagraph();
        para.createRun();
    }

    private static void addBulletList(XWPFDocument doc, java.util.List<String> items) {
        for (String item : items) {
            XWPFParagraph para = doc.createParagraph();
            para.setIndentationLeft(360); // 0.5 inch
            XWPFRun run = para.createRun();
            run.setText("• " + item);
            run.setFontSize(12);
            run.setFontFamily("Times New Roman");
        }
    }

    private static void addLabelValue(XWPFDocument doc, String label, String value) {
        XWPFParagraph para = doc.createParagraph();
        XWPFRun run = para.createRun();
        run.setText(label + " ");
        run.setBold(true);
        run.setFontSize(12);
        run = para.createRun();
        run.setText(value);
        run.setBold(false);
        run.setFontSize(12);
    }

    private static void setTableRow(XWPFTableRow row, String[] cells, boolean bold) {
        for (int i = 0; i < cells.length; i++) {
            XWPFTableCell cell = row.getCell(i);
            if (cell == null) cell = row.addNewTableCell();
            XWPFParagraph para = cell.getParagraphs().get(0);
            if (para.getRuns().size() == 0) {
                XWPFRun run = para.createRun();
                run.setText(cells[i]);
                run.setBold(bold);
                run.setFontSize(11);
                run.setFontFamily("Times New Roman");
            } else {
                para.getRuns().get(0).setText(cells[i], 0);
                para.getRuns().get(0).setBold(bold);
            }
            cell.setWidth("33%");
        }
    }
}