package Calculations;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class Payout {

    public static void main(String[] args) {
        // ✅ Update this file path to the correct location of your Excel file
    	String filePath = "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\EveryTicketCalculation\\2025-10-01 to 2025-10-15_Order_History.xlsx";



        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            List<Map<String, String>> records = new ArrayList<>();

            // ✅ Read header row
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("Error: Header row is missing in Excel file.");
                return;
            }

            int numColumns = headerRow.getLastCellNum();

            // ✅ Read all data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> record = new HashMap<>();

                for (int j = 0; j < numColumns; j++) {
                    Cell headerCell = headerRow.getCell(j);
                    if (headerCell == null) continue;

                    String header = headerCell.getStringCellValue().trim();
                    Cell cell = row.getCell(j);
                    String value = (cell == null) ? "" : getCellValueAsString(cell);
                    record.put(header, value);
                }

                records.add(record);
            }

            // ✅ Group by "Date"
            Map<String, List<Map<String, String>>> dateGroups = records.stream()
                    .filter(r -> r.get("Date") != null && !r.get("Date").isEmpty())
                    .collect(Collectors.groupingBy(r -> r.get("Date")));

            // ✅ Iterate each date group
            for (String date : dateGroups.keySet()) {
                System.out.println("\n📅 Date: " + date);

                List<Map<String, String>> dateRecords = dateGroups.get(date);

                // ✅ Group by "Payment Method"
                Map<String, List<Map<String, String>>> paymentGroups = dateRecords.stream()
                        .filter(r -> r.get("Payment Method") != null && !r.get("Payment Method").isEmpty())
                        .collect(Collectors.groupingBy(r -> r.get("Payment Method")));

                for (String paymentMethod : paymentGroups.keySet()) {
                    List<Map<String, String>> methodRecords = paymentGroups.get(paymentMethod);

                    int totalRecords = methodRecords.size();
                    double totalTicketPrice = 0;

                    for (Map<String, String> rec : methodRecords) {
                        try {
                            totalTicketPrice += Double.parseDouble(rec.get("Ticket Price"));
                        } catch (NumberFormatException e) {
                            // Ignore invalid numeric values
                        }
                    }

                    System.out.println("💳 Payment Method: " + paymentMethod);
                    System.out.println("🧾 Total Records: " + totalRecords);
                    System.out.println("💰 Total Ticket Price: " + totalTicketPrice);
                    System.out.println();
                }
            }

        } catch (IOException e) {
            System.err.println("❌ Error reading Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

   

	// ✅ Helper method to safely convert cell values to String
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BLANK:
            default:
                return "";
        }
    }
}
