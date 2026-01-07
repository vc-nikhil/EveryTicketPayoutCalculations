package payu;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class PayUReportGenerator {

    public static void main(String[] args) {
        String filePath = "transaction_8765164_1762172977119.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("Header row is empty!");
                return;
                //demo //
            }

            // Find column indices
            int dateCol = -1, modeCol = -1, amountCol = -1;
            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue().trim();
                if ("addedon".equals(header)) dateCol = cell.getColumnIndex();
                if ("mode".equals(header)) modeCol = cell.getColumnIndex();
                if ("amount".equals(header)) amountCol = cell.getColumnIndex();
            }

            if (dateCol == -1 || modeCol == -1 || amountCol == -1) {
                System.out.println("Required columns not found!");
                return;
            }

            // Map: Date -> (Mode -> Stats)
            Map<String, Map<String, Stats>> dateModeMap = new TreeMap<>();
            Map<String, Stats> dailyTotals = new TreeMap<>(); // For per-day total
            Stats grandTotal = new Stats();

            SimpleDateFormat inputDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            SimpleDateFormat outputDateFormat = new SimpleDateFormat("yyyy-MM-dd");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell dateCell = row.getCell(dateCol);
                Cell modeCell = row.getCell(modeCol);
                Cell amountCell = row.getCell(amountCol);

                if (dateCell == null || modeCell == null || amountCell == null) continue;

                String rawDate = getCellValueAsString(dateCell);
                String mode = getCellValueAsString(modeCell).trim().toUpperCase();
                double amount = 0.0;

                try {
                    amount = Double.parseDouble(getCellValueAsString(amountCell));
                } catch (Exception e) {
                    continue;
                }

                Date date;
                try {
                    date = inputDateFormat.parse(rawDate);
                } catch (Exception e) {
                    continue;
                }

                String formattedDate = outputDateFormat.format(date);
                double roundedAmount = rounded(amount);

                // Update date → mode map
                Stats modeStats = dateModeMap
                        .computeIfAbsent(formattedDate, k -> new HashMap<>())
                        .computeIfAbsent(mode, k -> new Stats());
                modeStats.add(roundedAmount);

                // Update daily total
                Stats dayTotal = dailyTotals.computeIfAbsent(formattedDate, k -> new Stats());
                dayTotal.add(roundedAmount);

                // Update grand total
                grandTotal.add(roundedAmount);
            }

            // Print results
            System.out.println("=== PayU Transaction Summary by Date and Payment Method ===\n");

            for (Map.Entry<String, Map<String, Stats>> dateEntry : dateModeMap.entrySet()) {
                String date = dateEntry.getKey();
                Map<String, Stats> modeMap = dateEntry.getValue();
                Stats dayTotal = dailyTotals.get(date);

                System.out.println("Date: " + date);
                System.out.println();

                for (Map.Entry<String, Stats> modeEntry : modeMap.entrySet()) {
                    String mode = modeEntry.getKey();
                    Stats stats = modeEntry.getValue();
                    System.out.printf("Payment Method: %s%n", mode);
                    System.out.printf("  Total Records: %d%n", stats.count);
                    System.out.printf("  Total Ticket Price: ₹%.2f%n", stats.totalAmount);
                    System.out.println();
                }

                // Daily Total
                System.out.printf("TOTAL FOR %s: %d transactions | ₹%.2f%n", date, dayTotal.count, dayTotal.totalAmount);
                System.out.println("------------------------------------------------------------");
            }

            // Grand Total
            System.out.printf("%nGRAND TOTAL: %d transactions | ₹%.2f%n", grandTotal.count, grandTotal.totalAmount);
            System.out.println("================================================================");

        } catch (Exception e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static double rounded(double value) {
        return Math.round(value * 100.0) / 100.0;
    }

    static class Stats {
        int count = 0;
        double totalAmount = 0.0;

        void add(double amount) {
            count++;
            totalAmount += amount;
        }
    }
}