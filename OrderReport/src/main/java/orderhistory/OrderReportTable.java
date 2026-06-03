package orderhistory;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.time.LocalDate;
import java.util.*;

public class OrderReportTable {

    private static class Stats {
        int count = 0;
        double totalPrice = 0.0;

        void add(double price) {
            count++;
            totalPrice += price;
        }

        String format() {
            return String.format("%d (%.0f)", count, totalPrice);
        }
    }

    private static final String[] PAYMENT_ORDER = {
            "UPI", "CC", "UPIPPI", "UPICC", "NB", "UPI SALE", "CARD", "CASH"
    };

    public static void main(String[] args) {

        String filePath = "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\OrderReport\\2026-05-01 to 2026-05-15_Order_History.xlsx";

        Map<LocalDate, Map<String, Stats>> dailyData = new TreeMap<>();
        Map<String, Stats> grandTotals = new HashMap<>();
        Map<LocalDate, Stats> dayTotals = new HashMap<>();
        Stats overallTotal = new Stats();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet("Order History");
            if (sheet == null) {
                System.out.println("Sheet not found!");
                return;
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // ✅ DATE (handles numeric + string)
                Cell dateCell = row.getCell(8);
                if (dateCell == null) continue;

                LocalDate date = null;

                try {
                    if (dateCell.getCellType() == CellType.NUMERIC &&
                            DateUtil.isCellDateFormatted(dateCell)) {

                        date = dateCell.getLocalDateTimeCellValue().toLocalDate();

                    } else if (dateCell.getCellType() == CellType.STRING) {

                        String dateStr = dateCell.getStringCellValue().trim();
                        if (!dateStr.isEmpty()) {
                            date = LocalDate.parse(dateStr.substring(0, 10));
                        }
                    }
                } catch (Exception e) {
                    continue;
                }

                if (date == null) continue;

                // ✅ PAYMENT
                Cell payCell = row.getCell(9);
                String payment = "Unknown";

                if (payCell != null) {
                    if (payCell.getCellType() == CellType.STRING) {
                        payment = payCell.getStringCellValue().trim();
                    } else if (payCell.getCellType() == CellType.NUMERIC) {
                        payment = String.valueOf((int) payCell.getNumericCellValue());
                    }
                }

                // ✅ PRICE
                Cell priceCell = row.getCell(5);
                double price = 0;

                if (priceCell != null) {
                    try {
                        if (priceCell.getCellType() == CellType.NUMERIC) {
                            price = priceCell.getNumericCellValue();
                        } else if (priceCell.getCellType() == CellType.STRING) {
                            price = Double.parseDouble(
                                    priceCell.getStringCellValue().replace(",", "").trim()
                            );
                        }
                    } catch (Exception ignored) {}
                }

                // ✅ AGGREGATION
                dailyData
                        .computeIfAbsent(date, k -> new HashMap<>())
                        .computeIfAbsent(payment, k -> new Stats())
                        .add(price);

                dayTotals.computeIfAbsent(date, k -> new Stats()).add(price);
                grandTotals.computeIfAbsent(payment, k -> new Stats()).add(price);
                overallTotal.add(price);
            }

        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        // ================= TABLE BUILD =================
        List<String[]> table = new ArrayList<>();

        // Header
        String[] header = new String[PAYMENT_ORDER.length + 2];
        header[0] = "Date";
        System.arraycopy(PAYMENT_ORDER, 0, header, 1, PAYMENT_ORDER.length);
        header[header.length - 1] = "Total";
        table.add(header);

        // Daily rows
        for (LocalDate date : dailyData.keySet()) {
            Map<String, Stats> dayMap = dailyData.get(date);
            Stats dayTotal = dayTotals.get(date);

            String[] row = new String[PAYMENT_ORDER.length + 2];
            row[0] = date.toString();

            for (int i = 0; i < PAYMENT_ORDER.length; i++) {
                Stats s = dayMap.get(PAYMENT_ORDER[i]);
                row[i + 1] = (s != null) ? s.format() : "0 (0)";
            }

            row[row.length - 1] = dayTotal.format();
            table.add(row);
        }

        // GRAND TOTAL
        String[] grandRow = new String[PAYMENT_ORDER.length + 2];
        grandRow[0] = "GRAND TOTAL";

        for (int i = 0; i < PAYMENT_ORDER.length; i++) {
            Stats s = grandTotals.getOrDefault(PAYMENT_ORDER[i], new Stats());
            grandRow[i + 1] = s.format();
        }

        grandRow[grandRow.length - 1] = overallTotal.format();
        table.add(grandRow);

        // FINAL TOTAL
        String totalStr = overallTotal.format();
        String[] finalRow = new String[PAYMENT_ORDER.length + 2];
        finalRow[0] = "TOTAL";
        Arrays.fill(finalRow, 1, finalRow.length, totalStr);
        table.add(finalRow);

        // PRINT TABLE
        printTable(table);
    }

    // ================= PRINTING =================

    private static void printTable(List<String[]> rows) {
        int cols = rows.get(0).length;
        int[] widths = new int[cols];

        // ✅ FIX: include bold text in width calculation
        for (String[] row : rows) {
            boolean isBold = row[0].equals("GRAND TOTAL") || row[0].equals("TOTAL");

            for (int c = 0; c < cols; c++) {
                String cell = row[c];
                if (isBold) {
                    cell = "**" + cell + "**";
                }
                widths[c] = Math.max(widths[c], cell.length());
            }
        }

        // Border line
        StringBuilder line = new StringBuilder("+");
        for (int w : widths) {
            line.append("-".repeat(w + 2)).append("+");
        }

        System.out.println(line);

        // Header
        printRow(rows.get(0), widths);
        System.out.println(line);

        // Data rows
        for (int i = 1; i < rows.size(); i++) {
            printRow(rows.get(i), widths);

            if (rows.get(i)[0].equals("GRAND TOTAL")) {
                System.out.println(line);
            }
        }

        System.out.println(line);
    }

    private static void printRow(String[] row, int[] widths) {
        System.out.print("|");

        boolean isBold = row[0].equals("GRAND TOTAL") || row[0].equals("TOTAL");

        for (int c = 0; c < row.length; c++) {
            String cell = row[c];

            if (isBold) {
                cell = "**" + cell + "**";
            }

            if (c == 0) {
                System.out.printf(" %-"+widths[c]+"s |", cell);
            } else {
                System.out.printf(" %"+widths[c]+"s |", cell);
            }
        }

        System.out.println();
    }
}