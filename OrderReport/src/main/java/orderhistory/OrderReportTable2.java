package orderhistory;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.time.LocalDate;
import java.util.*;

public class OrderReportTable2 {

    static class Stats {
        int count = 0;
        double amount = 0;
        double discount = 0;

        void add(double price, double disc) {
            count++;
            amount += price;
            discount += disc;
        }

        String format() {
            return String.format("%d(%.0f|%.0f)", count, amount, discount);
        }
    }

    private static final String[] PAYMENT_ORDER = {
            "UPI",
            "CC",
            "DC",
            "UPIPPI",
            "UPICC",
            "NB",
            "UPI SALE",
            "CARD",
            "CASH"
    };

    private static final int MAX_WIDTH = 18;

    public static void main(String[] args) {

        String filePath =
                "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\OrderReport\\2026-05-16 to 2026-05-31_Order_History.xlsx";

        Map<LocalDate, Map<String, Stats>> daily = new TreeMap<>();
        Map<String, Stats> grand = new HashMap<>();
        Map<LocalDate, Stats> dayTotal = new HashMap<>();
        Stats overall = new Stats();

        LocalDate minDate = null;
        LocalDate maxDate = null;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            Row header = sheet.getRow(0);
            Map<String, Integer> cols = new HashMap<>();

            for (Cell c : header) {
                cols.put(
                        c.getStringCellValue().toLowerCase().trim(),
                        c.getColumnIndex()
                );
            }

            System.out.println("Available Columns:");
            for (String k : cols.keySet()) {
                System.out.println(k);
            }

            int dateCol = findColumnFlexible(cols, "purchased", "date");
            int payCol = findColumnFlexible(cols, "payment");
            int priceCol = findColumnFlexible(cols, "ticket", "price", "amount");
            int discountCol = findColumnFlexible(cols, "discount", "disc");

            System.out.println(
                    "\nUsing columns -> Date:" + dateCol +
                    " Payment:" + payCol +
                    " Price:" + priceCol +
                    " Discount:" + discountCol
            );

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {

                Row row = sheet.getRow(r);
                if (row == null) continue;

                Cell dateCell = row.getCell(dateCol);
                if (dateCell == null) continue;

                LocalDate date;

                try {
                    if (dateCell.getCellType() == CellType.NUMERIC) {
                        date = dateCell.getLocalDateTimeCellValue().toLocalDate();
                    } else {
                        String d = dateCell.toString().trim();
                        date = LocalDate.parse(d.substring(0, 10));
                    }
                } catch (Exception e) {
                    continue;
                }

                if (minDate == null || date.isBefore(minDate))
                    minDate = date;

                if (maxDate == null || date.isAfter(maxDate))
                    maxDate = date;

                String payment = "UNKNOWN";

                Cell payCell = row.getCell(payCol);

                if (payCell != null) {

                    payment = payCell.toString()
                            .toUpperCase()
                            .trim()
                            .replace("_", " ")
                            .replaceAll("\\s+", " ");

                    if (payment.equals("CREDIT CARD")
                            || payment.equals("CREDITCARD")
                            || payment.equals("CC")) {
                        payment = "CC";
                    }

                    else if (payment.equals("DEBIT CARD")
                            || payment.equals("DEBITCARD")
                            || payment.equals("DC")) {
                        payment = "DC";
                    }

                    else if (payment.equals("NET BANKING")
                            || payment.equals("NETBANKING")
                            || payment.equals("NB")) {
                        payment = "NB";
                    }

                    else if (payment.equals("UPI PPI")) {
                        payment = "UPIPPI";
                    }

                    else if (payment.equals("UPI CC")) {
                        payment = "UPICC";
                    }

                    else if (payment.contains("UPI SALE")) {
                        payment = "UPI SALE";
                    }
                }

                double price = 0;

                Cell priceCell = row.getCell(priceCol);

                if (priceCell != null) {
                    try {
                        price =
                                (priceCell.getCellType() == CellType.NUMERIC)
                                        ? priceCell.getNumericCellValue()
                                        : Double.parseDouble(
                                                priceCell.toString()
                                                        .replace(",", "")
                                                        .trim()
                                        );
                    } catch (Exception ignored) {
                    }
                }

                double discount = 0;

                if (discountCol != -1) {

                    Cell discountCell = row.getCell(discountCol);

                    if (discountCell != null) {
                        try {
                            discount =
                                    (discountCell.getCellType() == CellType.NUMERIC)
                                            ? discountCell.getNumericCellValue()
                                            : Double.parseDouble(
                                                    discountCell.toString()
                                                            .replace(",", "")
                                                            .trim()
                                            );
                        } catch (Exception ignored) {
                        }
                    }
                }

                daily
                        .computeIfAbsent(date, k -> new HashMap<>())
                        .computeIfAbsent(payment, k -> new Stats())
                        .add(price, discount);

                dayTotal
                        .computeIfAbsent(date, k -> new Stats())
                        .add(price, discount);

                grand
                        .computeIfAbsent(payment, k -> new Stats())
                        .add(price, discount);

                overall.add(price, discount);
            }

        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        List<String[]> table = new ArrayList<>();

        String[] headerRow = new String[PAYMENT_ORDER.length + 3];

        headerRow[0] = "Date";

        System.arraycopy(
                PAYMENT_ORDER,
                0,
                headerRow,
                1,
                PAYMENT_ORDER.length
        );

        headerRow[headerRow.length - 2] = "Discount";
        headerRow[headerRow.length - 1] = "Total";

        table.add(headerRow);

        for (LocalDate d = minDate;
             !d.isAfter(maxDate);
             d = d.plusDays(1)) {

            Map<String, Stats> map =
                    daily.getOrDefault(d, new HashMap<>());

            Stats total =
                    dayTotal.getOrDefault(d, new Stats());

            String[] row =
                    new String[PAYMENT_ORDER.length + 3];

            row[0] = d.toString();

            for (int i = 0; i < PAYMENT_ORDER.length; i++) {

                Stats s = map.get(PAYMENT_ORDER[i]);

                row[i + 1] =
                        (s != null)
                                ? s.format()
                                : "0(0|0)";
            }

            row[row.length - 2] =
                    String.format("%.0f", total.discount);

            row[row.length - 1] =
                    total.format();

            table.add(row);
        }

        String[] grandRow =
                new String[PAYMENT_ORDER.length + 3];

        grandRow[0] = "GRAND TOTAL";

        for (int i = 0; i < PAYMENT_ORDER.length; i++) {

            grandRow[i + 1] =
                    grand.getOrDefault(
                            PAYMENT_ORDER[i],
                            new Stats()
                    ).format();
        }

        grandRow[grandRow.length - 2] =
                String.format("%.0f", overall.discount);

        grandRow[grandRow.length - 1] =
                overall.format();

        table.add(grandRow);

        String totalString = overall.format();

        String[] totalRow =
                new String[PAYMENT_ORDER.length + 3];

        totalRow[0] = "TOTAL";

        Arrays.fill(
                totalRow,
                1,
                totalRow.length,
                totalString
        );

        table.add(totalRow);

        printTable(table);
    }

    private static int findColumnFlexible(
            Map<String, Integer> map,
            String... keywords) {

        for (String key : map.keySet()) {

            for (String k : keywords) {

                if (key.contains(k)) {
                    return map.get(key);
                }
            }
        }

        return -1;
    }

    private static void printTable(List<String[]> rows) {

        int cols = rows.get(0).length;

        int[] widths = new int[cols];

        for (String[] row : rows) {

            for (int c = 0; c < cols; c++) {

                widths[c] =
                        Math.min(
                                Math.max(
                                        widths[c],
                                        row[c].length()
                                ),
                                MAX_WIDTH
                        );
            }
        }

        StringBuilder line = new StringBuilder("+");

        for (int w : widths) {
            line.append("-".repeat(w + 2)).append("+");
        }

        System.out.println(line);

        printRow(rows.get(0), widths);

        System.out.println(line);

        for (int i = 1; i < rows.size(); i++) {

            if ("GRAND TOTAL".equals(rows.get(i)[0])) {
                System.out.println(line);
            }

            printRow(rows.get(i), widths);

            if ("GRAND TOTAL".equals(rows.get(i)[0])) {
                System.out.println(line);
            }
        }

        System.out.println(line);
    }

    private static void printRow(
            String[] row,
            int[] widths) {

        System.out.print("|");

        for (int c = 0; c < row.length; c++) {

            String value = row[c];

            if (value.length() > widths[c]) {
                value =
                        value.substring(
                                0,
                                widths[c] - 2
                        ) + "..";
            }

            if (c == 0) {
                System.out.printf(
                        " %-"+widths[c]+"s |",
                        value
                );
            } else {
                System.out.printf(
                        " %"+widths[c]+"s |",
                        value
                );
            }
        }

        System.out.println();
    }
}