package pinelabreport;

import java.io.*;
import java.time.LocalDate;
import java.util.*;

public class PinelabReportGenerator {

    static class Stats {
        int count = 0;
        double amount = 0;

        void add(double price) {
            count++;
            amount += price;
        }

        String format() {
            return count + " (" + String.format("%.0f", amount) + ")";
        }
    }

    static class Record {
        String date;
        String paymentMethod;
        double ticketPrice;

        Record(String date, String paymentMethod, double ticketPrice) {
            this.date = date;
            this.paymentMethod = paymentMethod;
            this.ticketPrice = ticketPrice;
        }
    }

    private static final String[] PAYMENT_ORDER = {
            "UPI", "CARD", "CASH", "NET BANKING", "WALLET"
    };

    public static void main(String[] args) {

        String filePath = "C:\\Users\\Nikhil Sonawane\\eclipse-workspace\\PinelabReport\\transactionReport (2) (1).csv";

        try {
            List<Record> records = readCsv(filePath);

            if (records.isEmpty()) {
                System.out.println("No records found.");
                return;
            }

            Map<LocalDate, Map<String, Stats>> daily = new TreeMap<>();
            Map<String, Stats> grand = new HashMap<>();
            Map<LocalDate, Stats> dayTotal = new HashMap<>();

            LocalDate minDate = null;
            LocalDate maxDate = null;

            for (Record r : records) {

                LocalDate date = LocalDate.parse(r.date.split("T")[0]);
                String payment = r.paymentMethod.toUpperCase().trim();

                if (minDate == null || date.isBefore(minDate)) minDate = date;
                if (maxDate == null || date.isAfter(maxDate)) maxDate = date;

                daily
                        .computeIfAbsent(date, k -> new HashMap<>())
                        .computeIfAbsent(payment, k -> new Stats())
                        .add(r.ticketPrice);

                dayTotal.computeIfAbsent(date, k -> new Stats()).add(r.ticketPrice);
                grand.computeIfAbsent(payment, k -> new Stats()).add(r.ticketPrice);
            }

            // 🔥 FORCE END DATE (ADDED WITHOUT REMOVING ANY CODE)
            LocalDate forcedEndDate = LocalDate.parse("2026-04-30");
            if (maxDate == null || forcedEndDate.isAfter(maxDate)) {
                maxDate = forcedEndDate;
            }

            List<String[]> table = new ArrayList<>();

            String[] header = new String[PAYMENT_ORDER.length + 2];
            header[0] = "Date";
            System.arraycopy(PAYMENT_ORDER, 0, header, 1, PAYMENT_ORDER.length);
            header[header.length - 1] = "Total";
            table.add(header);

            for (LocalDate d = minDate; !d.isAfter(maxDate); d = d.plusDays(1)) {

                Map<String, Stats> map = daily.getOrDefault(d, new HashMap<>());
                Stats total = dayTotal.getOrDefault(d, new Stats());

                String[] row = new String[PAYMENT_ORDER.length + 2];
                row[0] = d.toString();

                for (int i = 0; i < PAYMENT_ORDER.length; i++) {
                    Stats s = map.get(PAYMENT_ORDER[i]);
                    row[i + 1] = (s != null) ? s.format() : "0 (0)";
                }

                row[row.length - 1] = total.format();
                table.add(row);
            }

            String[] g = new String[PAYMENT_ORDER.length + 2];
            g[0] = "GRAND TOTAL";

            Stats overall = new Stats();

            for (int i = 0; i < PAYMENT_ORDER.length; i++) {
                Stats s = grand.getOrDefault(PAYMENT_ORDER[i], new Stats());
                g[i + 1] = s.format();
                overall.count += s.count;
                overall.amount += s.amount;
            }

            g[g.length - 1] = overall.format();
            table.add(g);

            String totalStr = overall.format();
            String[] totalRow = new String[PAYMENT_ORDER.length + 2];
            totalRow[0] = "TOTAL";
            Arrays.fill(totalRow, 1, totalRow.length, totalStr);
            table.add(totalRow);

            printTable(table);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // -------- CSV READER (FIXED WITHOUT REMOVING ANY WORD) --------
    private static List<Record> readCsv(String filePath) throws IOException {
        List<Record> records = new ArrayList<>();
        boolean isFirstLine = true;

        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            String line;

            while ((line = br.readLine()) != null) {

                if (isFirstLine) {
                    isFirstLine = false;
                    continue;
                }

                String[] values = parseCsvLine(line);
                if (values.length < 23) continue;

                // 🔥 SAFE STATUS FILTER (ADDED)
                String status = "";

                if (values.length > 23) status = values[23].toUpperCase();
                else if (values.length > 22) status = values[22].toUpperCase();
                else if (values.length > 21) status = values[21].toUpperCase();

                if (!(status.contains("SUCCESS") || status.contains("CAPTURED"))) {
                    continue;
                }

                String method = values[0].trim();
                String priceStr = values[20].trim().replace(",", "");
                String dateStr = values[22].trim();

                double price;
                try {
                    price = Double.parseDouble(priceStr);
                } catch (Exception e) {
                    continue;
                }

                records.add(new Record(dateStr, method, price));
            }
        }
        return records;
    }

    private static String[] parseCsvLine(String line) {
        List<String> result = new ArrayList<>();
        StringBuilder field = new StringBuilder();
        boolean inQuotes = false;

        for (char c : line.toCharArray()) {
            if (c == '"') inQuotes = !inQuotes;
            else if (c == ',' && !inQuotes) {
                result.add(field.toString());
                field = new StringBuilder();
            } else field.append(c);
        }
        result.add(field.toString());

        return result.toArray(new String[0]);
    }

    private static void printTable(List<String[]> rows) {

        int cols = rows.get(0).length;
        int[] widths = new int[cols];

        for (String[] row : rows) {
            for (int c = 0; c < cols; c++) {
                widths[c] = Math.max(widths[c], row[c].length());
            }
        }

        StringBuilder line = new StringBuilder("+");
        for (int w : widths) line.append("-".repeat(w + 2)).append("+");

        System.out.println(line);

        for (int i = 0; i < rows.size(); i++) {

            // 🔥 LINE BEFORE GRAND TOTAL
            if (rows.get(i)[0].equals("GRAND TOTAL")) {
                System.out.println(line);
            }

            printRow(rows.get(i), widths);

            if (i == 0 || rows.get(i)[0].equals("GRAND TOTAL")) {
                System.out.println(line);
            }
        }

        System.out.println(line);
    }

    private static void printRow(String[] row, int[] widths) {
        System.out.print("|");

        for (int c = 0; c < row.length; c++) {
            if (c == 0) {
                System.out.printf(" %-"+widths[c]+"s |", row[c]);
            } else {
                System.out.printf(" %"+widths[c]+"s |", row[c]);
            }
        }

        System.out.println();
    }
}