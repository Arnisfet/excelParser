package org.example;

import com.aspose.cells.*;

import java.util.*;

public class SheetParser {
    private Map<String, String> headerColumns = new HashMap<>();
    private final Map<String, List<String>> rules = new HashMap<>();
    private Workbook header;
    private Map<String, List<String>> resultMap = new HashMap<>();
    private String fileUrl;

    public SheetParser() {
        try  {
            header = new Workbook("Header.xlsx");
            Worksheet worksheet = header.getWorksheets().get(0);
            Range range = worksheet.getCells().createRange("A1:X2");
            for (int row = range.getFirstRow(); row < range.getFirstRow() + range.getRowCount(); row++) {
                for (int col = range.getFirstColumn(); col < range.getFirstColumn() + range.getColumnCount(); col++) {
                    Cell cell;
                    if (!worksheet.getCells().get(row + 1, col).getStringValue().isEmpty())
                        cell = worksheet.getCells().get(row + 1, col);
                    else
                        cell = worksheet.getCells().get(row, col);
                    headerColumns.put(cell.getStringValue().toLowerCase().trim().replaceAll(" +", " "), getColumnLetter(col));
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    void matchingRules() {
        // Define the rules
        rules.put("Банк, предоставивший выписку", List.of("Банк, предоставивший выписку"));
        rules.put("вид (шифр) или ВО", List.of("вид (шифр) или ВО"));

        rules.put("номер документа", List.of("№ док."));
        rules.put("Дата совершения операции (dd.mm.yyyy) или дата проводки", List.of("Дата операции"));

        rules.put("наименование/ФИО получателя", List.of("Наименование  получателя"));
        rules.put("ИНН/КИО получателя", List.of("ИНН получателя"));
        rules.put("КПП получателя", List.of("КПП получателя"));
        rules.put("Номер счета получателя", List.of("№ счета получателя"));

        rules.put("наименование/ФИО плательщика", List.of("Наименование  плательщика"));
        rules.put("ИНН/КИО плательщика", List.of("ИНН плательщика"));
        rules.put("КПП плательщика", List.of("КПП плательщика"));
        rules.put("Номер счета плательщика", List.of("№ счета плательщика"));

        rules.put("По дебету", List.of("Дебет"));
        rules.put("По кредиту", List.of("Кредит"));

        rules.put("Назначение платежа", List.of("Назначение платежа"));

        rules.put("номер корреспондентского счета банка плательщика", List.of("номер корреспондентского счета  банка плательщика"));
        rules.put("наименование банка плательщика", List.of("Наименование Банка плательщика"));
        rules.put("БИК плательщика", List.of("БИК/SWIFT банка плат."));

        rules.put("номер корреспондентского счета банка получателя", List.of("номер корреспондентского счета банка получателя"));
        rules.put("наименование банка получателя", List.of("Наименование банка получателя"));
        rules.put("БИК получателя", List.of("БИК/SWIFT банка получ."));

        Map<String, List<String>> normalizedRules = new HashMap<>();

        rules.forEach((key, list) -> {
            String normalizedKey = key.trim().replaceAll(" +", " ");

            List<String> normalizedList = list.stream()
                    .map(value -> value.trim().replaceAll(" +", " "))
                    .toList();

            normalizedRules.put(normalizedKey, normalizedList);
        });

        rules.clear();
        rules.putAll(normalizedRules);
    }
    void parse(String filePth) {
        try {
            Workbook currentWorkbook = new Workbook(filePth);

            fileUrl = "file:///" + filePth.replace("\\", "/");
            WorksheetCollection worksheets = currentWorkbook.getWorksheets();

            for (int i = 0; i < worksheets.getCount(); i++) {
                Worksheet sheet = worksheets.get(i);
                System.out.println("Parsing sheet: " + sheet.getName());
                Cells cells = sheet.getCells();
                Set<String> unfoundKeys = new HashSet<>(rules.keySet());

                for (int row = 0; row <= cells.getMaxDataRow(); row++) {
                    for (int col = 0; col <= cells.getMaxDataColumn(); col++) {
                        Cell cell = cells.get(row, col);
                        String cellValue = cell.getStringValue().trim().replaceAll(" +", " ");

                        int finalRow = row;
                        int finalCol = col;

                        rules.forEach((key, patterns) ->
                                patterns.stream()
                                        .filter(cellValue::equalsIgnoreCase)
                                        .findFirst()
                                        .ifPresent(pattern -> {
                                            List<String> newValues = collectColumnValues(cells, finalRow + 2, finalCol);

                                            resultMap.merge(key.toLowerCase(), newValues, (existingValues, incomingValues) -> {
                                                existingValues.addAll(incomingValues);
                                                return existingValues;
                                            });

                                            unfoundKeys.remove(key);
                                        }));
                    }
                }

                if (!unfoundKeys.isEmpty()) {
                    System.out.println("For sheet: " + sheet.getName());
                    System.out.println("Columns not found for the following keys: " + String.join(", ", unfoundKeys));
                } else {
                    System.out.println("All columns were found for sheet: " + sheet.getName());
                }
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }



    private boolean consistencyCheck() {
        Integer referenceSize = null;

        for (Map.Entry<String, List<String>> entry : resultMap.entrySet()) {
            List<String> list = entry.getValue();

            if (list == null) {
                throw new IllegalStateException("List for key '" + entry.getKey() + "' is null");
            }
            if (referenceSize == null) {
                referenceSize = list.size();
            } else {
                if (list.size() != referenceSize) {
                    System.out.println("Inconsistent list sizes found. Key: " + entry.getKey() +
                            ", Size: " + list.size() +
                            ", Expected: " + referenceSize);
                    return false;
                }
            }
        }
        return true;
    }
    public void save() {
        if (!consistencyCheck())
            return ;
        try {
            Worksheet worksheet = header.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            // Iterate over the resultMap entries
            resultMap.forEach((key, list) -> {
                // Find the column letter for the current key
                String columnLetter = headerColumns.get(key.toLowerCase());
                if (columnLetter == null) {
                    System.out.println("No matching column found for key: " + key);
                    return;
                }

                int colIndex = getColumnIndexFromLetter(columnLetter);

                for (int i = 0; i < list.size(); i++) {
                    cells.get(2 + i, colIndex).putValue(list.get(i));

                    worksheet.getHyperlinks().add(2 + i, 0, 1, 1, fileUrl);

                    cells.get(2 + i, 0).setValue("Click here to view the source file");

                    Style style = cells.get(2 + i, 0).getStyle();
                    Font font = style.getFont();
                    font.setColor(Color.getBlue()); // Blue text
                    font.setUnderline(FontUnderlineType.SINGLE); // Underlined text
                    cells.get(2 + i, 0).setStyle(style); // Apply the style to the cell
                }

                System.out.println("Data for key '" + key + "' written to column " + columnLetter);
            });

            header.save("Updated_Header.xlsx");
            System.out.println("Workbook saved as 'Updated_Header.xlsx'.");

        } catch (Exception e) {
            throw new RuntimeException("Error while saving the workbook", e);
        }
    }

    private List<String> collectColumnValues(Cells cells, int startRow, int col) {
        List<String> list = new ArrayList<>();
        for (int row = startRow; row <= cells.getMaxDataRow(); row++) {
            Cell cell = cells.get(row, col);
            String cellValue = cell.getStringValue().trim();
            list.add(cellValue);
//            System.out.println("Value in column " + (col + 1) + ", row " + (row + 1) + ": " + cellValue);
        }
        return list;
    }

    public static String getColumnLetter(int columnIndex) {
        StringBuilder columnLetter = new StringBuilder();
        while (columnIndex >= 0) {
            columnLetter.insert(0, (char) ('A' + columnIndex % 26));
            columnIndex = columnIndex / 26 - 1;
        }
        return columnLetter.toString();
    }

    public static int getColumnIndexFromLetter(String columnLetter) {
        int columnIndex = 0;
        for (int i = 0; i < columnLetter.length(); i++) {
            columnIndex = columnIndex * 26 + (columnLetter.charAt(i) - 'A' + 1);
        }
        return columnIndex - 1;
    }

}
