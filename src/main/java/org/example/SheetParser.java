package org.example;

import com.aspose.cells.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SheetParser {
    private Map<String, String> headerColumns = new HashMap<>();
    private final Map<String, List<String>> rules = new HashMap<>();
    private Workbook header;
    private Map<String, List<String>> resultMap = new HashMap<>();

    public SheetParser() {
        try  {
            header = new Workbook("Header.xlsx");
            Worksheet worksheet = header.getWorksheets().get(0);
            Range range = worksheet.getCells().createRange("A1:Q2");
            for (int row = range.getFirstRow(); row < range.getFirstRow() + range.getRowCount(); row++) {
                for (int col = range.getFirstColumn(); col < range.getFirstColumn() + range.getColumnCount(); col++) {
                    Cell cell;
                    if (!worksheet.getCells().get(row + 1, col).getStringValue().isEmpty())
                        cell = worksheet.getCells().get(row + 1, col);
                    else
                        cell = worksheet.getCells().get(row, col);
                    headerColumns.put(cell.getStringValue(), getColumnLetter(col));
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    void matchingRules() {
        rules.put("номер документа", List.of("№ док."));
    }
    void parse(String workbook) {
        try  {
        Workbook currentSheet = new Workbook(workbook);
        int sheetsNumber = currentSheet.getWorksheets().getCount();

        Worksheet first = currentSheet.getWorksheets().get(0);
            Cells cells = first.getCells();

            for (int row = 0; row <= cells.getMaxDataRow(); row++) {
                for (int col = 0; col <= cells.getMaxDataColumn(); col++) {
                    Cell cell = cells.get(row, col);
                    String cellValue = cell.getStringValue().trim();

                    int finalRow = row;
                    int finalCol = col;
                    rules.forEach((key, patterns) ->
                            patterns.stream()
                                    .filter(cellValue::equalsIgnoreCase)
                                    .findFirst()
                                    .ifPresent(pattern -> {
                                        System.out.println("Match found for field '" + key + "' at: " + cellValue);

                                        List<String> list = collectColumnValues(cells, finalRow + 2, finalCol);
                                        resultMap.put(key, list);
                                            }));
                }
            }
            System.out.println();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private List<String> collectColumnValues(Cells cells, int startRow, int col) {
        List<String> list = new ArrayList<>();
        for (int row = startRow; row <= cells.getMaxDataRow(); row++) {
            Cell cell = cells.get(row, col);
            String cellValue = cell.getStringValue().trim();
            list.add(cellValue);
            System.out.println("Value in column " + (col + 1) + ", row " + (row + 1) + ": " + cellValue);
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
}
