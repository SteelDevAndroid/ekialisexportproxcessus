package fr.ekialis.processus;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Complète la première feuille, sans modifier le classeur source. */
public class ProcessusExcelExporter {
    public int export(Path input, Path output) throws IOException {
        Path destination = output.toAbsolutePath().normalize();
        if (input.toAbsolutePath().normalize().equals(destination) || Files.exists(destination)) {
            throw new IllegalArgumentException("Le fichier de sortie doit être nouveau et différent du fichier source.");
        }

        try (var stream = Files.newInputStream(input); var workbook = new XSSFWorkbook(stream)) {
            if (workbook.getNumberOfSheets() == 0) {
                throw new IllegalArgumentException("Le classeur ne contient aucune feuille.");
            }
            int count = complete(workbook.getSheetAt(0), workbook.getCreationHelper().createFormulaEvaluator());
            // Écrire seulement après validation complète, et ne jamais remplacer un fichier existant.
            Path temporary = Files.createTempFile(destination.getParent(), ".process-pr-", ".xlsx");
            try {
                try (var target = Files.newOutputStream(temporary)) {
                    workbook.write(target);
                }
                Files.move(temporary, destination);
            } finally {
                Files.deleteIfExists(temporary);
            }
            return count;
        }
    }

    private int complete(Sheet sheet, FormulaEvaluator evaluator) {
        DataFormatter formatter = new DataFormatter(Locale.ROOT);
        Row header = sheet.getRow(0);
        if (header == null) {
            throw new IllegalArgumentException("La première ligne doit contenir les en-têtes Niveau et Processus.");
        }
        int levelColumn = column(header, "Niveau", formatter, evaluator);
        int labelColumn = column(header, "Processus", formatter, evaluator);
        if (levelColumn < 0 || labelColumn < 0) {
            throw new IllegalArgumentException("Colonnes obligatoires manquantes : Niveau et Processus.");
        }
        int resultColumn = outputColumn(sheet, "pr", formatter, evaluator);
        int scaColumn = outputColumn(sheet, "sca", formatter, evaluator);
        int pr1Column = outputColumn(sheet, "pr1", formatter, evaluator);
        int pr2Column = outputColumn(sheet, "pr2", formatter, evaluator);
        int pr3Column = outputColumn(sheet, "pr3", formatter, evaluator);
        Set<Integer> outputColumns = Set.of(resultColumn, scaColumn, pr1Column, pr2Column, pr3Column);

        String levelOne = null;
        String levelTwo = null;
        int count = 0;
        for (int index = 1; index <= sheet.getLastRowNum(); index++) {
            Row row = sheet.getRow(index);
            if (row == null || empty(row, outputColumns, formatter, evaluator)) {
                continue;
            }
            String level = formatter.formatCellValue(row.getCell(levelColumn), evaluator).strip();
            String label = formatter.formatCellValue(row.getCell(labelColumn), evaluator);
            if (label.isBlank()) {
                throw invalid(row, "libellé Processus vide");
            }
            String path = switch (level) {
                case "1" -> {
                    levelOne = label;
                    levelTwo = null;
                    yield label;
                }
                case "2" -> {
                    if (levelOne == null) {
                        throw invalid(row, "niveau 2 sans parent de niveau 1");
                    }
                    levelTwo = label;
                    yield levelOne + "#" + label;
                }
                case "3" -> {
                    if (levelOne == null || levelTwo == null) {
                        throw invalid(row, "niveau 3 sans parent de niveau 2 dans le processus courant");
                    }
                    yield levelOne + "#" + levelTwo + "#" + label;
                }
                default -> throw invalid(row, "Niveau doit valoir 1, 2 ou 3 (valeur : " + level + ")");
            };
            setValue(row, resultColumn, path);
            setValue(row, scaColumn, company(path));
            String[] parts = parts(path);
            setValue(row, pr1Column, parts[0]);
            setValue(row, pr2Column, parts.length > 1 ? parts[1] : "");
            setValue(row, pr3Column, parts.length > 2 ? parts[2] : "");
            count++;
        }
        return count;
    }

    private int outputColumn(Sheet sheet, String name, DataFormatter formatter, FormulaEvaluator evaluator) {
        Row header = sheet.getRow(0);
        int index = column(header, name, formatter, evaluator);
        if (index < 0) {
            index = Math.max(header.getLastCellNum(), 0);
            for (Row row : sheet) {
                index = Math.max(index, row.getLastCellNum());
            }
        }
        setValue(header, index, name);
        return index;
    }

    private void setValue(Row row, int column, String value) {
        Cell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setBlank();
        cell.setCellValue(value);
    }

    private String company(String path) {
        // ESC couvre également ESCOTA, et COF couvre également COFIROUTE.
        for (String code : List.of("ASF", "ESC", "COF", "VA")) {
            if (path.contains(code)) {
                return code;
            }
        }
        return "";
    }

    private String[] parts(String path) {
        for (String prefix : List.of("VA - ", "ESC - ", "COF - ", "ASF - ")) {
            if (path.startsWith(prefix)) {
                path = path.substring(prefix.length());
                break;
            }
        }
        // La limite préserve tous les séparateurs suivants dans pr3, ainsi que les parties vides.
        return path.split(" - |#", 3);
    }

    private int column(Row header, String name, DataFormatter formatter, FormulaEvaluator evaluator) {
        int found = -1;
        for (Cell cell : header) {
            if (name.equalsIgnoreCase(formatter.formatCellValue(cell, evaluator).strip())) {
                if (found >= 0) {
                    throw new IllegalArgumentException("Colonne en double : " + name);
                }
                found = cell.getColumnIndex();
            }
        }
        return found;
    }

    private boolean empty(Row row, Set<Integer> outputColumns, DataFormatter formatter, FormulaEvaluator evaluator) {
        for (Cell cell : row) {
            if (!outputColumns.contains(cell.getColumnIndex())
                    && !formatter.formatCellValue(cell, evaluator).isBlank()) {
                return false;
            }
        }
        return true;
    }

    private IllegalArgumentException invalid(Row row, String reason) {
        return new IllegalArgumentException("Ligne " + (row.getRowNum() + 1) + " : " + reason);
    }
}
