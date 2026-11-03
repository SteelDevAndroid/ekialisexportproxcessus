package fr.ekialis.processus;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ProcessusExcelExporterTest {
    @TempDir
    Path directory;

    private final ProcessusExcelExporter exporter = new ProcessusExcelExporter();

    @Test
    void buildsPathsInRowOrderAndPreservesWorkbook() throws Exception {
        Path input = create(List.of("1|Racine A", "2|Enfant A", "3|Détail A", "3|Détail B",
                "2|Enfant B", "3|Détail C", "1|Racine B", "2|Enfant C", "3|Détail D"), false);
        byte[] original = Files.readAllBytes(input);
        Path output = directory.resolve("output.xlsx");

        assertThat(exporter.export(input, output)).isEqualTo(9);

        try (var stream = Files.newInputStream(output); var workbook = new XSSFWorkbook(stream)) {
            var sheet = workbook.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(3).getStringCellValue()).isEqualTo("pr");
            assertThat(java.util.stream.IntStream.rangeClosed(1, 9)
                    .mapToObj(i -> sheet.getRow(i).getCell(3).getStringCellValue()).toList())
                    .containsExactly("Racine A", "Racine A#Enfant A", "Racine A#Enfant A#Détail A",
                            "Racine A#Enfant A#Détail B", "Racine A#Enfant B", "Racine A#Enfant B#Détail C",
                            "Racine B", "Racine B#Enfant C", "Racine B#Enfant C#Détail D");
            assertThat(sheet.getRow(1).getCell(2).getCellFormula()).isEqualTo("1+1");
            assertThat(sheet.getRow(1).getCell(1).getCellStyle().getWrapText()).isTrue();
            assertThat(workbook.getSheet("Autre").getRow(0).getCell(0).getStringCellValue()).isEqualTo("Conserver");
        }
        assertThat(Files.readAllBytes(input)).isEqualTo(original);
    }

    @Test
    void recalculatesExistingPrIncludingFormulaCellsWithoutAddingColumn() throws Exception {
        Path input = create(List.of("1|Racine"), true);
        Path output = directory.resolve("output.xlsx");
        exporter.export(input, output);
        try (var stream = Files.newInputStream(output); var workbook = new XSSFWorkbook(stream)) {
            var sheet = workbook.getSheetAt(0);
            assertThat(sheet.getRow(0).getLastCellNum()).isEqualTo((short) 4);
            assertThat(sheet.getRow(1).getCell(3).getCellType()).isEqualTo(CellType.STRING);
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("Racine");
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"2|Orphelin", "3|Orphelin", "4|Invalide", "|Sans niveau", "1|"})
    void rejectsInvalidRowsWithoutProducingOutput(String row) throws Exception {
        Path input = create(List.of(row), false);
        Path output = directory.resolve("output.xlsx");
        assertThatThrownBy(() -> exporter.export(input, output))
                .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("Ligne 2");
        assertThat(output).doesNotExist();
    }

    @Test
    void discardsLevelTwoParentWhenANewRootStarts() throws Exception {
        Path input = create(List.of("1|A", "2|B", "1|C", "3|D"), false);
        assertThatThrownBy(() -> exporter.export(input, directory.resolve("output.xlsx")))
                .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("Ligne 5");
    }

    @Test
    void skipsBlankRowsWithoutLosingParents() throws Exception {
        Path input = create(List.of("1|A", "2|B", "", "3|C"), false);
        Path output = directory.resolve("output.xlsx");
        assertThat(exporter.export(input, output)).isEqualTo(3);
        try (var stream = Files.newInputStream(output); var workbook = new XSSFWorkbook(stream)) {
            assertThat(workbook.getSheetAt(0).getRow(4).getCell(3).getStringCellValue()).isEqualTo("A#B#C");
        }
    }

    @Test
    void refusesToOverwriteSourceOrExistingOutput() throws Exception {
        Path input = create(List.of("1|A"), false);
        byte[] original = Files.readAllBytes(input);
        assertThatThrownBy(() -> exporter.export(input, input)).isInstanceOf(IllegalArgumentException.class);
        Path output = directory.resolve("output.xlsx");
        Files.writeString(output, "existant");
        assertThatThrownBy(() -> exporter.export(input, output)).isInstanceOf(IllegalArgumentException.class);
        assertThat(Files.readAllBytes(input)).isEqualTo(original);
        assertThat(Files.readString(output)).isEqualTo("existant");
    }

    @Test
    void rejectsMissingHeaders() throws Exception {
        Path input = directory.resolve("input.xlsx");
        try (var workbook = new XSSFWorkbook(); var stream = Files.newOutputStream(input)) {
            workbook.createSheet().createRow(0).createCell(0).setCellValue("Inconnu");
            workbook.write(stream);
        }
        assertThatThrownBy(() -> exporter.export(input, directory.resolve("output.xlsx")))
                .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("Colonnes obligatoires");
    }

    private Path create(List<String> rows, boolean existingPr) throws Exception {
        Path input = directory.resolve("input.xlsx");
        try (var workbook = new XSSFWorkbook(); var stream = Files.newOutputStream(input)) {
            var sheet = workbook.createSheet("Processus");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue(" Niveau ");
            header.createCell(1).setCellValue("PROCESSUS");
            header.createCell(2).setCellValue("Autre colonne");
            if (existingPr) {
                header.createCell(3).setCellValue("pr");
            }
            var style = workbook.createCellStyle();
            style.setWrapText(true);
            for (int i = 0; i < rows.size(); i++) {
                var row = sheet.createRow(i + 1);
                if (rows.get(i).isEmpty()) {
                    continue;
                }
                String[] fields = rows.get(i).split("\\|", -1);
                // Exercer les niveaux Excel numériques et les niveaux stockés comme texte.
                if (i % 2 == 0 && !fields[0].isEmpty()) {
                    row.createCell(0).setCellValue(Integer.parseInt(fields[0]));
                } else {
                    row.createCell(0).setCellValue(fields[0]);
                }
                row.createCell(1).setCellValue(fields[1]);
                row.getCell(1).setCellStyle(style);
                row.createCell(2).setCellFormula("1+1");
                if (existingPr) {
                    row.createCell(3).setCellFormula("\"Ancien\"");
                }
            }
            workbook.createSheet("Autre").createRow(0).createCell(0).setCellValue("Conserver");
            workbook.write(stream);
        }
        return input;
    }
}
