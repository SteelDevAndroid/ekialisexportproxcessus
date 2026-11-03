package fr.ekialis.processus;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
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
            assertThat(sheet.getRow(0).getLastCellNum()).isEqualTo((short) 8);
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

    @ParameterizedTest
    @MethodSource("details")
    void derivesCompanyAndPartsFromPr(String label, String sca, String pr1, String pr2, String pr3) throws Exception {
        Path input = create(List.of("1|" + label), false);
        Path output = directory.resolve("output.xlsx");
        exporter.export(input, output);
        try (var stream = Files.newInputStream(output); var workbook = new XSSFWorkbook(stream)) {
            var sheet = workbook.getSheetAt(0);
            assertThat(java.util.stream.IntStream.range(3, 8)
                    .mapToObj(i -> sheet.getRow(0).getCell(i).getStringCellValue()).toList())
                    .containsExactly("pr", "sca", "pr1", "pr2", "pr3");
            assertThat(java.util.stream.IntStream.range(3, 8)
                    .mapToObj(i -> sheet.getRow(1).getCell(i).getStringCellValue()).toList())
                    .containsExactly(label, sca, pr1, pr2, pr3);
        }
    }

    static Stream<Arguments> details() {
        return Stream.of(
                Arguments.of("ASF - Assister - Appeler - Informer", "ASF", "Assister", "Appeler", "Informer"),
                Arguments.of("ESC - A#B - C#D", "ESC", "A", "B", "C#D"),
                Arguments.of("COF - A - B#C - D", "COF", "A", "B", "C - D"),
                Arguments.of("VA - A#B#C#D", "VA", "A", "B", "C#D"),
                Arguments.of("ESCOTA - A - B", "ESC", "ESCOTA", "A", "B"),
                Arguments.of("COFIROUTE#A#B", "COF", "COFIROUTE", "A", "B"),
                Arguments.of("Sans société", "", "Sans société", "", ""),
                Arguments.of("A - B", "", "A", "B", ""),
                Arguments.of("A#B", "", "A", "B", ""),
                Arguments.of("A - #C", "", "A", "", "C"),
                Arguments.of("A##", "", "A", "", ""),
                Arguments.of("ASF - ", "ASF", "", "", ""),
                Arguments.of("A#ESCOTA - C", "ESC", "A", "ESCOTA", "C"),
                Arguments.of("A#COFIROUTE", "COF", "A", "COFIROUTE", ""),
                Arguments.of("A#VA", "VA", "A", "VA", ""),
                Arguments.of("VA - A#COF#ESC#ASF", "ASF", "A", "COF", "ESC#ASF"),
                Arguments.of("ASF-A#B", "ASF", "ASF-A", "B", ""),
                Arguments.of("asf - A", "", "asf", "A", ""));
    }

    @Test
    void derivesDetailsFromFullHierarchyAndRecalculatesAllGeneratedColumns() throws Exception {
        Path input = create(List.of("1|ESC - Assister", "2|Appeler", "3|Informer"), false);
        Path firstOutput = directory.resolve("first.xlsx");
        Path secondOutput = directory.resolve("second.xlsx");
        exporter.export(input, firstOutput);
        try (var stream = Files.newInputStream(firstOutput); var workbook = new XSSFWorkbook(stream)) {
            var sheet = workbook.getSheetAt(0);
            assertThat(sheet.getRow(3).getCell(3).getStringCellValue())
                    .isEqualTo("ESC - Assister#Appeler#Informer");
            assertThat(java.util.stream.IntStream.range(4, 8)
                    .mapToObj(i -> sheet.getRow(3).getCell(i).getStringCellValue()).toList())
                    .containsExactly("ESC", "Assister", "Appeler", "Informer");
            // Les colonnes calculées existantes, même avec des formules, sont remplacées.
            for (int i = 3; i < 8; i++) {
                sheet.getRow(3).getCell(i).setCellFormula("\"Ancien\"");
            }
            try (var target = Files.newOutputStream(firstOutput)) {
                workbook.write(target);
            }
        }
        assertThat(exporter.export(firstOutput, secondOutput)).isEqualTo(3);
        try (var stream = Files.newInputStream(secondOutput); var workbook = new XSSFWorkbook(stream)) {
            var sheet = workbook.getSheetAt(0);
            assertThat(sheet.getRow(0).getLastCellNum()).isEqualTo((short) 8);
            assertThat(java.util.stream.IntStream.range(3, 8)
                    .mapToObj(i -> sheet.getRow(3).getCell(i).getStringCellValue()).toList())
                    .containsExactly("ESC - Assister#Appeler#Informer", "ESC", "Assister", "Appeler", "Informer");
        }
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
