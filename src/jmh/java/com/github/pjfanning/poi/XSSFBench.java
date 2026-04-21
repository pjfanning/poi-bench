package com.github.pjfanning.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.Benchmark;
import org.openjdk.jmh.annotations.Setup;
import org.openjdk.jmh.infra.Blackhole;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * JMH benchmarks for Apache POI XSSFWorkbook.
 *
 * <p>Benchmark variants:
 * <ul>
 *   <li><b>writeWorkbook</b> – Creates an XSSFWorkbook with one sheet containing 1,000 rows
 *       of 1,000 cells each, serialises it to a {@link ByteArrayOutputStream}, and consumes
 *       the resulting byte array. This measures workbook creation and write throughput.</li>
 *   <li><b>readWorkbook</b> – Opens the workbook previously written by {@link #createWorkbookBytes()}
 *       and traverses every sheet, row, and cell. This measures workbook parsing and traversal
 *       throughput.</li>
 * </ul>
 */
public class XSSFBench extends BenchmarkLauncher {

    private static final int ROW_COUNT = 100;
    private static final int CELL_COUNT = 100;

    private byte[] workbookBytes;

    @Setup
    public void setup() throws IOException {
        workbookBytes = createWorkbookBytes();
    }

    /**
     * Creates an XSSFWorkbook with one sheet of {@value ROW_COUNT} rows × {@value CELL_COUNT}
     * cells, writes it to a {@link ByteArrayOutputStream}, and returns the raw bytes.
     *
     * @return the serialised workbook as a byte array
     * @throws IOException if writing fails
     */
    byte[] createWorkbookBytes() throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            for (int r = 0; r < ROW_COUNT; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < CELL_COUNT; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue("abcdef");
                }
            }
            ByteArrayOutputStream out = new ByteArrayOutputStream(48 * 1024);
            workbook.write(out);
            return out.toByteArray();
        }
    }

    /**
     * Benchmark 1: measures how long it takes to create a workbook with multiple rows × multiple cells
     * and serialise it to a byte array.
     */
    @Benchmark
    public void writeWorkbook(Blackhole blackhole) throws IOException {
        blackhole.consume(createWorkbookBytes());
    }

    /**
     * Benchmark 2: measures how long it takes to open the pre-built workbook byte array and
     * traverse every sheet, row, and cell.
     */
    @Benchmark
    public void readWorkbook(Blackhole blackhole) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(workbookBytes))) {
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        blackhole.consume(cell);
                    }
                }
            }
        }
    }
}
