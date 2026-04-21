package com.github.pjfanning.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.Benchmark;
import org.openjdk.jmh.annotations.Param;
import org.openjdk.jmh.annotations.Setup;
import org.openjdk.jmh.infra.Blackhole;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * JMH benchmarks for Apache POI XSSFWorkbook.
 *
 * <p>The {@code cellCount} parameter is varied across multiple values to expose the
 * O(n²) → O(n) improvement in xmlbeans' {@code arraySetterHelper}
 * (see <a href="https://github.com/apache/xmlbeans/pull/30">xmlbeans#30</a>).
 * When writing an XLSX file, POI calls {@code arraySetterHelper} once per row; the inner
 * loop iterates over the {@code cellCount} cells in that row.  With the old code each
 * iteration re-scanned from the first child, making the total work O(cellCount²) per row.
 * The fixed code collects all existing children up-front (O(n) once) and then uses O(1)
 * list indexing, so the benefit grows dramatically with larger {@code cellCount} values.
 *
 * <p>Benchmark variants:
 * <ul>
 *   <li><b>writeWorkbook</b> – Creates an XSSFWorkbook with {@code rowCount} rows ×
 *       {@code cellCount} cells, serialises it to a {@link ByteArrayOutputStream}, and
 *       consumes the resulting byte array.  <em>This is the benchmark that exercises
 *       {@code arraySetterHelper} and will show the speedup from xmlbeans#30.</em></li>
 *   <li><b>readWorkbook</b> – Opens the workbook written during {@link #setup()} and
 *       traverses every sheet, row, and cell.  This measures read/parse throughput and
 *       is not affected by the {@code arraySetterHelper} change.</li>
 * </ul>
 */
public class XSSFBench extends BenchmarkLauncher {

    /**
     * Number of rows per sheet.  Kept small so that total cell count stays reasonable
     * even at the largest {@code cellCount} value.
     */
    private static final int ROW_COUNT = 10;

    /**
     * Number of cells per row.  Varied via {@code @Param} to expose the O(n²) → O(n)
     * scaling behaviour fixed in xmlbeans#30.  At small values (e.g. 100) the quadratic
     * penalty is negligible; at large values (e.g. 5000) it dominates and the improvement
     * becomes clearly measurable.
     */
    @Param({"100", "500", "2000", "5000"})
    public int cellCount;

    private byte[] workbookBytes;

    @Setup
    public void setup() throws IOException {
        workbookBytes = createWorkbookBytes();
    }

    /**
     * Creates an XSSFWorkbook with one sheet of {@value ROW_COUNT} rows × {@code cellCount}
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
                for (int c = 0; c < cellCount; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue("abcdef");
                }
            }
            ByteArrayOutputStream out = new ByteArrayOutputStream(64 * 1024);
            workbook.write(out);
            return out.toByteArray();
        }
    }

    /**
     * Benchmark 1: measures how long it takes to create a workbook and serialise it to a
     * byte array.  <strong>This is the primary benchmark for xmlbeans#30</strong> — the
     * {@code arraySetterHelper} hot-path is exercised during {@code workbook.write()}.
     * Throughput should scale linearly with {@code cellCount} on the patched xmlbeans,
     * and quadratically (i.e. drop much faster) on the unpatched version.
     */
    @Benchmark
    public void writeWorkbook(Blackhole blackhole) throws IOException {
        blackhole.consume(createWorkbookBytes());
    }

    /**
     * Benchmark 2: measures how long it takes to open the pre-built workbook byte array and
     * traverse every sheet, row, and cell.  {@code arraySetterHelper} is not called on the
     * read path, so this benchmark is unaffected by the xmlbeans#30 change.
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
