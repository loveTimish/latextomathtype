package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.nio.file.Path;

/** Compares one standard DOCX with one generated DOCX. */
public final class MathTypeDocxComparatorCli {

    private MathTypeDocxComparatorCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("usage: MathTypeDocxComparatorCli <standard.docx> <generated.docx> <report.json>");
            System.exit(2);
        }
        Path standard = Path.of(args[0]).toAbsolutePath();
        Path generated = Path.of(args[1]).toAbsolutePath();
        Path reportPath = Path.of(args[2]).toAbsolutePath();
        MathTypeDocxComparator.ComparisonReport report =
            new MathTypeDocxComparator().compare(standard, generated);
        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT)
            .writeValue(reportPath.toFile(), report);
        System.out.printf("objects=%d valid=%d exactStructure=%d report=%s%n",
            report.generatedObjectCount(), report.validGeneratedOleCount(),
            report.normalizedStructureEqualCount(), reportPath);
        if (!report.passesExactStructureGate()) {
            System.exit(1);
        }
    }
}
