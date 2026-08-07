package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.nio.file.Path;

/** Writes normalized records for every MathType OLE in one DOCX. */
public final class MathTypeDocxInspectCli {

    private MathTypeDocxInspectCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("usage: MathTypeDocxInspectCli <input.docx> <out.json>");
            System.exit(2);
        }
        Path input = Path.of(args[0]).toAbsolutePath();
        Path output = Path.of(args[1]).toAbsolutePath();
        MathTypeDocxComparator.InspectionReport report = new MathTypeDocxComparator().inspect(input);
        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT).writeValue(output.toFile(), report);
        System.out.printf("objects=%d report=%s%n", report.formulas().size(), output);
    }
}
