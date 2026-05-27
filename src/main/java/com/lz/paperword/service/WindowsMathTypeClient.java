package com.lz.paperword.service;

import com.lz.paperword.config.WindowsMathTypeProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;

/**
 * Calls an external Windows service that converts LaTeX markers in a draft DOCX
 * into real MathType equations through local Word + MathType automation.
 */
@Component
public class WindowsMathTypeClient {

    private static final Logger log = LoggerFactory.getLogger(WindowsMathTypeClient.class);
    private static final String DOCX_CONTENT_TYPE =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private final WindowsMathTypeProperties properties;

    public WindowsMathTypeClient(WindowsMathTypeProperties properties) {
        this.properties = properties;
    }

    public byte[] convert(byte[] draftDocx) throws IOException {
        HttpClient client = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(Math.max(5, properties.getTimeoutSeconds() / 2)))
            .build();

        HttpRequest.Builder requestBuilder = HttpRequest.newBuilder()
            .uri(URI.create(properties.getEndpoint()))
            .timeout(Duration.ofSeconds(Math.max(10, properties.getTimeoutSeconds())))
            .header("Content-Type", DOCX_CONTENT_TYPE)
            .header("Accept", DOCX_CONTENT_TYPE)
            .POST(HttpRequest.BodyPublishers.ofByteArray(draftDocx));

        if (properties.getApiKey() != null && !properties.getApiKey().isBlank()) {
            requestBuilder.header("X-API-Key", properties.getApiKey().trim());
        }

        HttpResponse<byte[]> response;
        try {
            response = client.send(requestBuilder.build(), HttpResponse.BodyHandlers.ofByteArray());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("Windows MathType conversion interrupted", e);
        }

        if (response.statusCode() < 200 || response.statusCode() >= 300) {
            throw new IOException("Windows MathType conversion failed, HTTP status: " + response.statusCode());
        }
        if (response.body() == null || response.body().length == 0) {
            throw new IOException("Windows MathType conversion returned empty document");
        }

        log.info("Windows MathType conversion completed, size: {} bytes", response.body().length);
        return response.body();
    }
}
