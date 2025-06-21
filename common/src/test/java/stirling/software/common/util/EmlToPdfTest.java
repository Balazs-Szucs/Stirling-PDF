package stirling.software.common.util;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Base64;

import org.apache.pdfbox.pdmodel.PDDocument;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyBoolean;
import static org.mockito.ArgumentMatchers.anyString;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import static org.mockito.Mockito.mockStatic;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import org.mockito.junit.jupiter.MockitoExtension;

import stirling.software.common.model.api.converters.EmlToPdfRequest;
import stirling.software.common.service.CustomPDFDocumentFactory;

@DisplayName("EML to PDF Conversion with Angus Mail/Jakarta Mail")
class EmlToPdfTest {

    private void testEmailConversion(String emlContent, String[] expectedContent, boolean includeAttachments) throws IOException {
        byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
        EmlToPdfRequest request = includeAttachments ? createRequestWithAttachments() : createBasicRequest();

        String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

        assertNotNull(htmlResult);
        for (String expected : expectedContent) {
            assertTrue(htmlResult.contains(expected), "HTML should contain: " + expected);
        }
    }

    @Nested
    @DisplayName("Core EML Parsing")
    class CoreParsingTests {
        @Test
        @DisplayName("Should parse simple text email correctly")
        void parseSimpleTextEmail() throws IOException {
            String emlContent = createSimpleTextEmail(
                    "sender@example.com",
                    "recipient@example.com",
                    "Simple Test Subject",
                    "This is a simple plain text email body.");

            testEmailConversion(emlContent, new String[] {
                "Simple Test Subject",
                "sender@example.com",
                "recipient@example.com",
                "This is a simple plain text email body",
                "<!DOCTYPE html>"
            }, false);
        }

        @Test
        @DisplayName("Should parse email with missing Subject and To headers")
        void parseEmailWithMissingHeaders() throws IOException {
            String emlContent = createEmailWithCustomHeaders();

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            assertTrue(htmlResult.contains("sender@example.com"));
            assertTrue(htmlResult.contains("This is an email body"));
            assertTrue(htmlResult.contains("<title></title>") ||
                       htmlResult.contains("<title>No Subject</title>"));
        }

        @Test
        @DisplayName("Should parse HTML email with styling")
        void parseHtmlEmailWithStyling() throws IOException {
            String htmlBody = "<html><head><style>.header{color:blue;font-weight:bold;}" +
                              ".content{margin:10px;}.footer{font-size:12px;}</style></head>" +
                              "<body><div class=\"header\">Important Notice</div>" +
                              "<div class=\"content\">This is <strong>HTML content</strong> with styling.</div>" +
                              "<div class=\"footer\">Best regards</div></body></html>";

            String emlContent = createHtmlEmail(
                    "html@example.com", "user@example.com", "HTML Email Test", htmlBody);

            testEmailConversion(emlContent, new String[] {
                "HTML Email Test",
                "Important Notice",
                "HTML content",
                "font-weight: bold"
            }, false);
        }

        @Test
        @DisplayName("Should parse multipart email with attachments")
        void parseMultipartEmailWithAttachments() throws IOException {
            String boundary = "----=_Part_" + System.currentTimeMillis();
            String emlContent = createMultipartEmailWithAttachment(
                    "multipart@example.com",
                    "user@example.com",
                    "Multipart Email Test",
                    "This email has both text content and an attachment.",
                    boundary,
                    "document.txt",
                    "Sample attachment content");

            testEmailConversion(emlContent, new String[] {
                "Multipart Email Test",
                "This email has both text content"
            }, true);
        }
    }

    @Nested
    @DisplayName("Email Encoding Support")
    class EncodingTests {

        @Test
        @DisplayName("Should handle international characters and UTF-8")
        void handleInternationalCharacters() throws IOException {
            String bodyWithIntlChars = "Hello! 你好 Привет مرحبا Hëllö Thañks!";
            String emlContent = createSimpleTextEmail(
                "intl@example.com",
                "user@example.com",
                "International Characters Test",
                bodyWithIntlChars);

            testEmailConversion(emlContent, new String[] {
                "你好", "Привет", "مرحبا", "Hëllö"
            }, false);
        }

        @Test
        @DisplayName("Should decode quoted-printable content correctly")
        void decodeQuotedPrintableContent() throws IOException {
            String content = "From: sender@example.com\n" +
                    "To: recipient@example.com\n" +
                    "Subject: Quoted-Printable Test\n" +
                    "Date: " + getTimestamp() + "\n" +
                    "MIME-Version: 1.0\n" +
                    "Content-Type: text/plain; charset=UTF-8\n" +
                    "Content-Transfer-Encoding: quoted-printable\n\n" +
                    "This is quoted=20printable content with special chars: =C3=A9=C3=A0=C3=A8.";

            testEmailConversion(content, new String[] {
                "Quoted-Printable Test",
                "This is quoted printable content with special chars: éàè."
            }, false);
        }

        @Test
        @DisplayName("Should decode Base64 content")
        void decodeBase64Content() throws IOException {
            String originalText = "This is Base64 encoded content: éàü ñ";
            String encodedContent = Base64.getEncoder().encodeToString(
                    originalText.getBytes(StandardCharsets.UTF_8));

            String content = "From: sender@example.com\n" +
                    "To: recipient@example.com\n" +
                    "Subject: Base64 Test\n" +
                    "Date: " + getTimestamp() + "\n" +
                    "MIME-Version: 1.0\n" +
                    "Content-Type: text/plain; charset=UTF-8\n" +
                    "Content-Transfer-Encoding: base64\n\n" +
                    encodedContent;

            testEmailConversion(content, new String[] {
                "Base64 Test", "Base64 encoded content"
            }, false);
        }

        @Test
        @DisplayName("Should correctly handle inline images with CID references")
        void handleInlineImages() throws IOException {
            String boundary = "----=_Part_CID_1234567890";
            String cid = "image123@example.com";
            String htmlBody = "<html><body><p>Here is an image:</p><img src=\"cid:" + cid + "\"></body></html>";
            String imageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

            String emlContent = createEmailWithInlineImage(htmlBody, boundary, cid, imageBase64);

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createRequestWithAttachments();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            assertTrue(htmlResult.contains("data:image/png;base64," + imageBase64));
            assertFalse(htmlResult.contains("cid:" + cid));
        }
    }

    @Nested
    @DisplayName("HTML Output Quality")
    class HtmlOutputTests {

        @Test
        @DisplayName("Should generate valid HTML structure")
        void generateValidHtmlStructure() throws IOException {
            String emlContent = createSimpleTextEmail(
                    "structure@test.com",
                    "user@test.com",
                    "HTML Structure Test",
                    "Testing HTML structure output");

            testEmailConversion(emlContent, new String[] {
                "<!DOCTYPE html>", "<html", "</html>", "HTML Structure Test"
            }, false);

            // Verify substantial content
            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, createBasicRequest());
            assertTrue(htmlResult.length() > 100, "HTML should have substantial content");
        }

        @Test
        @DisplayName("Should preserve safe CSS and remove problematic styles")
        void handleCssStylesCorrectly() throws IOException {
            String styledHtml = "<html><head><style>" +
                    ".safe { color: blue; font-size: 14px; }" +
                    ".problematic { position: fixed; word-break: break-all; }" +
                    ".good { margin: 10px; padding: 5px; }" +
                    "</style></head><body>" +
                    "<div class=\"safe\">Safe styling</div>" +
                    "<div class=\"problematic\">Problematic styling</div>" +
                    "<div class=\"good\">Good styling</div>" +
                    "</body></html>";

            String emlContent = createHtmlEmail("css@test.com", "user@test.com", "CSS Test", styledHtml);

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            // Should preserve safe CSS
            assertTrue(htmlResult.contains("color: blue"));
            assertTrue(htmlResult.contains("font-size: 14px"));
            assertTrue(htmlResult.contains("margin: 10px"));

            // Should remove problematic CSS
            assertFalse(htmlResult.contains("position: fixed"));
        }

        @Test
        @DisplayName("Should handle complex nested HTML structures")
        void handleComplexNestedHtml() throws IOException {
            String complexHtml = "<html><head><title>Complex Email</title></head><body>" +
                    "<div class=\"container\"><header><h1>Email Header</h1></header><main><section>" +
                    "<p>Paragraph with <a href=\"https://example.com\">link</a></p><ul>" +
                    "<li>List item 1</li><li>List item 2 with <em>emphasis</em></li></ul><table>" +
                    "<tr><td>Cell 1</td><td>Cell 2</td></tr><tr><td>Cell 3</td><td>Cell 4</td></tr>" +
                    "</table></section></main></div></body></html>";

            String emlContent = createHtmlEmail(
                    "complex@test.com", "user@test.com", "Complex HTML Test", complexHtml);

            testEmailConversion(emlContent, new String[] {
                "Email Header", "List item 2", "Cell 3", "example.com"
            }, false);

            // Verify substantial content
            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, createBasicRequest());
            assertTrue(htmlResult.length() > 300, "HTML should have substantial content");
        }
    }

    @Nested
    @DisplayName("Error Handling & Edge Cases")
    class ErrorHandlingTests {

        @Test
        @DisplayName("Should reject null input")
        void rejectNullInput() {
            EmlToPdfRequest request = createBasicRequest();

            Exception exception = assertThrows(
                    IllegalArgumentException.class,
                    () -> EmlToPdf.convertEmlToHtml(null, request));
            assertTrue(exception.getMessage().contains("EML file is empty or null"));
        }

        @Test
        @DisplayName("Should reject empty input")
        void rejectEmptyInput() {
            EmlToPdfRequest request = createBasicRequest();

            Exception exception = assertThrows(
                    IllegalArgumentException.class,
                    () -> EmlToPdf.convertEmlToHtml(new byte[0], request));
            assertTrue(exception.getMessage().contains("EML file is empty or null"));
        }

        @Test
        @DisplayName("Should handle malformed EML gracefully")
        void handleMalformedEmlGracefully() {
            String malformedEml = """
                    From: sender@test.com
                    Subject: Malformed EML
                    This line breaks header format
                    Content-Type: text/plain

                    Body content""";

            byte[] emlBytes = malformedEml.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();

            try {
                String result = EmlToPdf.convertEmlToHtml(emlBytes, request);
                assertNotNull(result, "Result should not be null");
                assertFalse(result.isEmpty(), "Result should not be empty");
            } catch (Exception e) {
                assertNotNull(e.getMessage(), "Exception message should not be null");
                assertFalse(e.getMessage().isEmpty(), "Exception message should not be empty");
            }
        }

        @Test
        @DisplayName("Should reject invalid EML format")
        void rejectInvalidEmlFormat() {
            byte[] invalidEml = "This is definitely not an EML file".getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();

            Exception exception = assertThrows(
                    IllegalArgumentException.class,
                    () -> EmlToPdf.convertEmlToHtml(invalidEml, request));
            assertTrue(exception.getMessage().contains("Invalid EML file format"));
        }
    }

    @Nested
    @DisplayName("Advanced Parsing Tests (Jakarta Mail)")
    class AdvancedParsingTests {

        @Test
        @DisplayName("Should successfully parse email using advanced parser")
        void initializeAngusMailSession() {
            assertDoesNotThrow(() -> {
                String emlContent = createSimpleTextEmail(
                        "angus@test.com",
                        "user@test.com",
                        "Angus Mail Test",
                        "Testing Angus Mail integration.");

                byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
                EmlToPdfRequest request = createBasicRequest();

                String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);
                assertNotNull(htmlResult);
                assertTrue(htmlResult.contains("Angus Mail Test"));
            });
        }

        @Test
        @DisplayName("Should parse complex MIME structures and select HTML part")
        void parseComplexMimeStructures() throws IOException {
            String boundary = "----=_Advanced_1234567890";
            String textBody = "This is the plain text part.";
            String htmlBody = "<html><body><h1>This is the HTML part</h1></body></html>";
            String emlContent = createMultipartAlternativeEmail(textBody, htmlBody, boundary);

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createRequestWithAttachments();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            assertTrue(htmlResult.contains("This is the HTML part"));
            assertFalse(htmlResult.contains("This is the plain text part"));
        }
    }

    @Nested
    @DisplayName("Additional Correctness Tests")
    class AdditionalCorrectnessTests {

        @Test
        @DisplayName("Should handle email with only an attachment and no body")
        void handleAttachmentOnlyEmail() throws IOException {
            String boundary = "----=_Part_AttachmentOnly_1234567890";
            String emlContent = createMultipartEmailWithAttachment(
                    "sender@example.com",
                    "recipient@example.com",
                    "Attachment Only Test",
                    "",
                    boundary,
                    "data.bin",
                    "binary data");

            testEmailConversion(emlContent, new String[] {
                "Attachment Only Test", "data.bin", "No content available"
            }, true);
        }

        @Test
        @DisplayName("Should handle mixed inline and regular attachments")
        void handleMixedAttachments() throws IOException {
            String boundary = "----=_Part_MixedAttachments_1234567890";
            String cid = "inline_image@example.com";
            String htmlBody = "<html><body><img src=\"cid:" + cid + "\"></body></html>";
            String imageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR42mNkAAIAAAoAAb6A/yoAAAAASUVORK5CYII=";
            String attachmentText = "This is a text attachment.";

            String emlContent = createEmailWithMixedAttachments(
                    htmlBody, boundary, cid, imageBase64, attachmentText);

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createRequestWithAttachments();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            assertTrue(htmlResult.contains("data:image/png;base64," + imageBase64));
            assertTrue(htmlResult.contains("text.txt"));
        }

        @Test
        @DisplayName("Should handle non-standard but valid character sets like ISO-8859-1")
        void handleIso88591Charset() throws IOException {
            String subject = "Subject with special characters: ñ é ü";
            String body = "Body with special characters: ñ é ü";

            String emlContent = createSimpleTextEmailWithCharset(
                    "sender@example.com",
                    "recipient@example.com",
                    subject,
                    body,
                    "ISO-8859-1");

            byte[] emlBytes = emlContent.getBytes(StandardCharsets.ISO_8859_1);
            EmlToPdfRequest request = createBasicRequest();

            String htmlResult = EmlToPdf.convertEmlToHtml(emlBytes, request);

            assertNotNull(htmlResult);
            assertTrue(htmlResult.contains(subject));
            assertTrue(htmlResult.contains(body));
        }

        @Test
        @DisplayName("Should handle emails with extremely long lines")
        void handleLongLines() throws IOException {
            StringBuilder longLine = new StringBuilder("This is a very long line: ");
            for (int i = 0; i < 1000; i++) {
                longLine.append("word").append(i).append(" ");
            }

            String emlContent = createSimpleTextEmail(
                    "sender@example.com",
                    "recipient@example.com",
                    "Long Line Test",
                    longLine.toString());

            testEmailConversion(emlContent, new String[] {
                "Long Line Test", "This is a very long line", "word999"
            }, false);
        }

        @ParameterizedTest
        @ValueSource(strings = {"windows-1252", "ISO-8859-2", "KOI8-R", "Shift_JIS"})
        @DisplayName("Should handle various character encodings")
        void handleVariousEncodings(String charset) throws IOException {
            String subject = "Encoding Test";
            String body = "Testing " + charset + " encoding";

            String emlContent = createSimpleTextEmailWithCharset(
                    "sender@example.com",
                    "recipient@example.com",
                    subject,
                    body,
                    charset);

            testEmailConversion(emlContent, new String[] {subject, body}, false);
        }
    }

    @Nested
    @ExtendWith(MockitoExtension.class)
    @DisplayName("PDF Conversion Tests")
    class PdfConversionTests {

        @Mock private CustomPDFDocumentFactory mockPdfDocumentFactory;

        @Mock private PDDocument mockPdDocument;

        @Test
        @DisplayName("Should convert EML to PDF without attachments when not requested")
        void convertEmlToPdfWithoutAttachments() throws Exception {
            String emlContent =
                    createSimpleTextEmail("from@test.com", "to@test.com", "Subject", "Body");
            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();
            byte[] fakePdfBytes = "fake-pdf".getBytes();

            try (MockedStatic<FileToPdf> fileToPdf = mockStatic(FileToPdf.class)) {
                fileToPdf
                        .when(
                                () ->
                                        FileToPdf.convertHtmlToPdf(
                                                anyString(),
                                                any(),
                                                any(byte[].class),
                                                anyString(),
                                                anyBoolean()))
                        .thenReturn(fakePdfBytes);

                byte[] resultPdf =
                        EmlToPdf.convertEmlToPdf(
                                "weasyprint",
                                request,
                                emlBytes,
                                "test.eml",
                                false,
                                mockPdfDocumentFactory);

                assertArrayEquals(fakePdfBytes, resultPdf);
                fileToPdf.verify(
                        () ->
                                FileToPdf.convertHtmlToPdf(
                                        anyString(),
                                        any(),
                                        any(byte[].class),
                                        anyString(),
                                        anyBoolean()));
                verify(mockPdfDocumentFactory, never()).load(any(byte[].class));
            }
        }

        @Test
        @DisplayName("Should convert EML to PDF with attachments when requested")
        void convertEmlToPdfWithAttachments() throws Exception {
            // Arrange
            String boundary = "----=_Part_1234567890";
            String emlContent = createMultipartEmailWithAttachment(
                    "multipart@example.com",
                    "user@example.com",
                    "Multipart Email Test",
                    "This email has both text content and an attachment.",
                    boundary,
                    "document.txt",
                    "Sample attachment content");
            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createRequestWithAttachments(); // includeAttachments = true
            byte[] fakePdfBytes = "fake-pdf".getBytes();

            try (MockedStatic<FileToPdf> fileToPdf = mockStatic(FileToPdf.class)) {
                fileToPdf
                        .when(
                                () ->
                                        FileToPdf.convertHtmlToPdf(
                                                anyString(),
                                                any(),
                                                any(byte[].class),
                                                anyString(),
                                                anyBoolean()))
                        .thenReturn(fakePdfBytes);

                try (MockedStatic<EmlToPdf> emlToPdf = mockStatic(EmlToPdf.class, invocation -> {
                    String methodName = invocation.getMethod().getName();
                    return switch (methodName) {
                        case "shouldAttachFiles" -> true;
                        case "attachFilesToPdf" -> fakePdfBytes;
                        default -> invocation.callRealMethod();
                    };
                })) {
                    byte[] resultPdf =
                            EmlToPdf.convertEmlToPdf(
                                    "weasyprint",
                                    request,
                                    emlBytes,
                                    "test.eml",
                                    false,
                                    mockPdfDocumentFactory);

                    assertArrayEquals(fakePdfBytes, resultPdf);

                    fileToPdf.verify(
                            () ->
                                    FileToPdf.convertHtmlToPdf(
                                            anyString(),
                                            any(),
                                            any(byte[].class),
                                            anyString(),
                                            anyBoolean()));
                }
            }
        }

        @Test
        @DisplayName("Should handle errors during EML to PDF conversion")
        void handleErrorsDuringConversion() throws Exception {
            // Arrange
            String emlContent =
                    createSimpleTextEmail("from@test.com", "to@test.com", "Subject", "Body");
            byte[] emlBytes = emlContent.getBytes(StandardCharsets.UTF_8);
            EmlToPdfRequest request = createBasicRequest();
            String errorMessage = "Conversion failed";

            try (MockedStatic<FileToPdf> fileToPdf = mockStatic(FileToPdf.class)) {
                fileToPdf
                        .when(
                                () ->
                                        FileToPdf.convertHtmlToPdf(
                                                anyString(),
                                                any(),
                                                any(byte[].class),
                                                anyString(),
                                                anyBoolean()))
                        .thenThrow(new IOException(errorMessage));

                // Act & Assert
                IOException exception = assertThrows(
                        IOException.class,
                        () -> EmlToPdf.convertEmlToPdf(
                                "weasyprint",
                                request,
                                emlBytes,
                                "test.eml",
                                false,
                                mockPdfDocumentFactory));

                assertTrue(exception.getMessage().contains(errorMessage));
            }
        }
    }

    // Helper methods
    private String getTimestamp() {
        // Use a fixed date-time for deterministic test results
        java.time.ZonedDateTime fixedDateTime = java.time.ZonedDateTime.of(2023, 1, 1, 12, 0, 0, 0, java.time.ZoneId.of("GMT"));
        return java.time.format.DateTimeFormatter.RFC_1123_DATE_TIME.format(fixedDateTime);
    }

    private String createSimpleTextEmail(String from, String to, String subject, String body) {
        return createSimpleTextEmailWithCharset(from, to, subject, body, "UTF-8");
    }

    private String createSimpleTextEmailWithCharset(String from, String to, String subject, String body, String charset) {
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nContent-Type: text/plain; charset=%s\nContent-Transfer-Encoding: 8bit\n\n%s",
                from, to, subject, getTimestamp(), charset, body);
    }

    private String createEmailWithCustomHeaders() {
        return String.format(
                "From: sender@example.com\nDate: %s\nContent-Type: text/plain; charset=UTF-8\nContent-Transfer-Encoding: 8bit\n\n%s",
                getTimestamp(), "This is an email body with some headers missing.");
    }

    private String createHtmlEmail(String from, String to, String subject, String htmlBody) {
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nContent-Type: text/html; charset=UTF-8\nContent-Transfer-Encoding: 8bit\n\n%s",
                from, to, subject, getTimestamp(), htmlBody);
    }

    private String createMultipartEmailWithAttachment(String from, String to, String subject, String body,
            String boundary, String filename, String attachmentContent) {
        String encodedContent = Base64.getEncoder().encodeToString(attachmentContent.getBytes(StandardCharsets.UTF_8));
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nContent-Type: multipart/mixed; boundary=\"%s\"\n\n" +
                "--%s\nContent-Type: text/plain; charset=UTF-8\nContent-Transfer-Encoding: 8bit\n\n%s\n\n" +
                "--%s\nContent-Type: text/plain; charset=UTF-8\nContent-Disposition: attachment; filename=\"%s\"\n" +
                "Content-Transfer-Encoding: base64\n\n%s\n\n--%s--",
                from, to, subject, getTimestamp(), boundary, boundary, body, boundary, filename, encodedContent, boundary);
    }

    private String createMultipartAlternativeEmail(String textBody, String htmlBody, String boundary) {
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nMIME-Version: 1.0\nContent-Type: multipart/alternative; boundary=\"%s\"\n\n" +
                "--%s\nContent-Type: text/plain; charset=UTF-8\nContent-Transfer-Encoding: 7bit\n\n%s\n\n" +
                "--%s\nContent-Type: text/html; charset=UTF-8\nContent-Transfer-Encoding: 7bit\n\n%s\n\n--%s--",
                "sender@example.com", "receiver@example.com", "Multipart/Alternative Test", getTimestamp(),
                boundary, boundary, textBody, boundary, htmlBody, boundary);
    }

    private String createEmailWithInlineImage(String htmlBody, String boundary, String contentId, String base64Image) {
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nContent-Type: multipart/related; boundary=\"%s\"\n\n" +
                "--%s\nContent-Type: text/html; charset=UTF-8\nContent-Transfer-Encoding: 8bit\n\n%s\n\n" +
                "--%s\nContent-Type: image/png\nContent-Transfer-Encoding: base64\nContent-ID: <%s>\n" +
                "Content-Disposition: inline; filename=\"image.png\"\n\n%s\n\n--%s--",
                "sender@example.com", "receiver@example.com", "Inline Image Test", getTimestamp(),
                boundary, boundary, htmlBody, boundary, contentId, base64Image, boundary);
    }

    private String createEmailWithMixedAttachments(String htmlBody, String boundary, String contentId,
            String base64Image, String attachmentBody) {
        String encodedAttachment = Base64.getEncoder().encodeToString(attachmentBody.getBytes(StandardCharsets.UTF_8));
        return String.format(
                "From: %s\nTo: %s\nSubject: %s\nDate: %s\nContent-Type: multipart/mixed; boundary=\"%s\"\n\n" +
                "--%s\nContent-Type: multipart/related; boundary=\"related-%s\"\n\n" +
                "--related-%s\nContent-Type: text/html; charset=UTF-8\nContent-Transfer-Encoding: 8bit\n\n%s\n\n" +
                "--related-%s\nContent-Type: image/png\nContent-Transfer-Encoding: base64\nContent-ID: <%s>\n" +
                "Content-Disposition: inline; filename=\"image.png\"\n\n%s\n\n--related-%s--\n\n" +
                "--%s\nContent-Type: text/plain; charset=UTF-8\nContent-Disposition: attachment; filename=\"%s\"\n" +
                "Content-Transfer-Encoding: base64\n\n%s\n\n--%s--",
                "sender@example.com", "receiver@example.com", "Mixed Attachments Test", getTimestamp(),
                boundary, boundary, boundary, boundary, htmlBody, boundary, contentId, base64Image, boundary,
                boundary, "text.txt", encodedAttachment, boundary);
    }

    private EmlToPdfRequest createBasicRequest() {
        EmlToPdfRequest request = new EmlToPdfRequest();
        request.setIncludeAttachments(false);
        return request;
    }

    private EmlToPdfRequest createRequestWithAttachments() {
        EmlToPdfRequest request = new EmlToPdfRequest();
        request.setIncludeAttachments(true);
        request.setMaxAttachmentSizeMB(10);
        return request;
    }
}
