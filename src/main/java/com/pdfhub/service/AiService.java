//package com.pdfhub.service;
//
//import com.fasterxml.jackson.databind.ObjectMapper;
//import com.fasterxml.jackson.databind.node.ObjectNode;
//import okhttp3.*;
//import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.text.PDFTextStripper;
//import org.springframework.beans.factory.annotation.Value;
//import org.springframework.core.io.Resource;
//import org.springframework.stereotype.Service;
//import org.springframework.web.multipart.MultipartFile;
//
//import com.google.cloud.translate.Translate;
//import com.google.cloud.translate.TranslateOptions;
//import com.google.cloud.translate.Translation;
//
//
//import java.io.IOException;
//import java.io.InputStream;
//
//@Service
//public class AiService {
//
//    @Value("${spring.ai.openai.api-key}")
//    private String API_KEY;
//    private static final String OPENAI_URL = "https://api.openai.com/v1/chat/completions";
//
//
//    public Resource summarizePdf(MultipartFile file) throws IOException {
//         String textData=file.getBytes().toString();
//        return null;
//    }
//
//    public Resource translatePdf(MultipartFile file) {
//        return null;
//    }
//
//    public Resource generateQuestions(MultipartFile file) {
//        return null;
//    }
//
//
//
//
////    public String extractTextFromPdf(File pdfFile) throws IOException {
////
////        PDDocument document = new PDDocument();
////        document.save(pdfFile);
////        PDFTextStripper pdfStripper = new PDFTextStripper();
////        String text = pdfStripper.getText(document);
////        document.close();
////        return text;
////    }
//
//    public String extractTextFromPdf(InputStream inputStream) throws IOException {
//        PDDocument document = new PDDocument();
//        document.save(inputStream.toString());// Load the PDF from InputStream
//        PDFTextStripper pdfStripper = new PDFTextStripper();
//        String text = pdfStripper.getText(document);
//        document.close();
//        return text;
//    }
//
//    // Basic summarization method: Shortens the text by splitting into paragraphs or sentences.
////    public String summarizeText(String text) {
////        // Custom prompt for summarization
////        String prompt = "You are a text summarizer. Summarize the following text in a concise and clear manner, keeping only the key points:\n" + text;
////        OpenAiApi openAiApi=new OpenAiApi("ABC&*HBDIHOHDUNLUOD");
////        openAiApi.chatCompletionEntity()
////        // Build the OpenAI request for summarization
////        OpenAiApi.ChatCompletionRequest completionRequest = OpenAiApi.ChatCompletionRequest.ResponseFormat.builder()
////                .model("text-davinci-003")  // Use GPT-3 model or GPT-4 if available
////                .prompt(prompt)
////                .maxTokens(150)  // Limit summary length
////                .temperature(0.7)  // Control creativity and accuracy
////                .build();
////       OpenAiApi.ChatCompletionRequest.ResponseFormat data= completionRequest.responseFormat();
////        // Get the response from OpenAI
////
////        // Return the summarized text
////        return data.toString().isEmpty() ? "Summary not available" : data.toString();
////    }
//
//
////    private static final String API_KEY = "ABC&*HBDIHOHDUNLUOD";
//
//    public String summarizeText(String text) {
//        // Custom prompt for summarization
//        String prompt = "You are a text summarizer. Summarize the following text in a concise and clear manner, keeping only the key points count the words and precise its one by three:\n" + text;
//
//        // Create a JSON payload for the API request
//        ObjectMapper objectMapper = new ObjectMapper();
//        ObjectNode requestBody = objectMapper.createObjectNode();
//        requestBody.put("model", "gpt-3.5-turbo");
//        requestBody.putArray("messages").addObject()
//                .put("role", "user")
//                .put("content", prompt);
//        requestBody.put("max_tokens", countWordsByThree(text));
//        requestBody.put("temperature", 0.7);  // Control creativity and accuracy
//
//        // Initialize OkHttpClient to make the HTTP request
//        OkHttpClient client = new OkHttpClient();
//        RequestBody body = RequestBody.create(requestBody.toString(), MediaType.parse("application/json"));
//        Request request = new Request.Builder()
//                .url(OPENAI_URL)
//                .header("Authorization", "Bearer " + API_KEY)
//                .post(body)
//                .build();
//
//        try {
//            // Make the API request and get the response
//            Response response = client.newCall(request).execute();
//
//            if (response.isSuccessful()) {
//                String responseBody = response.body().string();
//
//                // Parse the response body to extract the summarized text
//                ObjectNode responseJson = (ObjectNode) objectMapper.readTree(responseBody);
//                String summarizedText = responseJson.path("choices").get(0).path("message").path("content").asText();
//
//                return summarizedText.isEmpty() ? "Summary not available" : summarizedText;
//            } else {
//                return "Error: " + response.message();
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//            return "Error occurred while fetching the summary";
//        }
//    }
//
//    public static int countWordsByThree(String text) {
//        // Check if the input text is null or empty
//        if (text == null || text.trim().isEmpty()) {
//            return 0;  // Return 0 if text is null or empty
//        }
//
//        // Split the text by one or more whitespace characters and get the resulting array
//        String[] words = text.trim().split("\\s+");
//
//        // Return the word count divided by 3, rounded to the nearest integer
//        return (int) Math.round((double) words.length / 3);
//    }
//
////    public String extractTextFromPdf(File pdfFile) throws IOException {
////        PDDocument document = PDDocument.load(pdfFile);
////        PDFTextStripper pdfStripper = new PDFTextStripper();
////        String text = pdfStripper.getText(document);
////        document.close();
////        return text;
////    }
//
//    // Method to translate text using Google Cloud Translation API
//    public String translateText(String text, String targetLanguage) {
//        Translate translate = TranslateOptions.getDefaultInstance().getService();
//        Translation translation = translate.translate(
//                text,
//                Translate.TranslateOption.targetLanguage(targetLanguage)
//        );
//        return translation.getTranslatedText();
//    }
//}