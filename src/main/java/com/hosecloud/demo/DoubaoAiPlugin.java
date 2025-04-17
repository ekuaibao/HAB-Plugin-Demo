package com.hosecloud.demo;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.hosecloud.demo.vo.DoubaoAiResult;
import com.hosecloud.hab.plugin.BaseTaskPlugin;
import com.hosecloud.hab.plugin.annotation.Execute;
import com.hosecloud.hab.plugin.annotation.JsonSchemaDefinition;
import com.hosecloud.hab.plugin.annotation.JsonSchemaProperty;
import com.hosecloud.hab.plugin.model.Log;
import lombok.Setter;
import org.pf4j.Extension;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Setter
@Extension
@JsonSchemaDefinition(
        title = "豆包AI对话节点",
        description = "调用豆包AI进行对话，支持文本和图片输入"
)
public class DoubaoAiPlugin extends BaseTaskPlugin {

    @JsonSchemaProperty(
            title = "API密钥",
            description = "豆包AI的API密钥",
            required = true,
            example = "91b363f5-7092-46fd-a0aa-d1b96ba5c780"
    )
    private String apiKey;

    @JsonSchemaProperty(
            title = "对话内容",
            description = "要发送给AI的对话内容",
            required = true,
            example = "这是哪里"
    )
    private String content;

    @JsonSchemaProperty(
            title = "图片URL",
            description = "要发送给AI的图片URL，可选",
            required = false,
            example = "https://example.com/image.jpg"
    )
    private String imageUrl;

    @Override
    public String getName() {
        return "task-dynamic-doubao-ai";
    }

    @Execute(
            description = "执行豆包AI对话",
            outputClass = DoubaoAiResult.class
    )
    public DoubaoAiResult run() {
        try {
            // 构建请求URL
            URL url = new URL("https://ark.cn-beijing.volces.com/api/v3/chat/completions");
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("POST");
            conn.setRequestProperty("Content-Type", "application/json");
            conn.setRequestProperty("Authorization", "Bearer " + apiKey);
            conn.setDoOutput(true);

            // 构建请求体
            Map<String, Object> requestBody = new HashMap<>();
            requestBody.put("model", "doubao-1-5-vision-pro-32k-250115");

            List<Map<String, Object>> messages = new ArrayList<>();
            Map<String, Object> message = new HashMap<>();
            message.put("role", "user");

            List<Map<String, Object>> contentList = new ArrayList<>();
            
            // 如果有图片URL，添加图片内容
            if (imageUrl != null && !imageUrl.trim().isEmpty()) {
                Map<String, Object> imageContent = new HashMap<>();
                imageContent.put("type", "image_url");
                Map<String, Object> imageUrlMap = new HashMap<>();
                imageUrlMap.put("url", imageUrl);
                imageContent.put("image_url", imageUrlMap);
                contentList.add(imageContent);
            }

            // 添加文本内容
            Map<String, Object> textContent = new HashMap<>();
            textContent.put("type", "text");
            textContent.put("text", content);
            contentList.add(textContent);

            message.put("content", contentList);
            messages.add(message);
            requestBody.put("messages", messages);

            // 发送请求
            ObjectMapper mapper = new ObjectMapper();
            String jsonBody = mapper.writeValueAsString(requestBody);
            try (var os = conn.getOutputStream()) {
                byte[] input = jsonBody.getBytes(StandardCharsets.UTF_8);
                os.write(input, 0, input.length);
            }

            // 读取响应
            StringBuilder response = new StringBuilder();
            try (BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                String responseLine;
                while ((responseLine = br.readLine()) != null) {
                    response.append(responseLine.trim());
                }
            }

            // 解析响应
            Map<String, Object> responseMap = mapper.readValue(response.toString(), new TypeReference<>() {
            });
            List<Map<String, Object>> choices = (List<Map<String, Object>>) responseMap.get("choices");
            Map<String, Object> firstChoice = choices.get(0);
            Map<String, Object> messageResponse = (Map<String, Object>) firstChoice.get("message");
            String aiContent = (String) messageResponse.get("content");

            // 记录日志
            executeLogs.add(Log.success("成功调用豆包AI并获取响应"));

            // 返回结果
            DoubaoAiResult result = new DoubaoAiResult();
            result.setMessage("调用成功");
            result.setContent(aiContent);
            return result;

        } catch (Exception e) {
            executeLogs.add(Log.failure("调用豆包AI失败: " + e.getMessage()));
            return new DoubaoAiResult("调用豆包AI失败: " + e.getMessage());
        }
    }
} 