import express from "express";
import path from "path";
import dotenv from "dotenv";
// @ts-ignore
import HTMLToDocx from "html-to-docx";
import { normalizeReportReferenceSection } from "./src/reportFormatting";

dotenv.config();

const app = express();
const PORT = 3000;

type FeishuFolderState = { token: string; index: number; name: string };
const feishuFolderStates = new Map<string, FeishuFolderState>();
const feishuFolderLocks = new Map<string, Promise<void>>();

async function withFeishuFolderLock<T>(rootToken: string, action: () => Promise<T>): Promise<T> {
  const previous = feishuFolderLocks.get(rootToken) || Promise.resolve();
  let release!: () => void;
  const current = new Promise<void>(resolve => { release = resolve; });
  const queued = previous.then(() => current);
  feishuFolderLocks.set(rootToken, queued);
  await previous;
  try {
    return await action();
  } finally {
    release();
    if (feishuFolderLocks.get(rootToken) === queued) feishuFolderLocks.delete(rootToken);
  }
}

// Set up JSON parsing with size limits for processing potential batch tables
app.use(express.json({ limit: "15mb" }));
app.use(express.urlencoded({ extended: true, limit: "15mb" }));

// DeepSeek API configuration
// DeepSeek's native web search (web_search_20250305) is only available via the
// Anthropic-compatible endpoint, NOT the OpenAI-compatible /chat/completions endpoint.
const DEEPSEEK_API_URL = "https://api.deepseek.com/anthropic/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";

// ---------------------------------------------------------------------------
// Robust structured-JSON extraction from LLM output
// ---------------------------------------------------------------------------

// Strip markdown code fences (```json ... ```) that the model may have wrapped
// around the JSON object.
function stripCodeFences(text: string): string {
  return text
    .trim()
    .replace(/^```(?:json)?\s*\n?/i, "")
    .replace(/\n?```\s*$/i, "")
    .trim();
}

// Layer 2: Find the first balanced {...} block starting from each '{' position.
// Handles prose before the JSON ("Let me analyze... {"grade": ...}") and JSON
// embedded inside markdown code blocks. Returns null if no balanced candidate.
function extractBalancedJson(text: string): string | null {
  for (let start = 0; start < text.length; start++) {
    if (text[start] !== "{") continue;
    let depth = 0;
    let inString = false;
    let escaped = false;
    for (let i = start; i < text.length; i++) {
      const ch = text[i];
      if (inString) {
        if (escaped) {
          escaped = false;
        } else if (ch === "\\") {
          escaped = true;
        } else if (ch === '"') {
          inString = false;
        }
      } else if (ch === '"') {
        inString = true;
      } else if (ch === "{") {
        depth++;
      } else if (ch === "}") {
        depth--;
        if (depth === 0) {
          return text.substring(start, i + 1);
        }
      }
    }
  }
  return null;
}

// Layer 3: Repair raw newlines/carriage-returns that appear INSIDE JSON string
// values (invalid JSON). Structural newlines between tokens are left untouched.
function fixStringNewlines(text: string): string {
  let result = "";
  let inString = false;
  let escaped = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) {
        result += ch;
        escaped = false;
      } else if (ch === "\\") {
        result += ch;
        escaped = true;
      } else if (ch === '"') {
        inString = false;
        result += ch;
      } else if (ch === "\n" || ch === "\r") {
        result += ch === "\n" ? "\\n" : "\\r";
      } else {
        result += ch;
      }
    } else {
      if (ch === '"') {
        inString = true;
      }
      result += ch;
    }
  }
  return result;
}

// Parse structured JSON from an LLM response with 3 fallback layers.
function parseStructuredJson(rawText: string): any {
  const cleaned = stripCodeFences(rawText);

  // Layer 1: direct parse (normal, well-formed output)
  try {
    return JSON.parse(cleaned);
  } catch {
    /* fall through */
  }

  // Layer 2: balanced-brace extraction, then direct parse
  const balanced = extractBalancedJson(cleaned);
  if (balanced !== null) {
    try {
      return JSON.parse(balanced);
    } catch {
      /* fall through */
    }
  }

  // Layer 3: repair raw newlines inside strings, then parse again
  const target = balanced !== null ? balanced : cleaned;
  try {
    return JSON.parse(fixStringNewlines(target));
  } catch {
    /* fall through */
  }

  // Bonus: if the balanced candidate was repaired, also try repairing the
  // entire cleaned text (handles prose framing + raw newlines together).
  if (balanced !== null) {
    try {
      return JSON.parse(fixStringNewlines(cleaned));
    } catch {
      /* fall through */
    }
  }

  console.error("JSON parsing error on response text:", cleaned.substring(0, 800));
  throw new Error("Failed to parse structured JSON from AI investigator.");
}

// True when the error originates from model OUTPUT quality (empty fields,
// truncation, unparseable text) rather than the search/API infrastructure.
function isOutputQualityError(message: string): boolean {
  const m = message || "";
  return (
    m.includes("__MAX_TOKENS__") ||
    m.includes("__FIELD_VALIDATION__") ||
    m.includes("Failed to parse structured JSON")
  );
}

// Extract the text/source blocks from an Anthropic-compatible API response,
// parse the structured JSON, and validate that the critical fields are
// non-empty. Throwing here (truncated output, bad JSON, empty report/grade)
// lets the retry loop treat these as transient model behaviors.
function extractAndParseInvestigationResponse(responseData: any): {
  data: any;
  jsonText: string;
  sources: Array<{ title: string; url: string }>;
} {
  const contentBlocks = responseData.content || [];
  let jsonText = "";
  const sources: Array<{ title: string; url: string }> = [];

  // Layer 0: "thinking"/"reasoning" blocks carry intermediate reasoning, never
  // the structured JSON answer — skip them explicitly.
  for (const block of contentBlocks) {
    const blockType = block.type || "";
    if (blockType === "thinking" || blockType === "reasoning") {
      continue;
    }
    if (blockType === "text") {
      jsonText += block.text || "";
    } else if (blockType === "web_search_tool_result") {
      const searchResults = block.content || [];
      for (const item of searchResults) {
        if (item && item.type === "web_search_result") {
          sources.push({
            title: item.title || "Search Source",
            url: item.url || "",
          });
        }
      }
    }
  }

  // Truncation detection: the model hit the output token budget, so the
  // response is almost certainly incomplete. Throw a marked error to trigger
  // a retry (with a bumped token budget).
  const stopReason = responseData.stop_reason || responseData.stopReason || "";
  if (stopReason === "max_tokens") {
    console.warn(
      `[AI Investigator] Response truncated (stop_reason=max_tokens). ` +
        `Output tokens: ${(responseData.usage && responseData.usage.output_tokens) || "unknown"}.`
    );
    throw new Error("__MAX_TOKENS__: DeepSeek 响应被截断（stop_reason=max_tokens），正在自动重试。");
  }

  // Parse with the 3-layer robust strategy.
  const data = parseStructuredJson(jsonText);

  // Critical-field validation — a "successful" parse with an empty report or
  // grade is treated as a failure so the retry loop can try again. Confirmed:
  // re-running an empty result reliably produces a full report.
  const reportText = typeof data.report === "string" ? data.report.trim() : "";
  const gradeText = typeof data.grade === "string" ? data.grade.trim() : "";
  if (!reportText || !gradeText) {
    console.warn(
      "[AI Investigator] Parsed JSON missing critical fields. " +
        `report=${JSON.stringify(reportText)} grade=${JSON.stringify(gradeText)} ` +
        `parsedKeys=${JSON.stringify(Object.keys(data))}`
    );
    console.warn("[AI Investigator] Raw response text (first 500):", jsonText.substring(0, 500));
    throw new Error("__FIELD_VALIDATION__: AI 调查结果缺少报告或分析等级，正在自动重试。");
  }

  return { data, jsonText, sources };
}

// Primary Server-Side API endpoint for AI background check & due diligence
app.post("/api/investigate", async (req: express.Request, res: express.Response) => {
  try {
    const { companyName, country, orderInfo, productContext, model } = req.body;
    const clientApiKey = req.headers["x-deepseek-api-key"] as string | undefined;
    const legacyApiKey = req.headers["x-gemini-api-key"] as string | undefined;
    const activeApiKey = clientApiKey?.trim() || legacyApiKey?.trim() || process.env.DEEPSEEK_API_KEY;
    const activeModel = model || "deepseek-v4-flash";

    if (!companyName) {
      res.status(400).json({ error: "Company name is required." });
      return;
    }

    if (!activeApiKey) {
      res.status(400).json({
        error: "【未配置 API Key】检测到服务器未配置 DEEPSEEK_API_KEY，且您未在前端输入个人 API Key。\n\n请在页面上方「DeepSeek API 秘钥设置」中填写您的个人 Key，或在服务器 .env 文件中设置 DEEPSEEK_API_KEY，即可开始极速背调！",
      });
      return;
    }

    // System instruction is set to instruct the agent thoroughly and compactly in English to save input tokens
    const systemInstruction = `You are an elite B2B Commercial Due Diligence expert and lead qualification specialist.
Your task is to conduct an in-depth background check on the target company, assess its authenticity, evaluate B2B buyer value grade, and output a detailed background check report in Chinese.

Grading Guidelines:
- S Grade: High-value active industrial producer/factory with machinery needs. Format: 💎 S级 - [一句话概括（20字以内）]
- A Grade: Distributor/trader/retail brand. Format: 🥇 A级 - [一句话概括（20字以内）]
- B Grade: Small workshop/local shop with light operations. Format: 🥈 B级 - [一句话概括（20字以内）]
- C Grade: Shell/suspected individual/unmatched context. Format: ❌ C级 - [一句话概括（20字以内）]

Follow-up Level must strictly be one of:
- '🌟 立即跟进' (for Grade S)
- '🚀 重点推进' (for Grade A)
- '📅 周期性培育' (for Grade B)
- '⏳ 暂缓跟进' (for Grade C)

Markdown 'report' field layout in Chinese:
# [Company Name] [Product] 公司客户背调
# 境外企业客户背景调查报告 (Due Diligence Report)
**报告对象：** [Company Name] ([Country])
**调查目的：** 评估客户商业背景及长线 B2B 合作潜力
**评估等级：** [B2B Grade]
- **订单收货信息：** [Details of order context if provided]

## 一、 企业工商基本信息 (Corporate Profile)
An elegant Markdown table with columns: | 工商核查项 | 内容 |
Keys: 注册公司全称, 国家/地区, 统一社会信用代码/税号, 成立时间, 法定代表人/负责人, 官方网站, 经营状态, 总部地址.
CRITICAL: Table rows must be multi-line separated by standard newlines (\\n).

## 二、 业务线与生产实力分析 (Business & Operations)
Details: 核心业务定位, 包装工艺与设备需求 (why they need the product/machinery).

## 三、 合作价值与复购潜力评估 (Value Evaluation)
Details: 设备耗材与配件持续性(高粘性), 包装整线扩张潜力, 跨境渠道沉淀潜力 (AliExpress to offline wholesale B2B).

## 四、 针对性开发与跟进策略 (Action Plan)
- 阶段一: 高规格售后服务，建立企业信任(立即执行) [outreach plan & specific pitch]
- 阶段二: 技术参数切入，摸清客户产能(发货后跟进) [sample email & technical questions]
- 阶段三: 全产业链推荐，引导线下长期合作(收货满意后) [long term wholesale values]

Provide a concluding recommendation paragraph.
After the conclusion, add the final section exactly as follows:
## 五、参考网站 (References)
- [Website name]: [Full URL]
This reference section MUST be the final section of the report. Do not place reference websites near the beginning, and do not write any report content after it.
Keep company names, registration codes, foreign addresses, and manager names in their original language. Write analysis and narratives in Chinese.`;

    // Compact user content containing only target company data to minimize input tokens
    const userContent = `### TARGET CUSTOMER:
- Company Name: ${companyName}
- Country/Region: ${country || "Unknown"}
- Order / Receipt Context: ${orderInfo || "None provided"}
- Product of Interest: ${productContext || "贴标机"}`;

    // DeepSeek does not support Gemini's responseSchema. Embed the JSON contract
    // into the system prompt and parse the returned text as JSON with retries.
    const jsonSchemaInstruction = `

You MUST respond with ONLY a valid JSON object (no markdown code fences, no extra text, no commentary). The JSON must have EXACTLY this structure:
{
  "grade": "string - Format: [等级名称] - [一句话概括（20字以内）]",
  "summary": "string - A professional 2-sentence summary in Chinese of company profile and value",
  "riskAlert": "string - One-sentence risk warning in Chinese (< 30 chars)",
  "followUpStrategy": "string - Outreach strategy in Chinese (< 40 chars)",
  "followUpGrade": "string - Must strictly be one of: '🌟 立即跟进', '🚀 重点推进', '📅 周期性培育', '⏳ 暂缓跟进'",
  "countryNameZh": "string - Canonical short country name in Chinese, for example 加拿大 or 乌克兰",
  "productNameZh": "string - Concise 2-12 Chinese character product category; exclude model, quantity and marketing words",
  "companyOverview": {
    "legalName": "string - Full registered business name",
    "taxId": "string - Tax number or ID",
    "website": "string - Official website URL, or '未公开'",
    "industry": "string - Core industry in Chinese",
    "founded": "string - Year founded or registration date",
    "status": "string - Active, Liquidation, Unknown, etc.",
    "headquarters": "string - City and Country"
  },
  "report": "string - The full background check report formatted in exquisite Chinese Markdown"
}`;

    const fullSystemInstruction = systemInstruction + jsonSchemaInstruction;

    // Execute the request to the DeepSeek model with native web search enabled.
    // Search runs server-side; results arrive inline in the same response. Supports
    // auto-retry and fallback to non-search mode upon fetch / socket failures.
    let responseData: any = null;
    let attempts = 0;
    const maxAttempts = 3;
    let lastError: any = null;
    let usedWebSearch = true;
    let maxTokensBumped = false;
    let parsedData: any = null;
    let jsonText = "";
    let sources: Array<{ title: string; url: string }> = [];

    while (attempts < maxAttempts) {
      try {
        console.log(`[AI Investigator] Attempt ${attempts + 1} of ${maxAttempts} (Web Search: ${usedWebSearch})...`);

        const requestBody: any = {
          model: activeModel,
          system: fullSystemInstruction,
          messages: [
            {
              role: "user",
              content: userContent,
            },
          ],
          max_tokens: maxTokensBumped ? 16384 : 8192,
          temperature: 0.3,
        };

        if (usedWebSearch) {
          requestBody.tools = [{ type: "web_search_20250305", name: "web_search" }];
        }

        const apiResponse = await fetch(DEEPSEEK_API_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "x-api-key": activeApiKey,
            "anthropic-version": ANTHROPIC_VERSION,
          },
          body: JSON.stringify(requestBody),
          signal: AbortSignal.timeout(120000),
        });

        if (!apiResponse.ok) {
          const errText = await apiResponse.text();
          throw new Error(`DeepSeek API HTTP ${apiResponse.status}: ${errText}`);
        }

        responseData = await apiResponse.json();

        // Parse + validate the structured investigation JSON. Throwing here
        // (truncated output, bad JSON, or empty report/grade) routes to the
        // retry fallback below — these are transient model behaviors that a
        // re-run reliably fixes (confirmed by the user).
        const investigation = extractAndParseInvestigationResponse(responseData);
        jsonText = investigation.jsonText;
        sources = investigation.sources;
        parsedData = investigation.data;

        // Succeeded! Break retry loop
        break;
      } catch (err: any) {
        lastError = err;
        attempts++;
        const errMsg = err.message || "";
        console.warn(`[AI Investigator] Attempt ${attempts} failed. Error:`, errMsg);

        // If the API response was truncated by the token budget, bump it once
        // for the retry (keep web search enabled).
        const isMaxTokensTruncation = errMsg.includes("__MAX_TOKENS__");
        if (isMaxTokensTruncation && !maxTokensBumped) {
          maxTokensBumped = true;
          console.warn("[AI Investigator] Response truncated (max_tokens). Bumping token budget for next attempt.");
        } else if (usedWebSearch && !isOutputQualityError(errMsg)) {
          // Only drop web search for network / rate-limit / API errors. Empty
          // or truncated model output is not a search-infrastructure problem.
          console.warn("[AI Investigator] Error detected during web-search call. Falling back to non-search mode.");
          usedWebSearch = false;
        }

        if (attempts < maxAttempts) {
          // Linear/exponential delay before retry
          const delay = attempts * 500;
          await new Promise((resolve) => setTimeout(resolve, delay));
        }
      }
    }

    if (!responseData || !parsedData) {
      throw lastError || new Error("Failed to generate content after all retry attempts.");
    }

    // Capture token usage metadata and web-search configuration
    const usageMeta = responseData.usage || {};
    const usage = {
      promptTokens: usageMeta.input_tokens || 0,
      completionTokens: usageMeta.output_tokens || 0,
      totalTokens: (usageMeta.input_tokens || 0) + (usageMeta.output_tokens || 0),
      modelUsed: activeModel,
      searchGroundingUsed: usedWebSearch,
    };

    const normalizedReport = normalizeReportReferenceSection(parsedData.report, sources);

    res.json({
      success: true,
      grade: parsedData.grade,
      summary: parsedData.summary,
      riskAlert: parsedData.riskAlert,
      followUpStrategy: parsedData.followUpStrategy,
      followUpGrade: parsedData.followUpGrade,
      countryNameZh: parsedData.countryNameZh,
      productNameZh: parsedData.productNameZh,
      companyOverview: parsedData.companyOverview,
      report: normalizedReport,
      sources: sources,
      usage: usage,
    });
  } catch (error: any) {
    console.error("Investigation error detail:", error);
    
    let rawErrorMessage = error.message || "";
    let friendlyError = "连接 AI 服务器或发起背调调用时遇到未知障碍。请检查网络后重试。";
    
    const errStr = typeof error === 'object' ? JSON.stringify(error) : String(error);
    const combinedErrorText = (friendlyError + " " + rawErrorMessage + " " + errStr).toLowerCase();
    
    if (
      combinedErrorText.includes("429") || 
      combinedErrorText.includes("quota") || 
      combinedErrorText.includes("limit") || 
      combinedErrorText.includes("exhausted") ||
      combinedErrorText.includes("rate")
    ) {
      friendlyError = "【⚠️ DeepSeek API 调用额度已超限/访问过于频繁】\n\n检测到您的 DeepSeek API Key 已达当前套餐调用限制或额度耗尽。\n\n建议您：\n\n1. 稍等 1-2 分钟后重试（DeepSeek Key 通常有分钟级频率限制）；\n2. 在页面右上角「API Key 设置」中更新您的 DEEPSEEK_API_KEY，或切换为充值付费 API key，以享受完全不受限的高并发、千级批量订单实时深度背调！";
    } else if (
      combinedErrorText.includes("api_key_invalid") ||
      combinedErrorText.includes("api key is invalid") ||
      combinedErrorText.includes("invalid api key") ||
      combinedErrorText.includes("key not found") ||
      combinedErrorText.includes("authentication") ||
      combinedErrorText.includes("unauthorized") ||
      combinedErrorText.includes("401") ||
      combinedErrorText.includes("403")
    ) {
      friendlyError = "【❌ DeepSeek API Key 无效或未正确配置】\n\n您的 DEEPSEEK_API_KEY 可能填写有误或已失效。请在页面右上角「API Key 设置」重新检查、修改并保存您的 API 秘钥，确保其处于激活状态。";
    } else if (combinedErrorText.includes("__max_tokens__")) {
      friendlyError = "【⚠️ 报告生成被截断】AI 调查引擎的输出超出长度限制，多次重试后仍未完成完整报告。\n\n建议您：\n1. 稍后重试该订单；\n2. 在页面右上角「API Key 设置」中切换模型后重新核查；\n3. 若持续出现，请检查网络状态。";
    } else if (combinedErrorText.includes("__field_validation__")) {
      friendlyError = "【⚠️ AI 调查结果为空】AI 调查引擎连续多次返回空报告或缺失分析等级。\n\n这通常是模型瞬态输出异常。建议您：\n1. 稍后重试该订单（通常再次核查即可获得完整报告）；\n2. 在页面右上角「API Key 设置」中切换模型后重新核查；\n3. 若持续出现，请检查 API 配额与网络状态。";
    } else if (rawErrorMessage) {
      friendlyError = `【系统异常】AI 调查引擎反馈：${rawErrorMessage}`;
    }

    res.status(500).json({
      success: false,
      error: friendlyError,
    });
  }
});

// Helper function to build standard multipart/form-data payload with raw Buffers,
// bypassing any Node.js version/fetch/FormData/File/Blob serialization quirks.
function buildFeishuMultipartBody(
  boundary: string,
  fileName: string,
  parentType: string,
  parentNode: string,
  size: number,
  fileBuffer: Buffer,
  mimeType: string
): Buffer {
  const parts: Buffer[] = [];

  // append file_name
  parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file_name"\r\n\r\n${fileName}\r\n`));

  // append parent_type
  parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="parent_type"\r\n\r\n${parentType}\r\n`));

  // append parent_node
  parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="parent_node"\r\n\r\n${parentNode}\r\n`));

  // append size
  parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="size"\r\n\r\n${size}\r\n`));

  // append file
  parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${fileName}"\r\nContent-Type: ${mimeType}\r\n\r\n`));
  parts.push(fileBuffer);
  parts.push(Buffer.from(`\r\n--${boundary}--\r\n`));

  return Buffer.concat(parts);
}

// Feishu Cloud Docs Upload Proxy API
app.post("/api/feishu/upload", async (req: express.Request, res: express.Response) => {
  try {
    const {
      appId, appSecret, folderToken, currentSubfolderToken, currentSubfolderIndex,
      fileName, fileContentBase64, markdownContent, companyName, documentTitle
    } = req.body;

    if (!appId || !appSecret || !folderToken || !fileName || !fileContentBase64) {
      res.status(400).json({ success: false, error: "缺少必要的飞书配置参数或文件内容。" });
      return;
    }

    // Smart Folder Token Parser: if user pastes a full URL, automatically extract the folder token
    let parsedFolderToken = folderToken.trim();
    if (parsedFolderToken.includes("/folder/")) {
      const match = parsedFolderToken.match(/\/folder\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        parsedFolderToken = match[1];
      }
    } else if (parsedFolderToken.startsWith("http")) {
      // Try to get last part of the URL before any query parameters
      try {
        const urlObj = new URL(parsedFolderToken);
        const parts = urlObj.pathname.split("/");
        const lastPart = parts[parts.length - 1] || parts[parts.length - 2];
        if (lastPart && lastPart.length > 10) {
          parsedFolderToken = lastPart;
        }
      } catch (e) {
        // Fallback if URL parsing failed
      }
    }

    // 1. 获取 tenant_access_token
    const tokenRes = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        app_id: appId.trim(),
        app_secret: appSecret.trim(),
      }),
    });

    const tokenJson: any = await tokenRes.json();
    if (tokenJson.code !== 0 || !tokenJson.tenant_access_token) {
      let friendlyError = "获取飞书 Access Token 失败。";
      if (tokenJson.msg && tokenJson.msg.toLowerCase().includes("invalid")) {
        friendlyError += " App ID 或 App Secret 似乎有误，请在飞书开放平台核对，并确认应用已创建并「发布上线」。";
      } else {
        friendlyError += ` 飞书返回消息: ${tokenJson.msg || "错误代码 " + tokenJson.code}`;
      }
      res.status(400).json({
        success: false,
        error: friendlyError,
      });
      return;
    }

    const tenantAccessToken = tokenJson.tenant_access_token;
    const rootFolderToken = parsedFolderToken;
    const folderLimitCode = 1062507;
    const isFolderLimitError = (code: unknown, message?: unknown) =>
      Number(code) === folderLimitCode || String(message || "").toLowerCase().includes("parent node out of sibling num");

    const createSubfolder = async (index: number): Promise<FeishuFolderState> => {
      const date = new Date();
      const datePart = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}`;
      const name = `背调报告-${datePart}-${String(index).padStart(3, "0")}`;
      const folderRes = await fetch("https://open.feishu.cn/open-apis/drive/v1/files/create_folder", {
        method: "POST",
        headers: { Authorization: `Bearer ${tenantAccessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ name, folder_token: rootFolderToken }),
      });
      const folderJson: any = await folderRes.json();
      const token = folderJson.data?.token || folderJson.data?.folder_token;
      if (folderJson.code !== 0 || !token) {
        const error: any = new Error(folderJson.msg || "创建飞书报告子文件夹失败");
        error.code = folderJson.code;
        error.errorType = isFolderLimitError(folderJson.code, folderJson.msg)
          ? "folder_node_limit_exceeded"
          : "folder_create_failed";
        throw error;
      }
      const state = { token, index, name };
      console.log(`[Feishu Folder] Created subfolder name=${name}, index=${index}.`);
      return state;
    };

    const ensureSubfolder = async (): Promise<FeishuFolderState> => withFeishuFolderLock(rootFolderToken, async () => {
      const cached = feishuFolderStates.get(rootFolderToken);
      if (cached) return cached;
      const suppliedToken = String(currentSubfolderToken || "").trim();
      const suppliedIndex = Math.max(1, Number(currentSubfolderIndex) || 1);
      if (suppliedToken) {
        const restored = { token: suppliedToken, index: suppliedIndex, name: `背调报告-${String(suppliedIndex).padStart(3, "0")}` };
        feishuFolderStates.set(rootFolderToken, restored);
        console.log(`[Feishu Folder] Restored subfolder index=${suppliedIndex} from client state.`);
        return restored;
      }
      const created = await createSubfolder(1);
      feishuFolderStates.set(rootFolderToken, created);
      return created;
    });

    const rotateSubfolder = async (fullToken: string): Promise<FeishuFolderState> => withFeishuFolderLock(rootFolderToken, async () => {
      const cached = feishuFolderStates.get(rootFolderToken);
      // Another concurrent request may already have replaced the full folder.
      if (cached && cached.token !== fullToken) return cached;
      const nextIndex = Math.max(cached?.index || Number(currentSubfolderIndex) || 1, 1) + 1;
      const created = await createSubfolder(nextIndex);
      feishuFolderStates.set(rootFolderToken, created);
      console.warn(`[Feishu Folder] Folder index=${nextIndex - 1} is full; switched to index=${nextIndex}.`);
      return created;
    });

    let activeFolder = await ensureSubfolder();
    let folderSwitchCount = 0;
    let folderAutoSwitched = false;

    // 2. 将 HTML 转换为真正的 docx 二进制流，并构建 multipart/form-data
    // 飞书云空间 import_tasks 对于原生的 .docx 文件有极佳的转换支持（可直接转为在线编辑的 /docx/ 文档）
    let buffer: Buffer;
    let uploadFileName = fileName;
    let mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    let fileExtensionForImport = "docx";

    // 统一处理文件后缀为 .docx
    if (uploadFileName.endsWith(".doc")) {
      uploadFileName = uploadFileName.slice(0, -4) + ".docx";
    } else if (!uploadFileName.endsWith(".docx")) {
      uploadFileName = uploadFileName + ".docx";
    }

    try {
      const rawHtml = Buffer.from(fileContentBase64, "base64").toString("utf-8").replace(/^\ufeff/, "");
      // 使用 html-to-docx 插件，将 HTML 的排版和内容编译为标准的 Office XML .docx 压缩包
      buffer = await HTMLToDocx(rawHtml, null, {
        table: { row: { cantSplit: true } },
        footer: true,
        header: true,
        pageNumber: true,
      });
      console.log(`Successfully converted HTML report to binary .docx file (${buffer.length} bytes) for ${uploadFileName}`);
    } catch (docxError) {
      console.error("html-to-docx conversion failed, falling back to raw html upload:", docxError);
      // 后备容错逻辑：若 docx 转换库抛错，则回退到上传原始的 HTML/MD
      buffer = Buffer.from(fileContentBase64, "base64");
      mimeType = "text/html";
      fileExtensionForImport = "html";
      if (uploadFileName.endsWith(".docx")) {
        uploadFileName = uploadFileName.slice(0, -5) + ".html";
      } else if (uploadFileName.endsWith(".doc")) {
        uploadFileName = uploadFileName.slice(0, -4) + ".html";
      }
    }

    const uploadSourceFile = async (): Promise<{ fileToken: string; url: string }> => {
      const boundary = "----WebKitFormBoundary" + Math.random().toString(36).substring(2, 15);
      const multipartBody = buildFeishuMultipartBody(
        boundary, uploadFileName, "explorer", activeFolder.token, buffer.length, buffer, mimeType
      );
      const uploadRes = await fetch("https://open.feishu.cn/open-apis/drive/v1/files/upload_all", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tenantAccessToken}`,
          "Content-Type": `multipart/form-data; boundary=${boundary}`,
        },
        body: multipartBody,
      });
      const uploadJson: any = await uploadRes.json();
      if (uploadJson.code !== 0 || !uploadJson.data?.file_token) {
        const error: any = new Error(uploadJson.msg || "飞书文件上传失败");
        error.code = uploadJson.code;
        throw error;
      }
      return { fileToken: uploadJson.data.file_token, url: uploadJson.data.url || "" };
    };

    const sanitizeDocumentTitle = (value: unknown): string => String(value || "")
      .replace(/[\u0000-\u001f\u007f]/g, " ")
      .replace(/[\\/:*?"<>|]/g, " ")
      .replace(/\s+/g, " ")
      .replace(/\s*[-–—]+\s*/g, "-")
      .trim()
      .slice(0, 120)
      .replace(/[\s-]+$/g, "")
      .trim();

    let docTitle = sanitizeDocumentTitle(documentTitle);
    if (!docTitle && companyName) {
      docTitle = `${companyName} 公司客户背调 境外企业客户背景调查报告 (Due Diligence Report)`;
    } else if (!docTitle) {
      docTitle = uploadFileName.replace(/\.(docx|doc|html)$/i, "");
    }

    const maxImportAttempts = 3;
    const nonRetryableImportCodes = new Set([1069906, 1069907, 1069908, 1069909, 1069910, 1069911, 1069912, 1069913, 1069914]);
    const nonRetryableUploadCodes = new Set([1061003, 1061005, 1061011, 1061033, 1061041, 1061042, 1061043, 1061044, 99991663]);
    let finalDocUrl = "";
    let latestFileUrl = "";
    let latestFileToken = "";
    let importAttempts = 0;
    let importFallback = true;
    let importError = "conversion_failed";
    let importJobStatus: number | null = null;
    let importFileType = "file";
    let stopWithoutRetry = false;
    let lastAttemptError = "";
    let terminalJobError = "";

    const classifyResult = (resultObj: any): "docx" | "fallback" | "failed" | "processing" => {
      const resultUrl = typeof resultObj?.url === "string" ? resultObj.url : "";
      const resultType = String(resultObj?.file_type || resultObj?.type || "").toLowerCase();
      importJobStatus = resultObj?.job_status ?? null;
      importFileType = resultUrl.includes("/docx/") ? "docx" : resultUrl.includes("/file/") ? "file" : resultType || "unknown";
      // URL is the authoritative artifact check and also keeps compatibility
      // with older responses that reported job_status=0 for completed imports.
      if (resultUrl.includes("/docx/")) {
        finalDocUrl = resultUrl;
        return "docx";
      }
      if (resultObj?.job_status === 3) return "failed";
      if (resultObj?.job_status === 2) {
        if (resultUrl.includes("/file/") || resultType === "file" || !resultUrl) return "fallback";
        if (resultType === "docx" && resultUrl) {
          finalDocUrl = resultUrl;
          return "docx";
        }
        return "fallback";
      }
      return "processing";
    };

    for (let importAttempt = 1; importAttempt <= maxImportAttempts; importAttempt++) {
      importAttempts = importAttempt;
      if (importAttempt === 2) await new Promise(resolve => setTimeout(resolve, 2000));
      if (importAttempt === 3) await new Promise(resolve => setTimeout(resolve, 5000));

      try {
        const uploaded = await uploadSourceFile();
        latestFileToken = uploaded.fileToken;
        latestFileUrl = uploaded.url;
        console.log(`[Feishu Import] ${uploadFileName} attempt ${importAttempt}/${maxImportAttempts}: source uploaded.`);

        const importRes = await fetch("https://open.feishu.cn/open-apis/drive/v1/import_tasks", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${tenantAccessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            file_extension: fileExtensionForImport,
            file_token: latestFileToken,
            type: "docx",
            file_name: docTitle,
            point: { mount_type: 1, mount_key: activeFolder.token },
          }),
        });
        const importJson: any = await importRes.json();
        if (importJson.code !== 0 || !importJson.data?.ticket) {
          importError = `import_create_${importJson.code || "unknown"}`;
          console.warn(`[Feishu Import] ${uploadFileName} attempt ${importAttempt}: create failed code=${importJson.code}, msg=${importJson.msg || ""}`);
          if (isFolderLimitError(importJson.code, importJson.msg)) {
            if (folderSwitchCount >= 3) {
              importError = "folder_node_limit_exceeded";
              stopWithoutRetry = true;
              break;
            }
            const fullToken = activeFolder.token;
            activeFolder = await rotateSubfolder(fullToken);
            folderSwitchCount++;
            folderAutoSwitched = true;
            importAttempt--;
            continue;
          }
          if (nonRetryableImportCodes.has(importJson.code)) stopWithoutRetry = true;
          if (stopWithoutRetry) break;
          continue;
        }

        const ticket = importJson.data.ticket;
        let outcome: "docx" | "fallback" | "failed" | "processing" = "processing";
        for (let pollAttempt = 1; pollAttempt <= 30; pollAttempt++) {
          await new Promise(resolve => setTimeout(resolve, 2000));
          try {
            const statusRes = await fetch(`https://open.feishu.cn/open-apis/drive/v1/import_tasks/${ticket}`, {
              headers: { Authorization: `Bearer ${tenantAccessToken}` },
            });
            const statusJson: any = await statusRes.json();
            if (statusJson.code !== 0 || !statusJson.data?.result) {
              console.warn(`[Feishu Import] ${uploadFileName} ticket=${ticket}: poll ${pollAttempt} failed code=${statusJson.code}, msg=${statusJson.msg || ""}`);
              continue;
            }
            outcome = classifyResult(statusJson.data.result);
            terminalJobError = String(statusJson.data.result.job_error_msg || "");
            if (outcome !== "processing") {
              console.log(`[Feishu Import] ${uploadFileName} ticket=${ticket}: status=${importJobStatus}, type=${importFileType}, outcome=${outcome}, error=${statusJson.data.result.job_error_msg || ""}`);
              break;
            }
          } catch (pollError: any) {
            console.warn(`[Feishu Import] ${uploadFileName} ticket=${ticket}: transient poll error: ${pollError?.message || pollError}`);
          }
        }

        if (outcome === "processing") {
          await new Promise(resolve => setTimeout(resolve, 5000));
          try {
            const finalStatusRes = await fetch(`https://open.feishu.cn/open-apis/drive/v1/import_tasks/${ticket}`, {
              headers: { Authorization: `Bearer ${tenantAccessToken}` },
            });
            const finalStatusJson: any = await finalStatusRes.json();
            if (finalStatusJson.code === 0 && finalStatusJson.data?.result) {
              outcome = classifyResult(finalStatusJson.data.result);
              terminalJobError = String(finalStatusJson.data.result.job_error_msg || "");
            }
          } catch (finalPollError: any) {
            console.warn(`[Feishu Import] ${uploadFileName} ticket=${ticket}: final confirmation failed: ${finalPollError?.message || finalPollError}`);
          }
        }

        if (outcome === "docx") {
          importFallback = false;
          importError = "";
          console.log(`飞书文档导入任务转换成功: ${finalDocUrl}`);
          break;
        }
        if (isFolderLimitError(undefined, terminalJobError)) {
          if (folderSwitchCount >= 3) {
            importError = "folder_node_limit_exceeded";
            lastAttemptError = terminalJobError;
            break;
          }
          const fullToken = activeFolder.token;
          activeFolder = await rotateSubfolder(fullToken);
          folderSwitchCount++;
          folderAutoSwitched = true;
          importAttempt--;
          terminalJobError = "";
          continue;
        }
        if (outcome === "processing") {
          importError = "conversion_timeout";
          stopWithoutRetry = true;
          console.warn(`[Feishu Import] ${uploadFileName} ticket=${ticket}: still processing after final confirmation; no new task will be created.`);
          break;
        }
        importError = outcome === "fallback" ? "conversion_fallback" : "conversion_failed";
      } catch (attemptError: any) {
        importError = `upload_or_import_${attemptError?.code || "error"}`;
        lastAttemptError = attemptError?.message || String(attemptError);
        console.error(`[Feishu Import] ${uploadFileName} attempt ${importAttempt}: ${attemptError?.message || attemptError}`);
        if (attemptError?.errorType === "folder_create_failed") {
          importError = "folder_create_failed";
          break;
        }
        if (isFolderLimitError(attemptError?.code, attemptError?.message)) {
          if (folderSwitchCount >= 3) {
            importError = "folder_node_limit_exceeded";
            lastAttemptError = "连续切换 3 个子文件夹后仍然达到容量上限";
            break;
          }
          const fullToken = activeFolder.token;
          activeFolder = await rotateSubfolder(fullToken);
          folderSwitchCount++;
          folderAutoSwitched = true;
          importAttempt--;
          continue;
        }
        if (nonRetryableUploadCodes.has(attemptError?.code)) break;
      }
    }

    if (["folder_node_limit_exceeded", "folder_create_failed"].includes(importError) || (!finalDocUrl && !latestFileUrl)) {
      res.status(400).json({
        success: false,
        error: importError === "folder_node_limit_exceeded"
          ? `飞书目录已满，自动切换子目录后仍失败: ${lastAttemptError || "请检查根目录权限"}`
          : importError === "folder_create_failed"
            ? `创建飞书报告子目录失败: ${lastAttemptError || "请检查根目录权限"}`
            : `上传至飞书失败: ${lastAttemptError || "所有上传尝试均失败"}`,
        errorType: importError,
        subfolderToken: activeFolder.token,
        subfolderIndex: activeFolder.index,
        subfolderName: activeFolder.name,
      });
      return;
    }
    if (importFallback) finalDocUrl = latestFileUrl;
    res.json({
      success: true,
      fileToken: latestFileToken,
      url: finalDocUrl,
      urlType: importFallback ? "file" : "docx",
      importAttempts,
      importFallback,
      importError,
      importJobStatus,
      importFileType,
      subfolderToken: activeFolder.token,
      subfolderIndex: activeFolder.index,
      subfolderName: activeFolder.name,
      folderAutoSwitched,
    });
  } catch (error: any) {
    console.error("Feishu upload API error:", error);
    res.status(error?.errorType ? 400 : 500).json({
      success: false,
      error: `服务器端处理飞书上传出错: ${error.message || error}`,
      errorType: error?.errorType || "document_upload_failed",
    });
  }
});

// Feishu Cloud Sheets Upload & Convert Proxy API
app.post("/api/feishu/upload-sheet", async (req: express.Request, res: express.Response) => {
  try {
    const { appId, appSecret, folderToken, fileName, fileContentBase64 } = req.body;

    if (!appId || !appSecret || !folderToken || !fileName || !fileContentBase64) {
      res.status(400).json({ success: false, error: "缺少必要的飞书配置参数或文件内容。" });
      return;
    }

    // Smart Folder Token Parser: if user pastes a full URL, automatically extract the folder token
    let parsedFolderToken = folderToken.trim();
    if (parsedFolderToken.includes("/folder/")) {
      const match = parsedFolderToken.match(/\/folder\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        parsedFolderToken = match[1];
      }
    } else if (parsedFolderToken.startsWith("http")) {
      try {
        const urlObj = new URL(parsedFolderToken);
        const parts = urlObj.pathname.split("/");
        const lastPart = parts[parts.length - 1] || parts[parts.length - 2];
        if (lastPart && lastPart.length > 10) {
          parsedFolderToken = lastPart;
        }
      } catch (e) {
        // Fallback
      }
    }

    // 1. 获取 tenant_access_token
    const tokenRes = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        app_id: appId.trim(),
        app_secret: appSecret.trim(),
      }),
    });

    const tokenJson: any = await tokenRes.json();
    if (tokenJson.code !== 0 || !tokenJson.tenant_access_token) {
      let friendlyError = "获取飞书 Access Token 失败。";
      if (tokenJson.msg && tokenJson.msg.toLowerCase().includes("invalid")) {
        friendlyError += " App ID 或 App Secret 似乎有误，请在飞书开放平台核对，并确认应用已创建并「发布上线」。";
      } else {
        friendlyError += ` 飞书返回消息: ${tokenJson.msg || "错误代码 " + tokenJson.code}`;
      }
      res.status(400).json({
        success: false,
        error: friendlyError,
      });
      return;
    }

    const tenantAccessToken = tokenJson.tenant_access_token;

    // 2. 转换 Base64 文件内容并构建 multipart/form-data
    const buffer = Buffer.from(fileContentBase64, "base64");
    const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    const boundary = "----WebKitFormBoundary" + Math.random().toString(36).substring(2, 15);
    const multipartBody = buildFeishuMultipartBody(
      boundary,
      fileName,
      "explorer",
      parsedFolderToken,
      buffer.length,
      buffer,
      mimeType
    );

    // 3. 上传原始电子表格文件
    const uploadRes = await fetch("https://open.feishu.cn/open-apis/drive/v1/files/upload_all", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${tenantAccessToken}`,
        "Content-Type": `multipart/form-data; boundary=${boundary}`,
      },
      body: multipartBody,
    });

    const uploadJson: any = await uploadRes.json();
    if (uploadJson.code !== 0 || !uploadJson.data) {
      const code = uploadJson.code;
      const msg = uploadJson.msg || "";
      let friendlyError = `上传表格至飞书失败 (Code: ${code}): `;

      if (code === 1061033 || code === 1061004 || msg.toLowerCase().includes("permission") || msg.toLowerCase().includes("forbidden") || code === 99991663) {
        friendlyError += "您的自建应用【无权访问该目标文件夹】。请检查以下两点：\n" +
          "1. 文件夹未授权给自建应用：请在浏览器中打开该飞书文件夹，点击右上角「分享」并搜索您的自建应用名称（即您在飞书后台创建的应用名字），将其添加为协作者，并勾选「可编辑」或「可管理」权限。\n" +
          "2. 权限未开通或未发布：请在飞书开放平台 (open.feishu.cn) 应用后台 -> 「开发配置」 -> 「权限管理」中搜索并勾选「云空间」相关的读写文件权限（如 `drive:file`），并必须在「版本管理与发布」中创建一个新版本并发布上线。";
      } else if (code === 1061005) {
        friendlyError += "上传参数有误，请确保目标文件夹 Token 正确且未包含非法字符。";
      } else if (code === 1061011 || msg.toLowerCase().includes("not found")) {
        friendlyError += "找不到指定的飞书文件夹，请核对您的 Folder Token。";
      } else {
        friendlyError += `${msg || "未指明的飞书服务器错误"}`;
      }

      res.status(400).json({
        success: false,
        error: friendlyError,
      });
      return;
    }

    const fileToken = uploadJson.data.file_token;

    // 4. 创建导入转换在线表格任务
    const importRes = await fetch("https://open.feishu.cn/open-apis/drive/v1/import_tasks", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${tenantAccessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        file_token: fileToken,
        file_extension: "xlsx",
        type: "sheet",
        file_name: fileName.endsWith(".xlsx") ? fileName.slice(0, -5) : fileName,
        point: {
          mount_type: 1, // 挂载到文件夹
          mount_key: parsedFolderToken,
        },
      }),
    });

    const importJson: any = await importRes.json();
    if (importJson.code !== 0 || !importJson.data?.ticket) {
      res.status(400).json({
        success: false,
        error: `创建飞书在线表格转换任务失败 (Code: ${importJson.code}): ${importJson.msg || "未知错误"}`,
      });
      return;
    }

    const ticket = importJson.data.ticket;

    // 5. 轮询导入转换任务状态直至完成 (最多轮询 30 次，每次等 2 秒，最高 60 秒)
    let onlineSheetUrl = "";
    let attempts = 0;
    const maxAttempts = 30;

    while (attempts < maxAttempts) {
      await new Promise((resolve) => setTimeout(resolve, 2000));
      attempts++;

      const statusRes = await fetch(`https://open.feishu.cn/open-apis/drive/v1/import_tasks/${ticket}`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${tenantAccessToken}`,
        },
      });

      const statusJson: any = await statusRes.json();
      if (statusJson.code === 0 && statusJson.data && statusJson.data.result) {
        const resultObj = statusJson.data.result;
        const jobStatus = resultObj.job_status !== undefined ? resultObj.job_status : (resultObj.status !== undefined ? resultObj.status : statusJson.data.job_status);
        if (jobStatus === 0) {
          // 转换成功
          onlineSheetUrl = resultObj.url;
          break;
        } else if (jobStatus === 2) {
          // 转换失败
          res.status(400).json({
            success: false,
            error: `飞书在线表格转换任务执行失败，错误消息: ${resultObj.job_error_msg || "未知原因"}`,
          });
          return;
        }
      } else if (statusJson.code !== 0) {
        console.error("查询飞书表格导入状态任务失败:", statusJson.msg);
        break;
      }
    }

    if (!onlineSheetUrl) {
      res.status(408).json({
        success: false,
        error: "飞书在线表格转换任务轮询超时，请稍后在对应的飞书云空间文件夹中查看已生成的表格。",
      });
      return;
    }

    res.json({
      success: true,
      url: onlineSheetUrl,
    });
  } catch (error: any) {
    console.error("Feishu sheet upload & convert API error:", error);
    res.status(500).json({
      success: false,
      error: `服务器端处理飞书表格转换出错: ${error.message || error}`,
    });
  }
});

// Feishu Bitable (Multi-dimensional Table) Sync API
app.post("/api/feishu/sync-bitable", async (req: express.Request, res: express.Response) => {
  try {
    const { appId, appSecret, bitableUrl, records, mode } = req.body;

    const normalizeName = (str: string): string => {
      if (typeof str !== "string") return "";
      return str
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, "") // Remove zero-width spaces, BOM, and non-breaking space
        .replace(/\s+/g, "")                         // Strip all spaces
        .replace(/[\(\)\{\}\[\]（）｛｝【】\/\\-_:：,，.。.?？!！]/g, "") // Strip all punctuation, parentheses, and brackets completely
        .toLowerCase()
        .trim();
    };

    if (!appId || !appSecret || !bitableUrl || !records) {
      res.status(400).json({ success: false, error: "缺少必要的飞书配置参数或同步数据。" });
      return;
    }

    // 1. Parse bitableUrl to extract token and tableId
    let token = "";
    let tableId = "";
    let viewId = "";
    
    const tableMatch = bitableUrl.match(/[?&]table=([a-zA-Z0-9_-]+)/);
    if (tableMatch) {
      tableId = tableMatch[1];
    }

    // Feishu links use `view=vew...`, while the records API expects `view_id`.
    // Accept both forms so copied browser links and explicit API-style links work.
    const viewMatch = bitableUrl.match(/[?&](?:view|view_id)=([a-zA-Z0-9_-]+)/);
    if (viewMatch) {
      viewId = viewMatch[1];
    }
    
    if (bitableUrl.includes("/wiki/")) {
      const wikiMatch = bitableUrl.match(/\/wiki\/([a-zA-Z0-9_-]+)/);
      if (wikiMatch) {
        token = wikiMatch[1];
      }
    } else if (bitableUrl.includes("/base/")) {
      const baseMatch = bitableUrl.match(/\/base\/([a-zA-Z0-9_-]+)/);
      if (baseMatch) {
        token = baseMatch[1];
      }
    } else if (bitableUrl.includes("/bitable/")) {
      const bitableMatch = bitableUrl.match(/\/bitable\/([a-zA-Z0-9_-]+)/);
      if (bitableMatch) {
        token = bitableMatch[1];
      }
    } else {
      const pathParts = bitableUrl.split('?')[0].split('/');
      token = pathParts[pathParts.length - 1] || pathParts[pathParts.length - 2] || "";
    }

    if (!token) {
      res.status(400).json({
        success: false,
        error: "多维表格/知识空间 Token 解析失败，请检查您填写的 URL 是否正确且包含有效的路径段。"
      });
      return;
    }

    // 2. 获取 tenant_access_token
    const tokenRes = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        app_id: appId.trim(),
        app_secret: appSecret.trim(),
      }),
    });

    const tokenJson: any = await tokenRes.json();
    if (tokenJson.code !== 0 || !tokenJson.tenant_access_token) {
      let friendlyError = "获取飞书 Access Token 失败。";
      if (tokenJson.msg && tokenJson.msg.toLowerCase().includes("invalid")) {
        friendlyError += " App ID 或 App Secret 似乎有误，请在飞书开放平台核对，并确认应用已创建并「发布上线」。";
      } else {
        friendlyError += ` 飞书返回消息: ${tokenJson.msg || "错误代码 " + tokenJson.code}`;
      }
      res.status(400).json({
        success: false,
        error: friendlyError,
      });
      return;
    }

    const tenantAccessToken = tokenJson.tenant_access_token;

    // 3. App Token Strategy:
    // By default, we use the token extracted from the URL as-is (this is critical for Wiki-based multi-dimensional tables,
    // since the collaborator permission is added to the Wiki page itself, which corresponds to the Wiki Token.
    // If we resolve it to the underlying obj_token immediately, we will get a 91403 (No Permission) error).
    //
    // However, if using the extracted token fails with a permission error (like 91403, 1254302, 1061033, 1254003),
    // and the token starts with "wik", we will automatically attempt to resolve it to the underlying "obj_token" (bas...)
    // and try the operation again as a fallback. This covers both Wiki-level and Doc-level collaborator permission setups!
    let appToken = token;
    let hasAttemptedWikiResolve = false;
    let resolvedObjToken: string | null = null;

    // Helper to resolve Wiki Token to Bitable App Token (obj_token) if needed
    const resolveWikiTokenToObjToken = async (): Promise<string | null> => {
      if (!token.startsWith("wik")) return null;
      if (resolvedObjToken) return resolvedObjToken;
      try {
        console.log(`[Bitable Sync] Attempting to resolve Wiki Token ${token} to Bitable obj_token using get_node API as fallback...`);
        const wikiRes = await fetch(`https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token=${token}`, {
          method: "GET",
          headers: {
            Authorization: `Bearer ${tenantAccessToken}`,
          },
        });
        const wikiJson: any = await wikiRes.json();
        if (wikiJson.code === 0 && wikiJson.data?.node?.obj_token) {
          resolvedObjToken = wikiJson.data.node.obj_token;
          console.log(`[Bitable Sync] Resolved Wiki Token ${token} successfully. obj_token: ${resolvedObjToken}`);
          return resolvedObjToken;
        } else {
          console.warn(`[Bitable Sync] get_node failed (code: ${wikiJson.code}, msg: ${wikiJson.msg})`);
        }
      } catch (err) {
        console.error("[Bitable Sync] Error during get_node resolution:", err);
      }
      return null;
    };

    // Helper to check if a Feishu error is permission-related
    const isPermissionCode = (code: number) => {
      return [91403, 1254302, 1061033, 1061004, 99991663].includes(code);
    };

    // 4. If tableId is not extracted from URL, try to fetch the first table from the Bitable automatically
    if (!tableId) {
      try {
        console.log(`[Bitable Sync] No tableId found in URL. Attempting to fetch list of tables using token: ${appToken}...`);
        let tablesRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables`, {
          method: "GET",
          headers: {
            Authorization: `Bearer ${tenantAccessToken}`,
          },
        });
        let tablesJson: any = await tablesRes.json();

        if (isPermissionCode(tablesJson.code) && token.startsWith("wik") && !hasAttemptedWikiResolve) {
          hasAttemptedWikiResolve = true;
          const resolved = await resolveWikiTokenToObjToken();
          if (resolved) {
            console.log(`[Bitable Sync] Retrying table list fetch with resolved obj_token: ${resolved}`);
            appToken = resolved;
            tablesRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables`, {
              method: "GET",
              headers: {
                Authorization: `Bearer ${tenantAccessToken}`,
              },
            });
            tablesJson = await tablesRes.json();
          }
        }

        if (tablesJson.code === 0 && tablesJson.data?.items && tablesJson.data.items.length > 0) {
          tableId = tablesJson.data.items[0].table_id;
          console.log(`[Bitable Sync] Automatically detected first tableId: ${tableId}`);
        } else {
          console.warn(`[Bitable Sync] Failed to fetch table list (code: ${tablesJson.code}, msg: ${tablesJson.msg})`);
        }
      } catch (tablesErr) {
        console.error("[Bitable Sync] Error fetching Bitable tables:", tablesErr);
      }
    }

    if (!tableId) {
      res.status(400).json({
        success: false,
        error: "未在链接中检测到 table= 参数，且无法自动获取多维表格的子表。请确保在浏览器中打开具体的子表，并复制完整的地址栏链接（例如：.../wiki/...?table=tble...）"
      });
      return;
    }

    // 3.5 Fetch target table's fields to perform intelligent auto-mapping of fields
    let bitableFieldNames: string[] = [];
    const bitableFieldTypes: Record<string, number> = {};
    try {
      console.log(`[Bitable Sync] Fetching fields for table ${tableId} using token: ${appToken}...`);
      let fieldsRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/fields?page_size=100`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${tenantAccessToken}`,
        },
      });
      let fieldsJson: any = await fieldsRes.json();

      if (isPermissionCode(fieldsJson.code) && token.startsWith("wik") && !hasAttemptedWikiResolve) {
        hasAttemptedWikiResolve = true;
        const resolved = await resolveWikiTokenToObjToken();
        if (resolved) {
          console.log(`[Bitable Sync] Retrying fields fetch with resolved obj_token: ${resolved}`);
          appToken = resolved;
          fieldsRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/fields?page_size=100`, {
            method: "GET",
            headers: {
              Authorization: `Bearer ${tenantAccessToken}`,
            },
          });
          fieldsJson = await fieldsRes.json();
        }
      }

      if (fieldsJson.code === 0 && fieldsJson.data?.items) {
        bitableFieldNames = fieldsJson.data.items.map((item: any) => {
          const name = item.field_name;
          bitableFieldTypes[normalizeName(name)] = item.type;
          return name;
        });
        console.log(`[Bitable Sync] Successfully fetched Bitable field schema. Total fields: ${bitableFieldNames.length}.`);
        console.log(`[Bitable Sync] Bitable field types metadata (For reference, 1=Text, 2=Number, 5=DateTime, 7=Checkbox, 13=Phone, 15=Url, etc.):`, JSON.stringify(bitableFieldTypes));
      } else {
        console.warn(`[Bitable Sync] Failed to fetch Bitable fields (code: ${fieldsJson.code}, msg: ${fieldsJson.msg}). Fallback to raw field mapping.`);
      }
    } catch (fieldsErr) {
      console.error("[Bitable Sync] Error fetching Bitable fields, proceeding with raw fields:", fieldsErr);
    }

    const keyAliases: Record<string, string[]> = {
      companyName: ["公司名称", "公司", "客商名称", "客户名称", "企业名称", "Company Name", "Company", "Name", "客户", "客商", "买家名称"],
      country: ["国家", "国家/地区", "地区", "国家地区", "Country", "Country/Region", "Region", "国籍", "收货国家"],
      productContext: ["意向产品", "采购产品", "产品", "商品", "经营范围", "Product", "Product Context", "Goods", "产品名称", "商品信息"],
      grade: ["背景评判等级", "背景分析等级", "背调等级", "背调级别", "评估等级", "信用评级", "客户等级", "评级", "Grade", "Rating", "等级"],
      summary: ["报告概要", "调查概要", "背调摘要", "摘要", "结论", "Summary", "Overview", "报告汇总"],
      feishuUrl: ["背景报告文档", "背景调查报告文档", "背调报告文档", "飞书文档链接", "飞书链接", "报告链接", "Word链接", "云文档", "飞书文档", "标题", "Report URL", "Document Link", "Link", "背调报告", "云链接", "背景报告", "背调结果", "报告"],
      orderInfo: ["详细订单", "订单详情", "订单信息", "详细信息", "收货详情", "Order Info", "Order", "订单详情/税号信息", "订单备注"],
      completedAt: ["背调时间", "背调完成时间", "调查时间", "调查完成时间", "Completed At", "Investigation Time", "Investigation Completed At"],
      riskAlert: ["风险提示", "风险", "信用风险", "风险预警", "Risk Alert", "Risk"],
      followUpStrategy: ["跟进策略", "开发策略", "销售策略", "开发跟进策略", "Follow-up Strategy", "Strategy", "跟进方案"],
      followUpGrade: ["跟进等级", "跟进级别", "Follow-up Level", "Follow-up Grade", "FollowUpGrade", "FollowUpLevel", "跟进"]
    };

    let mappingAnalysisList: any[] = [];
    let processedRecords = records;

    let coerceValueByBitableType = (fieldName: string, value: any): any => value;

    if (bitableFieldNames.length > 0) {
      const bitableCleanMap: Record<string, string> = {};
      bitableFieldNames.forEach(name => {
        bitableCleanMap[normalizeName(name)] = name;
      });

      coerceValueByBitableType = (fieldName: string, value: any): any => {
        if (value === null || value === undefined) return undefined;
        const normName = normalizeName(fieldName);
        const fieldType = bitableFieldTypes[normName];

        // 1. Skip fields that have no known type, are read-only, or require complex lookup/upload
        // Bitable Types:
        // 1: Text, 2: Number, 3: SingleSelect, 4: MultiSelect, 5: DateTime, 7: Checkbox, 11: User, 13: Phone, 15: Url, 17: Attachment, etc.
        // Skip explicitly read-only types (e.g. Formula 20, Lookup 22, System fields 1001-1004, AutoNumber 99001)
        const readOnlyTypes = [20, 22, 23, 24, 25, 26, 1001, 1002, 1003, 1004, 1005, 99001];
        if (fieldType !== undefined && readOnlyTypes.includes(fieldType)) {
          console.log(`[Bitable Sync] Omitting field "${fieldName}" because its type (${fieldType}) is read-only.`);
          return undefined;
        }

        // Handle Attachment (type 17) gracefully
        if (fieldType === 17) {
          if (typeof value === "string" || !Array.isArray(value)) {
            console.log(`[Bitable Sync] Omitting plain string value for Attachment field "${fieldName}" (type 17) to prevent Feishu Bitable API validation errors. Please use a Link/Url (type 15) or Text (type 1) field instead.`);
            return undefined;
          }
        }

        // 1.5 Coerce User/Person (type 11)
        if (fieldType === 11) {
          if (Array.isArray(value)) {
            return value.map(item => {
              if (typeof item === "object" && item !== null) return item;
              const s = String(item).trim();
              if (s.includes("@")) return { email: s };
              return { name: s };
            });
          }
          if (typeof value === "object" && value !== null) {
            return [value];
          }
          const s = String(value).trim();
          if (s === "") return undefined;
          if (s.includes("@")) return [{ email: s }];
          return [{ name: s }];
        }

        // 2. Coerce Text (type 1)
        if (fieldType === 1) {
          if (typeof value === "object") {
            try {
              return JSON.stringify(value);
            } catch (e) {
              return String(value);
            }
          }
          return String(value);
        }

        // 3. Coerce Number (type 2)
        if (fieldType === 2) {
          if (typeof value === "number") return isNaN(value) ? undefined : value;
          if (typeof value === "string") {
            const cleanStr = value.trim().replace(/[^\d.-]/g, ""); // remove characters like currency, spaces, commas
            if (cleanStr === "") return undefined;
            const num = Number(cleanStr);
            return isNaN(num) ? undefined : num;
          }
          if (typeof value === "boolean") return value ? 1 : 0;
          return undefined;
        }

        // 4. Coerce SingleSelect (type 3)
        if (fieldType === 3) {
          if (typeof value === "object") {
            try {
              return JSON.stringify(value).trim();
            } catch (e) {
              return String(value).trim();
            }
          }
          const strVal = String(value).trim();
          return strVal === "" ? undefined : strVal;
        }

        // 5. Coerce MultiSelect (type 4)
        if (fieldType === 4) {
          if (Array.isArray(value)) {
            return value.map(val => String(val).trim()).filter(Boolean);
          }
          if (typeof value === "string") {
            // Split by commas, semicolons, slash, or newlines
            return value.split(/[,,;；，\n/]/).map(s => s.trim()).filter(Boolean);
          }
          const strVal = String(value).trim();
          return strVal === "" ? undefined : [strVal];
        }

        // 6. Coerce DateTime (type 5)
        if (fieldType === 5) {
          if (value instanceof Date) {
            return value.getTime();
          }
          if (typeof value === "number") {
            // If it's a 10-digit timestamp (seconds), convert to 13-digit (milliseconds)
            if (value > 0 && value < 9999999999) return value * 1000;
            return value;
          }
          if (typeof value === "string") {
            const cleanStr = value.trim();
            if (cleanStr === "") return undefined;
            const timestamp = Date.parse(cleanStr);
            if (!isNaN(timestamp)) {
              return timestamp;
            }
            const num = Number(cleanStr);
            if (!isNaN(num) && num > 0) {
              if (num < 9999999999) return num * 1000;
              return num;
            }
          }
          return undefined;
        }

        // 7. Coerce Checkbox (type 7)
        if (fieldType === 7) {
          if (typeof value === "boolean") return value;
          if (typeof value === "number") return value !== 0;
          if (typeof value === "string") {
            const lowerStr = value.trim().toLowerCase();
            return ["true", "yes", "1", "是", "对"].includes(lowerStr);
          }
          return !!value;
        }

        // 8. Coerce Phone (type 13)
        if (fieldType === 13) {
          return String(value).trim();
        }

        // 9. Coerce URL/Link (type 15)
        if (fieldType === 15) {
          if (typeof value === "string") {
            let cleanStr = value.trim();
            if (cleanStr === "") return undefined;
            if (!cleanStr.startsWith("http://") && !cleanStr.startsWith("https://")) {
              cleanStr = "https://" + cleanStr;
            }
            try {
              new URL(cleanStr); // Validate URL format
              return {
                text: value.trim(),
                link: cleanStr
              };
            } catch (e) {
              console.warn(`[Bitable Sync] Omitted invalid URL for field "${fieldName}": ${value}`);
              return undefined;
            }
          }
          return value;
        }

        // Default fallback: check for empty string
        if (value === "") return undefined;
        return value;
      };

      console.log(`[Bitable Sync] Bitable normalized field map (for matching):`, JSON.stringify(bitableCleanMap, null, 2));

      mappingAnalysisList = [];

      const getFieldTypeName = (typeCode?: number): string => {
        if (typeCode === undefined) return "未知";
        const names: Record<number, string> = {
          1: "多行文本",
          2: "数字/货币",
          3: "单选",
          4: "多选",
          5: "日期时间",
          7: "复选框",
          11: "人员",
          13: "电话",
          15: "超链接/URL",
          17: "附件",
          18: "单向关联",
          19: "查找引用",
          20: "公式",
          21: "双向关联",
          1001: "创建时间",
          1002: "修改时间",
          1003: "创建人",
          1004: "修改人"
        };
        return names[typeCode] || `类型 ${typeCode}`;
      };

      processedRecords = records.map((record: any, recordIdx: number) => {
        const rawFields = record.fields || {};
        const mappedFields: Record<string, any> = {};
        const mappingLog: string[] = [];

        // Helper to safely assign coerced value
        const assignCoerced = (bitableField: string, originalVal: any) => {
          const coerced = coerceValueByBitableType(bitableField, originalVal);
          if (coerced !== undefined) {
            mappedFields[bitableField] = coerced;
            return true;
          }
          return false;
        };

        for (const [key, val] of Object.entries(rawFields)) {
          const normKey = normalizeName(key);

          // 1. Check exact / normalized match in Bitable
          if (bitableCleanMap[normKey]) {
            const bField = bitableCleanMap[normKey];
            const success = assignCoerced(bField, val);
            if (success) {
              mappingLog.push(`Exact match: Excel column "${key}" -> Bitable field "${bField}" (value: ${JSON.stringify(mappedFields[bField])})`);
            } else {
              mappingLog.push(`Exact match failed coercion: Excel column "${key}" -> Bitable field "${bField}" (value was empty or coerced to undefined)`);
            }
            if (recordIdx === 0) {
              mappingAnalysisList.push({
                excelField: key,
                bitableField: bField,
                matchType: "exact",
                status: success ? "success" : "coercion_failed",
                fieldTypeName: getFieldTypeName(bitableFieldTypes[normKey])
              });
            }
            continue;
          }

          // 2. Check if the key is a standard key or one of its aliases
          let foundMatch = false;
          for (const [stdKey, aliases] of Object.entries(keyAliases)) {
            const normStdKey = normalizeName(stdKey);
            const normAliases = aliases.map(normalizeName);

            if (normKey === normStdKey || normAliases.includes(normKey)) {
              // This field represents stdKey. Find if any alias of stdKey actually exists in Bitable.
              const possibleBitableNames = [stdKey, ...aliases];
              for (const candidate of possibleBitableNames) {
                const normCand = normalizeName(candidate);
                if (bitableCleanMap[normCand]) {
                  const bField = bitableCleanMap[normCand];
                  const success = assignCoerced(bField, val);
                  if (success) {
                    mappingLog.push(`Alias match (stdKey: ${stdKey}): Excel column "${key}" -> Bitable field "${bField}" (value: ${JSON.stringify(mappedFields[bField])})`);
                    foundMatch = true;
                  } else {
                    mappingLog.push(`Alias match failed coercion (stdKey: ${stdKey}): Excel column "${key}" -> Bitable field "${bField}" (value was empty/coerced to undefined)`);
                    foundMatch = true;
                  }
                  if (recordIdx === 0) {
                    mappingAnalysisList.push({
                      excelField: key,
                      bitableField: bField,
                      matchType: "alias",
                      status: success ? "success" : "coercion_failed",
                      fieldTypeName: getFieldTypeName(bitableFieldTypes[normCand])
                    });
                  }
                  break;
                }
              }
            }
            if (foundMatch) break;
          }

          if (foundMatch) continue;

          // 3. Substring match fallback (normalized)
          for (const bitableFieldName of bitableFieldNames) {
            const normBitableName = normalizeName(bitableFieldName);
            if (normKey.includes(normBitableName) || normBitableName.includes(normKey)) {
              const bField = bitableFieldName;
              const success = assignCoerced(bField, val);
              if (success) {
                mappingLog.push(`Substring match: Excel column "${key}" -> Bitable field "${bField}" (value: ${JSON.stringify(mappedFields[bField])})`);
                foundMatch = true;
              } else {
                mappingLog.push(`Substring match failed coercion: Excel column "${key}" -> Bitable field "${bField}" (value was empty/coerced to undefined)`);
                foundMatch = true;
              }
              if (recordIdx === 0) {
                mappingAnalysisList.push({
                  excelField: key,
                  bitableField: bField,
                  matchType: "substring",
                  status: success ? "success" : "coercion_failed",
                  fieldTypeName: getFieldTypeName(bitableFieldTypes[normBitableName])
                });
              }
              break;
            }
          }

          if (!foundMatch) {
            mappingLog.push(`No match: Excel column "${key}" did not match any Bitable field (value: ${JSON.stringify(val)})`);
            if (recordIdx === 0) {
              mappingAnalysisList.push({
                excelField: key,
                bitableField: null,
                matchType: "none",
                status: "unmatched",
                fieldTypeName: "无"
              });
            }
          }
        }

        // Only use a general text field when this record has no successful mapping at all,
        // or when the target table is effectively a single-text-field table. This prevents
        // report summaries from being injected into a standard table's "备注" field.
        const generalTextFieldCandidates = [
          "文本", "内容", "详情", "备注", "报告", "调查结果", "说明", "信息",
          "text", "content", "details", "notes", "summary", "report", "description", "desc"
        ];

        const textFields = bitableFieldNames.filter(fieldName =>
          bitableFieldTypes[normalizeName(fieldName)] === 1
        );
        const hasSuccessfulMapping = Object.keys(mappedFields).length > 0;
        const isSingleTextFieldTable = textFields.length === 1;

        // Find a general field only when fallback is actually allowed.
        let targetGeneralField: string | null = null;
        if (!hasSuccessfulMapping) {
          for (const candidate of generalTextFieldCandidates) {
            const matchedName = bitableFieldNames.find(f => normalizeName(f) === normalizeName(candidate));
            if (matchedName && !mappedFields[matchedName]) {
              targetGeneralField = matchedName;
              break;
            }
          }
        }

        // Preserve fallback behavior for a table whose only field is text.
        if (!targetGeneralField && isSingleTextFieldTable) {
          const singleField = textFields[0];
          if (!mappedFields[singleField]) {
            targetGeneralField = singleField;
          }
        }

        if (targetGeneralField) {
          // Construct a highly readable, elegant text block summarizing all available fields
          const summaryParts: string[] = [];
          if (rawFields.companyName || rawFields.country || rawFields.productContext || rawFields.grade || rawFields.completedAt || rawFields.feishuUrl || rawFields.summary || rawFields.orderInfo) {
            if (rawFields.companyName) summaryParts.push(`【公司名称】 ${rawFields.companyName}`);
            if (rawFields.country) summaryParts.push(`【国家/地区】 ${rawFields.country}`);
            if (rawFields.productContext) summaryParts.push(`【意向产品】 ${rawFields.productContext}`);
            if (rawFields.grade) summaryParts.push(`【背调等级】 ${rawFields.grade}`);
            if (rawFields.completedAt) summaryParts.push(`【背调时间】 ${rawFields.completedAt}`);
            if (rawFields.feishuUrl) summaryParts.push(`【报告链接】 ${rawFields.feishuUrl}`);
            if (rawFields.summary) summaryParts.push(`\n【报告概要】\n${rawFields.summary}`);
            if (rawFields.orderInfo) summaryParts.push(`\n【订单详情】\n${rawFields.orderInfo}`);
          } else {
            for (const [k, v] of Object.entries(rawFields)) {
              if (v !== null && v !== undefined && String(v).trim() !== "") {
                summaryParts.push(`【${k}】 ${String(v).trim()}`);
              }
            }
          }

          const fullTextSummary = summaryParts.join("\n");
          if (fullTextSummary) {
            mappedFields[targetGeneralField] = fullTextSummary;
            mappingLog.push(`Smart fallback: Auto-mapped formatted summary to general field "${targetGeneralField}" because no standard fields were mapped`);
          }
        }

        // Print the detailed mapping log for the first record
        if (recordIdx === 0 && mappingLog.length > 0) {
          console.log(`[Bitable Sync] Mapping Analysis for Record #1:\n  - ` + mappingLog.join("\n  - "));
        }

        // Return the record with mapped fields
        return { fields: mappedFields };
      });

      // If none of the records got any fields mapped, fallback to mapping the full summary to the first field
      const anyFieldsMapped = processedRecords.some((r: any) => r.fields && Object.keys(r.fields).length > 0);
      if (!anyFieldsMapped && bitableFieldNames.length > 0) {
        // Find the first field that is of type Text (type 1) as fallback to prevent type errors. Default to the first field if none found.
        let fallbackField = bitableFieldNames[0];
        for (const name of bitableFieldNames) {
          const type = bitableFieldTypes[name.toLowerCase().trim()];
          if (type === 1) { // Text
            fallbackField = name;
            break;
          }
        }
        console.log(`[Bitable Sync] No fields matched standard schema. Falling back to mapping full formatted summary to the Bitable text field: "${fallbackField}"`);
        processedRecords = records.map((record: any) => {
          const rawFields = record.fields || {};
          
          const summaryParts: string[] = [];
          if (rawFields.companyName || rawFields.country || rawFields.productContext || rawFields.grade || rawFields.completedAt || rawFields.feishuUrl || rawFields.summary || rawFields.orderInfo) {
            if (rawFields.companyName) summaryParts.push(`【公司名称】 ${rawFields.companyName}`);
            if (rawFields.country) summaryParts.push(`【国家/地区】 ${rawFields.country}`);
            if (rawFields.productContext) summaryParts.push(`【意向产品】 ${rawFields.productContext}`);
            if (rawFields.grade) summaryParts.push(`【背调等级】 ${rawFields.grade}`);
            if (rawFields.completedAt) summaryParts.push(`【背调时间】 ${rawFields.completedAt}`);
            if (rawFields.feishuUrl) summaryParts.push(`【报告链接】 ${rawFields.feishuUrl}`);
            if (rawFields.summary) summaryParts.push(`\n【报告概要】\n${rawFields.summary}`);
            if (rawFields.orderInfo) summaryParts.push(`\n【订单详情】\n${rawFields.orderInfo}`);
          } else {
            for (const [k, v] of Object.entries(rawFields)) {
              if (v !== null && v !== undefined && String(v).trim() !== "") {
                summaryParts.push(`【${k}】 ${String(v).trim()}`);
              }
            }
          }

          const fullTextSummary = summaryParts.join("\n");
          return {
            fields: {
              [fallbackField]: fullTextSummary
            }
          };
        });
      }
    }

    // --- NEW: Function 2 - Precise Update Mode ---
    if (mode === "update") {
      console.log(`[Bitable Sync] Mode is update. Fetching existing records from table ${tableId}${viewId ? `, view ${viewId}` : " (all records; no view configured)"} for matching...`);
      if (!viewId) {
        console.warn("[Bitable Sync] No view ID was found in the configured Bitable URL. Function 2 will scan the entire table; use a URL containing view=vew... to limit matching to the exported view.");
      }
      let allBitableRecords: any[] = [];
      let hasMore = true;
      let pageToken = "";
      let fetchCount = 0;
      const seenPageTokens = new Set<string>();
      let paginationError = "";

      while (hasMore) {
        const requestPageToken = pageToken;
        const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records?page_size=500${viewId ? `&view_id=${encodeURIComponent(viewId)}` : ""}${requestPageToken ? `&page_token=${encodeURIComponent(requestPageToken)}` : ""}`;

        try {
          const listRes = await fetch(url, {
            headers: {
              Authorization: `Bearer ${tenantAccessToken}`,
            }
          });

          if (!listRes.ok) {
            paginationError = `读取多维表格记录第 ${fetchCount + 1} 页失败：HTTP ${listRes.status} ${listRes.statusText}`;
            break;
          }

          const json: any = await listRes.json();
          if (json.code !== 0) {
            paginationError = `读取多维表格记录第 ${fetchCount + 1} 页失败 (Code: ${json.code})：${json.msg || "未知错误"}`;
            break;
          }
          if (!json.data || !Array.isArray(json.data.items)) {
            paginationError = `读取多维表格记录第 ${fetchCount + 1} 页失败：飞书响应缺少 data.items。`;
            break;
          }

          allBitableRecords.push(...json.data.items);
          fetchCount++;
          hasMore = json.data.has_more === true;
          const nextPageToken = typeof json.data.page_token === "string" ? json.data.page_token.trim() : "";
          console.log(`[Bitable Sync] Fetch page=${fetchCount}, view=${viewId || "ALL"}, pageRecords=${json.data.items.length}, total=${allBitableRecords.length}, hasMore=${hasMore}.`);

          if (!hasMore) {
            pageToken = "";
            break;
          }
          if (!nextPageToken) {
            paginationError = `读取多维表格记录第 ${fetchCount} 页后，飞书返回 has_more=true 但没有下一页 page_token。`;
            break;
          }
          if (nextPageToken === requestPageToken || seenPageTokens.has(nextPageToken)) {
            paginationError = `读取多维表格记录时检测到重复 page_token，已在第 ${fetchCount} 页停止以避免无限循环。`;
            break;
          }

          seenPageTokens.add(nextPageToken);
          pageToken = nextPageToken;
        } catch (error) {
          paginationError = `读取多维表格记录第 ${fetchCount + 1} 页发生异常：${error instanceof Error ? error.message : String(error)}`;
          break;
        }
      }

      if (paginationError) {
        console.error(`[Bitable Sync] Pagination aborted after ${fetchCount} completed page(s) and ${allBitableRecords.length} record(s): ${paginationError}`);
        res.status(502).json({
          success: false,
          error: `${paginationError}\n\n为避免使用不完整的飞书记录产生错误匹配，本次同步未执行。`
        });
        return;
      }

      console.log(`[Bitable Sync] Finished pagination: pages=${fetchCount}, records=${allBitableRecords.length}, hasMore=${hasMore}.`);
      
      // Define strict helpers for update-mode identifiers. Never use a generic
      // "id" alias here: it can incorrectly select 商品ID or 记录 ID as an order.
      const getOrderNumber = (fields: Record<string, any>): string => {
        const orderAliases = ["订单号", "订单编号", "orderid", "orderno", "ordernumber", "parentorderid", "交易号"];
        for (const [k, v] of Object.entries(fields)) {
          const normK = normalizeName(k);
          if (orderAliases.some(alias => normK === normalizeName(alias))) {
            if (v !== null && v !== undefined && String(v).trim() !== "") {
              return String(v).trim();
            }
          }
        }
        return "";
      };

      const getRecordId = (fields: Record<string, any>): string => {
        const recordIdAliases = ["记录 ID", "记录ID", "record_id", "recordId", "飞书记录ID"];
        for (const [k, v] of Object.entries(fields)) {
          const normK = normalizeName(k);
          if (recordIdAliases.some(alias => normK === normalizeName(alias))) {
            if (v !== null && v !== undefined && String(v).trim() !== "") {
              return String(v).trim();
            }
          }
        }
        return "";
      };

      // Detect grade & url column names in Bitable schema
      const findPreferredField = (exactNames: string[], keywords: string[], excluded: string[] = []): string => {
        const excludedNormalized = new Set(excluded.filter(Boolean).map(normalizeName));
        for (const exactName of exactNames) {
          const matched = bitableFieldNames.find(name =>
            !excludedNormalized.has(normalizeName(name)) && normalizeName(name) === normalizeName(exactName)
          );
          if (matched) return matched;
        }
        return bitableFieldNames.find(name => {
          const normName = normalizeName(name);
          return !excludedNormalized.has(normName) && keywords.some(kw =>
            normName.includes(normalizeName(kw)) || normalizeName(kw).includes(normName)
          );
        }) || "";
      };

      const gradeKeywords = ["背景评判等级", "背景分析等级", "背调等级", "评估等级", "信用评级", "客户等级", "grade", "rating"];
      const bitableGradeField = findPreferredField(
        ["背景评判等级", "背景分析等级", "背调等级", "评估等级"],
        gradeKeywords
      );

      const feishuUrlKeywords = ["标题", "背景报告文档", "背景调查报告文档", "飞书文档链接", "飞书链接", "报告链接", "word链接", "云文档", "报告", "飞书文档", "reporturl", "documentlink", "link", "背调报告", "云链接", "背调结果"];
      const bitableUrlField = findPreferredField(
        ["背景报告文档", "背景调查报告文档", "标题", "报告链接"],
        feishuUrlKeywords,
        [bitableGradeField]
      );

      const riskKeywords = ["风险提示", "风险", "信用风险", "风险预警", "riskalert", "risk", "warning"];
      const bitableRiskField = findPreferredField(
        ["风险提示", "风险预警", "信用风险"],
        riskKeywords,
        [bitableGradeField, bitableUrlField]
      );

      const strategyKeywords = ["跟进策略", "开发策略", "销售策略", "开发跟进策略", "followupstrategy", "strategy", "跟进方案"];
      const bitableStrategyField = findPreferredField(
        ["跟进策略", "开发跟进策略", "开发策略", "跟进方案"],
        strategyKeywords,
        [bitableGradeField, bitableUrlField, bitableRiskField]
      );

      const followUpGradeKeywords = ["跟进等级", "跟进级别", "跟进评分", "followupgrade", "followuplevel", "跟进"];
      const bitableFollowUpGradeField = findPreferredField(
        ["跟进等级", "跟进级别", "跟进评分"],
        followUpGradeKeywords,
        [bitableGradeField, bitableUrlField, bitableRiskField, bitableStrategyField]
      );

      console.log(`[Bitable Sync Update] Target update columns: Grade Column="${bitableGradeField}", Doc URL Column="${bitableUrlField}", Risk Column="${bitableRiskField}", Strategy Column="${bitableStrategyField}", FollowUpGrade Column="${bitableFollowUpGradeField}"`);

      const updateRecords: any[] = [];
      let updatedCount = 0;
      let unmatchedCount = 0;
      let conflictCount = 0;
      let recordIdMatchedCount = 0;
      let orderFallbackMatchedCount = 0;
      const unmatchedOrders: string[] = [];
      const matchDiagnostics: Array<{ reason: string; recordId: string; orderNumber: string }> = [];
      const matchReasonCounts: Record<string, number> = {};

      const existingByRecordId = new Map<string, any>();
      const existingByOrderNumber = new Map<string, any[]>();
      for (const item of allBitableRecords) {
        if (item.record_id) existingByRecordId.set(String(item.record_id).trim(), item);
        const itemOrder = getOrderNumber(item.fields || {});
        if (itemOrder) {
          const matches = existingByOrderNumber.get(itemOrder) || [];
          matches.push(item);
          existingByOrderNumber.set(itemOrder, matches);
        }
      }

      const skipRecord = (reason: string, recordId: string, orderNumber: string, isConflict = false) => {
        unmatchedCount++;
        if (isConflict) conflictCount++;
        matchReasonCounts[reason] = (matchReasonCounts[reason] || 0) + 1;
        unmatchedOrders.push(orderNumber || recordId || "未知行");
        if (matchDiagnostics.length < 20) matchDiagnostics.push({ reason, recordId, orderNumber });
      };

      for (const record of records) {
        const rawFields = record.fields || {};
        const orderNum = getOrderNumber(rawFields);
        const inputRecordId = getRecordId(rawFields);
        let matchedBitableRec: any = null;
        let matchMethod: "record_id" | "order_fallback" | "" = "";

        if (inputRecordId) {
          if (!orderNum) {
            skipRecord("missing_order_number", inputRecordId, orderNum, true);
            continue;
          }
          matchedBitableRec = existingByRecordId.get(inputRecordId) || null;
          if (!matchedBitableRec) {
            skipRecord("record_id_not_found", inputRecordId, orderNum, true);
            continue;
          }
          const existingOrderNum = getOrderNumber(matchedBitableRec.fields || {});
          if (!existingOrderNum || existingOrderNum !== orderNum) {
            skipRecord("record_id_order_conflict", inputRecordId, orderNum, true);
            continue;
          }
          matchMethod = "record_id";
        } else if (orderNum) {
          const orderMatches = existingByOrderNumber.get(orderNum) || [];
          if (orderMatches.length === 0) {
            skipRecord("order_number_not_found", inputRecordId, orderNum);
            continue;
          }
          if (orderMatches.length > 1) {
            skipRecord("duplicate_order_number", inputRecordId, orderNum, true);
            continue;
          }
          matchedBitableRec = orderMatches[0];
          matchMethod = "order_fallback";
        } else {
          skipRecord("missing_identifiers", inputRecordId, orderNum);
          continue;
        }

        if (matchedBitableRec) {
          const recordId = matchedBitableRec.record_id;
          const fieldsToUpdate: Record<string, any> = {};

          // Extract value for grade
          if (bitableGradeField) {
            const gradeVal = rawFields.grade || rawFields["背景评判等级"] || rawFields["背景分析等级"] || rawFields["背调等级"] || rawFields["评估等级"] || "";
            if (gradeVal) {
              const coercedGrade = coerceValueByBitableType(bitableGradeField, gradeVal);
              if (coercedGrade !== undefined) {
                fieldsToUpdate[bitableGradeField] = coercedGrade;
              }
            }
          }

          // Extract value for url
          if (bitableUrlField) {
            const urlVal = rawFields.feishuUrl || rawFields["背景报告文档"] || rawFields["标题"] || rawFields["报告链接"] || rawFields["飞书文档链接"] || "";
            if (urlVal) {
              const coercedUrl = coerceValueByBitableType(bitableUrlField, urlVal);
              if (coercedUrl !== undefined) {
                fieldsToUpdate[bitableUrlField] = coercedUrl;
              }
            }
          }

          // Extract value for riskAlert
          if (bitableRiskField) {
            const riskVal = rawFields.riskAlert || rawFields["风险提示"] || rawFields["风险"] || "";
            if (riskVal) {
              const coercedRisk = coerceValueByBitableType(bitableRiskField, riskVal);
              if (coercedRisk !== undefined) {
                fieldsToUpdate[bitableRiskField] = coercedRisk;
              }
            }
          }

          // Extract value for followUpStrategy
          if (bitableStrategyField) {
            const strategyVal = rawFields.followUpStrategy || rawFields["跟进策略"] || rawFields["开发策略"] || "";
            if (strategyVal) {
              const coercedStrategy = coerceValueByBitableType(bitableStrategyField, strategyVal);
              if (coercedStrategy !== undefined) {
                fieldsToUpdate[bitableStrategyField] = coercedStrategy;
              }
            }
          }

          // Extract value for followUpGrade
          if (bitableFollowUpGradeField) {
            const followUpGradeVal = rawFields.followUpGrade || rawFields["跟进等级"] || rawFields["跟进级别"] || "";
            if (followUpGradeVal) {
              const coercedFollowUpGrade = coerceValueByBitableType(bitableFollowUpGradeField, followUpGradeVal);
              if (coercedFollowUpGrade !== undefined) {
                fieldsToUpdate[bitableFollowUpGradeField] = coercedFollowUpGrade;
              }
            }
          }

          if (Object.keys(fieldsToUpdate).length > 0) {
            updateRecords.push({
              record_id: recordId,
              fields: fieldsToUpdate
            });
            if (matchMethod === "record_id") recordIdMatchedCount++;
            if (matchMethod === "order_fallback") orderFallbackMatchedCount++;
          } else {
            skipRecord("no_updatable_fields", inputRecordId, orderNum);
          }
        }
      }

      console.log(`[Bitable Sync Update] Prepared=${updateRecords.length}, recordId=${recordIdMatchedCount}, orderFallback=${orderFallbackMatchedCount}, unmatched=${unmatchedCount}, conflicts=${conflictCount}.`);
      console.log(`[Bitable Sync Update] Match reason counts: ${JSON.stringify(matchReasonCounts)}`);
      if (matchDiagnostics.length > 0) {
        console.warn(`[Bitable Sync Update] Match diagnostics (showing ${matchDiagnostics.length} of ${unmatchedCount}):`);
        matchDiagnostics.forEach((diagnostic, index) => {
          console.warn(
            `[Bitable Sync Update] #${index + 1} reason=${diagnostic.reason}, recordId=${JSON.stringify(diagnostic.recordId || "")}, orderNumber=${JSON.stringify(diagnostic.orderNumber || "")}`
          );
        });
        if (unmatchedCount > matchDiagnostics.length) {
          console.warn(`[Bitable Sync Update] ${unmatchedCount - matchDiagnostics.length} additional unmatched record(s) omitted from detailed logs.`);
        }
      }

      if (updateRecords.length === 0) {
        res.status(400).json({
          success: false,
          error: `更新未执行：没有记录通过「记录 ID 定位 + 订单号校验」。\n\n` +
            `• 多维表格中现有记录总数：${allBitableRecords.length}\n` +
            `• 冲突记录数：${conflictCount}\n` +
            `• 未匹配标识：${unmatchedOrders.slice(0, 5).join(", ")}${unmatchedOrders.length > 5 ? "..." : ""}\n\n` +
            `【排查建议】：请确认 Excel 的「记录 ID」来自当前目标多维表格，并核对「订单号」完全一致。`,
          recordIdMatchedCount,
          orderFallbackMatchedCount,
          conflictCount,
          unmatchedCount,
          matchDiagnostics,
          matchReasonCounts
        });
        return;
      }

      const chunkSize = 100;
      // Batch update in chunks of 100
      for (let i = 0; i < updateRecords.length; i += chunkSize) {
        const chunk = updateRecords.slice(i, i + chunkSize);
        
        let updateRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_update`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${tenantAccessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            records: chunk
          })
        });

        let updateJson: any = await updateRes.json();

        if (isPermissionCode(updateJson.code) && token.startsWith("wik") && !hasAttemptedWikiResolve) {
          hasAttemptedWikiResolve = true;
          const resolved = await resolveWikiTokenToObjToken();
          if (resolved) {
            appToken = resolved;
            updateRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_update`, {
              method: "POST",
              headers: {
                Authorization: `Bearer ${tenantAccessToken}`,
                "Content-Type": "application/json",
              },
              body: JSON.stringify({
                records: chunk
              })
            });
            updateJson = await updateRes.json();
          }
        }

        if (updateJson.code !== 0) {
          res.status(400).json({
            success: false,
            error: `批量更新记录至飞书失败 (Code: ${updateJson.code}): ${updateJson.msg || "未知错误"}`
          });
          return;
        }

        if (updateJson.data?.records) {
          updatedCount += updateJson.data.records.length;
        }
      }

      res.json({
        success: true,
        updatedCount,
        recordIdMatchedCount,
        orderFallbackMatchedCount,
        conflictCount,
        unmatchedCount,
        unmatchedOrders,
        matchDiagnostics,
        matchReasonCounts,
        mappingAnalysis: [
          {
            excelField: "背景评判等级",
            bitableField: bitableGradeField || "未找到匹配列",
            matchType: "exact",
            status: bitableGradeField ? "success" : "unmatched",
            fieldTypeName: "单选/文本"
          },
          {
            excelField: "背景报告文档",
            bitableField: bitableUrlField || "未找到匹配列",
            matchType: "exact",
            status: bitableUrlField ? "success" : "unmatched",
            fieldTypeName: "超链接"
          },
          {
            excelField: "风险提示",
            bitableField: bitableRiskField || "未找到匹配列",
            matchType: "exact",
            status: bitableRiskField ? "success" : "unmatched",
            fieldTypeName: "单行/多行文本"
          },
          {
            excelField: "跟进策略",
            bitableField: bitableStrategyField || "未找到匹配列",
            matchType: "exact",
            status: bitableStrategyField ? "success" : "unmatched",
            fieldTypeName: "单行/多行文本"
          },
          {
            excelField: "跟进等级",
            bitableField: bitableFollowUpGradeField || "未找到匹配列",
            matchType: "exact",
            status: bitableFollowUpGradeField ? "success" : "unmatched",
            fieldTypeName: "单行/多行文本"
          }
        ]
      });
      return;
    }

    // --- CRITICAL PRE-FLIGHT FIX FOR BLANK ROWS ---
    // Filter out records that ended up with completely empty fields to prevent Bitable API from creating empty rows!
    const originalCount = processedRecords.length;
    processedRecords = processedRecords.filter((record: any) => record.fields && Object.keys(record.fields).length > 0);
    console.log(`[Bitable Sync] Empty-row filter: Filtered out ${originalCount - processedRecords.length} records with empty fields payload. Records remaining to sync: ${processedRecords.length}.`);

    if (processedRecords.length === 0) {
      const excelHeaders = records.length > 0 ? Object.keys(records[0].fields || {}) : [];
      res.status(400).json({
        success: false,
        error: `同步未执行：Excel 数据中的所有行在匹配多维表格时都成了「空值」或「无匹配字段」，已拦截。这避免了在飞书多维表格中插入空行。\n\n` +
          `• 您的多维表格现有字段：\n  ${bitableFieldNames.length > 0 ? bitableFieldNames.join(", ") : "(未获取到任何字段)"}\n\n` +
          `• 您上传的数据中包含的列名：\n  ${excelHeaders.length > 0 ? excelHeaders.join(", ") : "(无有效列名)"}\n\n` +
          `【解决方案】：请确保您在多维表格中新增了「背调报告文档」或「背景调查报告文档」或「云链接」等相应的列，且至少有一列能与 Excel 列名相匹配，以便数据可以正确落库。`
      });
      return;
    }

    // 4. Batch append records to Bitable
    if (processedRecords.length > 0) {
      console.log(`[Bitable Sync] Prepared ${processedRecords.length} records for syncing to Bitable.`);
      console.log(`[Bitable Sync] Sample record #1 fields payload to Feishu (Excel row values mapped and coerced correctly):`, JSON.stringify(processedRecords[0].fields, null, 2));
    }

    // Split into chunks of 100 to stay safely below Feishu's batch API limit
    const chunkSize = 100;
    let addedCount = 0;

    for (let i = 0; i < processedRecords.length; i += chunkSize) {
      const chunk = processedRecords.slice(i, i + chunkSize);
      
      let appendRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_create`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tenantAccessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          records: chunk
        })
      });

      let appendJson: any = await appendRes.json();

      if (isPermissionCode(appendJson.code) && token.startsWith("wik") && !hasAttemptedWikiResolve) {
        hasAttemptedWikiResolve = true;
        const resolved = await resolveWikiTokenToObjToken();
        if (resolved) {
          console.log(`[Bitable Sync] Retrying batch_create with resolved obj_token: ${resolved}`);
          appToken = resolved;
          appendRes = await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/batch_create`, {
            method: "POST",
            headers: {
              Authorization: `Bearer ${tenantAccessToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              records: chunk
            })
          });
          appendJson = await appendRes.json();
        }
      }

      if (appendJson.code !== 0) {
        let friendlyError = `同步表格记录至飞书失败 (Code: ${appendJson.code}): `;
        const code = appendJson.code;
        const msg = appendJson.msg || "";

        if (isPermissionCode(code) || msg.toLowerCase().includes("permission") || msg.toLowerCase().includes("forbidden") || code === 1254302) {
          friendlyError += "您的自建应用【无权访问该目标多维表格/知识空间】。请检查以下两点：\n\n" +
            "1. 【协同权限未添加】：请在浏览器中打开该多维表格或知识空间页面：\n" +
            "   - 如果是多维表格（以 /base/ 或 /bitable/ 开头）：点击右上角「分享」/「协作」，搜索您的自建应用名称（即您在飞书后台创建的应用名字），将其添加为协作者，并勾选「可编辑」或「可管理」权限。\n" +
            "   - 如果是知识空间文档（以 /wiki/ 开头）：请务必在「知识库页面」点击右上角「...」->「页面设置」/「分享」（或该 Wiki 页面右上角的「分享」），搜索您的自建应用名称，将其添加为协作者，并勾选「可编辑」或「可管理」权限。\n\n" +
            "2. 【读写权限未发布】：请在飞书开放平台 (open.feishu.cn) 应用后台 -> 「开发配置」 -> 「权限管理」中搜索并勾选「多维表格」相关的读写权限（如 `bitable:app`）以及「知识库」相关的读权限，并必须在「版本管理与发布」中创建一个新版本并发布上线。";
        } else if (code === 1061011 || msg.toLowerCase().includes("not found")) {
          friendlyError += "未找到该多维表格或子表，请核对您的链接或 Table ID (tble...) 是否正确。";
        } else if (code === 1254003) {
          friendlyError += "指定的子表 (table) 标识符不合法或不存在，请检查 URL 中 table=tble... 参数。";
        } else {
          friendlyError += `${appendJson.msg || "未知错误"}`;
        }

        res.status(400).json({
          success: false,
          error: friendlyError
        });
        return;
      }

      if (appendJson.data?.records) {
        addedCount += appendJson.data.records.length;
      }
    }

    res.json({
      success: true,
      addedCount,
      mappingAnalysis: mappingAnalysisList
    });

  } catch (error: any) {
    console.error("Feishu sync bitable API error:", error);
    res.status(500).json({
      success: false,
      error: `服务器端处理飞书多维表格同步出错: ${error.message || error}`,
    });
  }
});

// Configure Vite or production static file serving
const startServerInstance = async () => {
  if (process.env.NODE_ENV !== "production") {
    // Load Vite only in development. Production and portable builds do not
    // need Vite or any of its native development dependencies at runtime.
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req: express.Request, res: express.Response) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`AI Background Check Server running on port ${PORT}`);
  });
};

startServerInstance();
